"""
MC & S Plugin: Tax Return Deductions Processor
================================================
Plugin ID  : plugin_tax_return_processor
Version    : 1.1.0

FIXES (v1.1.0)
--------------
- Dynamic FY: calculated from current date, not hardcoded to FY2025
- Portable: recipient/name read from settings (user_email, user_name)
- Draft mode: respects context.draft_mode — creates draft instead of sending
- Image support: .jpg/.jpeg/.png/.tiff passed to Claude vision API
- PDF fallback: pdfminer.six used if pdftotext subprocess is unavailable
- Lessons injection: active lessons injected into analysis prompt
"""
import base64
import json
import os
import re
import traceback
from datetime import datetime
from io import BytesIO

import anthropic

from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
from config import get_setting, get_active_lessons


def _current_fy() -> str:
    now = datetime.now()
    year = now.year if now.month >= 7 else now.year - 1
    return f"FY{year + 1}"


def _fy_date_range(fy: str) -> str:
    year = int(fy[2:])
    return f"1 July {year - 1} – 30 June {year}"


EXTRACTION_SYSTEM_PROMPT = """\
You are a tax deductions extraction assistant for an Australian accounting firm.
Extract ALL expense/deduction items from the provided client documents.
For each item extract: description, amount (AUD), category_hint (D1-D10 or unknown), notes.
Return a JSON array of objects with keys: description, amount, category_hint, notes.
If amount unknown use null. Be thorough — extract every expense item mentioned.
"""

ANALYSIS_SYSTEM_PROMPT_TEMPLATE = """\
You are a senior Australian tax accountant.
You are analysing deduction items for an individual tax return ({fy}).
AUTHORITY HIERARCHY:
1. ITAA97 — Income Tax Assessment Act 1997
2. ATO Taxation Rulings (TR) and Practical Compliance Guidelines (PCG)
3. ATO Individual Tax Return Instructions {fy}
4. ATO website guidance
CRITICAL: For EVERY disallowed item provide a specific citation, e.g.:
- "s8-1(2)(b) ITAA97 — private or domestic in nature"
- "ATO Instructions {fy}, D3 — conventional clothing not deductible"
- "PCG 2023/1, cl 14 — fixed rate method covers internet"
- "TR 2022/1, para 45 — effective life of laptop is 4 years"
Return JSON: allowed[], disallowed[], flagged_for_review[], follow_up_questions[], d_totals{{D1..D10}}.
"""

ANALYSIS_USER_TEMPLATE = """\
CLIENT: {client_name}
OCCUPATION: {occupation}
FINANCIAL YEAR: {fy} ({fy_range})

EXTRACTED EXPENSE ITEMS:
{items_json}

PRIOR YEAR CONTEXT:
{prior_year_context}
{lessons_block}
Analyse each item and classify as allowed, disallowed, or flagged for review.
For every disallowed item, provide the specific ATO citation.
"""


class TaxReturnProcessorPlugin(AgentPlugin):
    name        = "Tax Return Deductions Processor"
    description = "Analyses client tax documents and generates a D1-D10 deductions report with ATO citations."
    detail      = (
        "Monitors for flagged DOCUMENTS_RECEIVED emails. Downloads attachments (including images), "
        "extracts expense items with Claude, analyses deductibility against ATO rules, "
        "and produces a Word report. Every disallowed item includes the specific ATO citation."
    )
    version = "1.1.0"
    icon    = "📋"
    requires_graph  = True
    requires_claude = True
    default_schedule = Schedule.every_minutes(5)
    _processed_ids: set

    def load(self, context: PluginContext) -> bool:
        self._processed_ids = set()
        if not context.graph:
            context.log("📋 Tax Return Processor: Microsoft 365 not connected.")
            return False
        if not context.claude:
            context.log("📋 Tax Return Processor: Claude AI not configured.")
            return False
        return True

    @classmethod
    def settings_schema(cls) -> list[dict]:
        return [
            {
                "key": "report_recipient",
                "label": "Report Recipient Email",
                "default": "",
                "type": "text",
                "help": "Email to send the report to. Defaults to the user_email setting.",
            },
            {
                "key": "monitor_folder",
                "label": "Folder to Monitor",
                "default": "Inbox",
                "type": "text",
                "help": "Outlook folder to check for flagged DOCUMENTS_RECEIVED emails.",
            },
            {
                "key": "max_per_run",
                "label": "Max Emails Per Run",
                "default": "5",
                "type": "number",
                "help": "Maximum number of tax return emails to process per run.",
            },
        ]

    def run(self, context: PluginContext) -> PluginResult:
        try:
            return self._do_run(context)
        except Exception as e:
            context.log(f"📋 Tax Return Processor: Error — {e}")
            traceback.print_exc()
            return PluginResult(success=False, error=str(e))

    def _do_run(self, context: PluginContext) -> PluginResult:
        graph     = context.graph
        log       = context.log
        folder    = self.get_plugin_setting("monitor_folder", "Inbox")
        recipient = (
            self.get_plugin_setting("report_recipient", "")
            or get_setting("user_email", "")
        )
        user_name = get_setting("user_name", "there")
        max_count = int(self.get_plugin_setting("max_per_run", "5"))
        fy        = _current_fy()
        fy_range  = _fy_date_range(fy)

        if not recipient:
            log("📋 Tax Return Processor: No report recipient — set user_email in Settings.")
            return PluginResult(success=False, error="No report recipient configured.")

        try:
            import requests as req
            params = {
                "$filter": "isRead eq false and flag/flagStatus eq 'flagged'",
                "$top": max_count,
                "$orderby": "receivedDateTime asc",
                "$select": "id,subject,from,receivedDateTime,body,bodyPreview,hasAttachments,categories",
            }
            url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder}/messages"
            r = req.get(url, headers=graph._headers(), params=params, timeout=30)
            r.raise_for_status()
            emails = r.json().get("value", [])
        except Exception as e:
            return PluginResult(success=False, error=f"Could not fetch emails: {e}")

        tax_emails = [
            e for e in emails
            if "Documents Received" in e.get("categories", [])
            and e["id"] not in self._processed_ids
        ]

        if not tax_emails:
            log("📋 Tax Return Processor: No flagged DOCUMENTS_RECEIVED emails to process.")
            return PluginResult(success=True, summary="No tax return emails to process.")

        log(f"📋 Tax Return Processor: Found {len(tax_emails)} email(s) to process.")
        result = PluginResult(success=True)

        for email in tax_emails:
            msg_id     = email["id"]
            subject    = email.get("subject", "(No Subject)")
            sender     = email.get("from", {}).get("emailAddress", {})
            from_name  = sender.get("name", "Client")
            from_email = sender.get("address", "")
            log(f"  Processing: \"{subject}\" from {from_name}")
            try:
                attachment_texts, image_blocks = self._get_attachment_content(graph, msg_id, log)
                body_text = re.sub(r"<[^>]+>", " ", email.get("body", {}).get("content", ""))
                body_text = re.sub(r"\s+", " ", body_text).strip()
                all_text = body_text + "\n\n" + "\n\n".join(attachment_texts)

                if len(all_text.strip()) < 50 and not image_blocks:
                    log(f"  ⚠ No usable content — skipping.")
                    self._processed_ids.add(msg_id)
                    result.items_skipped += 1
                    continue

                log(f"  Extracting expense items...")
                items = self._extract_items(context.claude, all_text[:6000], image_blocks)
                if not items:
                    log(f"  ⚠ No expense items found — skipping.")
                    self._processed_ids.add(msg_id)
                    result.items_skipped += 1
                    continue

                log(f"  Found {len(items)} item(s). Analysing deductibility...")
                client_name = from_name
                occupation  = self._guess_occupation(body_text, items)
                analysis    = self._analyse_items(
                    context.claude, client_name, occupation, items, fy, fy_range
                )

                log(f"  Generating deductions report...")
                report_bytes, report_filename = self._generate_report(
                    client_name, occupation, subject, analysis, fy, user_name
                )

                allowed_total    = sum(i.get("deductible_amount", 0) or 0 for i in analysis.get("allowed", []))
                disallowed_count = len(analysis.get("disallowed", []))
                flagged_count    = len(analysis.get("flagged_for_review", []))

                report_subject = f"Tax Return Deductions Report — {client_name} — {fy}"
                report_body = f"""
<p>Hi {user_name},</p>
<p>The Tax Return Deductions Processor has completed the analysis for
<strong>{client_name}</strong> ({fy}).</p>
<table style="border-collapse:collapse;font-family:Arial;font-size:13px">
  <tr><td style="padding:4px 12px 4px 0"><strong>Total Deductible Amount:</strong></td>
      <td style="padding:4px 0"><strong>${allowed_total:,.2f}</strong></td></tr>
  <tr><td style="padding:4px 12px 4px 0">Items Allowed:</td>
      <td style="padding:4px 0">{len(analysis.get('allowed', []))}</td></tr>
  <tr><td style="padding:4px 12px 4px 0">Items Disallowed (with ATO citations):</td>
      <td style="padding:4px 0">{disallowed_count}</td></tr>
  <tr><td style="padding:4px 12px 4px 0">Items Flagged for Your Review:</td>
      <td style="padding:4px 0">{flagged_count}</td></tr>
</table>
<p>The full report is attached. Each disallowed item includes the specific ATO citation.</p>
<p>Kind regards,<br>EVA — CoWorker</p>
"""
                if context.draft_mode:
                    import tempfile
                    with tempfile.NamedTemporaryFile(
                        suffix=os.path.splitext(report_filename)[1], delete=False
                    ) as tmp:
                        tmp.write(report_bytes)
                        tmp_path = tmp.name
                    try:
                        graph.create_draft_with_attachments(
                            recipient, report_subject, report_body,
                            attachment_paths=[tmp_path]
                        )
                        log(f"  📝 Draft created for {recipient} (draft_mode=True).")
                    finally:
                        try:
                            os.unlink(tmp_path)
                        except Exception:
                            pass
                else:
                    self._send_report_email(
                        graph, recipient, report_subject, report_body,
                        report_bytes, report_filename
                    )
                    log(f"  ✅ Report sent to {recipient}.")

                log(f"  ${allowed_total:,.2f} deductible, {disallowed_count} disallowed, {flagged_count} flagged.")
                graph.mark_as_read(msg_id)
                self._processed_ids.add(msg_id)
                result.actions_taken += 1
                result.drafts_created += 1
                self.log_activity(
                    source=self.name,
                    subject=f"Tax report: {client_name} {fy}",
                    category="tax_return",
                    action=f"Report {'drafted' if context.draft_mode else 'sent'} — ${allowed_total:,.2f} deductible",
                    draft_created=1,
                )
            except Exception as e:
                log(f"  ❌ Error processing {from_name}: {e}")
                traceback.print_exc()
                result.items_skipped += 1

        result.summary = f"{result.actions_taken} report(s) generated, {result.items_skipped} skipped."
        return result

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _get_attachment_content(self, graph, msg_id: str, log) -> tuple[list[str], list[dict]]:
        texts: list[str] = []
        image_blocks: list[dict] = []
        try:
            import requests as req
            url = f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/attachments"
            r = req.get(url, headers=graph._headers(), timeout=30)
            r.raise_for_status()
            for att in r.json().get("value", []):
                name = att.get("name", "").lower()
                b64  = att.get("contentBytes", "")
                if not b64:
                    continue
                raw = base64.b64decode(b64)
                if name.endswith(".pdf"):
                    texts.append(self._extract_pdf_text(raw, name))
                elif name.endswith((".txt", ".csv")):
                    texts.append(raw.decode("utf-8", errors="ignore"))
                elif name.endswith((".xlsx", ".xls")):
                    texts.append(self._extract_excel_text(raw, name))
                elif name.endswith((".jpg", ".jpeg", ".png", ".tiff", ".tif", ".webp")):
                    ext = os.path.splitext(name)[1].lower().lstrip(".")
                    mt  = {"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png",
                           "tiff":"image/tiff","tif":"image/tiff","webp":"image/webp"}.get(ext,"image/jpeg")
                    image_blocks.append({"type":"image","source":{"type":"base64","media_type":mt,"data":b64}})
                    log(f"    Image queued for vision: {name}")
                else:
                    log(f"    Skipping: {name}")
        except Exception as e:
            log(f"    ⚠ Could not fetch attachments: {e}")
        return texts, image_blocks

    def _extract_pdf_text(self, raw: bytes, name: str) -> str:
        # Try pdftotext first (poppler)
        try:
            import subprocess, tempfile
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
                f.write(raw); tmp = f.name
            res = subprocess.run(["pdftotext", tmp, "-"], capture_output=True, text=True, timeout=30)
            os.unlink(tmp)
            if res.returncode == 0 and res.stdout.strip():
                return f"[PDF: {name}]\n{res.stdout}"
        except Exception:
            pass
        # Fallback: pdfminer.six (pure Python, works on Windows)
        try:
            from pdfminer.high_level import extract_text as pm_extract
            return f"[PDF: {name}]\n{pm_extract(BytesIO(raw))}"
        except Exception:
            pass
        return f"[PDF: {name} — could not extract text]"

    def _extract_excel_text(self, raw: bytes, name: str) -> str:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(BytesIO(raw), data_only=True)
            lines = [f"[Excel: {name}]"]
            for sheet in wb.sheetnames:
                for row in wb[sheet].iter_rows(values_only=True):
                    t = "\t".join(str(c) for c in row if c is not None)
                    if t.strip():
                        lines.append(t)
            return "\n".join(lines)
        except Exception:
            return f"[Excel: {name} — could not extract text]"

    def _extract_items(self, claude: anthropic.Anthropic, text: str, image_blocks: list[dict]) -> list[dict]:
        content: list[dict] = []
        if text.strip():
            content.append({"type":"text","text":f"Extract all expense/deduction items:\n\n{text}"})
        content.extend(image_blocks)
        if image_blocks and not text.strip():
            content.append({"type":"text","text":"Extract all expense/deduction items visible in the image(s)."})
        if not content:
            return []
        resp = claude.messages.create(
            model=self.get_claude_model(), max_tokens=2000,
            system=EXTRACTION_SYSTEM_PROMPT,
            messages=[{"role":"user","content":content}],
        )
        raw = re.sub(r"^```(?:json)?\s*","",resp.content[0].text.strip())
        raw = re.sub(r"\s*```$","",raw)
        return json.loads(raw)

    def _analyse_items(self, claude, client_name, occupation, items, fy, fy_range) -> dict:
        lessons = get_active_lessons()
        lessons_block = ""
        if lessons:
            lessons_block = "LEARNED PREFERENCES (from prior accountant decisions):\n"
            lessons_block += "\n".join(f"- {l['lesson']}" for l in lessons) + "\n"
        prompt = ANALYSIS_USER_TEMPLATE.format(
            client_name=client_name, occupation=occupation or "Not specified",
            fy=fy, fy_range=fy_range, items_json=json.dumps(items, indent=2),
            prior_year_context="No prior year data available.",
            lessons_block=lessons_block,
        )
        resp = claude.messages.create(
            model=self.get_claude_model(), max_tokens=4000,
            system=ANALYSIS_SYSTEM_PROMPT_TEMPLATE.format(fy=fy),
            messages=[{"role":"user","content":prompt}],
        )
        raw = re.sub(r"^```(?:json)?\s*","",resp.content[0].text.strip())
        raw = re.sub(r"\s*```$","",raw)
        return json.loads(raw)

    def _guess_occupation(self, body_text: str, items: list[dict]) -> str:
        for pat in [r"occupation[:\s]+([A-Za-z\s]+)", r"employed as[:\s]+([A-Za-z\s]+)"]:
            m = re.search(pat, body_text, re.IGNORECASE)
            if m:
                return m.group(1).strip()[:50]
        return "Not specified"

    def _generate_report(self, client_name, occupation, email_subject, analysis, fy, reviewer_name) -> tuple[bytes, str]:
        try:
            from docx import Document
            from docx.enum.text import WD_ALIGN_PARAGRAPH
        except ImportError:
            return self._generate_text_report(client_name, analysis, fy)

        doc = Document()
        title = doc.add_heading(f"Tax Return Deductions Report — {fy}", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Client: {client_name}")
        doc.add_paragraph(f"Occupation: {occupation}")
        doc.add_paragraph(f"Date Prepared: {datetime.now().strftime('%d %B %Y')}")
        doc.add_paragraph(f"Source: {email_subject}")
        doc.add_paragraph()

        doc.add_heading("Allowed Deductions", 1)
        allowed = analysis.get("allowed", [])
        if allowed:
            tbl = doc.add_table(rows=1, cols=5); tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            for i, h in enumerate(["Description","Category","Amount","Work %","Deductible"]):
                hdr[i].text = h; hdr[i].paragraphs[0].runs[0].bold = True
            for item in allowed:
                row = tbl.add_row().cells
                row[0].text = item.get("description",""); row[1].text = item.get("d_category","")
                row[2].text = f"${item.get('amount',0) or 0:,.2f}"
                row[3].text = f"{item.get('work_pct',100)}%"
                row[4].text = f"${item.get('deductible_amount',0) or 0:,.2f}"
        else:
            doc.add_paragraph("No items allowed.")
        doc.add_paragraph()

        doc.add_heading("Deductions Summary (D1-D10)", 1)
        totals = analysis.get("d_totals", {})
        tbl2 = doc.add_table(rows=1, cols=2); tbl2.style = "Table Grid"
        tbl2.rows[0].cells[0].text = "Category"; tbl2.rows[0].cells[0].paragraphs[0].runs[0].bold = True
        tbl2.rows[0].cells[1].text = "Total Deductible"; tbl2.rows[0].cells[1].paragraphs[0].runs[0].bold = True
        for cat, amt in totals.items():
            if amt:
                r = tbl2.add_row().cells; r[0].text = cat; r[1].text = f"${amt:,.2f}"
        grand = sum(v for v in totals.values() if v)
        r = tbl2.add_row().cells; r[0].text = "TOTAL"; r[0].paragraphs[0].runs[0].bold = True
        r[1].text = f"${grand:,.2f}"; r[1].paragraphs[0].runs[0].bold = True
        doc.add_paragraph()

        doc.add_heading("Disallowed Items (ATO Citations)", 1)
        disallowed = analysis.get("disallowed", [])
        if disallowed:
            tbl3 = doc.add_table(rows=1, cols=4); tbl3.style = "Table Grid"
            hdr3 = tbl3.rows[0].cells
            for i, h in enumerate(["Description","Amount","Reason","ATO Citation"]):
                hdr3[i].text = h; hdr3[i].paragraphs[0].runs[0].bold = True
            for item in disallowed:
                row = tbl3.add_row().cells
                row[0].text = item.get("description","")
                row[1].text = f"${item.get('amount',0) or 0:,.2f}"
                row[2].text = item.get("reason","")
                row[3].text = item.get("citation","No citation provided")
        else:
            doc.add_paragraph("No items disallowed.")
        doc.add_paragraph()

        doc.add_heading(f"Items Flagged for {reviewer_name}'s Review", 1)
        flagged = analysis.get("flagged_for_review", [])
        if flagged:
            for item in flagged:
                p = doc.add_paragraph(style="List Bullet")
                p.add_run(item.get("description","")).bold = True
                p.add_run(f" — ${item.get('amount',0) or 0:,.2f}")
                doc.add_paragraph(f"  Question: {item.get('question','')}")
        else:
            doc.add_paragraph("No items flagged.")
        doc.add_paragraph()

        doc.add_heading("Suggested Client Follow-Up Questions", 1)
        for i, q in enumerate(analysis.get("follow_up_questions", []), 1):
            doc.add_paragraph(f"{i}. {q}")

        buf = BytesIO(); doc.save(buf); buf.seek(0)
        safe = re.sub(r"[^\w\s-]","",client_name).strip().replace(" ","_")
        return buf.read(), f"TaxReturn_{safe}_{fy}_{datetime.now().strftime('%Y%m%d')}.docx"

    def _generate_text_report(self, client_name, analysis, fy) -> tuple[bytes, str]:
        lines = [f"TAX RETURN DEDUCTIONS REPORT — {fy}", f"Client: {client_name}",
                 f"Date: {datetime.now().strftime('%d %B %Y')}", "", "=== ALLOWED DEDUCTIONS ==="]
        for item in analysis.get("allowed", []):
            lines.append(f"  {item.get('d_category','?')} | {item.get('description','')} | ${item.get('deductible_amount',0):,.2f}")
        lines += ["", "=== DISALLOWED ITEMS (ATO CITATIONS) ==="]
        for item in analysis.get("disallowed", []):
            lines += [f"  {item.get('description','')} | ${item.get('amount',0):,.2f}",
                      f"    Reason: {item.get('reason','')}", f"    Citation: {item.get('citation','No citation')}"]
        lines += ["", "=== FLAGGED FOR REVIEW ==="]
        for item in analysis.get("flagged_for_review", []):
            lines.append(f"  {item.get('description','')} — {item.get('question','')}")
        lines += ["", "=== FOLLOW-UP QUESTIONS ==="]
        for i, q in enumerate(analysis.get("follow_up_questions", []), 1):
            lines.append(f"  {i}. {q}")
        safe = re.sub(r"[^\w\s-]","",client_name).strip().replace(" ","_")
        return "\n".join(lines).encode("utf-8"), f"TaxReturn_{safe}_{fy}_{datetime.now().strftime('%Y%m%d')}.txt"

    def _send_report_email(self, graph, to_address, subject, body_html, attachment_bytes, attachment_name):
        import requests as req
        encoded = base64.b64encode(attachment_bytes).decode("utf-8")
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body_html},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
                "attachments": [{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment_name,
                    "contentBytes": encoded,
                    "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                }],
            },
            "saveToSentItems": True,
        }
        r = req.post("https://graph.microsoft.com/v1.0/me/sendMail",
                     headers=graph._headers(), json=payload, timeout=30)
        r.raise_for_status()
