# ATO_CORRESPONDENCE_AGENT.md — Claude Code Task List

Add an ATO Correspondence specialist agent that batch-processes daily ATO document downloads, classifies each document, and drafts client emails with the correct attachments. Remove SharePoint save from chat exports. Increase file upload limit to 50.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 4: Remove SharePoint save option from chat exports

**Files:** `frontend/client/src/pages/Chat.tsx`, `api_server.py`

**Problem:** Eliza (reception) doesn't want the SharePoint save option in chat exports. Remove it from the export dropdown.

**Changes required:**

1. In `Chat.tsx`, find the Export dropdown. Remove the "Save to Client Folder (SharePoint)" option entirely. Keep only:
   - Download Transcript (.docx)
   - Download Summary (.docx)
   - Download Recommendation (.docx)

2. Remove the `SharePointFolderPicker` component import and usage from `Chat.tsx` if it's only used here.

3. In `api_server.py`, keep the `/api/chat/export/sharepoint` endpoint code — don't delete it. Just remove it from the UI. It may be useful later or for other features.

4. Remove any SharePoint-related state from Chat.tsx (sharepoint modal visibility, selected folder path, etc.) if they're no longer referenced.

**Test:** Open AI Chat → Export dropdown should show only the three download options. No SharePoint option.

**Commit message:** `fix: remove SharePoint save option from chat export dropdown`

---

## Fix 2 of 4: Increase file upload limit to 50 files per message

**Files:** `api_server.py`, `frontend/client/src/pages/Chat.tsx`

**Problem:** The current limit is 5 files per message. ATO daily correspondence folders can contain 50+ documents that need to be processed in one batch.

**Changes required:**

### Backend — api_server.py

1. Find the `MAX_UPLOADS_PER_MESSAGE` constant (should be 5). Change to 50:
   ```python
   MAX_UPLOADS_PER_MESSAGE = 50
   ```

2. Find the `MAX_UPLOAD_SIZE` constant. Keep at 25MB per file — ATO PDFs are small (typically 50-500KB each).

3. Check the Claude API call in `/api/chat`. When 50 PDF documents are attached, the total content may exceed Claude's context window. Add a safeguard:
   - For each PDF, extract only the first 5 pages or 10,000 characters of text
   - If a PDF can't be read, skip it and note it in the response
   - Add a total content cap: if all file content blocks combined exceed 200,000 characters, truncate the oldest files and warn the user

4. For batch PDF processing, instead of sending all PDFs as native document blocks (which would be huge), extract text from each PDF and send as text blocks with the filename:
   ```python
   # For batch mode (>10 files), extract text instead of sending raw PDFs
   if len(attached_files) > 10:
       for file_ref in attached_files:
           if file_ref["type"] == ".pdf":
               # Extract text from PDF using PyPDF2 or pdfplumber
               text = _extract_pdf_text(file_path, max_pages=5)
               file_content_blocks.append({
                   "type": "text",
                   "text": f"<uploaded_file name=\"{file_ref['name']}\">\n{text[:10000]}\n</uploaded_file>"
               })
   ```

5. Add `PyPDF2` to requirements.txt if not already present (for PDF text extraction).

### Frontend — Chat.tsx

6. Update the file upload limit in the UI:
   - Change the max files check from 5 to 50
   - Update any UI text that mentions "max 5 files" to "max 50 files"
   - The file chip display area may need to scroll if many files are uploaded — make it a scrollable container with a max height

7. When many files are being uploaded, show a progress counter: "Uploading 23 of 47 files..."

8. Add a "Upload Folder" option alongside the paperclip button — allows selecting multiple files at once via a file picker with multi-select:
   ```html
   <input type="file" multiple accept=".pdf,.docx,.xlsx,.csv,.txt,.jpg,.png" />
   ```

**Test:** Upload 20+ PDF files in one message. All should upload and be processed.

**Commit message:** `feat: increase file upload limit to 50 files with batch PDF text extraction`

---

## Fix 3 of 4: Create the ATO Correspondence specialist agent

**Files:** `specialists/registry.py`, new file `specialists/prompts/ato_correspondence.md`, `frontend/client/src/pages/Chat.tsx`

**Changes required:**

### Create the specialist prompt

1. Create `specialists/prompts/ato_correspondence.md` with this content:

```markdown
# ATO Correspondence Processor - Master Prompt

**Version:** 1.0
**Audience:** MC & S reception staff
**Purpose:** Batch-process daily ATO correspondence downloads, classify documents, and draft client emails

---

## 1. Identity and Purpose

You are the ATO Correspondence Processor for MC & S Pty Ltd. Reception downloads the daily ATO correspondence folder and uploads all documents to you in one batch. Your job is to:

1. Read every uploaded document
2. Classify each one by type
3. Sort them by client
4. Draft an appropriate email for each client with the correct attachment references
5. Present everything in an organised, reviewable format

You operate in batch mode - you may receive 1 to 50 documents at once. Process them all systematically.

---

## 2. Hard Operating Rules

1. **Read every document.** Do not skip any uploaded file.
2. **Never invent data.** If you cannot read a document or extract a figure, say so clearly.
3. **Always identify the client.** Extract the taxpayer name, TFN (last 3 digits only for reference), and entity type from each document.
4. **Always identify the document type.** Classify into the categories below.
5. **Always extract the key figures.** Dollar amounts, dates, assessment periods.
6. **Draft in first person as the accountant.** Use the accountant's name from context. Never say "I'll let [accountant] know."
7. **Never use em dashes, en dashes, or smart quotes.**
8. **Group by client.** If one client has multiple documents, group them in one email.
9. **Flag anything unusual.** Amended assessments, large debts, compliance actions - these need the accountant's attention before sending.

---

## 3. Document Classification

Classify each document into one of these categories:

| Category | Description | Email Tone |
|----------|-------------|------------|
| **NOA - Refund** | Notice of Assessment showing a refund | Positive - "Great news, your refund of $X..." |
| **NOA - Payable** | Notice of Assessment showing tax owing | Neutral - "Your assessment shows $X owing..." |
| **NOA - Nil** | Notice of Assessment with nil balance | Neutral - "Your assessment is balanced..." |
| **NOA - Amended** | Amended assessment from ATO | Cautious - flag for accountant review first |
| **Activity Statement** | BAS/IAS assessment or notice | Informational - include period and outcome |
| **Payment Reminder** | ATO payment reminder or overdue notice | Urgent - include amount and due date |
| **Debt Notice** | Running balance or debt notification | Urgent - flag for accountant review |
| **Instalment Notice** | PAYG instalment rate or amount | Informational - include rate/amount and period |
| **Correspondence** | General ATO letter or notice | Depends on content |
| **Compliance Action** | Audit notice, review, or compliance activity | STOP - flag for accountant, do NOT draft email |
| **Unknown** | Cannot classify | Flag for manual review |

---

## 4. Processing Workflow

### Step 1 - Read and Classify

For each uploaded document:
- Extract: client name, TFN (last 3 only), entity type, document type, key figures, dates
- If a document is unreadable, note it as "Could not process: [filename]"

### Step 2 - Sort by Client

Group all documents by client name. One client may have multiple documents (e.g., individual NOA + company NOA).

### Step 3 - Present Summary

Before drafting any emails, present a summary table:

```
## ATO Correspondence Summary - [Date]

| # | Client | Document Type | Key Figure | Action |
|---|--------|--------------|------------|--------|
| 1 | Korkie, Gordon | NOA - Refund | $3,240 refund | Draft email |
| 2 | Korkie, Gordon | NOA - Payable (Company) | $12,100 owing | Draft email |
| 3 | Smith, Jane | Activity Statement Q3 | $1,800 credit | Draft email |
| 4 | Ahmed, Omar | Compliance Action | Audit notice | FLAG - Accountant review |
| 5 | [unreadable] | Unknown | - | Manual review needed |

Total: 5 documents, 4 clients, 1 flagged for review
```

### Step 4 - Draft Emails

After presenting the summary, ask: "Shall I draft the emails for all clients, or would you like to review the flagged items first?"

When instructed to draft, produce each email in this format:

```
---
### Email 1 of 4: Korkie, Gordon
**To:** [need client email - please provide or check XPM]
**Subject:** Your Tax Assessment - FY2025
**Attachments:** [filename1.pdf, filename2.pdf]

Hi Gordon,

[Appropriate email body based on document type and figures]

[No sign-off - signature is added automatically]
---
```

### Step 5 - Client Email Lookup

For each client, you need their email address to create the draft. Handle this by:
- Asking reception to provide the email address
- Noting: "Client email needed for: [list of clients]"
- Reception can type the emails or look them up in XPM

Once emails are provided, the drafts can be created.

---

## 5. Email Templates by Document Type

### NOA - Refund
```
Hi [First Name],

We have received your Notice of Assessment from the ATO for the [year] financial year.

Your assessment shows a refund of $[amount]. The ATO will process this refund to your nominated bank account, which typically takes 5-14 business days.

Please find your Notice of Assessment attached for your records.

If you have any questions, please don't hesitate to get in touch.
```

### NOA - Payable
```
Hi [First Name],

We have received your Notice of Assessment from the ATO for the [year] financial year.

Your assessment shows a balance payable of $[amount]. The due date for payment is [date].

Payment can be made via:
- BPAY: Biller Code [code], Reference [ref] (shown on the NOA)
- ATO online services at my.gov.au
- Phone: 1800 815 886

Please find your Notice of Assessment attached for your records. Let us know if you would like to discuss payment options.
```

### NOA - Nil
```
Hi [First Name],

We have received your Notice of Assessment from the ATO for the [year] financial year.

Your assessment is balanced - there is no amount payable or refundable.

Please find your Notice of Assessment attached for your records.
```

### Activity Statement
```
Hi [First Name],

We have received your [BAS/IAS] assessment from the ATO for the [period] period.

[If credit: Your assessment shows a credit of $[amount], which will be refunded to your nominated account.]
[If debit: Your assessment shows a balance of $[amount] owing. The due date is [date].]

Please find the assessment attached for your records.
```

### Payment Reminder
```
Hi [First Name],

We have received a payment reminder from the ATO regarding your [tax type] account.

The outstanding amount is $[amount] and was due on [date]. We recommend making payment as soon as possible to avoid additional interest charges.

Please find the ATO notice attached. Let us know if you need assistance with payment arrangements.
```

---

## 6. Flagging Rules

**STOP and flag for accountant review (do NOT draft an email) when:**
- Document is a compliance action, audit notice, or review notification
- Amended assessment with a significant change (>$5,000 difference)
- Debt notice over $50,000
- Any document mentioning penalties, prosecution, or legal action
- Any document you cannot confidently classify

**Flag format:**
```
⚠️ FLAGGED FOR REVIEW: [Client Name]
Document: [filename]
Type: [classification]
Reason: [why it needs review]
Suggested action: [what the accountant should do]
```

---

## 7. Batch Output Format

When drafting multiple emails, number them clearly and separate with horizontal rules. At the end, provide a checklist:

```
## Draft Summary

- [ ] Email 1: Korkie, Gordon - NOA Refund $3,240 - NEED EMAIL ADDRESS
- [ ] Email 2: Smith, Jane - Activity Statement credit $1,800 - NEED EMAIL ADDRESS
- [x] Email 3: Chen, Wei - NOA Payable $890 - chen@email.com provided
- ⚠️ FLAGGED: Ahmed, Omar - Compliance action - DO NOT SEND, accountant review required

Provide client email addresses and I'll create the drafts in Outlook.
```

---

## 8. House Style

- Plain Australian English
- Calibri 11pt (handled automatically by the system)
- No em dashes, en dashes, or smart quotes
- No exclamation marks
- No "Kind regards" or sign-off (signature added automatically)
- Reference dollar amounts with $ and commas: $12,300
- Reference dates as: 30 April 2026
- Keep emails under 150 words unless the situation requires more detail
```

### Register the agent

2. In `specialists/registry.py`, add the ATO Correspondence agent. Place it in the `documents` category:
   ```python
   SpecialistAgent(
       id="ato_correspondence",
       name="ATO Correspondence",
       description="Batch-process daily ATO downloads - classifies documents, drafts client emails",
       icon="📬",
       category="documents",
       system_prompt=_load_prompt("ato_correspondence.md"),
       supports_files=True,
       file_types=[".pdf", ".docx", ".xlsx", ".csv", ".txt", ".jpg", ".png"],
       model_preference="sonnet",
   )
   ```

3. In `Chat.tsx`, add example prompts:
   ```typescript
   ato_correspondence: [
       "Process today's ATO correspondence batch",
       "I've uploaded the daily ATO folder - please classify and draft emails",
       "Sort these ATO documents by client and draft responses",
   ],
   ```

**Test:** Upload 5-10 test ATO PDFs. The agent should classify each one, present a summary table, then draft appropriate emails for each client.

**Commit message:** `feat: add ATO Correspondence batch processor agent for daily ATO document handling`

---

## Fix 4 of 4: Enable email draft creation from the ATO Correspondence chat

**Files:** `api_server.py`, `frontend/client/src/pages/Chat.tsx`

**Problem:** After the ATO Correspondence agent drafts emails in the chat, reception needs to actually CREATE those drafts in Outlook with the correct PDF attachments. Currently the chat can only display text — it can't trigger email actions.

**Changes required:**

### Backend — api_server.py

1. Add an endpoint to create an email draft from chat:
   ```python
   @app.route("/api/chat/create-draft", methods=["POST"])
   def create_draft_from_chat():
       """Create an Outlook draft email from AI Chat content."""
       data = request.json or {}
       to_address = data.get("to", "")
       subject = data.get("subject", "")
       body_html = data.get("body", "")
       attachment_file_ids = data.get("attachment_ids", [])  # File IDs from chat uploads
       
       if not to_address or not subject or not body_html:
           return jsonify({"ok": False, "error": "Missing to, subject, or body"}), 400
       
       graph = _get_graph()
       if not graph:
           return jsonify({"ok": False, "error": "Not connected to Microsoft"}), 500
       
       try:
           # If there are attachments, use create_draft_with_attachments
           if attachment_file_ids:
               # Resolve file IDs to file paths
               attachments = []
               for file_id in attachment_file_ids:
                   # Find the file in CHAT_UPLOAD_DIR
                   for f in CHAT_UPLOAD_DIR.iterdir():
                       if f.stem == file_id:
                           attachments.append({
                               "path": str(f),
                               "name": data.get("attachment_names", {}).get(file_id, f.name),
                           })
                           break
               
               draft_id = graph.create_draft_with_attachments(
                   to_address=to_address,
                   subject=subject,
                   body_html=body_html,
                   attachments=attachments,
               )
           else:
               draft_id = graph.create_draft(
                   to_address=to_address,
                   subject=subject,
                   body_html=body_html,
               )
           
           return jsonify({"ok": True, "draft_id": draft_id})
       except Exception as e:
           logger.error(f"Failed to create draft from chat: {e}", exc_info=True)
           return jsonify({"ok": False, "error": str(e)}), 500
   ```

### Frontend — Chat.tsx

2. Add a "Create Draft" button that appears in assistant messages when the ATO Correspondence agent has drafted an email. The button should detect email-like content in the response (looking for "To:", "Subject:", and email body patterns).

3. When clicked, it shows a small confirmation modal:
   ```
   ┌──────────────────────────────────────┐
   │  Create Draft in Outlook             │
   │                                      │
   │  To: [pre-filled, editable]          │
   │  Subject: [pre-filled, editable]     │
   │  Attach: [checkbox list of uploaded  │
   │           PDFs for this client]      │
   │                                      │
   │  [Cancel]         [Create Draft]     │
   └──────────────────────────────────────┘
   ```

4. The "Create Draft" button calls `POST /api/chat/create-draft` and shows a toast: "Draft created in Outlook for [client name]"

5. After creating a draft, mark it in the UI so reception knows it's done — change the button to "✓ Draft Created" (greyed out).

6. For batch mode (multiple emails drafted in one response), each email section should have its own "Create Draft" button.

### Rebuild frontend

```
cd frontend && pnpm build && xcopy /E /I /Y dist\public ..\frontend_dist\
```

**Test:**
1. Select ATO Correspondence agent
2. Upload 5 ATO PDFs
3. Say "Process these"
4. Agent classifies and drafts emails
5. Click "Create Draft" next to each email
6. Check Outlook Drafts — each draft should have the correct To, Subject, body, and PDF attachment
7. Flagged items should NOT have a Create Draft button

**Commit message:** `feat: enable email draft creation from ATO Correspondence chat with attachment forwarding`

---

## Done — Post-fix checklist

- [ ] SharePoint save option removed from export dropdown
- [ ] Can upload 50 files in one message
- [ ] ATO Correspondence agent appears under DOCUMENTS category
- [ ] Upload 10+ ATO PDFs → agent classifies and sorts by client
- [ ] Summary table shows all documents grouped by client
- [ ] Flagged items (compliance actions, large debts) are marked for review
- [ ] Each drafted email has a "Create Draft" button
- [ ] Create Draft creates the email in Outlook with correct attachment
- [ ] Draft uses the correct template (refund vs payable vs nil etc.)
- [ ] Agent count is now 12
- [ ] Rebuild installer: `build_installer.bat`
