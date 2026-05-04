# ATTACHMENT_AWARENESS.md — Claude Code Task List

Fix Smart Responder so it knows when emails have attachments instead of telling clients their files are missing.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Include attachment metadata in Smart Responder prompt

**Files:** `plugins/plugin_smart_responder.py`, `graph_client.py`

**Problem:** When a client sends an email with attachments (PDFs, documents, etc.), the Smart Responder only sees the email body text. The body often says things like "Please find attached..." but Claude has no way to know whether attachments are actually present. It then drafts a response saying "it appears the attachments may not have come through" — which is wrong and embarrassing to send to a client.

**Example of the bug:**
- Client emails: "Attached are letters confirming the Stamp duty refunded from S.R.O." with 2 PDF attachments
- Smart Responder drafts: "However, it appears the attachments may not have come through with your email. Can you please attach?"
- The attachments ARE there — Claude just doesn't know about them

**Changes required:**

### Part A — Fetch attachment metadata in graph_client.py

1. Add a method to fetch attachment metadata for a message (if one doesn't already exist):
   ```python
   def get_attachment_metadata(self, message_id: str) -> list:
       """Fetch attachment names and sizes for a message (metadata only, not content)."""
       url = f"{self.base_url}/me/messages/{message_id}/attachments?$select=name,size,contentType"
       result = self._make_request("GET", url)
       if result and "value" in result:
           return [
               {
                   "name": att.get("name", "unknown"),
                   "size": att.get("size", 0),
                   "content_type": att.get("contentType", ""),
               }
               for att in result["value"]
           ]
       return []
   ```

2. If a method like this already exists (check for `get_attachments`, `fetch_attachments`, `list_attachments`), use it instead.

### Part B — Include attachment info in the Smart Responder prompt

3. In `plugins/plugin_smart_responder.py`, find where the email is processed and the Claude prompt is built. After fetching the email body, check for attachments:
   ```python
   # Check for attachments
   attachment_info = ""
   has_attachments = email.get("hasAttachments", False)
   if has_attachments:
       attachments = context.graph.get_attachment_metadata(message_id)
       if attachments:
           att_list = ", ".join(
               f"{att['name']} ({att['size'] // 1024} KB)" for att in attachments
           )
           attachment_info = (
               f"\n\nATTACHMENT INFO: This email has {len(attachments)} attachment(s): {att_list}. "
               f"The attachments are confirmed present in the email — do NOT tell the sender "
               f"their attachments are missing or ask them to re-send."
           )
   ```

4. Append `attachment_info` to the email context that gets sent to Claude. This should go right after the email body in the prompt:
   ```python
   # When building the user message for Claude:
   email_context = f"<email_body>\n{email_body}\n</email_body>{attachment_info}"
   ```

5. Also add a general instruction to the Smart Responder's system prompt:
   ```
   ATTACHMENT HANDLING:
   - If the email metadata confirms attachments are present, acknowledge receipt of the files by name.
   - Never tell the sender their attachments are missing when the system confirms they are attached.
   - If the email mentions attachments but no attachment metadata is provided, you may ask the sender to confirm.
   - When acknowledging attachments, reference them naturally: "Thank you for sending through the [filename] documents."
   ```

### Part C — Handle the case where email mentions attachments but has none

6. For completeness, also handle the reverse case — email body says "attached" but `hasAttachments` is False:
   ```python
   if not has_attachments and any(word in email_body.lower() for word in ["attached", "attachment", "enclosed", "find attached"]):
       attachment_info = (
           "\n\nATTACHMENT INFO: The sender mentions attachments in their email, "
           "but no attachments were found on this message. It may be appropriate "
           "to politely ask the sender to re-send with the files attached."
       )
   ```

   This is the ONLY case where Claude should mention missing attachments.

### Part D — Apply to other plugins that process emails

7. Check these plugins for the same issue — if they read email bodies and might draft responses mentioning attachments:
   - `plugin_noa_processor.py` — this one actually downloads attachments, so it knows they're there. Verify.
   - `plugin_engagement_letter.py` — check if it processes emails with attachments
   - `plugin_meeting_prep.py` — check if it summarises email content including attachment references

   For each, ensure the plugin either:
   - Already handles attachments correctly (like NOA Processor which downloads them), OR
   - Gets attachment metadata added to its prompt context

**Test:**

1. Send a test email to elio@mcands.com.au with a PDF attached and body text saying "Please find attached my documents"
2. Wait for Smart Responder to draft a reply (within 60 seconds)
3. Open the draft — it should acknowledge the attachment by name: "Thank you for sending through the [filename]"
4. It should NOT say "the attachments may not have come through"

5. Send a second test email with body text saying "I've attached the receipts" but with NO actual attachment
6. The draft should politely note: "You mentioned attachments but I wasn't able to see any files attached to your email — could you please re-send?"

**Commit message:** `fix: include attachment metadata in Smart Responder prompt so AI knows files are attached`

---

## Post-fix checklist

- [ ] Email with attachments → draft acknowledges files by name
- [ ] Email with attachments → draft does NOT say files are missing
- [ ] Email mentioning attachments but with none → draft politely asks to re-send
- [ ] Email with no attachment mentions → draft doesn't mention attachments at all
- [ ] NOA Processor still handles its PDF attachments correctly
- [ ] Rebuild installer when ready: `build_installer.bat`
