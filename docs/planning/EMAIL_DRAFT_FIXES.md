# EMAIL_DRAFT_FIXES.md — Claude Code Task List

Fix two issues with Smart Responder email drafts: URLs not clickable, and missing subject lines.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 2: Convert plain-text URLs to clickable hyperlinks in email drafts

**Files:** `plugins/plugin_smart_responder.py`, optionally `graph_client.py`

**Problem:** When the Smart Responder drafts an email that includes a URL (e.g., the Tax Return Checklist link from the Knowledge Base), the URL appears as plain text in the HTML email. The recipient has to copy and paste it into a browser instead of clicking it. This looks unprofessional and confusing for clients.

**Changes required:**

1. Create a URL linkifier utility. Either add it to `plugin_smart_responder.py` or to a shared utility like `prompt_utils.py`:
   ```python
   import re
   
   def linkify_urls(text: str) -> str:
       """Convert plain-text URLs to clickable HTML hyperlinks.
       
       Skips URLs that are already inside an href="" or src="" attribute
       to avoid double-wrapping existing links.
       """
       # Match URLs not already inside href= or src=
       # Negative lookbehind for href=" or src="
       url_pattern = r'(?<!href=")(?<!href=\')(?<!src=")(?<!src=\')(https?://[^\s<>"\']+)'
       return re.sub(url_pattern, r'<a href="\1">\1</a>', text)
   ```

2. In `plugins/plugin_smart_responder.py`, find where the draft body is passed to `create_draft()`. Before that call, apply the linkifier:
   ```python
   # After Claude generates the draft body, before creating the draft:
   draft_body = linkify_urls(draft_body)
   ```

3. Centralised alternative: instead of doing this per-plugin, add the linkification inside `graph_client.py`'s `create_draft` method so ALL plugins benefit:
   ```python
   def create_draft(self, to_address, subject, body_html, reply_to_id=None):
       # Linkify any plain-text URLs in the body
       body_html = linkify_urls(body_html)
       # ... rest of existing logic
   ```
   This is the better approach — do it here so every plugin's drafts get clickable links automatically.

4. Also add a hint to the Smart Responder's system prompt to encourage Claude to use descriptive link text rather than raw URLs:
   ```
   When referencing URLs, present them as clickable links with context. For example, write "You can access our Tax Return Checklist here: [URL]" rather than just pasting a raw URL. The system will automatically convert URLs to clickable hyperlinks.
   ```

5. Apply `linkify_urls` to `create_draft_with_attachments` as well if it has a separate body handling path.

**Test:** Send a test email that triggers a Knowledge Base response containing a URL. The draft should show the URL as a blue, clickable hyperlink in Outlook.

**Commit message:** `fix: convert plain-text URLs to clickable hyperlinks in all email drafts`

---

## Fix 2 of 2: Fix missing subject line in threaded reply drafts

**Files:** `graph_client.py`

**Problem:** A draft reply sent from CoWorker shows "(No subject)" in Outlook. When using the `createReply` Graph API endpoint, the subject should be automatically set to "RE: [original subject]" by Microsoft. If it's showing "(No subject)", something is overwriting it during the PATCH step that sets the body.

**Changes required:**

1. In `graph_client.py`, find the `_create_threaded_reply_draft` method (or wherever `createReply` is called followed by a PATCH).

2. Check the PATCH payload. If it includes a `subject` field set to empty or None, that's overwriting the auto-generated "RE: ..." subject:
   ```python
   # BAD — this overwrites the subject:
   patch_data = {
       "subject": subject,  # If subject is None or empty, this blanks it
       "body": {
           "contentType": "HTML",
           "content": full_body
       }
   }
   
   # GOOD — only set body, leave subject alone:
   patch_data = {
       "body": {
           "contentType": "HTML",
           "content": full_body
       }
   }
   ```

3. The `createReply` endpoint automatically sets the subject to "RE: [original subject]". The subsequent PATCH should ONLY update the body — never touch the subject. Remove `subject` from the PATCH payload entirely.

4. If the `subject` parameter is still needed for standalone drafts (no `reply_to_id`), keep it in the standalone path only:
   ```python
   if reply_to_id:
       # Threaded reply — createReply sets the subject automatically
       # PATCH only updates the body
       patch_data = {
           "body": {"contentType": "HTML", "content": full_body}
       }
   else:
       # Standalone new email — we set the subject
       draft_data = {
           "subject": subject,
           "body": {"contentType": "HTML", "content": full_body},
           ...
       }
   ```

5. Also check `create_draft_with_attachments` for the same issue — if it uses createReply followed by a PATCH, the PATCH should not include the subject.

6. While investigating, also check: is the Smart Responder passing an empty subject to `create_draft()`? If so, fix that too — it should pass the original email's subject:
   ```python
   # In plugin_smart_responder.py, where create_draft is called:
   # Make sure subject is taken from the original email, not hardcoded to ""
   subject = original_email.get("subject", "")
   context.graph.create_draft(
       to_address=sender_email,
       subject=subject,  # Should be the original subject
       body_html=draft_body,
       reply_to_id=message_id,
   )
   ```

**Test:** Send a test email with a clear subject line. Let Smart Responder draft a reply. The draft in Outlook should show "RE: [original subject]", not "(No subject)".

**Commit message:** `fix: preserve subject line in threaded reply drafts — don't overwrite createReply subject`

---

## Done — Post-fix checklist

- [ ] Send test email with a subject line → draft shows "RE: [subject]"
- [ ] Send test email that triggers a KB response with a URL → URL is clickable in the draft
- [ ] Check other plugin drafts (debtor follow-up, BAS reminder) → URLs are also clickable
- [ ] Rebuild installer if needed: `build_installer.bat`
