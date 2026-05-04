# DRAFT_FONT_FIX.md — Claude Code Task List

Fix draft email font consistency so manual edits by the accountant match the AI-drafted text.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Consistent font and color in draft emails for manual editing

**Files:** `graph_client.py`

**Problem:** When CoWorker creates a draft email and the accountant opens it in Outlook to add a few lines or edit, the new text they type uses Outlook's default font (usually Calibri 11pt) which doesn't match the AI-drafted text. The draft looks inconsistent — two different fonts in the same email. The accountant has to manually reformat every time they edit a draft.

**Root cause:** The HTML body sent via Graph API doesn't set an explicit default font style that Outlook's editor inherits for new content. Outlook only inherits the font of the element the cursor is in.

**Changes required:**

1. In `graph_client.py`, add a constant and helper near the top of the file:
   ```python
   # Standard MC&S email font style — Outlook inherits this for manual edits
   MCS_EMAIL_STYLE = (
       'font-family: Calibri, Arial, Helvetica, sans-serif; '
       'font-size: 11pt; '
       'color: #000000; '
       'line-height: 1.5;'
   )
   
   def _wrap_email_body(body_html: str) -> str:
       """Wrap email body in a styled container so Outlook inherits the font for manual edits."""
       return f'<div style="{MCS_EMAIL_STYLE}">{body_html}</div>'
   ```

2. Apply `_wrap_email_body()` to the body HTML BEFORE it gets sent to the Graph API, but AFTER the signature is appended. The wrapping order should be:
   ```
   body_html (AI draft content)
   → _append_signature(body_html)  (adds signature)
   → _wrap_email_body(body_html)   (wraps everything in styled div)
   → send to Graph API
   ```

3. The style uses **Calibri 11pt black** — Outlook's default business email font. When the accountant clicks into the draft and starts typing, their new text will match.

4. Apply the wrapping in ALL draft/send paths:
   - `create_draft()` — standalone new emails
   - `_create_threaded_reply_draft()` — threaded replies
   - `create_draft_with_attachments()` — emails with attachments
   - `send_email()` — direct sends (approval queue auto-sends)

5. Also convert any `\n` newlines in the body to `<br>` tags if the body is plain text (no existing HTML tags):
   ```python
   if '<p>' not in body_html and '<br' not in body_html and '<div' not in body_html:
       body_html = body_html.replace('\n', '<br>\n')
   ```

6. For the threaded reply path: the styled `<div>` should only wrap the NEW content (AI draft + signature), not the quoted original email. The quoted original keeps its own formatting:
   ```html
   <div style="font-family: Calibri, ...">
     [AI drafted reply content]
     [Signature image]
   </div>
   [Quoted original email — left as-is by createReply]
   ```

**Test:**
1. Trigger Smart Responder to create a draft
2. Open the draft in Outlook
3. Click at the end of the AI-drafted text
4. Press Enter and type a new line: "I'll also need your bank statements."
5. The new text should be in the same font (Calibri 11pt black) as the drafted text

**Commit message:** `fix: wrap draft emails in consistent Calibri 11pt styling so manual edits match AI-drafted text`

---

## Post-fix checklist

- [ ] Open a draft in Outlook → type a new line → font matches the draft
- [ ] Threaded reply → new content has consistent font, quoted original is unchanged
- [ ] Signature still renders correctly below the body
- [ ] No double-wrapping if a draft is edited and re-saved
- [ ] Rebuild installer: `build_installer.bat`
