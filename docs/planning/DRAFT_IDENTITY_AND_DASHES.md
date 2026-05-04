# DRAFT_IDENTITY_AND_DASHES.md — Claude Code Task List

Fix two issues with Smart Responder email drafts: em dashes appearing despite the fix, and the AI drafting in third person ("I'll let Elio know") instead of AS the accountant.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 2: Smart Responder must draft as the accountant, not about them

**Files:** `plugins/plugin_smart_responder.py`, `config.py`

**Problem:** The Smart Responder monitors Elio's inbox. When a client emails Elio, the AI drafts a response saying "I'll let Elio know" — as if it's Elio's assistant, not Elio himself. This is wrong. The draft should be written AS Elio, in first person. The client should never know an AI drafted it.

**Example of the bug:**
- Client Jackie emails Elio: "Hi Elio that's all good thanks"
- Smart Responder drafts: "Thank you, Jackie — noted. I'll let Elio know."
- This implies someone OTHER than Elio is responding, which confuses the client

**What it should draft:**
- "Thank you, Jackie — noted. I'll take a look and follow up with you shortly."

**Changes required:**

1. The Smart Responder needs to know WHO the accountant is. Read the accountant's name from settings. In `config.py`, check if there's already a setting for the user's name (e.g., `user_name`, `accountant_name`, `display_name`). If not, it should be derived from:
   - The `outlook_mailbox` setting (e.g., `elio@mcands.com.au` → "Elio")
   - Or the Microsoft Graph `user_name` / `display_name` from the OAuth profile
   - Or a new setting `accountant_name` in Settings

2. Check if there's already a `user_name` or similar setting being stored. Search:
   ```
   grep -rn "user_name\|accountant_name\|display_name\|staff_name" config.py api_server.py graph_client.py plugins/plugin_smart_responder.py
   ```

3. If no name setting exists, add one:
   - In `config.py` DEFAULT_SETTINGS, add `accountant_name` with default `""`
   - On app startup or first Graph authentication, auto-populate it from the Microsoft Graph profile:
     ```python
     # After successful auth in graph_client.py:
     user_info = self.get_user_info()  # GET /me
     if user_info and user_info.get("displayName"):
         from config import get_setting, save_setting
         if not get_setting("accountant_name"):
             save_setting("accountant_name", user_info["displayName"])
     ```

4. In `plugins/plugin_smart_responder.py`, update the system prompt to include the accountant's identity:
   ```python
   from config import get_setting
   
   accountant_name = get_setting("accountant_name", "")
   mailbox = get_setting("outlook_mailbox", "")
   
   # Add to the system prompt:
   identity_prompt = f"""
   CRITICAL IDENTITY RULE:
   You ARE {accountant_name}. You are drafting emails as {accountant_name} from {mailbox}.
   Write in first person. Never refer to yourself in third person.
   Never say "I'll let {accountant_name} know" or "I'll pass this to {accountant_name}" — YOU are {accountant_name}.
   Never say "the team will" or "our office will" unless genuinely referring to other staff.
   
   Examples:
   WRONG: "I'll let Elio know about this."
   RIGHT: "Noted, I'll take a look at this."
   
   WRONG: "I'll pass this on to the accountant."
   RIGHT: "I'll review this and get back to you."
   
   WRONG: "Elio will be in touch."
   RIGHT: "I'll be in touch."
   """
   ```

5. Inject this identity prompt at the START of the system prompt, before any other instructions. It must be the first thing Claude sees.

6. Apply the same identity rule to ALL other plugins that draft emails:
   - `plugin_debtor_followup.py`
   - `plugin_engagement_letter.py`
   - `plugin_annual_review.py`
   - `plugin_bas_reminder.py`
   - `plugin_client_outreach.py`
   
   Each should read `accountant_name` and draft as that person in first person.

7. Add `accountant_name` to the Settings UI if it's not already there — as a text field in the Microsoft 365 section:
   ```
   Accountant Name: [Elio Scarton]
   (Auto-detected from Microsoft 365. This is how your name appears in AI-drafted emails.)
   ```

**Test:** Send a test email from a personal account to elio@mcands.com.au. The draft response should say "I'll" and "I" — never "Elio will" or "I'll let Elio know."

**Commit message:** `fix: Smart Responder drafts as the accountant in first person — never refers to them in third person`

---

## Fix 2 of 2: Remove all em dashes and en dashes from email drafts

**Files:** `graph_client.py`, `plugins/plugin_smart_responder.py`

**Problem:** Em dashes (—) and en dashes (–) still appear in draft emails despite being unprofessional for business correspondence. The fix needs to happen in TWO places: instruct Claude not to use them, AND strip them as a safety net before the email is sent.

**Example of the bug:**
- Draft shows: "Thank you, Jackie — noted."
- Should show: "Thank you, Jackie - noted." (or better: rephrase to avoid the dash entirely)

**Changes required:**

### Part A — Instruct Claude to never use them

1. In `plugins/plugin_smart_responder.py`, add to the system prompt:
   ```
   FORMATTING RULES:
   - Never use em dashes (—) or en dashes (–). Use a regular hyphen (-) or rephrase the sentence.
   - Never use semicolons in client emails. Use a full stop and start a new sentence.
   - Keep sentences short and clear.
   ```

2. Add the same instruction to ALL other email-drafting plugins:
   - `plugin_debtor_followup.py`
   - `plugin_engagement_letter.py`
   - `plugin_annual_review.py`
   - `plugin_bas_reminder.py`
   - `plugin_client_outreach.py`
   - `plugin_morning_briefing.py`

### Part B — Strip dashes as a safety net in graph_client.py

3. In `graph_client.py`, find the `_wrap_email_body` function (or wherever the body HTML is processed centrally before ALL email paths). Add dash replacement:
   ```python
   def _wrap_email_body(body_html: str) -> str:
       # Remove em dashes and en dashes
       body_html = body_html.replace('\u2014', '-')  # em dash —
       body_html = body_html.replace('\u2013', '-')  # en dash –
       body_html = body_html.replace('&mdash;', '-')
       body_html = body_html.replace('&ndash;', '-')
       
       # ... existing wrapping logic
   ```

4. Make sure this runs on EVERY email path — create_draft, send_email, create_draft_with_attachments, _create_threaded_reply_draft. If `_wrap_email_body` is already called in all paths, adding it there is sufficient.

5. Also strip smart quotes while we're at it — they cause encoding issues in some email clients:
   ```python
   body_html = body_html.replace('\u201c', '"')  # left double quote "
   body_html = body_html.replace('\u201d', '"')  # right double quote "
   body_html = body_html.replace('\u2018', "'")  # left single quote '
   body_html = body_html.replace('\u2019', "'")  # right single quote '
   ```

**Test:** Send a test email. The draft should contain no em dashes, en dashes, or smart quotes. Check by searching the draft text for — and –.

**Commit message:** `fix: strip em dashes, en dashes, and smart quotes from all email drafts`

---

## Done — Post-fix checklist

- [ ] Send test email → draft says "I'll review this" not "I'll let Elio know"
- [ ] Draft never refers to the accountant in third person
- [ ] `accountant_name` setting is auto-populated from Microsoft Graph profile
- [ ] `accountant_name` appears in Settings UI under Microsoft 365
- [ ] No em dashes (—) in any draft
- [ ] No en dashes (–) in any draft
- [ ] No smart quotes (" " ' ') in any draft
- [ ] All 6 email-drafting plugins have the identity and formatting rules
- [ ] Rebuild installer: `build_installer.bat`
