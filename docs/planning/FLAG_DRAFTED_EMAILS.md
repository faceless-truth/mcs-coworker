# FLAG_DRAFTED_EMAILS.md — Claude Code Task List

Flag emails in Outlook when a draft reply is created so the accountant sees which emails need attention.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Flag emails when draft reply is created

**Files:** `graph_client.py`, `plugins/plugin_smart_responder.py`

**Problem:** When the Smart Responder creates a draft reply, the original email stays in the inbox looking like any other read email. The accountant has no visual indicator that a draft has been created and is ready to review. The old approach (marking as unread) caused an infinite loop of duplicate drafts. Flagging is the correct Outlook pattern — it puts a visible flag icon on the email without triggering re-processing.

**Changes required:**

### Part A — Add flag method to graph_client.py

1. Add a method to flag/unflag an email:
   ```python
   def flag_email(self, message_id: str, flag_status: str = "flagged") -> bool:
       """Flag an email in Outlook.
       
       Args:
           message_id: The Graph API message ID
           flag_status: 'flagged', 'notFlagged', or 'complete'
       
       Returns:
           True if successful, False otherwise
       """
       url = f"{self.base_url}/me/messages/{message_id}"
       result = self._make_request("PATCH", url, json={
           "flag": {"flagStatus": flag_status}
       })
       return result is not None
   ```

2. This uses the same `_make_request` path as all other working email operations — no auth issues.

### Part B — Flag after draft creation in Smart Responder

3. In `plugins/plugin_smart_responder.py`, find where the draft is successfully created. AFTER the draft creation AND after marking the email as read, add the flag:
   ```python
   # After successful draft creation:
   context.graph.mark_as_read(message_id)       # Prevent reprocessing
   context.graph.flag_email(message_id)          # Visual indicator for accountant
   self._mark_as_processed(message_id, ...)      # SQLite tracking safety net
   ```

4. Only flag when a draft was ACTUALLY created. Do NOT flag for:
   - `NO_REPLY` results (email was skipped — no draft exists)
   - Errors or failures (draft wasn't created)
   - Emails that were already processed (skipped by the duplicate check)

5. The flag should happen after `mark_as_read` but before `_mark_as_processed`. If flagging fails (transient Graph API error), log a warning but don't crash the plugin — the draft was still created successfully.
   ```python
   try:
       context.graph.flag_email(message_id)
       logger.info(f"Flagged email {message_id} — draft ready for review")
   except Exception as e:
       logger.warning(f"Could not flag email {message_id}: {e}")
   ```

### Part C — Apply to other plugins that create drafts

6. Check these plugins — if they create drafts in response to inbound emails, they should also flag the original:
   - `plugin_noa_processor.py` — processes NOA emails and drafts client responses. Flag the NOA email after drafting.
   - `plugin_engagement_letter.py` — if it drafts in response to an email, flag it.
   - `plugin_bas_reminder.py` — these are outbound reminders, not replies to inbound. Do NOT flag.
   - `plugin_debtor_followup.py` — outbound follow-ups. Do NOT flag.
   - `plugin_annual_review.py` — outbound outreach. Do NOT flag.
   - `plugin_client_outreach.py` — outbound. Do NOT flag.

   Rule: only flag when the plugin is RESPONDING to an inbound email with a draft reply. Don't flag outbound drafts that the plugin initiates on its own.

7. For each plugin that needs flagging, add the same pattern:
   ```python
   # After successful draft creation for an inbound email:
   try:
       context.graph.flag_email(original_message_id)
   except Exception:
       pass  # Non-fatal
   ```

**Test:**
1. Send a test email to elio@mcands.com.au
2. Wait for Smart Responder to create a draft (within 60 seconds)
3. Open Outlook inbox — the original email should have a flag icon
4. Open Drafts — the draft reply should be there
5. Send another email that triggers NO_REPLY — it should NOT be flagged
6. The email should be marked as read (not unread) — no duplicate drafts

**Commit message:** `feat: flag emails in Outlook when draft reply is created for easy review`

---

## Post-fix checklist

- [ ] Emails with drafts show a flag icon in Outlook inbox
- [ ] Emails with NO_REPLY result are NOT flagged
- [ ] Emails are still marked as read (no duplicate draft loop)
- [ ] NOA Processor also flags NOA emails after drafting
- [ ] Outbound-only plugins (debtor, BAS, annual review) do NOT flag anything
- [ ] Flag failure doesn't crash the plugin
- [ ] Rebuild installer: `build_installer.bat`
