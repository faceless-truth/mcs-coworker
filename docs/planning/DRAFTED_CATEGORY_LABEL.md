# DRAFTED_CATEGORY_LABEL.md — Claude Code Task List

Add a "Drafted" category label to emails when a draft reply is created, so the accountant sees a clear text indicator alongside the flag.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Add "Drafted" category label to flagged emails

**Files:** `graph_client.py`, `plugins/plugin_smart_responder.py`, `plugins/plugin_noa_processor.py`

**Problem:** When the Smart Responder creates a draft reply and flags the original email, the flag alone is just a red icon. Accountants want to see the word "Drafted" on the email so it's immediately obvious why it's flagged. Outlook categories support text labels that display directly on the email in the inbox list.

**Changes required:**

### Part A — Add category method to graph_client.py

1. Add a method to set a category on an email:
   ```python
   def add_category(self, message_id: str, category: str = "Drafted") -> bool:
       """Add a category label to an email in Outlook.
       
       Categories show as coloured text labels in the inbox.
       Outlook auto-creates the category if it doesn't exist.
       The user can customise the colour in Outlook settings.
       
       Args:
           message_id: The Graph API message ID
           category: The category name to add (default: "Drafted")
       
       Returns:
           True if successful, False otherwise
       """
       # First get existing categories so we don't overwrite them
       url = f"{self.base_url}/me/messages/{message_id}?$select=categories"
       current = self._make_request("GET", url)
       categories = current.get("categories", []) if current else []
       
       # Only add if not already present
       if category not in categories:
           categories.append(category)
       
       # Update the message
       url = f"{self.base_url}/me/messages/{message_id}"
       result = self._make_request("PATCH", url, json={"categories": categories})
       return result is not None
   ```

2. Also add a method to remove a category (for future use):
   ```python
   def remove_category(self, message_id: str, category: str = "Drafted") -> bool:
       """Remove a category label from an email."""
       url = f"{self.base_url}/me/messages/{message_id}?$select=categories"
       current = self._make_request("GET", url)
       categories = current.get("categories", []) if current else []
       
       if category in categories:
           categories.remove(category)
           url = f"{self.base_url}/me/messages/{message_id}"
           result = self._make_request("PATCH", url, json={"categories": categories})
           return result is not None
       return True
   ```

### Part B — Apply to Smart Responder

3. In `plugins/plugin_smart_responder.py`, find where `flag_email` is called after draft creation. Add the category call right after:
   ```python
   # After successful draft creation:
   context.graph.mark_as_read(message_id)
   context.graph.flag_email(message_id)
   try:
       context.graph.add_category(message_id, "Drafted")
   except Exception as e:
       logger.warning(f"Could not add Drafted category to {message_id}: {e}")
   ```

4. Only add the category when a draft was ACTUALLY created — same rule as flagging. Do NOT add for NO_REPLY results.

### Part C — Apply to NOA Processor

5. In `plugins/plugin_noa_processor.py`, find where the plugin flags emails after drafting a client response. Add the same category call:
   ```python
   try:
       context.graph.add_category(message_id, "Drafted")
   except Exception:
       pass
   ```

### Part D — Apply to any other plugin that flags emails

6. Search for other plugins that call `flag_email`:
   ```
   grep -rn "flag_email" plugins/*.py
   ```
   Add `add_category(message_id, "Drafted")` after every `flag_email` call.

7. The category call should always be wrapped in try/except — if it fails, the flag and draft still exist. The category is a nice-to-have visual indicator, not a critical operation.

**Test:**
1. Send a test email to elio@mcands.com.au
2. Wait for Smart Responder to create a draft
3. Open Outlook inbox — the original email should show:
   - A red flag icon
   - A "Drafted" text label/category
4. The word "Drafted" should be visible directly in the inbox list view
5. If the category colour isn't ideal, right-click it in Outlook → "Set Color" to choose green/blue/etc.

**Commit message:** `feat: add Drafted category label to flagged emails for clear visual indicator`

---

## Post-fix checklist

- [ ] Emails with drafts show both flag AND "Drafted" label
- [ ] "Drafted" text is visible in the Outlook inbox list view
- [ ] NO_REPLY emails do NOT get the "Drafted" label
- [ ] Category failure doesn't crash the plugin
- [ ] NOA Processor also adds "Drafted" label
- [ ] Rebuild installer: `build_installer.bat`

## Optional — Customise category colour in Outlook

After the first "Drafted" label appears, the accountant can set a preferred colour:
1. Right-click the "Drafted" label on any email
2. Click "Set Color"
3. Choose green (recommended — means "ready to go")
4. This colour persists for all future "Drafted" labels on that machine
