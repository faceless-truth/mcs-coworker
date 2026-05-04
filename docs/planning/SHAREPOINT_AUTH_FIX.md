# SHAREPOINT_AUTH_FIX.md — Claude Code Task List

The SharePoint connection fails because SharePoint methods use is_authenticated() which checks the broken MSAL disk cache. Email operations work because they go through _make_request() which uses the in-memory token. Fix: make SharePoint use the same path.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Wire all SharePoint methods through _make_request

**Files:** `graph_client.py`

**Problem:** 

Email operations work because they call `self._make_request("GET", url)` which uses the in-memory access token directly. No disk cache involved.

SharePoint operations fail because `test_sharepoint_connection()` starts with `if not self.is_authenticated()` which calls MSAL's `acquire_token_silent()` which reads from the disk cache file — which was never written. The function returns "Not authenticated" before ever trying the Graph API.

The actual Graph API call to SharePoint would work fine — it's the same token, same permissions. The gate is the problem.

**Changes required:**

1. In `test_sharepoint_connection()`:
   - REMOVE the `if not self.is_authenticated()` check entirely
   - Just call `get_sharepoint_site_id()` directly
   - If it returns None, the error message from `_make_request` will tell us why
   - Add `print("[SharePoint Test] Starting connection test...")` at the top

2. In `get_sharepoint_site_id()`:
   - Must use ONLY `self._make_request("GET", url)` for the API call
   - NO `is_authenticated()` checks
   - NO raw `requests.get()` calls
   - NO separate `_get_token()` calls
   - Add `print(f"[SharePoint] Looking up site: {url}")` before the call
   - Add `print(f"[SharePoint] Result: {result}")` after the call

3. In `get_sharepoint_drive_id()`:
   - Must use ONLY `self._make_request("GET", url)`
   - NO other token handling

4. In `upload_to_sharepoint()`:
   - Must use ONLY `self._make_request()` for all API calls
   - For file upload PUT: `self._make_request("PUT", url, data=file_content, headers=extra_headers)`
   - If `_make_request` doesn't support PUT or custom headers for file upload, add PUT support to `_make_request`
   - NO raw `requests.put()` calls
   - NO separate token retrieval

5. In `list_sharepoint_folder()`:
   - Must use ONLY `self._make_request("GET", url)`

6. In `sharepoint_folder_exists()`:
   - Must use ONLY `self._make_request("GET", url)`

7. **Verify by searching the file:** After changes, run:
   ```
   grep -n "is_authenticated\|acquire_token_silent\|_get_token\|requests\.get\|requests\.put\|requests\.post" graph_client.py
   ```
   
   The ONLY acceptable hits should be:
   - The `is_authenticated()` method definition itself
   - The `_get_token()` method definition itself (used internally by `_make_request`)
   - Inside `_make_request` itself
   
   NO SharePoint method should directly call any of these. They should all go through `_make_request()` exclusively.

8. **Check _make_request supports all needed HTTP methods.** It needs to handle:
   - GET (site lookup, folder listing)
   - POST (create upload session)
   - PUT (file upload content)
   - PATCH (already used for email updates)
   
   If it only handles GET/POST/PATCH, add PUT support. The method should accept a `data` parameter for raw binary content (file uploads) and a way to set Content-Type to `application/octet-stream`.

9. **Check _make_request handles the colon URL format.** SharePoint site lookup uses:
   ```
   https://graph.microsoft.com/v1.0/sites/mcandscomau.sharepoint.com:/sites/MCS354
   ```
   The colon after the hostname is valid but some URL libraries may encode it. Make sure `_make_request` passes the URL as-is without re-encoding.

**Test:**

Start the app, sign in, then click Test Connection in Settings. The console should show:
```
[SharePoint Test] Starting connection test...
[SharePoint] Looking up site: https://graph.microsoft.com/v1.0/sites/mcandscomau.sharepoint.com:/sites/MCS354
[SharePoint] Result: {'id': '...', 'name': 'MCS354', ...}
```

If it shows `Result: None`, check the `_make_request` logs for the HTTP status code — it will be 403 (scope issue) or 404 (URL issue), which gives us a clear answer.

**Commit message:** `fix: wire all SharePoint methods through _make_request — same token path as working email`

---

## Post-fix

If Test Connection shows green:
- Test "Save to Client Folder" from AI Chat with a test client name
- Rebuild installer: `build_installer.bat`

If Test Connection still fails:
- The console prints will show exactly what URL was called and what came back
- 403 = token doesn't have SharePoint scopes (need to re-sign in with new scopes)
- 404 = URL format wrong (the print will show the exact URL)
- None with no error = _make_request is silently swallowing the response
