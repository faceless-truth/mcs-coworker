# SHAREPOINT_FIX.md — Claude Code Task List

Fix SharePoint integration properly. Hardcode the config, fix the auth.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Hardcode SharePoint config and fix auth

**Files:** `graph_client.py`, `api_server.py`, `frontend/client/src/pages/Settings.tsx`

**Problem:** SharePoint test connection always fails with "not authenticated" even though the app successfully reads and sends emails via the same Microsoft Graph API. The settings-based SharePoint config never saves properly, and the token retrieval path for SharePoint differs from the working email path.

**Changes required:**

### Part A — Hardcode SharePoint config

1. In `graph_client.py`, add constants at the top of the file (near GRAPH_SCOPES):
   ```python
   # MC&S SharePoint — hardcoded, same for all accountants
   SHAREPOINT_SITE_URL = "https://mcandscomau.sharepoint.com/sites/MCS354"
   SHAREPOINT_LIBRARY = "Shared Documents"
   SHAREPOINT_CLIENT_BASE = "Server/Clients"
   ```

2. Update `_get_sharepoint_config()` to return the hardcoded values instead of reading from settings:
   ```python
   def _get_sharepoint_config(self) -> dict:
       return {
           "site_url": SHAREPOINT_SITE_URL,
           "library": SHAREPOINT_LIBRARY,
           "client_base": SHAREPOINT_CLIENT_BASE,
       }
   ```

3. Remove any `get_setting("sharepoint_site_url")`, `get_setting("sharepoint_library")`, `get_setting("sharepoint_client_base")` calls throughout the file.

### Part B — Fix auth to use the exact same token path as email

4. Print GRAPH_SCOPES at module load so we can verify:
   ```python
   print(f"[SharePoint] GRAPH_SCOPES = {GRAPH_SCOPES}")
   ```
   Verify it contains `Sites.Read.All`, `Sites.ReadWrite.All`, and `Files.ReadWrite.All` alongside the Mail scopes. If any are missing, add them.

5. Find how the WORKING email methods get their token. Trace the code path from `create_draft()` or `fetch_unread_emails()` — how do they make authenticated Graph API calls? They likely go through `_make_request()` which calls `_get_token()` which calls `acquire_token_silent()`.

6. Now trace how `get_sharepoint_site_id()` and `test_sharepoint_connection()` get their token. Do they use the same `_make_request()`? Or do they use raw `requests.get()` with a separate `_get_token()` call?

7. Make them ALL use the EXACT same path. If `_make_request` works for email, use `_make_request` for SharePoint too:
   ```python
   def get_sharepoint_site_id(self) -> Optional[str]:
       config = self._get_sharepoint_config()
       from urllib.parse import urlparse
       parsed = urlparse(config["site_url"])
       hostname = parsed.hostname
       site_path = parsed.path
       
       # Use _make_request — the SAME method that works for email
       url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
       result = self._make_request("GET", url)
       
       if result and "id" in result:
           return result["id"]
       
       # Fallback: search by name
       site_name = site_path.rstrip("/").split("/")[-1]
       url = f"https://graph.microsoft.com/v1.0/sites?search={site_name}"
       result = self._make_request("GET", url)
       if result and "value" in result and result["value"]:
           return result["value"][0].get("id")
       
       return None
   ```

8. Rewrite `test_sharepoint_connection()` to also use `_make_request()` instead of raw requests:
   ```python
   def test_sharepoint_connection(self) -> dict:
       if not self.is_authenticated():
           return {"ok": False, "error": "Not signed into Microsoft. Please sign in first."}
       
       # Step 1: Find site
       site_id = self.get_sharepoint_site_id()
       if not site_id:
           return {"ok": False, "error": "Could not find SharePoint site at " + SHAREPOINT_SITE_URL}
       
       # Step 2: Find drive
       drive_id = self.get_sharepoint_drive_id(site_id)
       if not drive_id:
           return {"ok": False, "error": f"Could not find document library '{SHAREPOINT_LIBRARY}'"}
       
       # Step 3: Check client base folder
       config = self._get_sharepoint_config()
       url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{config['client_base']}:/children?$top=1"
       result = self._make_request("GET", url)
       if result and "value" in result:
           return {"ok": True, "message": f"Connected to SharePoint. Found client folders in /{config['client_base']}/"}
       
       return {"ok": False, "error": f"Connected to site but could not find /{config['client_base']}/ folder"}
   ```

9. Also make sure `_make_request` supports the colon-separated URL format (`/sites/hostname:/path`). Some implementations might URL-encode the colon. Test by adding a print:
   ```python
   # In _make_request, at the top:
   print(f"[Graph API] {method} {url}")
   ```

### Part C — Verify scopes include SharePoint

10. If GRAPH_SCOPES does NOT already contain the SharePoint scopes, the token will never have SharePoint access. Find GRAPH_SCOPES and ensure it is:
    ```python
    GRAPH_SCOPES = [
        "Mail.ReadWrite",
        "Mail.Send",
        "User.Read",
        "Files.Read.All",
        "Files.ReadWrite.All",
        "Sites.Read.All",
        "Sites.ReadWrite.All",
    ]
    ```

11. IMPORTANT: If scopes were changed, the user MUST delete their MSAL cache and re-sign in for the new scopes to take effect. Add a log message on startup:
    ```python
    print(f"[Auth] Requesting scopes: {GRAPH_SCOPES}")
    ```

### Part D — Clean up frontend

12. In `Settings.tsx`, remove the SharePoint configuration section entirely (Site URL, Document Library, Client Folder Base fields). These are now hardcoded.

13. Keep the Test Connection button but simplify it — move it to a single line:
    ```
    SharePoint: [Test Connection] ✓ Connected / ✗ Not connected
    ```

14. Remove the `/api/sharepoint/debug` endpoint from `api_server.py` if it still exists.

15. Remove the settings save/load code for `sharepoint_site_url`, `sharepoint_library`, `sharepoint_client_base` from the frontend API calls.

### Part E — Rebuild

```
cd frontend && pnpm build && xcopy /E /I /Y dist\public ..\frontend_dist\
```

**Test:**

Start the app. The console should print:
```
[Auth] Requesting scopes: ['Mail.ReadWrite', 'Mail.Send', 'User.Read', 'Files.Read.All', 'Files.ReadWrite.All', 'Sites.Read.All', 'Sites.ReadWrite.All']
[SharePoint] GRAPH_SCOPES = [...]
```

After sign-in, click Test Connection in Settings. Should show green.

If it still fails after this fix, the console will show the exact Graph API URL and response code, which tells us definitively whether it's a scope issue (403), URL issue (404), or something else.

**Commit message:** `fix: hardcode SharePoint config for MC&S, fix auth to use same token path as email`

---

## After the fix

If Test Connection works:
- Delete the old MSAL cache and re-sign in if scopes were changed
- Test "Save to Client Folder" from AI Chat
- Rebuild installer for distribution

If Test Connection still fails:
- Check the console output for the exact Graph API URL and status code
- If 403: scopes aren't in the token — delete MSAL cache, re-sign in
- If 404: URL format is wrong — the console print will show what was sent
- If 401: token retrieval failed — check _make_request output
