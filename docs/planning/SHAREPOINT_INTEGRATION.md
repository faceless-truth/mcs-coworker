# SHAREPOINT_INTEGRATION.md — Setup Guide + Claude Code Task List

SharePoint integration for MCS CoWorker — save AI Chat exports, email correspondence, and documents directly to client folders on SharePoint.

---

## Part 1: Azure Portal Setup (Manual — do this first)

### Step 1: Add SharePoint permissions to your Azure App Registration

1. Go to **portal.azure.com**
2. Navigate to **Azure Active Directory → App registrations**
3. Find your existing **MCS CoWorker** app registration (the one already used for Outlook/Graph)
4. Click **API permissions** in the left sidebar
5. Click **+ Add a permission**
6. Select **Microsoft Graph**
7. Select **Delegated permissions**
8. Search for and add these permissions:
   - `Sites.Read.All` — read SharePoint sites and document libraries
   - `Sites.ReadWrite.All` — read and write files to SharePoint
   - `Files.ReadWrite.All` — read and write files in OneDrive/SharePoint
9. Click **Add permissions**
10. Click **Grant admin consent for MC&S** (the green button at the top) — you need to be a tenant admin for this
11. Verify all permissions show a green tick under "Status"

### Step 2: Find your SharePoint site details

1. Open your SharePoint site in a browser (e.g., `https://mcands.sharepoint.com/sites/ClientFiles` or similar)
2. Note the URL — you'll need the **site hostname** and **site path**:
   - Hostname: `mcands.sharepoint.com`
   - Site path: `/sites/ClientFiles` (or whatever your site is called)
3. Find the **document library** name where client folders live:
   - Usually called "Documents" or "Shared Documents" or "Client Files"
   - Click into it and note the name from the URL or breadcrumb

### Step 3: Verify the folder structure

Confirm your SharePoint client folder structure looks like:
```
/Documents/Clients/Korkie, Gordon/
/Documents/Clients/Korkie, Gordon/Korkie Family Trust/
/Documents/Clients/Korkie, Gordon/Korkie Holdings Pty Ltd/
/Documents/Clients/Smith, Jane/
```

Note the exact path from the document library root to the client folders. For example:
- If client folders are at: `Documents > Clients > Korkie, Gordon`
- Then the base path is: `Clients` (relative to the document library)

### Step 4: Update CoWorker Settings

After the code changes below are deployed:
1. Open CoWorker → Settings
2. Find the new **SharePoint** section
3. Enter:
   - **SharePoint Site URL**: `https://mcands.sharepoint.com/sites/ClientFiles`
   - **Document Library**: `Documents` (or whatever yours is called)
   - **Client Folder Base**: `Clients` (the folder within the library that contains all client folders)
4. Click **Test Connection** — should show a green tick if permissions are correct
5. Save

---

## Part 2: Code Changes (Claude Code task file)

Work through each fix in order.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

### Fix 1 of 4: Add SharePoint Graph API methods to graph_client.py

**Files:** `graph_client.py`

**Problem:** The Graph client currently only handles Outlook operations. Need to add SharePoint file read/write capabilities using the same authenticated session.

**Changes required:**

1. Add SharePoint configuration reader:
   ```python
   def _get_sharepoint_config(self) -> dict:
       """Read SharePoint config from settings."""
       return {
           "site_url": get_setting("sharepoint_site_url", ""),
           "library": get_setting("sharepoint_library", "Documents"),
           "client_base": get_setting("sharepoint_client_base", "Clients"),
       }
   ```

2. Add method to resolve the SharePoint site ID:
   ```python
   def get_sharepoint_site_id(self) -> Optional[str]:
       """Resolve the SharePoint site ID from the configured URL."""
       config = self._get_sharepoint_config()
       if not config["site_url"]:
           return None
       
       # Parse the URL: https://mcands.sharepoint.com/sites/ClientFiles
       # Graph API: GET /sites/{hostname}:/{site-path}
       from urllib.parse import urlparse
       parsed = urlparse(config["site_url"])
       hostname = parsed.hostname  # mcands.sharepoint.com
       site_path = parsed.path     # /sites/ClientFiles
       
       url = f"{self.base_url}/sites/{hostname}:{site_path}"
       result = self._make_request("GET", url)
       if result and "id" in result:
           return result["id"]
       return None
   ```

3. Add method to get the document library drive ID:
   ```python
   def get_sharepoint_drive_id(self, site_id: str) -> Optional[str]:
       """Get the drive ID for the configured document library."""
       config = self._get_sharepoint_config()
       url = f"{self.base_url}/sites/{site_id}/drives"
       result = self._make_request("GET", url)
       if result and "value" in result:
           for drive in result["value"]:
               if drive.get("name", "").lower() == config["library"].lower():
                   return drive["id"]
           # Fallback: return the first drive
           if result["value"]:
               return result["value"][0]["id"]
       return None
   ```

4. Add method to check if a client folder exists:
   ```python
   def sharepoint_folder_exists(self, client_name: str, entity_name: str = None) -> bool:
       """Check if a client folder exists in SharePoint."""
       site_id = self.get_sharepoint_site_id()
       if not site_id:
           return False
       drive_id = self.get_sharepoint_drive_id(site_id)
       if not drive_id:
           return False
       
       config = self._get_sharepoint_config()
       folder_path = f"{config['client_base']}/{client_name}"
       if entity_name:
           folder_path += f"/{entity_name}"
       
       url = f"{self.base_url}/drives/{drive_id}/root:/{folder_path}"
       result = self._make_request("GET", url)
       return result is not None and "id" in result
   ```

5. Add method to upload a file to a client folder:
   ```python
   def upload_to_sharepoint(
       self,
       file_content: bytes,
       filename: str,
       client_name: str,
       entity_name: str = None,
       subfolder: str = None,
   ) -> Optional[str]:
       """
       Upload a file to a client's SharePoint folder.
       
       Returns the SharePoint file URL if successful, None if failed.
       
       Path structure:
         /{library}/{client_base}/{client_name}/{entity_name}/{subfolder}/{filename}
       Example:
         /Documents/Clients/Korkie, Gordon/Korkie Family Trust/Tax Returns/recommendation_2026-04-28.docx
       """
       site_id = self.get_sharepoint_site_id()
       if not site_id:
           logger.error("SharePoint site ID not found — check sharepoint_site_url setting")
           return None
       
       drive_id = self.get_sharepoint_drive_id(site_id)
       if not drive_id:
           logger.error("SharePoint drive ID not found — check sharepoint_library setting")
           return None
       
       config = self._get_sharepoint_config()
       
       # Build the folder path
       folder_path = config["client_base"]
       if client_name:
           folder_path += f"/{client_name}"
       if entity_name:
           folder_path += f"/{entity_name}"
       if subfolder:
           folder_path += f"/{subfolder}"
       
       # Upload using the simple upload API (for files < 4MB)
       # For larger files, use the upload session API
       file_path = f"{folder_path}/{filename}"
       
       if len(file_content) < 4 * 1024 * 1024:  # < 4MB
           url = f"{self.base_url}/drives/{drive_id}/root:/{file_path}:/content"
           headers = {"Content-Type": "application/octet-stream"}
           result = self._make_request("PUT", url, data=file_content, headers=headers)
       else:
           # Create upload session for large files
           url = f"{self.base_url}/drives/{drive_id}/root:/{file_path}:/createUploadSession"
           session = self._make_request("POST", url, json={
               "item": {"@microsoft.graph.conflictBehavior": "rename"}
           })
           if not session or "uploadUrl" not in session:
               logger.error("Failed to create upload session")
               return None
           
           # Upload in chunks
           upload_url = session["uploadUrl"]
           chunk_size = 10 * 1024 * 1024  # 10MB chunks
           for i in range(0, len(file_content), chunk_size):
               chunk = file_content[i:i + chunk_size]
               end = min(i + chunk_size, len(file_content))
               headers = {
                   "Content-Range": f"bytes {i}-{end - 1}/{len(file_content)}",
                   "Content-Type": "application/octet-stream",
               }
               import requests
               resp = requests.put(upload_url, data=chunk, headers=headers)
               if resp.status_code not in (200, 201, 202):
                   logger.error(f"Upload chunk failed: {resp.status_code}")
                   return None
               result = resp.json() if resp.status_code in (200, 201) else None
       
       if result and "webUrl" in result:
           logger.info(f"Uploaded to SharePoint: {file_path}")
           return result["webUrl"]
       
       return None
   ```

6. Add method to list files in a client folder:
   ```python
   def list_sharepoint_folder(self, client_name: str, entity_name: str = None) -> list:
       """List files in a client's SharePoint folder."""
       site_id = self.get_sharepoint_site_id()
       if not site_id:
           return []
       drive_id = self.get_sharepoint_drive_id(site_id)
       if not drive_id:
           return []
       
       config = self._get_sharepoint_config()
       folder_path = f"{config['client_base']}/{client_name}"
       if entity_name:
           folder_path += f"/{entity_name}"
       
       url = f"{self.base_url}/drives/{drive_id}/root:/{folder_path}:/children"
       result = self._make_request("GET", url)
       
       if result and "value" in result:
           return [
               {
                   "name": item["name"],
                   "size": item.get("size", 0),
                   "modified": item.get("lastModifiedDateTime", ""),
                   "url": item.get("webUrl", ""),
                   "is_folder": "folder" in item,
               }
               for item in result["value"]
           ]
       return []
   ```

7. Add a test connection method:
   ```python
   def test_sharepoint_connection(self) -> dict:
       """Test the SharePoint connection and return status."""
       try:
           site_id = self.get_sharepoint_site_id()
           if not site_id:
               return {"ok": False, "error": "Could not find SharePoint site. Check the URL in settings."}
           
           drive_id = self.get_sharepoint_drive_id(site_id)
           if not drive_id:
               return {"ok": False, "error": "Could not find document library. Check the library name in settings."}
           
           config = self._get_sharepoint_config()
           # Try to list the client base folder
           url = f"{self.base_url}/drives/{drive_id}/root:/{config['client_base']}:/children?$top=1"
           result = self._make_request("GET", url)
           if result and "value" in result:
               return {"ok": True, "message": f"Connected. Found client folders in /{config['client_base']}/"}
           
           return {"ok": False, "error": f"Could not access /{config['client_base']}/ folder. Check the client folder base path."}
       except Exception as e:
           return {"ok": False, "error": str(e)}
   ```

8. Ensure `_make_request` supports PUT method (for file uploads). If it only handles GET/POST/PATCH, add PUT support.

**Test:** `python -c "from graph_client import GraphClient; print('SharePoint methods added OK')"`

**Commit message:** `feat(sharepoint/1of4): add SharePoint Graph API methods for file read/write to client folders`

---

### Fix 2 of 4: Add SharePoint API endpoints

**Files:** `api_server.py`

**Changes required:**

1. Add SharePoint settings endpoints:
   ```python
   @app.route("/api/sharepoint/test", methods=["POST"])
   def test_sharepoint():
       """Test the SharePoint connection."""
       graph = app.config.get("graph_client")
       if not graph:
           return jsonify({"ok": False, "error": "Graph client not initialised"})
       result = graph.test_sharepoint_connection()
       return jsonify(result)
   
   @app.route("/api/sharepoint/folders", methods=["GET"])
   def list_sharepoint_folders():
       """List client folders in SharePoint."""
       client_name = request.args.get("client", "")
       entity_name = request.args.get("entity", "")
       graph = app.config.get("graph_client")
       if not graph:
           return jsonify({"ok": False, "error": "Graph client not initialised"})
       files = graph.list_sharepoint_folder(client_name or None, entity_name or None)
       return jsonify({"ok": True, "files": files})
   ```

2. Add a "Save to SharePoint" export endpoint:
   ```python
   @app.route("/api/chat/export/sharepoint", methods=["POST"])
   def export_to_sharepoint():
       """Export a chat conversation directly to a client's SharePoint folder."""
       data = request.json or {}
       messages = data.get("messages", [])
       agent_name = data.get("agent_name", "Chat")
       agent_id = data.get("agent_id", "general")
       export_type = data.get("export_type", "transcript")
       client_name = data.get("client_name", "")
       entity_name = data.get("entity_name", "")
       
       if not client_name:
           return jsonify({"ok": False, "error": "Client name is required to save to SharePoint"}), 400
       
       graph = app.config.get("graph_client")
       if not graph:
           return jsonify({"ok": False, "error": "Graph client not initialised"})
       
       # Generate the document (reuse existing export logic)
       # ... build the docx using the same _markdown_to_docx or transcript logic ...
       
       from datetime import datetime
       timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
       filename = f"{agent_id}_{export_type}_{timestamp}.docx"
       
       # Get the docx bytes
       # (reuse the existing export_chat logic to build the BytesIO buffer)
       buffer = _build_export_document(messages, agent_name, export_type, client_name, entity_name)
       
       if not buffer:
           return jsonify({"ok": False, "error": "Failed to generate document"}), 500
       
       # Upload to SharePoint
       subfolder = "CoWorker Exports"  # Creates a subfolder for CoWorker files
       url = graph.upload_to_sharepoint(
           file_content=buffer.read(),
           filename=filename,
           client_name=client_name,
           entity_name=entity_name,
           subfolder=subfolder,
       )
       
       if url:
           return jsonify({"ok": True, "url": url, "filename": filename})
       else:
           return jsonify({"ok": False, "error": "Failed to upload to SharePoint. Check connection settings."}), 500
   ```

3. Refactor the existing export logic into a reusable `_build_export_document()` function that both the download and SharePoint endpoints can use.

**Commit message:** `feat(sharepoint/2of4): add SharePoint API endpoints for test, browse, and export`

---

### Fix 3 of 4: Add SharePoint settings and "Save to Client Folder" button

**Files:** `frontend/client/src/pages/Settings.tsx`, `frontend/client/src/pages/Chat.tsx`

**Changes required:**

### Settings.tsx

1. Add a **SharePoint** section after the Microsoft 365 section:
   ```
   ┌─────────────────────────────────────────────────────────────┐
   │  SharePoint — Client File Storage                           │
   │  Save AI Chat exports and correspondence to client folders  │
   │                                                             │
   │  Site URL:        [https://mcands.sharepoint.com/sites/...] │
   │  Document Library: [Documents                             ] │
   │  Client Folder Base: [Clients                             ] │
   │                                                             │
   │  [Test Connection]  ✓ Connected. Found client folders.      │
   └─────────────────────────────────────────────────────────────┘
   ```

2. The "Test Connection" button calls `POST /api/sharepoint/test` and shows the result (green tick or red error).

3. All three fields save via the existing `/api/settings` PATCH endpoint:
   - `sharepoint_site_url`
   - `sharepoint_library`
   - `sharepoint_client_base`

### Chat.tsx

4. Add a fourth option to the Export dropdown:
   ```
   📥 Export ▾
   ├── Download Transcript (.docx)
   ├── Download Summary (.docx)
   ├── Download Recommendation (.docx)
   └── 📁 Save to Client Folder (SharePoint)
   ```

5. The "Save to Client Folder" option:
   - Only enabled when a client name has been entered in the client field
   - Only enabled when SharePoint settings are configured (check if `sharepoint_site_url` is set)
   - When clicked, calls `POST /api/chat/export/sharepoint` with the conversation, export_type="recommendation" (or "transcript" if no structured content)
   - Shows a loading state: "Saving to SharePoint..."
   - On success: shows a toast with "Saved to SharePoint" and a clickable link to the file
   - On failure: shows the error message

6. The button should be greyed out with a tooltip if:
   - No client name entered: "Enter a client name to save to SharePoint"
   - SharePoint not configured: "Configure SharePoint in Settings first"

### Rebuild frontend

```
cd frontend && pnpm build && xcopy /E /I /Y dist\public ..\frontend_dist\
```

**Commit message:** `feat(sharepoint/3of4): add SharePoint settings UI and Save to Client Folder export option`

---

### Fix 4 of 4: Auto-file Smart Responder correspondence to SharePoint

**Files:** `plugins/plugin_smart_responder.py`, `plugins/plugin_correspondence_logger.py`

**Problem:** When Smart Responder drafts a reply to a client, a copy should be saved to the client's SharePoint folder for the file record. This makes the correspondence log accessible alongside other client documents.

**Changes required:**

1. In `plugins/plugin_smart_responder.py`, after a successful draft creation:
   ```python
   # Save a copy to SharePoint if configured
   sharepoint_url = get_setting("sharepoint_site_url", "")
   if sharepoint_url and client_name:
       try:
           from client_utils import normalise_client_name
           normalised = normalise_client_name(client_name)
           
           # Build a simple text file with the email details
           from datetime import datetime
           timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
           content = f"""Subject: {subject}
Date: {datetime.now().strftime('%d %B %Y %H:%M')}
From: elio@mcands.com.au
To: {sender_email}
Status: Draft created by CoWorker

---

{draft_body_text}
"""
           context.graph.upload_to_sharepoint(
               file_content=content.encode("utf-8"),
               filename=f"correspondence_{timestamp}.txt",
               client_name=normalised,
               subfolder="CoWorker Correspondence",
           )
       except Exception as e:
           logger.warning(f"SharePoint auto-file failed (non-fatal): {e}")
   ```

2. This should be a **best-effort** operation — if SharePoint upload fails, the draft is still created successfully. The SharePoint save is a bonus, not a requirement.

3. Only auto-file if SharePoint is configured (the `sharepoint_site_url` setting is non-empty). If not configured, skip silently.

**Test:** 
1. Configure SharePoint in Settings
2. Send a test email → let Smart Responder create a draft
3. Check SharePoint → client folder should have a new `CoWorker Correspondence/` subfolder with the correspondence file

**Commit message:** `feat(sharepoint/4of4): auto-file Smart Responder correspondence to client SharePoint folders`

---

## Post-setup checklist

- [ ] Azure portal: SharePoint permissions added and admin consent granted
- [ ] Settings: SharePoint URL, library, and base path configured
- [ ] Test Connection shows green tick
- [ ] AI Chat export → "Save to Client Folder" works for a test client
- [ ] File appears in the correct SharePoint folder: `Clients/Korkie, Gordon/CoWorker Exports/`
- [ ] Smart Responder auto-files correspondence (if SharePoint configured)
- [ ] SharePoint upload failures don't crash plugins (best-effort)
- [ ] Rebuild installer: `build_installer.bat`
