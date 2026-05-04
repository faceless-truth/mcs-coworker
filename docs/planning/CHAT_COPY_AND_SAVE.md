# CHAT_COPY_AND_SAVE.md — Claude Code Task List

Three improvements to the AI Chat: enable text selection/copy, auto-save all conversations to memory, and add a SharePoint folder picker for choosing where to save exports.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 3: Enable text selection and copy in chat messages

**Files:** `frontend/client/src/pages/Chat.tsx`, possibly global CSS

**Problem:** The chat message area doesn't allow users to select and copy text. Accountants need to copy snippets from AI responses — a specific section reference, a dollar amount, a paragraph to paste into an email or document.

**Changes required:**

1. Find the chat message container elements in `Chat.tsx`. Check if there's any CSS that prevents text selection:
   ```css
   /* Remove if found: */
   user-select: none;
   -webkit-user-select: none;
   pointer-events: none;
   ```

2. Ensure all message text elements have selectable text:
   ```css
   user-select: text;
   -webkit-user-select: text;
   cursor: text;
   ```

3. If pywebview blocks text selection globally or blocks the right-click context menu, override specifically for the chat area:
   ```css
   .chat-messages * {
       user-select: text !important;
       -webkit-user-select: text !important;
   }
   ```

4. Verify that code blocks and tables in AI responses are also selectable and copyable.

**Test:** Open AI Chat, get a response. Click and drag to select text. Ctrl+C. Paste into Notepad.

**Commit message:** `fix: enable text selection and copy in chat messages`

---

## Fix 2 of 3: Auto-save all chat conversations to memory

**Files:** `api_server.py`, `frontend/client/src/pages/Chat.tsx`

**Problem:** Chat conversations are only saved to memory if a client name is provided. All conversations should be auto-saved — they contain research, analysis, and advice that should be searchable later.

**Changes required:**

### Backend — api_server.py

1. Add an endpoint to save a full conversation:
   ```python
   @app.route("/api/chat/save-conversation", methods=["POST"])
   def save_conversation():
       data = request.json or {}
       messages = data.get("messages", [])
       agent_id = data.get("agent_id", "general")
       agent_name = data.get("agent_name", "General Chat")
       client_name = data.get("client_name", "")
       entity_name = data.get("entity_name", "")
       
       store = _get_memory_store()
       if not store or not messages:
           return jsonify({"ok": True, "saved": 0})
       
       from client_utils import normalise_client_name
       normalised = normalise_client_name(client_name) if client_name else "_general"
       
       conversation_parts = []
       for msg in messages:
           role = msg.get("role", "")
           content = msg.get("content", "")[:500]
           if role == "user":
               conversation_parts.append(f"Q: {content}")
           elif role == "assistant":
               conversation_parts.append(f"A: {content}")
       
       full_summary = f"[{agent_name} conversation]\n" + "\n".join(conversation_parts)
       
       store.store_client_interaction(
           client_name=normalised,
           entity_name=entity_name,
           interaction_type="ai_chat_conversation",
           summary=full_summary[:2000],
           metadata={
               "agent_id": agent_id,
               "agent_name": agent_name,
               "message_count": len(messages),
           },
       )
       
       return jsonify({"ok": True, "saved": 1})
   ```

### Frontend — Chat.tsx

2. When switching agents, auto-save the outgoing conversation:
   ```typescript
   const handleAgentSwitch = (newAgentId: string) => {
       if (messages.length > 1) {
           fetch('/api/chat/save-conversation', {
               method: 'POST',
               headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
               body: JSON.stringify({
                   messages, agent_id: selectedAgentId, agent_name: selectedAgentName,
                   client_name: clientName, entity_name: entityName,
               }),
           });
       }
       setSelectedAgentId(newAgentId);
       setMessages([]);
   };
   ```

3. Save on window/app close:
   ```typescript
   useEffect(() => {
       const handleBeforeUnload = () => {
           if (messages.length > 1) {
               navigator.sendBeacon('/api/chat/save-conversation',
                   JSON.stringify({
                       messages, agent_id: selectedAgentId, agent_name: selectedAgentName,
                       client_name: clientName, entity_name: entityName,
                   })
               );
           }
       };
       window.addEventListener('beforeunload', handleBeforeUnload);
       return () => window.removeEventListener('beforeunload', handleBeforeUnload);
   }, [messages, selectedAgentId, clientName, entityName]);
   ```

**Test:** Have a conversation. Switch agents. Check Memory page — conversation should appear.

**Commit message:** `feat: auto-save all chat conversations to memory on agent switch and app close`

---

## Fix 3 of 3: SharePoint folder picker for choosing save location

**Files:** `api_server.py`, `frontend/client/src/pages/Chat.tsx`

**Problem:** The "Save to Client Folder" export currently saves to a hardcoded `CoWorker Exports` subfolder. Accountants need to choose WHERE in the client's SharePoint folder tree to save — e.g., under `Tax Returns/2026`, `Correspondence`, `BAS`, or any custom subfolder.

**Changes required:**

### Backend — api_server.py

1. Add an endpoint to browse SharePoint folders for a specific client:
   ```python
   @app.route("/api/sharepoint/browse", methods=["GET"])
   def browse_sharepoint():
       """Browse folders within a client's SharePoint directory."""
       client_name = request.args.get("client", "")
       path = request.args.get("path", "")  # Subfolder path within client folder
       
       graph = _get_graph()
       if not graph:
           return jsonify({"ok": False, "error": "Not connected"})
       
       if not client_name:
           return jsonify({"ok": False, "error": "Client name required"})
       
       from client_utils import normalise_client_name
       normalised = normalise_client_name(client_name)
       
       # List folder contents
       full_path = normalised
       if path:
           full_path += f"/{path}"
       
       items = graph.list_sharepoint_folder(full_path)
       
       # Separate folders and files
       folders = [item for item in items if item.get("is_folder", False)]
       files = [item for item in items if not item.get("is_folder", False)]
       
       return jsonify({
           "ok": True,
           "current_path": path or "/",
           "client_name": normalised,
           "folders": folders,
           "files": files,
       })
   ```

2. Update the `/api/chat/export/sharepoint` endpoint to accept a `subfolder` parameter from the user:
   ```python
   subfolder = data.get("subfolder", "CoWorker Exports")  # User-chosen or default
   ```

### Frontend — Chat.tsx

3. When the user clicks "Save to Client Folder (SharePoint)", instead of immediately saving, show a **folder picker modal**:

   ```
   ┌─────────────────────────────────────────────────────┐
   │  Save to SharePoint — Korkie, Gordon                │
   │                                                     │
   │  📁 BAS                                             │
   │  📁 Correspondence                                  │
   │  📁 CoWorker Exports                                │
   │  📁 Korkie Family Trust                             │
   │     📁 2025                                         │
   │     📁 2026                                         │
   │  📁 Korkie Holdings Pty Ltd                         │
   │  📁 Tax Returns                                     │
   │     📁 2025                                         │
   │     📁 2026                                         │
   │                                                     │
   │  Current path: /Tax Returns/2026                    │
   │                                                     │
   │  [New Folder]        [Cancel]  [Save Here]          │
   └─────────────────────────────────────────────────────┘
   ```

4. The folder picker should:
   - Fetch `/api/sharepoint/browse?client=Korkie,Gordon` on open to list the top-level folders
   - When a folder is clicked, fetch `/api/sharepoint/browse?client=Korkie,Gordon&path=Tax Returns` to navigate into it
   - Show a breadcrumb trail at the top: `Korkie, Gordon / Tax Returns / 2026`
   - Allow clicking breadcrumb segments to navigate back up
   - Show a "New Folder" button that creates a new subfolder (POST to a new endpoint)
   - Show "Cancel" and "Save Here" buttons
   - "Save Here" calls `/api/chat/export/sharepoint` with the chosen subfolder path
   - Default selection: highlight "CoWorker Exports" if it exists

5. Create the folder picker as a reusable React component `SharePointFolderPicker`:
   ```typescript
   interface SharePointFolderPickerProps {
       clientName: string;
       entityName?: string;
       onSelect: (path: string) => void;
       onCancel: () => void;
       isOpen: boolean;
   }
   ```

6. Add a "New Folder" endpoint:
   ```python
   @app.route("/api/sharepoint/create-folder", methods=["POST"])
   def create_sharepoint_folder():
       data = request.json or {}
       client_name = data.get("client_name", "")
       path = data.get("path", "")  # Full path within client folder
       
       graph = _get_graph()
       if not graph or not client_name:
           return jsonify({"ok": False, "error": "Missing client name or not connected"})
       
       from client_utils import normalise_client_name
       normalised = normalise_client_name(client_name)
       
       # Create folder via Graph API
       # POST /drives/{drive_id}/root:/{folder_path}:/children
       # Body: {"name": "New Folder", "folder": {}}
       folder_name = path.split("/")[-1] if "/" in path else path
       parent_path = "/".join(path.split("/")[:-1]) if "/" in path else ""
       
       full_parent = normalised
       if parent_path:
           full_parent += f"/{parent_path}"
       
       result = graph.create_sharepoint_folder(full_parent, folder_name)
       return jsonify({"ok": bool(result), "path": path})
   ```

7. Add `create_sharepoint_folder` method to `graph_client.py`:
   ```python
   def create_sharepoint_folder(self, parent_path: str, folder_name: str) -> bool:
       """Create a new folder in SharePoint."""
       site_id = self.get_sharepoint_site_id()
       if not site_id:
           return False
       drive_id = self.get_sharepoint_drive_id(site_id)
       if not drive_id:
           return False
       
       config = self._get_sharepoint_config()
       full_path = f"{config['client_base']}/{parent_path}"
       
       url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{full_path}:/children"
       result = self._make_request("POST", url, json={
           "name": folder_name,
           "folder": {},
           "@microsoft.graph.conflictBehavior": "fail",
       })
       return result is not None
   ```

### Export type selection in the picker

8. Before saving, let the user choose the export type within the modal:
   ```
   ┌─────────────────────────────────────────────────┐
   │  Export as:                                      │
   │  ● Transcript  ○ Summary  ○ Recommendation      │
   │                                                  │
   │  Current path: /Tax Returns/2026                 │
   │                                                  │
   │  [Cancel]                    [Save Here]         │
   └─────────────────────────────────────────────────┘
   ```

9. If SharePoint upload fails, fall back to saving to Downloads and show a toast:
   ```
   "SharePoint upload failed — saved to Downloads instead: filename.docx"
   ```

### Rebuild frontend

```
cd frontend && pnpm build && xcopy /E /I /Y dist\public ..\frontend_dist\
```

**Test:**
1. Open AI Chat, enter client name "Korkie, Gordon", have a conversation
2. Click Export → Save to Client Folder
3. Folder picker modal opens showing the client's SharePoint folders
4. Navigate into "Tax Returns" → "2026"
5. Click "Save Here"
6. Toast: "Saved to SharePoint: Korkie, Gordon/Tax Returns/2026/gst_transcript_2026-04-30.docx"
7. Check SharePoint — file should be in that folder
8. If SharePoint fails — file saves to Downloads with an error toast

**Commit message:** `feat: add SharePoint folder picker for choosing export save location with folder browsing`

---

## Done — Post-fix checklist

- [ ] Can select and copy any text in chat messages
- [ ] Switching agents saves conversation to memory
- [ ] Closing app saves conversation to memory
- [ ] Memory page shows all chat conversations (searchable)
- [ ] Export → Save to Client Folder opens folder picker
- [ ] Can browse client's SharePoint folder tree
- [ ] Can create new folders from the picker
- [ ] Can choose export type (Transcript/Summary/Recommendation) in the picker
- [ ] Successful save shows toast with full path
- [ ] Failed save falls back to Downloads with error message
- [ ] Rebuild installer: `build_installer.bat`
