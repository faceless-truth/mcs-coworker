# EXPORT_DOWNLOAD_FIX.md — Claude Code Task List

Fix chat exports to save directly to the user's Downloads folder instead of using blob downloads that pywebview silently drops.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Save chat exports directly to Downloads folder

**Files:** `api_server.py`, `frontend/client/src/pages/Chat.tsx`

**Problem:** When the user clicks Export → Download Transcript/Summary/Recommendation, the frontend generates a blob URL and tries to trigger a browser download. This works in Chrome/Firefox but pywebview silently drops blob downloads — the file is generated but never saved. The user sees "exported" but nothing appears in their Downloads folder.

**Changes required:**

### Backend — api_server.py

1. Find the `/api/chat/export` POST endpoint. After generating the docx BytesIO buffer, save it directly to the user's Downloads folder:
   ```python
   from pathlib import Path
   
   # Save to Downloads folder
   downloads_dir = Path.home() / "Downloads"
   downloads_dir.mkdir(parents=True, exist_ok=True)
   filepath = downloads_dir / filename
   
   # Write the docx file
   buffer.seek(0)
   filepath.write_bytes(buffer.read())
   
   # Return the path so the frontend can show a confirmation
   return jsonify({
       "ok": True,
       "path": str(filepath),
       "filename": filename,
   })
   ```

2. Remove or keep the `send_file()` response as a fallback — but the primary response should be the JSON with the file path. If other clients (not pywebview) need the blob download, you can check for an `Accept` header or a query parameter to decide which response format to use. For simplicity, just return JSON always since CoWorker only runs in pywebview.

3. Make sure this works for ALL three export types (transcript, summary, recommendation). The summary export calls Claude to generate the summary first — after that, the same save-to-Downloads logic applies.

### Frontend — Chat.tsx

4. Update ALL export handlers (`handleExport` or similar) to:
   ```typescript
   const handleExport = async (exportType: 'transcript' | 'summary' | 'recommendation') => {
       try {
           // Show loading state for summary (it calls Claude)
           if (exportType === 'summary') {
               // show loading toast
           }
           
           const response = await fetch('/api/chat/export', {
               method: 'POST',
               headers: {
                   'Authorization': `Bearer ${token}`,
                   'Content-Type': 'application/json',
               },
               body: JSON.stringify({
                   messages: messages,
                   agent_name: selectedAgentName,
                   agent_id: selectedAgentId,
                   export_type: exportType,
                   client_name: clientName || null,
                   entity_name: entityName || null,
               }),
           });
           
           const result = await response.json();
           
           if (result.ok) {
               // Show success toast with the filename
               // e.g. "Saved to Downloads: gst_transcript_2026-04-30_1323.docx"
               showToast(`Saved to Downloads: ${result.filename}`, 'success');
           } else {
               showToast(`Export failed: ${result.error}`, 'error');
           }
       } catch (error) {
           showToast('Export failed', 'error');
       }
   };
   ```

5. Remove ALL blob-related code from the export handlers:
   - No `response.blob()`
   - No `URL.createObjectURL()`
   - No `document.createElement('a')` with download attribute
   - No `URL.revokeObjectURL()`
   
   Just POST, read JSON, show toast.

6. The "Save to Client Folder (SharePoint)" export option should continue to work as before — it saves to SharePoint, not Downloads. No change needed there.

### Rebuild frontend

```
cd frontend && pnpm build && xcopy /E /I /Y dist\public ..\frontend_dist\
```

**Test:**
1. Open AI Chat, have a conversation with any agent
2. Click Export → Download Transcript
3. Check `C:\Users\Elio\Downloads\` — file should be there with name like `general_chat_transcript_2026-04-30_1330.docx`
4. Toast should say "Saved to Downloads: general_chat_transcript_2026-04-30_1330.docx"
5. Open the docx — proper formatting with headings, MC&S footer
6. Test Download Summary — should take a moment (Claude generates it) then save to Downloads
7. Test Download Recommendation — should save to Downloads (or show "no structured recommendation" if none exists)

**Commit message:** `fix: save chat exports directly to Downloads folder — pywebview does not support blob downloads`

---

## Post-fix checklist

- [ ] Transcript export saves to Downloads folder
- [ ] Summary export saves to Downloads folder
- [ ] Recommendation export saves to Downloads folder
- [ ] Toast notification shows the filename
- [ ] No silent failures — if export fails, error toast appears
- [ ] SharePoint export still works separately
- [ ] Rebuild installer: `build_installer.bat`
