# MEMORY_DETAIL_VIEW.md — Claude Code Task List

Add clickable memory entries that expand to show full conversation details, recommendations, and context.

Repository: `C:\Users\Elio\mcs-coworker`
Branch: `main`

---

## Fix 1 of 1: Add expandable detail view for memory entries

**Files:** `frontend/client/src/pages/Memory.tsx`, `api_server.py`, `memory_store.py`

**Problem:** The Memory page shows a list of interactions with a short summary preview, but clicking on an entry does nothing. For AI chat conversations, the summary is truncated and you can't see what was actually discussed or recommended. Accountants need to click into a memory entry and see the full content — especially for specialist agent conversations that contain detailed analysis, recommendations, and advice.

**Changes required:**

### Part A — Backend: return full content on demand

1. Add an endpoint to get the full detail of a single memory entry:
   ```python
   @app.route("/api/memory/<doc_id>", methods=["GET"])
   def get_memory_detail(doc_id):
       """Get full details of a single memory entry."""
       store = _get_memory_store()
       if not store:
           return jsonify({"ok": False, "error": "Memory not available"})
       
       # Try each collection
       for collection_name in ["interactions", "lessons", "documents"]:
           try:
               col = store._get_collection(collection_name)
               result = col.get(
                   ids=[doc_id],
                   include=["documents", "metadatas"],
               )
               if result and result["ids"]:
                   return jsonify({
                       "ok": True,
                       "entry": {
                           "id": result["ids"][0],
                           "content": result["documents"][0] if result["documents"] else "",
                           "metadata": result["metadatas"][0] if result["metadatas"] else {},
                           "collection": collection_name,
                       }
                   })
           except Exception:
               continue
       
       return jsonify({"ok": False, "error": "Entry not found"}), 404
   ```

2. Update the `/api/chat/save-conversation` endpoint to store MORE content — currently it truncates to 2000 chars. Increase to 10000 chars so full conversations are preserved:
   ```python
   full_summary = f"[{agent_name} conversation]\n" + "\n".join(conversation_parts)
   store.store_client_interaction(
       ...
       summary=full_summary[:10000],  # Store more of the conversation
       ...
   )
   ```

3. Also store the full conversation as structured metadata so it can be rendered properly:
   ```python
   metadata={
       "agent_id": agent_id,
       "agent_name": agent_name,
       "message_count": len(messages),
       "full_messages": json.dumps(messages[:50]),  # Store the full message array as JSON
   }
   ```

### Part B — Frontend: clickable expandable memory entries

4. In `Memory.tsx`, make each memory entry clickable. When clicked, expand it to show the full content in a detail panel:

   ```
   ┌──────────────────────────────────────────────────────────────────┐
   │ 🔵 Abdallah, Rame    email_draft    32m ago                     │
   │ Draft reply RE: R&D Application - Sure thing Rame...            │
   └──────────────────────────────────────────────────────────────────┘
   
   ↓ Click to expand ↓
   
   ┌──────────────────────────────────────────────────────────────────┐
   │ 🔵 Abdallah, Rame    email_draft    32m ago              [✕]   │
   ├──────────────────────────────────────────────────────────────────┤
   │                                                                  │
   │ Type: email_draft                                                │
   │ Agent: Smart Responder                                           │
   │ Date: 1 May 2026, 12:32 PM                                      │
   │                                                                  │
   │ Full Content:                                                    │
   │ ─────────────                                                    │
   │ Draft reply RE: R&D Application                                  │
   │                                                                  │
   │ Sure thing Rame, I'll hold off on the company return             │
   │ for now and will be in touch once we're ready to                 │
   │ proceed.                                                         │
   │                                                                  │
   └──────────────────────────────────────────────────────────────────┘
   ```

5. For `ai_chat_conversation` entries, render the full conversation with proper formatting:

   ```
   ┌──────────────────────────────────────────────────────────────────┐
   │ 🔵 _General    ai_chat_conversation    1h ago            [✕]   │
   ├──────────────────────────────────────────────────────────────────┤
   │                                                                  │
   │ Agent: Division 7A Specialist                                    │
   │ Date: 1 May 2026, 11:15 AM                                      │
   │ Messages: 8                                                      │
   │                                                                  │
   │ Conversation:                                                    │
   │ ─────────────                                                    │
   │ 👤 You: Business was sold in 2018 for $1.2m. The                │
   │    company still has a $200k loan to the director...             │
   │                                                                  │
   │ 🤖 Div 7A Specialist: Based on the facts provided,              │
   │    this loan would be caught by Division 7A under                │
   │    s109D of ITAA 1936...                                         │
   │                                                                  │
   │ 👤 You: What about the benchmark rate?                           │
   │                                                                  │
   │ 🤖 Div 7A Specialist: The current benchmark rate                │
   │    for complying loans under s109N is...                         │
   │                                                                  │
   └──────────────────────────────────────────────────────────────────┘
   ```

6. Implementation approach — two options, use Option A:

   **Option A — Inline expansion (recommended):**
   - Click a memory entry → it expands in place showing the full content below
   - Click again or click ✕ → collapses back to the summary
   - Fetch full content from `/api/memory/<doc_id>` on first click, cache it
   - Only one entry can be expanded at a time (clicking another collapses the previous)

7. The expanded view should:
   - Show all metadata fields: type, agent, date, client name, entity name
   - Show the full content/summary text with proper line breaks
   - For ai_chat_conversation: parse the `full_messages` metadata and render as a chat-like conversation with alternating user/assistant messages
   - For email_draft: show the draft content
   - For email (inbound/outbound): show the email summary
   - Include a "Copy" button that copies the full content to clipboard
   - Include an "Export" button that downloads the entry as a .docx file (reuse the existing export logic)

8. Styling for the expanded view:
   - Light background (slightly different from the list background)
   - Rounded corners, subtle border
   - User messages aligned left with a person icon
   - Assistant messages aligned left with a bot icon and the agent name
   - Metadata shown in a muted smaller font at the top
   - Smooth expand/collapse animation (CSS transition on max-height)

### Part C — Render markdown in expanded content

9. AI chat conversations often contain markdown formatting (headings, bold, lists, tables, code blocks). The expanded view should render markdown properly:
   - Install or use an existing markdown renderer (check if `react-markdown` or similar is already in the frontend dependencies)
   - If not installed, add `react-markdown` to the frontend:
     ```
     cd frontend && pnpm add react-markdown
     ```
   - Render assistant messages through the markdown renderer
   - User messages can stay as plain text

### Rebuild frontend

```
cd frontend && pnpm build && xcopy /E /I /Y dist\public ..\frontend_dist\
```

**Test:**
1. Open Memory page — entries should show a hover cursor indicating they're clickable
2. Click an email_draft entry → expands showing full draft content
3. Click an ai_chat_conversation entry → expands showing the full conversation with user/assistant messages properly formatted
4. Click the ✕ or the entry again → collapses
5. Click "Copy" → full content copied to clipboard
6. Markdown headings, bold, and lists render properly in expanded assistant messages
7. Click a different entry → previous one collapses, new one expands

**Commit message:** `feat: add expandable detail view for memory entries with full conversation rendering`

---

## Post-fix checklist

- [ ] Memory entries are clickable with hover cursor
- [ ] Clicking expands to show full content
- [ ] AI chat conversations show formatted user/assistant messages
- [ ] Email drafts show the full draft text
- [ ] Markdown renders in assistant messages (headings, bold, lists)
- [ ] Copy button copies full content
- [ ] Metadata (type, agent, date) displayed clearly
- [ ] Smooth expand/collapse animation
- [ ] Only one entry expanded at a time
- [ ] Rebuild installer: `build_installer.bat`
