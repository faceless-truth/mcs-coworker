# Recon — mark-as-unread + SharePoint folder duplication

Investigation only. No fixes. Findings inform the follow-up task files.

Branch: `chore/recon-unread-and-sharepoint-dup`, branched off `main`
(the precondition "after `feat/chat-persistence` merges" had not been
met at recon time — see commit message for context).

All references are to current `main` plus this branch's two new files.
Line numbers were recorded against the same blobs and will be stable
unless `graph_client.py`, `client_utils.py`, etc. are rewritten.

---

## Stream A — Mark-as-unread after drafting

### A1. Where does CoWorker fetch and mark-read incoming emails?

**Single fetch entry point:** `graph_client.py:386` —
`GraphClient.fetch_unread_emails(folder, max_count)`.

- Endpoint: `GET /me/mailFolders/{folder}/messages` (note: folder-scoped,
  not the bare `/me/messages`).
- `$select` is used and explicitly lists
  `id,subject,from,receivedDateTime,body,bodyPreview,hasAttachments,toRecipients`.
  The full `body` payload is requested (not just `bodyPreview`), so the
  callers don't make a second per-message fetch to pull HTML.
- `$filter=isRead eq false`, `$orderby=receivedDateTime desc`, `$top=max_count`.
- Goes through raw `requests.get` with `self._headers()` and
  `r.raise_for_status()` — **not** routed through `_make_request`.
  (`_make_request` is reserved for SharePoint methods, per its docstring at
  `graph_client.py:1142`.)

**Callers of `fetch_unread_emails`:**
- `plugins/plugin_smart_responder.py:136` — main draft-creation loop.
- `plugins/plugin_engagement_letter.py:78` — engagement-letter trigger scan.
- One other plugin uses `fetch_emails_from_sender` (`graph_client.py:879`)
  for the FuseSign nudge path — out of scope here, no body retrieval beyond
  `$select` defaults.

**Does the fetch path mark-as-read as a side effect?** No, not from this
call. The list endpoint with `isRead eq false` filter and explicit `$select`
returns messages without flipping their read state — there is no PATCH or
POST around `fetch_unread_emails` and no client-side render. The read flip
that is happening today comes from a *different* Graph call: see A2.

There is also one explicit mark-read call site inside `graph_client.py`:
`mark_as_read(message_id)` at `graph_client.py:399` is a `PATCH
/me/messages/{id}` with `{"isRead": True}`. It is invoked from every
plugin that processes inbound mail — see grep results below — never as
a side effect, only deliberately:

| Site | Trigger |
|------|--------|
| `plugins/plugin_smart_responder.py:205` | `NO_REPLY_TOKEN` path — Claude says "no reply needed" |
| `plugins/plugin_smart_responder.py:228` | Immediately after successful `create_draft` |
| `plugins/plugin_engagement_letter.py:196` | After draft + FuseSign envelope (or fallback) |
| `plugins/plugin_meeting_prep.py:131`     | After meeting brief draft |
| `plugins/plugin_noa_processor.py:414`    | After NOA processed |
| `plugins/plugin_asic_returns.py:480`     | After ASIC reminder logged |

### A2. Where does draft creation happen?

All three draft-creation entry points live in `graph_client.py` and funnel
through one shared helper:

| Entry point | Location |
|-------------|----------|
| `create_draft(to, subject, body, reply_to_id)` | `graph_client.py:534` |
| `create_draft_with_attachments(...)` | `graph_client.py:955` |
| `create_draft_with_inline_image(...)` | `graph_client.py:1037` (delegates to `create_draft`) |
| `_create_threaded_reply_draft(reply_to_id, body)` | `graph_client.py:566` (private, called by the two above when `reply_to_id` is set) |

`_create_threaded_reply_draft` issues `POST /me/messages/{reply_to_id}/createReply`
at `graph_client.py:574`, then a follow-up `PATCH /me/messages/{draft_id}`
at line 590 to write the AI body above the quoted original.

**Side effect of `createReply`:** Microsoft Graph marks the original
inbound message as read when `createReply` is called against it. This is
documented inline at `graph_client.py:543-547` ("Do NOT mark the original
as unread — that creates an infinite loop where the next plugin run sees
it as unread again and drafts a duplicate reply") and again at
`graph_client.py:600` ("Graph's `createReply` marks the parent as read —
we undo that here. Non-fatal on failure"). So the read flip the user is
seeing happens here, not in `fetch_unread_emails`.

**Plugins each call `create_draft*` directly** — no plugin reimplements
`createReply`. Common helpers, no per-plugin duplication.

Plugin call sites (subset; same shared helper underneath):

```
plugin_smart_responder.py:217   create_draft(reply_to_id=...)
plugin_morning_briefing.py:157  create_draft (no reply context)
plugin_engagement_letter.py:166 create_draft (no reply context)
plugin_bas_reminder.py:156,172
plugin_debtor_followup.py:163,184
plugin_annual_review.py:172,185
plugin_client_outreach.py:390
plugin_fusesign_monitor.py:123
plugin_asic_returns.py:441,603  create_draft_with_attachments
plugin_noa_processor.py:382     create_draft_with_attachments
api_server.py:2530              create_draft (chat → draft)
```

### A3. Existing PATCH patterns

`graph_client.py` already has a working `mark_as_unread` helper —
**a fix does not need to add one**:

- `mark_as_unread(message_id)` at `graph_client.py:405-409`. Body
  `{"isRead": False}`. Routed via raw `requests.patch`, not `_make_request`.
- A wrapper `_reopen_original(message_id)` exists at
  `graph_client.py:598-605` that swallows exceptions around
  `mark_as_unread`. **It is dead code — not called anywhere in the
  repo** (verified: only declaration, zero usages). The comment block
  above it explains the original intent: undo the createReply read flip.

Other PATCH `/me/messages/...` patterns already in the codebase:
- `mark_as_read` (line 399) — body `{"isRead": True}`, raw `requests.patch`.
- `flag_email` (line 626) — body `{"flag": {"flagStatus": "flagged"}}`,
  routed via `_make_request` (returns False on transient error instead of
  raising).
- `add_category` / `remove_category` (lines 632, 655) — GET for current
  categories, then PATCH `{"categories": [...]}`, both via `_make_request`.
- `_create_threaded_reply_draft` (line 591) — PATCH the new draft's body.

So there are two PATCH styles in this file: raw `requests.patch` (used by
`mark_as_read` / `mark_as_unread` / draft-body update) and `_make_request`
(used by category/flag helpers). Neither style is universally "correct" —
worth noting because the fix should pick one consciously.

### A4. Failure-handling philosophy in the email loop

Smart Responder defends against duplicate drafts using a SQLite table
`smart_responder_processed`, **not** purely the Outlook `isRead` flag.
`_is_already_processed(message_id)` at `plugin_smart_responder.py:342`
returns True if a row exists for that Graph message id; the loop guards
the per-email body at line 175 with this check before doing anything
expensive. A row is written by `_mark_as_processed(message_id, draft_id,
action)` at line 250 (after a successful draft), at line 208 (after a
`NO_REPLY` decision), and is NOT written on Claude failure
(line 195-200) or draft-creation failure (line 321-325) — those branches
log, increment `errors`, and `continue`, leaving the row absent so the
next loop will retry. Old rows are pruned to 30 days by
`_cleanup_old_processed` at line 149.

Implication for the fix: the dedup loop concern documented at
`graph_client.py:543-547` ("don't mark unread or you'll loop forever") is
already addressed by this DB-side dedup. Marking unread after drafting
would not actually cause re-drafting under current code — the comment
predates the processed-table guard. The fix can mark unread on draft
*success only*; on failure the message stays in its current state
(read-after-`createReply`) and gets retried via the existing path.

---

## Stream B — SharePoint client folder duplication

### B1. Email processing concurrency

**Single in-process scheduler.** No Windows Task Scheduler entry, no
`--scheduled` CLI mode, no external cron. The scheduler is one
background daemon thread started inside the long-running pywebview
process:

- `main.py:123` — `loader.start_scheduler()` from the GUI bootstrap.
- `plugin_loader.py:522-541` — `start_scheduler` spawns a single
  `_scheduler_thread` (`plugin_loader.py:527-530`), running
  `_scheduler_loop` (line 675).
- Heartbeat ticks come from `EventBus.subscribe("heartbeat.tick", ...)`
  at `plugin_loader.py:294`. Both the heartbeat path and the loop path
  share `self._run_lock` (line 272) so a plugin can't fire twice
  concurrently. `self._inflight` (line 281) is a second-layer guard.
- Plugin executions go through a `ThreadPoolExecutor(max_workers=4)` at
  line 276 — within one process. No cross-process locking.
- No file lock, no DB row lock, no MSAL-backed multi-instance check
  (grep for `lock_file`, `pidfile`, `msvcrt`, `fcntl`, `portalocker`
  returns nothing).
- No Graph subscription / webhook path. Nothing in the repo emits or
  listens for `email.triage.complete` either — that EventBus topic is
  declared at `event_wiring.py:70` and subscribed at line 164 but
  never published, so the cross-plugin event chain wired around it is
  effectively dormant.

If a second instance is launched on Elio's machine (or on Harry/Ross's
when they install the same build), the two processes share the
Outlook mailbox via the same MSAL cache, but each has its own
in-process scheduler and SQLite DB (`MCS_DATA_DIR`-resolved per
install). The duplicate-folder symptom is therefore *unlikely* to be a
race between concurrent in-process runs of one machine; cross-machine
duplicates would still be possible if two installs both auto-file to
SharePoint for the same sender.

### B2. Folder creation call sites

Two distinct mechanisms, both in `graph_client.py`:

1. **Implicit creation as a side effect of upload.**
   `upload_to_sharepoint(file_content, filename, client_name,
   entity_name, subfolder)` at `graph_client.py:1271-1346`. Builds a
   path of the form
   `Server/Clients/{client_name}/{entity_name}/{subfolder}/{filename}`
   and either:
   - PUTs to `/drives/{id}/root:/{file_path}:/content` for files <4 MB
     (line 1306), or
   - POSTs an upload session to `:/createUploadSession` for larger
     files (line 1311), with `@microsoft.graph.conflictBehavior:
     "rename"` set on the session item (line 1313).

   **No existence check is performed on the parent folders before
   upload** — Graph creates intermediate path segments implicitly,
   keyed by the verbatim string the caller supplied. The simple-upload
   path doesn't even pass a conflictBehavior, so two callers with two
   slightly different `client_name` strings produce two folders side
   by side with no warning.

2. **Explicit creation via the in-app folder picker.**
   `create_sharepoint_folder(parent_path, folder_name)` at
   `graph_client.py:1410-1435`. Routed via `_make_request` POST to
   `/drives/{id}/root:/{full_parent}:/children` with body
   `{"name": ..., "folder": {}, "@microsoft.graph.conflictBehavior":
   "fail"}`. Conflict-behavior `fail` means a duplicate at the *exact*
   verbatim name is rejected by Graph itself. Triggered by the API
   route at `api_server.py:2694-2725` (`POST
   /api/sharepoint/create-folder`) when the user clicks "create
   folder" in the SharePoint picker.

There is **no `create_folder` / `ensure_folder` / `_get_or_create_folder`
helper** for the implicit path. `sharepoint_folder_exists(client_name,
entity_name)` at `graph_client.py:1254-1269` exists (GET
`/drives/{id}/root:/{folder_path}` and check for an `id`) but is **not
called by `upload_to_sharepoint` before uploading** — verified by
re-reading lines 1271-1346.

**Triggers of `upload_to_sharepoint`:**

| Site | Trigger | Name source |
|------|---------|-------------|
| `plugins/plugin_smart_responder.py:313` | After successful draft of a reply, files a copy of the draft body under `<client_name>/CoWorker Correspondence/correspondence_<ts>.txt`. | `normalise_client_name(from_name or sender)` (line 277) — every inbound email's sender becomes a candidate folder. |
| `api_server.py:2768`                    | `POST /api/chat/export/sharepoint` — saves a chat transcript to a chosen client. | `normalise_client_name(client_name)` (line 2753) where `client_name` is what the user typed in the chat sidebar. |

The Smart Responder path is the dominant write source: every inbound
email that gets a draft causes a `client_name` derivation from `from_name
or sender`. If the inbound `from_name` is "GORDON KORKIE" the
normaliser produces "Korkie, Gordon"; if it is "Gordon J. Korkie" it
produces "Korkie, Gordon J." — these become two distinct folders.

### B3. Normalisation logic

There is **no** dedicated folder-name sanitiser. Greps for
`_make_safe_folder_name`, `_safe_folder`, `sanitize_folder`,
`sanitise_folder` return zero matches. The recon brief's "user says
some normalisation already exists" expectation maps to **one function**:

**`normalise_client_name(name)` at `client_utils.py:37-84`.**

What it actually does (in this order):
1. `name = name.strip()`.
2. If `is_entity_name(name)` (`client_utils.py:31` — checks against
   suffixes like `pty ltd`, `trust`, `smsf`, `holdings`,
   `enterprises`, `group`, `partnership`, `incorporated`, `inc`,
   `foundation`, `estate` — `client_utils.py:21-28`), return
   `_title_case_entity(name)` which does a `.title()` and then patches
   `Smsf` → `SMSF` and `Pty Ltd` casing.
3. Else if `"@" in name`, replace it with the local part with `._-`
   collapsed to spaces and digits stripped (`_name_from_email`,
   `client_utils.py:87-94`).
4. If the resulting string contains `,`, split into `[surname, first]`
   on the first comma, `.title()` each side, return `"Surname, First"`.
5. Else `.split()` on whitespace; treat the **last whitespace token**
   as the surname; everything before it is the first/middle name;
   return `"Surname.title(), First.title()"`.

What it does **not** do:
- No diacritic / Unicode normalisation (no `unicodedata.normalize`).
- No collapse of repeated whitespace (`"Korkie  Gordon"` and
  `"Korkie Gordon"` end up identical only because `split()` discards
  empties — but `"Korkie ,  Gordon"` after the comma split becomes
  `"Korkie ,  , Gordon"` then `.title()`, with the trailing space
  preserved differently across paths).
- No punctuation stripping. Trailing periods (`"Gordon J."` →
  `"Korkie, Gordon J."`), apostrophes (`"O'Brien"` stays `O'Brien`),
  hyphenated surnames (`"Smith-Jones"` is one token so they stay
  joined — this one is fine).
- Multi-word surnames are explicitly *not* handled (docstring at
  `client_utils.py:49-52`): `"Van Der Berg, Hans"` only round-trips
  if it's already in `Surname, First` form.

**Where it is and is not applied:**

`normalise_client_name` is consistently called on the **caller's side**
before `upload_to_sharepoint(client_name=...)`. Every plugin that
auto-files passes a normalised name (smart_responder line 277,
engagement_letter line 186, bas_reminder line 188, debtor_followup line
195, annual_review line 195, fusesign_monitor lines 134/202,
correspondence_logger lines 366/423, meeting_prep lines 122/198,
noa_processor line 431). The two API-route entry points
(`api_server.py:2452, 2569, 2674, 2713, 2753`) also normalise.

But `upload_to_sharepoint` itself does **no second-pass normalisation**
and **no existence check against existing folder names** at any stage.
And `SharePointIndexer._index_client_folder` at `sharepoint_indexer.py:159`
treats `client_folder["name"]` (the *verbatim* SharePoint folder name) as
the `client_name` metadata key it writes into the memory store — so the
indexer's view of "client Korkie, Gordon" is whatever the folder is
literally called, while plugin-driven memory writes use the normalised
form. That's an asymmetry inside MemoryStore, not directly the folder
duplication, but worth flagging.

**Most likely cause of duplication, pinned to code:** the asymmetry is
not between two code paths inside the same function — it is between the
**current normaliser's output** and **legacy folder names that were
created when the normaliser produced different output (or didn't
exist)**. Concrete failure modes already reachable today:

- An inbound email's `from_name` containing a middle initial
  (`"Gordon J. Korkie"`) produces `"Korkie, Gordon J."`, while the
  pre-existing folder is `"Korkie, Gordon"`. `upload_to_sharepoint`
  writes to a **new** sibling folder.
- An inbound email from `gordon.korkie@gmail.com` parses via
  `_name_from_email` to `"gordon korkie"` → `"Korkie, Gordon"`. That
  matches the canonical form but only because the local part is
  lowercase and dot-separated. `gj.korkie@gmail.com` would yield
  `"Korkie, Gj"` — separate folder.
- A folder created manually before this normaliser shipped (`"Gordon
  Korkie"` in raw form) won't match the function's output `"Korkie,
  Gordon"`. Auto-filing from any plugin therefore creates a parallel
  `"Korkie, Gordon"` folder, leaving the legacy one orphaned.

The audit script at `scripts/audit_sharepoint_duplicates.py` (B6)
groups all current folders under `Server/Clients/` by
`normalise_client_name(name)` so we can see how many of these
collisions actually exist today.

### B4. Morning brief structure

- File: `plugins/plugin_morning_briefing.py`.
- Build entry: `MorningBriefingPlugin.run(context)` at line 91.
- Body assembly: two private compilers — `_compile_reception_briefing`
  (line 201) and `_compile_accountant_briefing` (a few lines below).
  Each builds a python list of section strings (`fusesign_section`,
  `asic_section`, `debtor_section`, `noa_section`, etc.) and joins
  them, returning `(text, actionable_count)`. Each section is a
  pre-formatted plain-text block — bullet lines like
  `f"  • {client} — {name} (due: {due_str or 'no date'})"` (line 505).
- Output format: **plain text wrapped in a single `<pre>`** at line
  150-153:
  `f"<pre style='font-family:Calibri,sans-serif;font-size:14px'>{briefing_text}</pre>"`.
  Sent as the email `body_html`.
- Delivery: per recipient, either `context.graph.create_draft(email,
  subject, body_html)` (line 157) when `context.draft_mode` is on, or
  `context.graph.send_email(email, subject, body_html)` (line 159)
  otherwise. There is **no in-app render path** — the brief is purely
  email. It is also persisted to MemoryStore for later recall (line
  168-176) but that's an archive, not a UI surface.
- Composition is freeform-ish but stable: the compiler picks which
  sections to include based on `reception_mode` and `is_monday`,
  appends pre-built strings to a list, joins them. There is no
  template engine, no per-section approval slot.

**Critical for any "approve folder creation in the brief" idea:** the
brief is **read-only**. There is no existing inline-button / mailto-action
/ structured-reply mechanism. The only "interactivity" is
`_gather_pending_approvals` at line 466-484 which is a textual *summary*
of the in-app approval queue (`approval_queue.py`) — count + the first
five descriptions as bullet points (lines 475-481). The actual
approve/reject UI lives in the React app, not in the email. So adding
an approval mechanism to the brief itself is a meaningfully larger feature
than reusing what's there. If the eventual fix wants approval-style
gating, it should probably hook into `approval_queue.submit(...)` (used
e.g. by the engagement letter plugin at `plugin_engagement_letter.py:152`)
and surface it through the existing in-app approval queue, *not* invent
mailto-callback handling in the brief.

### B5. New-client detection trigger

There is **no upstream "this is a new client we haven't seen before"
gate** before folder creation. Folder creation is purely a downstream
side effect of `upload_to_sharepoint` — no code path asks "do we know
this client?" before that call.

What does exist is fragmented and partial:

- **Engagement letter trigger phrases** at
  `plugins/plugin_engagement_letter.py:40-44` (`"engagement letter",
  "letter of engagement", "new client", "onboard", "sign up", ...`) —
  these are matched against *email subject and bodyPreview* (line 87),
  not against any client-existence database. The plugin then does an
  XPM search at line 116 (`context.gateway.xpm.list_clients(search=
  client_name, limit=1)`), but only to enrich the engagement letter's
  context (entity_type, ABN). The XPM lookup result is *not* used to
  decide "create a SharePoint folder for this person".
- **EventBus topic `email.category.new_client`** declared at
  `event_wiring.py:79` (`Events.NEW_CLIENT_DETECTED`) and emitted
  inside the `on_triage_complete` handler at line 150. But the parent
  topic `email.triage.complete` (line 70) has **zero publishers** in
  the repo (greps for `emit.*triage`, `publish.*triage`,
  `emit.*EMAIL_TRIAGE`, `publish.*EMAIL_TRIAGE` all return no
  matches). The legacy `plugin_triage` was retired (see CLAUDE.md note
  about `plugin_smart_responder` "replaces the retired
  triage/ross/elio/reply plugins") and the smart responder doesn't
  publish triage events. So this NEW_CLIENT_DETECTED chain is
  effectively dormant — declared but unwired.
- **XPM nightly sync** does not exist in this codebase — no scheduled
  job pulls the XPM client list down and seeds SharePoint. The XPM
  client roster is consulted ad-hoc per email (engagement letter,
  meeting prep) but never enumerated to build a known-clients set.
- **Manual creation** via `POST /api/sharepoint/create-folder`
  (`api_server.py:2694`) is the only deliberate creation path. The
  current SharePoint picker UI exposes this for the user, but only as
  a sub-folder-under-an-existing-client operation (the `client_name`
  field in the request body must already correspond to an existing
  parent folder).

So the new-client trigger is *implicit*: an email arrives, smart
responder writes a draft, it auto-files the draft to
`Server/Clients/<normalised>/CoWorker Correspondence/...`, and Graph
creates the folder. Whether the sender is a real client, a vendor, a
spam reply, or a never-onboarded prospect is invisible to that path.

### B6. Audit — actual count of duplicate folders

Script committed at `scripts/audit_sharepoint_duplicates.py`. Behaviour:

- Reuses `GraphClient.get_sharepoint_site_id` and
  `get_sharepoint_drive_id` — no new auth path.
- Lists immediate child folders of `Server/Clients/` (the
  `SHAREPOINT_CLIENT_BASE` constant from `graph_client.py:99`),
  paging through `@odata.nextLink` until exhausted, with `$select` of
  `id,name,folder` to pull `folder.childCount` cheaply.
- Groups every folder by `normalise_client_name(folder.name)`.
- Renders a markdown report with: total folders, number of
  collision groups, and the top 20 groups ordered by group size
  (verbatim names + `childCount` per member).
- Output path: `docs/recon/sharepoint_duplicates_audit.md`.

The script header explicitly documents that it is one-off, has no
tests, and is not integrated into the running app. **Not run during
recon** per the task brief — Elio runs it manually on his machine
where the MSAL token cache lives, and the resulting report gets
committed in a separate follow-up.

---

## Summary observations (no fix proposals)

1. `mark_as_unread` is already implemented at
   `graph_client.py:405-409`; a `_reopen_original` wrapper at line
   598 exists but is dead code.
2. The duplicate-draft loop concern that motivated *not* marking
   unread (`graph_client.py:543-547`) is now defended by the
   `smart_responder_processed` SQLite table — that comment is stale
   relative to the current dedup strategy.
3. `upload_to_sharepoint` at `graph_client.py:1271` performs **no
   pre-upload existence check** against the client folder name and
   passes no `conflictBehavior` for the simple-upload PUT — the
   verbatim normalised string becomes a folder by Graph default.
4. There is **no folder-name sanitiser** distinct from
   `normalise_client_name`. The duplication mode pinned to code is
   asymmetry between the *current* normaliser output and **legacy
   folder names** (manually created or created by older code), not
   between two code paths in the live codebase.
5. The morning brief is a plain-text-wrapped-in-`<pre>` email
   delivered via `create_draft` / `send_email`. There is no inline
   approve/reject mechanism — adding folder-creation gating to the
   brief would be a new feature, not an extension of an existing
   pattern. The in-app approval queue (`approval_queue.py`) is the
   nearest existing primitive and is already used by the engagement
   letter plugin.
6. The `email.triage.complete` → `NEW_CLIENT_DETECTED` event chain
   wired in `event_wiring.py` has no publisher in the current
   codebase. New-client detection in any meaningful "do we know this
   sender?" sense does not exist before folder creation occurs.
