# MCS CoWorker — Aptos Draft Font Change

**Claude Code prompt** — paste into Claude Code at the root of `C:\Users\Elio\mcs-coworker`.

---

Change MCS CoWorker so that all plugin-created Outlook draft email bodies render in Aptos (with Calibri as fallback). Plan the changes first — list every file you'll touch and why — before writing any code.

## Goal

Every Outlook draft created by any plugin must render in Aptos 11pt with Calibri/sans-serif fallback. This applies to bodies created by Smart Email Responder, NOA Processor, BAS Reminder Drafter, Debtor Follow-Up, ASIC Annual Return Handler, Engagement Letter Generator, and any other current or future plugin that creates a draft via Graph API.

## Approach

Inline CSS only. Do NOT use a `<style>` block — many email clients strip them. Wrap each draft body in:

```html
<div style="font-family: Aptos, Calibri, sans-serif; font-size: 11pt; color: #000000;">
  ...body html...
</div>
```

The wrapping must happen in ONE place, not be repeated across plugins. Find where draft bodies are currently assembled — most likely in `graph_client.py` (the function that calls Graph's `POST /me/messages` or similar) or a helper used from `plugin_base.py` / `PluginContext`. If body assembly is currently scattered across plugins (each plugin building its own HTML and passing it straight to Graph), introduce a single helper:

```python
graph_client.format_draft_body(body_html: str) -> str
```

…that wraps the body and is called by every plugin draft path. Update plugins to use it. This is the only refactor allowed in this change — don't restructure anything else.

## Configurability

Make font family and size settings, not hard-coded constants. Add to SQLite settings table:

| Setting | Default |
| --- | --- |
| `draft_font_family` | `Aptos, Calibri, sans-serif` |
| `draft_font_size` | `11pt` |
| `draft_font_color` | `#000000` |

Read these values inside `format_draft_body()` so changes take effect on the next draft without restart.

Surface them in the Settings tab as three small inputs under a "Draft Formatting" section. Validate the font-size input loosely (must end in `pt` or `px` and start with a number) — if invalid, fall back to default and show a one-line error.

## What NOT to touch

- The signature handling — that's a separate upcoming change. Leave the existing static-image signature insertion exactly as it is. Just make sure that when `format_draft_body` wraps the body, the signature still appears below it correctly (the signature is appended outside the Aptos div, which is fine — it will render in whatever font the signature image/HTML specifies).
- AI Chat specialist agents — this is plugin-draft-only.
- Body content itself — don't reformat the AI-generated text, only wrap it.
- Any HTML escaping logic — don't change how plugins build their inner HTML.

## Edge cases to handle

1. If a plugin passes a body that's already wrapped in `<html>`/`<body>` tags (some templates may do this), don't double-wrap. Detect a leading `<html` or `<body` and wrap only the inner content, or wrap the whole thing in the div before `<html>` — pick whichever produces valid output in Outlook desktop preview.
2. If the body is plain text (no HTML), convert newlines to `<br>` first, then wrap. Detect plain text by absence of any `<` character.
3. Empty body — pass through untouched, no wrapping.

## Tests

Add unit tests in the existing test suite for `format_draft_body` covering:

- HTML body gets wrapped with correct inline style
- Plain text body gets newlines converted then wrapped
- Already-html-wrapped body doesn't double-wrap
- Empty body returns empty
- Custom font-family setting from SQLite is honoured
- Invalid font-size setting falls back to default

Run the existing test suite to confirm nothing else breaks.

## Acceptance criteria

- A draft created by Smart Email Responder, NOA Processor, BAS Reminder, and Debtor Follow-Up all render in Aptos 11pt when opened in Outlook desktop.
- Falls back cleanly to Calibri on a machine that somehow lacks Aptos.
- Changing `draft_font_family` in Settings to "Calibri" and creating a new draft produces Calibri output without restarting the app.
- All existing tests still pass; new `format_draft_body` tests pass.
- No changes to signature handling, AI body generation, or anything else.

## Deliverable

A commit on a feature branch (e.g. `feat/aptos-draft-font`) with:

- The `format_draft_body` helper
- Updated plugin draft paths calling it
- New SQLite settings + migration if you add columns
- Settings tab UI for the three inputs
- New tests
- A short PR description summarising the change

Stop and ask before doing anything outside this scope.

---

## Post-implementation checklist (for Elio)

- [ ] Pull branch locally, run `build.bat`, smoke test by drafting one email via Smart Email Responder
- [ ] Open the test draft in **Outlook desktop** — confirm Aptos renders
- [ ] Open the same draft in **Outlook web** — confirm Aptos renders (web is pedantic; if it works here, desktop is fine)
- [ ] Change `draft_font_family` to `Calibri` in Settings, draft another email, confirm font changed without restart
- [ ] Merge to `main`, tag a release (likely v2.3 if v2.2 is current — confirm before tagging)
- [ ] Zip the dist folder, publish GitHub Release
- [ ] Notify Harry, Ross, Eliza to reinstall

Once shipped and verified, move on to the dynamic signature change.
