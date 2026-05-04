# MCS CoWorker — Dynamic Per-Accountant Signatures

**Design doc + Claude Code prompt** — paste at the root of `C:\Users\Elio\mcs-coworker`.

**Depends on:** `feat/aptos-draft-font` merged to `main` (v2.3 shipped). This change builds on `format_draft_body` and the existing signature insertion path in `graph_client.py`.

---

## Goal

Replace the current static-image signature with a dynamic, fully-HTML signature that:

- Pulls the right accountant's name, title, phone, and email from a central staff table — resolved automatically by matching the install's M365 signed-in user.
- Renders the firm logo and social icons inline (base64) so they show in Outlook desktop, Outlook web, and Outlook iOS without "show pictures" warnings.
- Has working clickable links for website, social platforms, and Google reviews.
- Has the firm-wide privacy/disclaimer block in proper formatted text (not an image).
- Is editable centrally — Elio updates one row in Settings, that staff member's signature changes everywhere.
- Falls back gracefully to the legacy image if anything goes wrong.

## Locked decisions

| Decision | Choice |
| --- | --- |
| Architecture | Central HTML template + per-person variables resolved at draft time |
| Per-person fields | name, title (nullable), email |
| Firm constants | logo, address, phone, website, social URLs, privacy text — editable in Settings |
| Image strategy | Base64-inlined PNG (logo + social icons); no remote hosting |
| Sign-in matching | Graph API `/me` → email → match to staff table row |
| HTML layout | Table-based (Outlook desktop is Word-rendered; tables are non-negotiable for email reliability) |
| Admin model | Centralised — Elio populates all 9 rows once via Settings; new hires added by Elio |
| Email display | Not shown in signature — only used as M365 matching key |
| LinkedIn | Out of v1, structurally easy to add later |
| Legacy image | Kept as fallback for at least 2 releases, then removed |

## Staff data (pre-seed)

| Name | Title | Email |
| --- | --- | --- |
| Elio Scarton | CPA, Tax Agent | elio@mcands.com.au |
| Vince Mercuri | *(blank)* | vince@mcands.com.au |
| Angelo Covelli | *(blank)* | angelo@mcands.com.au |
| Ross Mercuri | CA | ross@mcands.com.au |
| Eliza Lewis | Reception | reception@mcands.com.au |
| Brooke Austin | *(blank)* | brooke@mcands.com.au |
| Harry Gan | CPA, SMSF Auditor | harry@mcands.com.au |
| Lyn Karman | *(blank)* | lyn@mcands.com.au |
| Louise Boyd | *(blank)* | louise@mcands.com.au |

All staff share the firm phone `(03) 9794 0000` — stored as a settings key, not a per-row column.

## Firm constants (template defaults)

| Field | Value |
| --- | --- |
| Company | MC&S Pty Ltd |
| Phone | (03) 9794 0000 |
| Website (display) | mcands.com.au |
| Website (link) | https://www.mcands.com.au |
| Address line 1 | 23 Timor Circuit, Keysborough, Vic 3173 |
| Address line 2 | PO BOX 4440, Dandenong South, VIC, 3164 |
| Instagram | https://www.instagram.com/mcsaccounting |
| Facebook | https://www.facebook.com/mcandsaccounting |
| Google Review | https://www.google.com/maps/place//data=!4m3!3m2!1s0x6ad6140f023d542b:0x4be4c0f80d96d34b!12e1?source=g.page.m.ia._&laa=nmx-review-solicitation-ia2 |
| Privacy text | *(see below)* |

**Privacy text (verbatim):**

> This email and any attachments are confidential and may be subject to copyright, legal or some other professional privilege. They are intended solely for the attention and use of the named addressee(s). They may only be copied, distributed or disclosed with the consent of the copyright owner. If you have received this email by mistake or by breach of the confidentiality clause, please notify the sender immediately by return email and delete or destroy all copies of the email. Any confidentiality, privilege or copyright is not waived or lost because this email has been sent to you by mistake. Liability limited by a scheme approved under Professional Standards Legislation.

## Schema

### New table: `staff_signatures`

```sql
CREATE TABLE IF NOT EXISTS staff_signatures (
  id           INTEGER PRIMARY KEY AUTOINCREMENT,
  name         TEXT NOT NULL,
  title        TEXT,                          -- nullable; rendered only if non-empty
  email        TEXT NOT NULL UNIQUE,          -- M365 sign-in match key, lowercased
  enabled      INTEGER NOT NULL DEFAULT 1,    -- 0 = skip this row when matching
  created_at   REAL NOT NULL,
  updated_at   REAL NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_staff_signatures_email
  ON staff_signatures(LOWER(email));
```

If `config.py` already has a Staff or Notify table that holds names + emails, **investigate first** — extend it with the new columns rather than creating a parallel table. Pick whichever is cleaner. Either way, this is the only schema change.

### New settings keys (firm constants — editable)

Add to defaults dict in `init_db()` (idempotent via `INSERT OR IGNORE`):

| Key | Default |
| --- | --- |
| `signature_company` | MC&S Pty Ltd |
| `signature_phone` | (03) 9794 0000 |
| `signature_website_display` | mcands.com.au |
| `signature_website_url` | https://www.mcands.com.au |
| `signature_address_line1` | 23 Timor Circuit, Keysborough, Vic 3173 |
| `signature_address_line2` | PO BOX 4440, Dandenong South, VIC, 3164 |
| `signature_instagram_url` | https://www.instagram.com/mcsaccounting |
| `signature_facebook_url` | https://www.facebook.com/mcandsaccounting |
| `signature_linkedin_url` | *(empty by default)* |
| `signature_google_review_url` | *(the long URL above)* |
| `signature_privacy_text` | *(the privacy paragraph above)* |
| `signature_mode` | `dynamic` (`dynamic` / `legacy_image` / `disabled`) |

`signature_mode = legacy_image` keeps current behaviour; `dynamic` is new; `disabled` appends nothing. Default for new installs: `dynamic`.

## Image assets

Bundle these in the repo at `assets/signature/`:

| File | Source | Size |
| --- | --- | --- |
| `logo.png` | The 300×300 logo Elio supplied | ~11KB |
| `instagram.png` | 24×24 brand-coloured icon | ~1.5KB |
| `facebook.png` | 24×24 brand-coloured icon | ~1.5KB |

**Sourcing icons:** use simple-icons (https://simpleicons.org/) — MIT/CC0 — and rasterise the SVG to 24×24 PNG with the official brand colour as the background. Or generate two clean flat-colour PNG buttons in code; either approach is fine. Avoid Font Awesome (licensing) and avoid SVG inline (Outlook desktop on older Windows installs renders it inconsistently).

**Encoding:** read each file once at app startup, base64-encode, cache the data URI strings in memory (`signature_builder._image_cache`). Don't re-read on every draft.

**PyInstaller:** add `assets/signature/*.png` to `build.spec` `datas` so they're bundled in the frozen exe. Resolve at runtime via the standard `_MEIPASS` pattern already used elsewhere.

## New module: `signature_builder.py`

Public API:

```python
def build_signature_html(user_email: str | None) -> str:
    """
    Build the full HTML signature for the given M365 user email.

    - If user_email is None or doesn't match any enabled staff row,
      returns the legacy image signature (or empty string if mode=disabled).
    - If signature_mode != 'dynamic', returns legacy image / empty.
    - Cached image data URIs are loaded on first call.
    """
```

Caller is `graph_client._append_signature` (or whatever currently appends the static image). Replace the body of that function so it:

1. Fetches the M365 signed-in user email via Graph `/me` (cached per session — single call at startup is enough; doesn't change mid-session).
2. Calls `signature_builder.build_signature_html(email)`.
3. Returns the result, which the existing draft-assembly pipeline appends after the body.

The Aptos-wrapped body and the signature HTML are concatenated as today. The signature defines its own font styling so it renders correctly regardless of body font.

## HTML template structure

Table-based, inline styles only, no `<style>` blocks. Approximate structure (Claude Code can refine the exact markup for cross-client compatibility — what matters is the layout matches Elio's current signature):

```html
<table cellpadding="0" cellspacing="0" border="0" style="font-family: Aptos, Calibri, sans-serif; font-size: 10pt; color: #000000;">
  <tr>
    <!-- Logo column -->
    <td style="padding-right: 16px; vertical-align: top;">
      <img src="data:image/png;base64,{LOGO_B64}" width="80" height="80" alt="MC&S" style="display: block;" />
    </td>

    <!-- Details column -->
    <td style="vertical-align: top;">
      <div style="font-size: 11pt; font-weight: bold;">{NAME}</div>
      <!-- Title line — only rendered if title is non-empty -->
      {TITLE_LINE}
      <div>{COMPANY}</div>
      <div>{FIRM_PHONE}&nbsp;&nbsp;|&nbsp;&nbsp;<a href="{WEBSITE_URL}" style="color: #1A4A6E; text-decoration: none;">{WEBSITE_DISPLAY}</a></div>
      <div style="font-size: 9pt; color: #555555;">{ADDRESS_LINE1}</div>
      <div style="font-size: 9pt; color: #555555;">{ADDRESS_LINE2}</div>

      <!-- Social row -->
      <div style="padding-top: 6px;">
        <a href="{INSTAGRAM_URL}"><img src="data:image/png;base64,{IG_B64}" width="20" height="20" alt="Instagram" style="display: inline-block; margin-right: 4px;" /></a>
        <a href="{FACEBOOK_URL}"><img src="data:image/png;base64,{FB_B64}" width="20" height="20" alt="Facebook" style="display: inline-block; margin-right: 4px;" /></a>
        <!-- LinkedIn slot, only rendered if URL is set -->
        {LINKEDIN_ICON}
      </div>

      <!-- Google review CTA -->
      <div style="padding-top: 8px; font-size: 10pt;">
        Love what we do? <a href="{GOOGLE_REVIEW_URL}" style="color: #1A4A6E;">Leave us a Google review here</a>
      </div>
    </td>
  </tr>
</table>

<!-- Privacy block, full-width, smaller, italicised -->
<div style="margin-top: 12px; font-family: Aptos, Calibri, sans-serif; font-size: 8pt; color: #666666; font-style: italic; line-height: 1.4;">
  {PRIVACY_TEXT}
</div>
```

**Title rendering rule:** if `title` is null or empty, omit the title `<div>` entirely (no blank line). If present:

```html
<div style="font-size: 9pt; color: #444444;">{TITLE}</div>
```

**LinkedIn slot:** identical pattern — only render the icon if `signature_linkedin_url` is non-empty in settings.

## Sign-in matching

In `graph_client.py`, on first authenticated call (or at app startup right after MSAL token acquisition):

```python
# pseudocode
me = graph_client.get_me()  # GET /v1.0/me
session_user_email = me.get("mail") or me.get("userPrincipalName")
session_user_email = session_user_email.lower() if session_user_email else None
```

Cache `session_user_email` for the process lifetime. `signature_builder.build_signature_html` uses this to look up the staff row:

```sql
SELECT name, title, email
FROM staff_signatures
WHERE LOWER(email) = ? AND enabled = 1
LIMIT 1;
```

If no match → fall back to legacy image (or empty if `signature_mode = disabled`). Log a single warning to the activity log on the first miss per session: `"Signature: M365 user X not found in staff_signatures, using legacy image."` Don't spam the log on every draft.

## Settings UI

New "Email Signature" section in the Settings tab, below "Microsoft 365 / Signature". Three sub-sections:

### 1. Mode picker

Radio: Dynamic (recommended) / Legacy Image / Disabled. Persists to `signature_mode`.

### 2. Staff signatures table

Editable table. Columns: Name, Title, Email, Enabled (checkbox), Actions (Edit / Delete). Below: an "Add staff member" button. CRUD via new endpoints:

- `GET /api/staff-signatures` → list
- `POST /api/staff-signatures` → create
- `PUT /api/staff-signatures/<id>` → update
- `DELETE /api/staff-signatures/<id>` → soft-delete (`enabled = 0`) or hard-delete; pick soft-delete to preserve history

Whitelist these endpoints behind the existing local API token. Validate email format on save.

### 3. Firm constants

Editable text inputs for company name, phone, website (display + URL), address lines, social URLs (Instagram, Facebook, LinkedIn, Google Review), privacy text (multiline). Save persists to settings keys.

### 4. Live preview

Below all of the above: a "Preview your signature" panel that renders `build_signature_html(current_user_email)` in an iframe (or sanitised HTML container). Helps Elio verify what each accountant will see before drafting any test emails. Add a dropdown to preview as a different staff member (`build_signature_html(other_email)`).

### 5. Refresh

A "Refresh from M365" button that re-calls Graph `/me` and updates the cached `session_user_email`. Useful if someone re-signs in mid-session.

## Migration / rollout

1. **Schema migration** — `init_db()` creates `staff_signatures` table on first start (idempotent).
2. **Pre-seed** — on first start, if `staff_signatures` is empty AND `signature_mode` defaults to `dynamic`, insert the 9 rows above. Detect "first start of dynamic mode" via a one-shot setting `signature_seeded = 1`. Subsequent starts are no-ops.
3. **Asset bundling** — `assets/signature/*.png` added to repo and `build.spec`.
4. **Default mode for existing installs** — `dynamic` if not set; `legacy_image` is preserved if a user previously chose it. New installs default to `dynamic`.
5. **Smoke test order** — Elio's machine first; then 1–2 staff (Harry, Ross); then everyone reinstalls.

## Edge cases

| Scenario | Behaviour |
| --- | --- |
| M365 not signed in | Legacy image fallback, log once per session |
| M365 user not in `staff_signatures` | Legacy image fallback, log once per session |
| Staff row exists but `enabled = 0` | Legacy image fallback |
| Logo PNG missing from bundle | Skip logo `<img>`, render text-only signature, log error once |
| Social icon PNG missing | Skip that icon link, others still render |
| Title is null/empty | Omit title `<div>` entirely (no blank line) |
| LinkedIn URL not set | Omit LinkedIn icon entirely |
| Google review URL not set | Omit "Leave us a Google review" line entirely |
| `signature_mode = disabled` | Append nothing (just body, no signature) |
| Two installs, same Outlook account | Both get the same dynamic signature — no duplicate-signature problem |

## Tests

Add `tests/test_signature_builder.py` covering:

- Match by email returns correct name/title in HTML
- Title null → no title `<div>` in output
- Email not in table → legacy image returned
- Mode = `legacy_image` → bypasses dynamic path
- Mode = `disabled` → returns empty string
- LinkedIn URL set vs unset → icon present vs absent
- Privacy text is rendered with italic + small font styling
- Image data URIs load correctly from bundled assets
- Pre-seed inserts 9 rows on first run, no duplicates on second run

Plus 1–2 integration tests confirming `graph_client._append_signature` (or its replacement) calls `signature_builder.build_signature_html` with the correct email.

Run existing test suite — must still pass.

## Acceptance criteria

- A draft created by Smart Email Responder while signed in as Elio shows: MC&S logo, "Elio Scarton", "CPA, Tax Agent", "MC&S Pty Ltd", "(03) 9794 0000 | mcands.com.au" with website hyperlinked, both address lines, Instagram and Facebook icons (clickable, opening the right URLs), Google review link, full privacy text in italic 8pt grey.
- Same draft signed in as Vince shows the same template but with Vince's name, no title line (omitted entirely), and no other layout shift.
- Same draft signed in as Louise shows "Louise Boyd" with no title line.
- Updating the firm phone in Settings updates every staff member's signature on the next draft.
- Logo and icons render in Outlook desktop, Outlook web (no "show pictures" warning), and Outlook iOS without broken-image placeholders.
- Social icon clicks resolve to Instagram and Facebook firm pages.
- Google review link opens the review prompt.
- Mode toggle in Settings actually switches between dynamic / legacy / disabled without restart.
- Editing Elio's title in the Staff Signatures table and creating a new draft reflects the change immediately.
- An install with no M365 sign-in falls back to legacy image without errors.
- All existing tests pass; new signature tests pass.

## Out of scope

- LinkedIn icon (out of v1; structurally trivial to add when firm has a page).
- Per-machine `.htm` reading from `%APPDATA%\Microsoft\Signatures\` (we considered it; rejected in favour of central template).
- Mobile numbers (no field; can add later).
- Email line in signature display (intentionally omitted; matches current style).
- Signatures on emails sent directly from Outlook (Outlook handles those itself; this is plugin-draft-only).
- AI Chat specialist agents (not a draft path; unaffected).

## Deliverable

Branch `feat/dynamic-signatures` with:

- `signature_builder.py` (new module)
- `assets/signature/logo.png`, `instagram.png`, `facebook.png`
- `config.py` schema additions + setting defaults + pre-seed logic
- `graph_client.py` updated to use `signature_builder` + cache M365 session email
- `api_server.py` new `/api/staff-signatures` CRUD endpoints + whitelisted firm-constant settings keys
- `frontend/client/src/pages/Settings.tsx` "Email Signature" section (mode picker + staff table + firm constants + live preview)
- `build.spec` updated to bundle `assets/signature/*.png`
- `tests/test_signature_builder.py` covering the cases above
- A short PR description summarising the change

**Plan first** — list every file you'll touch and why before writing any code. Stop and ask before doing anything outside this scope.

---

## Post-implementation checklist (for Elio)

- [ ] Pull branch locally, run `build.bat`, install
- [ ] Open Settings → Email Signature → confirm 9 staff rows pre-seeded correctly
- [ ] Click "Preview as Elio" — verify logo, name, title, phone, address, social icons, Google review, privacy text all render
- [ ] Cycle preview through every staff member — confirm titles correctly show/hide for the four with no title (Vince, Angelo, Brooke, Lyn, Louise)
- [ ] Smoke test: draft email via Smart Email Responder → open in Outlook desktop → verify dynamic signature appears, links are clickable
- [ ] Same draft → open in Outlook web (outlook.office.com) → verify no "show pictures" warning, images render inline
- [ ] Same draft → forward to phone, open in Outlook iOS → verify rendering
- [ ] Click each social icon in a sent test email → confirm it goes to the right URL
- [ ] Click Google review link → confirm review prompt opens
- [ ] Edit your title in Settings, draft another email, confirm change applied without restart
- [ ] Toggle mode to "Legacy Image", confirm old image signature returns
- [ ] Toggle back to "Dynamic"
- [ ] Tag release v2.4 (assuming v2.3 is the Aptos release), zip dist, GitHub Release
- [ ] Stagger rollout: Elio → Harry & Ross → everyone else
- [ ] Tell each staff member, when they reinstall, to **disable Outlook's auto-signature on replies/forwards** (File → Options → Mail → Signatures → set Replies/forwards to "(none)") to prevent double signatures
