# BirdsEye — Claude context

Tampermonkey userscript for Scentbird CRM support agents. Single IIFE in `BirdsEye.user.js` (~6270 lines, vanilla JS, no build step). Current version: **8.10.2**. Architecture overview in `DEVELOPER.md` (line ranges there are pre-Okta and slightly stale).

## Runtime

- `@match https://scentbird.kustomerapp.com/*` — main toolbar/panels
- `@match https://crm.scentbird.com/login/callback*` — Okta popup handler only
- Default `@run-at` (idle). **Do not** add `document-start` — it caused infinite mutation-observer loops on Kustomer's React render.

## Auth (Okta SSO)

- Full PKCE Authorization Code flow. Constants: `OKTA_CLIENT_ID`, `OKTA_AUTHORIZE`, `OKTA_TOKEN`, `OKTA_REDIRECT_URI`, `OKTA_SCOPE`.
- Login: user clicks "🔐 Sign in with Okta" toolbar button → `startOktaLogin()` (line 455) opens popup → popup hits `/login/callback` → `handleOktaCallback()` (line 77) exchanges code for tokens.
- **CRM validates the `id_token` as the bearer, not the `access_token`.** `getAccessToken()` (line 65) returns `id_token`. Don't "fix" this.
- Token bundle in `GM_setValue('birdseye_okta_tokens')` = `{access_token, refresh_token, id_token}`. Pending PKCE state in `birdseye_okta_pending`.
- The callback branch only intercepts when our pending state matches the URL state — CRM's own login flow is left untouched.
- Refresh: `refreshTokens()` (line 360) is a singleton via `_refreshInFlight` Promise — coalesces concurrent 401s. Resolves with the new `id_token`.
- Re-auth: `_reauthRequired` flag (line 361) + `refreshOktaButtonState()` (line 495) flips toolbar button to red only on hard refresh failure.
- Toolbar button is **hidden** when signed in and no captcha is pending (state `'in'`). Visible for states: `signin` (no token), `reauth` (refresh dead), `captcha` (`_captchaRequired`). Click routes by state: captcha → `openCaptchaPopup()`; otherwise → `startOktaLogin()`.

## Network

- **All CRM GraphQL goes through `crmRequest({method, url, headers, data, onload, onerror})`** (line 418) — wraps `GM_xmlhttpRequest`, auto-injects `Authorization: Bearer <id_token>`, refreshes-and-retries once on 401.
- Don't send raw `GM_xmlhttpRequest` to crm.scentbird.com; always use `crmRequest`.
- `gqlMutate(operationName, query, variables, callback)` (line 1325) is the high-level mutation helper.
- **403 = DataDome bot challenge.** `handle403()` (line 565) sets `_captchaRequired` and flips the toolbar button orange ("⚠ CRM Captcha"). Clicking calls `openCaptchaPopup()` (line 574) — popup-blocker-safe because it's a direct click handler. Popup waits 5s then polls `/graphql` every 3s for up to 30s; once status ≠ 403, it auto-closes and toolbar resets. Intermediate state: "⏳ Solving challenge…".
- **`crmRequest` spoofs same-origin CRM headers** on every call: `Origin: https://crm.scentbird.com`, `Referer: https://crm.scentbird.com/`, `Sec-Fetch-Site: same-origin`, `Sec-Fetch-Mode: cors`, `Sec-Fetch-Dest: empty`. DataDome fingerprints these against its cookie baseline. Without the full set, even valid cookies get gated. Don't drop these without a replacement same-origin transport.
- Okta popup **warms DataDome**: after token exchange, popup navigates to `crm.scentbird.com/` and opener force-closes it after 5s, so CRM's JS challenge runs once per sign-in.
- `@connect`: `crm.scentbird.com`, `www.scentbird.com`, `api.scentbird.com`, `scentbird.okta.com`.

## GraphQL queries

- `ORDERS_QUERY` (line 1106) — fetches orders with items, tracking, and product info including `status` (used for plaques).
- `USER_DETAILS_QUERY` (line 1150) — includes full shipping address (`street1 street2 city region postcode country`). The info bar's identity line depends on this. Don't strip the address subfields.
- `COMMENTS_QUERY` (line 1175) — `commentList { id created author { firstName lastName email } comment zendeskUrl }`. `author` requires subselection; `ownerId` was removed from the schema. Self-comment detection uses `author.email` vs `_myAuthorEmail`.
- `QUEUE_QUERY` (line 1296) — includes `tradingItem { id sku status productInfo { ... } }`. Status drives plaques.

## Settings

- `SETTINGS_KEY = 'birdseye_settings'` (line 197). Defaults: `{ fontScale: 1.15, modalScale: 1.0, theme: 'dark' }`.
- `applySettings(s)` (line 213) sets `--sb-fs-scale`, `--sb-modal-scale`, and `document.documentElement.dataset.theme`. Runs at boot.
- Settings panel accessible from toolbar. Theme toggle button cycles dark ↔ light.

## UI / styling

- **Font scaling**: `<style id="sb-style">` injected at start defines `--sb-fs-9` … `--sb-fs-14` as `calc(Npx * var(--sb-fs-scale))`. Default scale `1.15`. All `font-size:` values in the file use these vars — no raw `font-size:Npx` anywhere. To rescale globally, change `--sb-fs-scale`.
- **Light theme**: `:root[data-theme="light"]` overrides all `--sb-*` tokens. Foreground colors are darkened versions of the dark-theme accents. Applied via `document.documentElement.dataset.theme = 'light'`.

## Panel / modal factories (line ~710)

- `createOverlay(id)` (line 710) — full-screen backdrop (`rgba(0,0,0,0.6)`), click-outside-to-close.
- `createModal({id, title, width, noScale})` (line 731) — centered dialog inside an overlay. Scrollable, `max-height:85vh`.
- `createPanel({id, title, width, right, maxHeight, centered, draggable, noScale, onClose})` (line 771) — fixed floating panel. `all:initial` defeats Kustomer CSS bleed. `centered:true` → viewport center. `draggable:true` → header-drag with viewport clamping. `noScale:true` → ignores `MODAL_SCALE`.
- **Replacement and Custom Shipment forms** use `createPanel` with `centered:true, draggable:true` — *not* `createModal`. They do not darken the background.
- `removePanel(id)` (line 2017) — destructor; called automatically before re-creating a panel with the same id.

Known panel ids: `sb-order-panel`, `sb-charges-panel`, `sb-cancel-panel`, `sb-search-panel`, `sb-queue-panel`, `sb-edit-customer-panel`, `sb-product-search-panel`, `sb-replacement-form`, `sb-settings-panel`.

## Theme tokens

```
--sb-bg-panel: #1e1e2e    --sb-bg-deep: #0f172a    --sb-bg-row: #1a1a2e
--sb-text: #e2e8f0        --sb-text-muted: #94a3b8  --sb-text-dim: #64748b
--sb-accent: #a5b4fc      --sb-success: #86efac     --sb-warning: #fcd34d
--sb-danger: #fca5a5      --sb-info: #60a5fa
```

Close button is always a pill (`Close ✕`, `border-radius:20px`). Sticky headers use negative margins (`margin:-14px -14px 10px`) to fill edge-to-edge.

## Adjacent UI helpers (line ~857)

- `renderUserRow(user)` (line 857) — small "First Last · email" header row.
- `makeStatusLog()` (line 899) — appendable status log, returns el with `.log(msg, color)`.
- `makeCopyBtn(textOrFn, label)` (line 911) — clipboard copy with ✔ feedback.
- `makeLabel(text)` (line 926) — section label for form panels.
- `trackingStatusColor(status)` (line 873) — returns color string for last tracking line (bold + colored).
- `makeProductStatusPlaque(status)` (line 882) — returns a colored pill for non-LIVE order items. Used in Last Orders and Queue panels.
- `formatShippingAddress(shipping)` (line 865) — returns `"street · city, region, postcode"` one-liner. Cached at `cachedCustomerCtx._shippingAddress`.

## User info bar (`renderUserInfoBar`, line 5791)

- Row 1 (identity): `name · email · shipping address`
- Row 2 (tags): status / plan / credits / location / warning tags
- "🎟 N credits" tag appears in plan group when subscription has unspent NEW credits.
- Address comes from `USER_DETAILS_QUERY` shipping subfields — not from a separate call.

## Panel behavior notes

- Last Orders panel hides orders with `status === 'UPGRADED'` (superseded shells, noise).
- Queue panel shows `makeProductStatusPlaque` on each item driven by `tradingItem.status`.
- Replacements button is disabled on orders that are already replacements (to prevent sending to CRM for an unknown reason — behavior is untested and possibly a silent validation bug on their end).

## Editing gotchas

- The file mixes literal `\uXXXX` escapes and real Unicode (`—`, `🔄`, etc.). If Edit's old_string fails to match, use a smaller unique chunk that avoids the unicode region, or `cat -A` the line to see actual bytes.
- MutationObserver watches `document.body` subtree for `childList`. Anything that mutates `textContent` inside the observer fires an infinite loop. `refreshOktaButtonState` guards with `dataset.authState` — only writes to DOM when state changes.
- For bulk `font-size:Npx` replacements, use `sed -i` (~160 occurrences); Edit one-at-a-time is too slow.
- No automated tests. Smoke-test on a real Kustomer ticket after changes.

## Release

- Bump `@version` at the top of the file (semver).
- Commit message style: `<version> — <one-line summary>` (see `git log` for examples).
- Push to `main`. Tampermonkey auto-update pulls from GitHub raw.

## Open follow-ups

- Several callsites match `order.orderNumber` substring for auto-comment cleanup. That broke once when CRM changed copy; worth a more robust matcher if it breaks again.
- 403 captcha popup does not auto-retry the originally failing request — user must re-click. Could be fixed by stashing failed `crmRequest` opts and replaying after the probe succeeds.
