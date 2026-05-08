# BirdsEye — Claude context

Tampermonkey userscript for Scentbird CRM support agents. Single IIFE in `BirdsEye.user.js` (~5700 lines, vanilla JS, no build step). For full architecture see `DEVELOPER.md` (line ranges in it are pre-Okta and slightly stale).

## Runtime

- `@match https://scentbird.kustomerapp.com/*` — main toolbar/panels
- `@match https://crm.scentbird.com/login/callback*` — Okta popup handler only
- Default `@run-at` (idle). **Do not** add `document-start` — it caused infinite mutation-observer loops on Kustomer's React render.
- Single companion script: `BirdsEye - Token Helper.user.js` (legacy; not loaded into the main script).

## Auth (Okta SSO)

- Full PKCE Authorization Code flow. Constants: `OKTA_CLIENT_ID`, `OKTA_AUTHORIZE`, `OKTA_TOKEN`, `OKTA_REDIRECT_URI`, `OKTA_SCOPE`.
- Login: user clicks "🔐 Sign in with Okta" toolbar button → `startOktaLogin()` opens popup → popup hits `/login/callback` → `handleOktaCallback()` exchanges code for tokens.
- **CRM validates the `id_token` as the bearer, not the `access_token`.** `getAccessToken()` returns `id_token`. Don't "fix" this.
- Token bundle in `GM_setValue('birdseye_okta_tokens')` = `{access_token, refresh_token, id_token}`. Pending PKCE in `birdseye_okta_pending`.
- The callback branch only intercepts when our pending state matches the URL state — so CRM's own login flow is left alone.
- Refresh: `refreshTokens()` is a singleton via `_refreshInFlight` Promise (coalesces concurrent 401s).
- Re-auth UI: `_reauthRequired` flag + `refreshOktaButtonState()` flips toolbar button to red "🔐 Sign in with Okta" only on hard refresh failure.
- Toolbar button is **hidden** when signed in and no captcha is needed (state `'in'`). Visible only when `signin` (no token), `reauth` (refresh dead), or `captcha` (`_captchaRequired`). Click action routes by state: captcha → open captcha popup; otherwise → `startOktaLogin()`.

## Network

- **All CRM GraphQL goes through `crmRequest({method, url, headers, data, onload, onerror})`** — wraps `GM_xmlhttpRequest`, auto-injects `Authorization: Bearer <id_token>`, refreshes-and-retries once on 401.
- Don't send raw `GM_xmlhttpRequest` to crm.scentbird.com; always use `crmRequest`.
- `gqlMutate(operationName, query, variables, callback)` is the high-level mutation helper.
- 403 = CRM captcha (DataDome). `handle403()` flips the toolbar button orange ("⚠ CRM Captcha") and sets `_captchaRequired`. Clicking the button calls `openCaptchaPopup()` (popup-blocker-safe because it's a click). The popup waits 5s before its first probe, then polls `/graphql` every 3s for up to 30s; once status ≠ 403, it auto-closes and the toolbar resets. Toolbar surfaces a "⏳ Solving challenge…" intermediate state while the probe is running.
- **`crmRequest` spoofs same-origin CRM headers** (`Origin`, `Referer`, `Sec-Fetch-Site: same-origin`, `Sec-Fetch-Mode: cors`, `Sec-Fetch-Dest: empty`) on every CRM call. DataDome cross-checks the request shape against its cookie's fingerprint baseline; without the full set, even valid cookies get gated. This is what stopped the persistent 403 issue on 2026-05-08. Don't drop these headers without a replacement same-origin transport.
- Okta sign-in popup also **warms DataDome**: after the token exchange succeeds the popup navigates to `crm.scentbird.com/` and the opener force-closes it after 5s, ensuring CRM's JS challenge runs once per login.
- `USER_DETAILS_QUERY` returns the full shipping address (`street1 street2 city region postcode country`), not just the id — the info bar's identity line depends on this.
- `@connect`: `crm.scentbird.com`, `www.scentbird.com`, `api.scentbird.com`, `scentbird.okta.com`.

## UI / styling

- **Font scaling**: a single `<style id="sb-style">` is injected at script start defining `--sb-fs-9` … `--sb-fs-14` as `calc(Npx * var(--sb-fs-scale))`. Default scale is `1.15`. All inline `font-size:` values use these vars — no raw `font-size:Npx` anywhere. To rescale globally, change one `--sb-fs-scale` value.
- **Modal/panel factories** at ~line 510:
  - `createOverlay(id)` — full-screen backdrop, click-outside-to-close.
  - `createModal({id, title, width})` — centered dialog inside an overlay.
  - `createPanel({id, title, width, right, maxHeight, centered, draggable, onClose})` — fixed floating panel. `centered:true` puts it in the viewport center; `draggable:true` adds header-drag with viewport clamping. `all:initial` defeats Kustomer CSS bleed.
- **Replacement and Custom Shipment forms** use `createPanel` with `centered:true, draggable:true` — *not* `createModal`. They don't darken the page. Drag from the title bar.
- Theme tokens: `#1e1e2e` panel bg, `#0f172a` deepest bg, `#e2e8f0` text, `#94a3b8` muted, `#6ee7b7` success, `#fca5a5` error, `#f59e0b` warning, `#a5b4fc` accent.
- Close button is always a pill (`Close ✕`, `border-radius:20px`).
- Status helpers near the formatters: `trackingStatusColor(status)` (used on the most-recent tracking line — bolded + color-coded) and `makeProductStatusPlaque(status)` (returns a pill for any non-LIVE order item, colored amber/red/gray).
- `formatShippingAddress(shipping)` returns one-liner; cached at `cachedCustomerCtx._shippingAddress` by `_fetchAndRenderSubBar`.
- User info bar (`renderUserInfoBar`) renders a top identity line (`name · email · shipping address`) above the tag groups, then a second row with status/plan/credits/location/warning tags. The "🎟 N credits" tag appears in the plan group when subscription has NEW credits.
- Last Orders panel hides orders with `status === 'UPGRADED'` (superseded shells; just noise).
- Queue panel uses `makeProductStatusPlaque` on each item (driven by `tradingItem.status` from QUEUE_QUERY).

## Editing gotchas

- The file mixes literal `\uXXXX` escapes and real Unicode (e.g. `—`, `🔄`). Edit's old_string must match the exact source representation; if a match fails, fall back to a smaller chunk that avoids the unicode region, or `cat -A` the line to see actual bytes.
- Mutation observer watches `document.body` subtree for `childList`. Anything that mutates `textContent` from inside the observer callback creates an infinite loop. `refreshOktaButtonState` guards against this with `dataset.authState`.
- For bulk `font-size:Npx` replacements use `sed -i` since there are ~160 across the file; the harness Edit tool one-at-a-time is too slow.
- No automated tests. Smoke-test on a real Kustomer ticket.

## Release

- Bump `@version` at the top of the file (semver).
- Commit message style: `<version> — <one-line summary>` (look at `git log` for examples).
- Push to `main`. Tampermonkey auto-update pulls from GitHub raw.

## Open follow-ups

- Several callsites still match `order.orderNumber` substring for auto-comment cleanup; that broke once already when CRM changed copy. Worth a more robust matcher if it breaks again.
- 403 captcha popup does not auto-retry the originally failing request — the user has to re-click. Could be wired up by stashing the failed `crmRequest` opts and replaying after the probe succeeds.
