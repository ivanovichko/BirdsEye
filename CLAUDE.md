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

## Network

- **All CRM GraphQL goes through `crmRequest({method, url, headers, data, onload, onerror})`** — wraps `GM_xmlhttpRequest`, auto-injects `Authorization: Bearer <id_token>`, refreshes-and-retries once on 401.
- Don't send raw `GM_xmlhttpRequest` to crm.scentbird.com; always use `crmRequest`.
- `gqlMutate(operationName, query, variables, callback)` is the high-level mutation helper.
- 403 = CRM captcha (still legacy `handle403()` UX; intentionally left alone).
- `@connect`: `crm.scentbird.com`, `www.scentbird.com`, `api.scentbird.com`, `scentbird.okta.com`.

## Editing gotchas

- The file mixes literal `✘`-style escapes and real Unicode (e.g. `—`). Edit's old_string must match the exact source representation; if a match fails, `cat -A` the line to see what's there.
- Mutation observer watches `document.body` subtree for `childList`. Anything that mutates `textContent` from inside the observer callback creates an infinite loop. `refreshOktaButtonState` guards against this with `dataset.authState`.
- No automated tests. Smoke-test on a real Kustomer ticket.

## Release

- Bump `@version` at the top of the file (semver).
- Commit message style: `<version> — <one-line summary>` (look at `git log` for examples).
- Push to `main`. Tampermonkey auto-update pulls from GitHub raw.

## Open follow-ups

- Several callsites still match `order.orderNumber` substring for auto-comment cleanup; that broke once already when CRM changed copy. Worth a more robust matcher if it breaks again.
