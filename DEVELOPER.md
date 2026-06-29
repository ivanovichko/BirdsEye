# BirdsEye — Developer Documentation

## Overview

BirdsEye is a Tampermonkey userscript for Scentbird CRM support agents running on `https://scentbird.kustomerapp.com/*`. It injects a persistent toolbar and floating panels that let agents query and mutate CRM data without leaving the Kustomer ticket view.

---

## Architecture

Single IIFE, ~5300 lines, vanilla JS, no frameworks or build step. Everything lives in one file to simplify Tampermonkey deployment. The script is structured in top-to-bottom sections:

| Lines | Section |
|-------|---------|
| 20–76 | Constants & state |
| 127–198 | Utility helpers |
| 201–380 | UI factories |
| 384–432 | Toolbar |
| 435–480 | Email resolution / Kustomer API |
| 464–707 | GraphQL queries & mutations (string constants) |
| 709–835 | Data fetchers |
| 838–993 | Mutation helpers |
| 995–1097 | Customer context & resolution |
| 1056–1241 | Order rendering |
| 1244–1365 | Charge explanation & apply user |
| 1366–3242 | Panels (Search, Token, Orders, Charges, Queue, Edit Customer, Cancel Sub) |
| 3242–3530 | Refund Charges panel |
| 3528–3687 | Refund confirm dialog |
| 3689–4399 | Replacement form |
| 4401–4458 | Fill Name button |
| 4460–4833 | Variable fill & slash autocomplete |
| 4865–4930 | Button registration |
| 4932–5103 | User info bar |
| 5105–5230 | Subscription bar fetcher & refresh logic |
| 5232–5320 | Mutation observer & initialization |

---

## API Endpoints

| Endpoint | Auth | Used for |
|----------|------|----------|
| `POST https://crm.scentbird.com/graphql` | `Authorization: Bearer <token>` | All CRM queries and mutations |
| `GET https://scentbird.api.kustomerapp.com/v1/customers/{id}` | Cookie (session) | Email + phone resolution per ticket |
| `POST https://api.scentbird.com/graphql` | CSRF token via hidden popup | Queue delete, address update |

---

## State Variables

```js
let BEARER_TOKEN           // CRM Bearer token — persisted in localStorage('sb_crm_token')
let tokenValid             // Boolean — set true after validateToken() succeeds
let cachedCustomerCtx      // { email, user, _kustomerId, _fraudStatus, _gender, _gwebLink, _shippingAddrId }
let userInfoLastEmail      // Last resolved email — prevents repeated API calls per ticket
let _userInfoTimer         // Debounce handle for refreshUserInfoBar()
let _refundCommentsPosted  // Set<string> — deduplicates refund comments within a session
let _totalRefundedCents    // Running refund total (cents) this session
let _totalRefundedCurrency // Currency of refunded charges
let _myOwnerId             // CRM ownerId discovered after first comment post
let toolbarEl              // Reference to injected toolbar element
let fillNameRegistered     // Guard — prevents duplicate Fill Name button registration
let _customerApiReady      // Boolean — customer API (api.scentbird.com) authenticated
let _customerCsrfToken     // CSRF token for customer API
```

### Context lifecycle

`cachedCustomerCtx` is populated by `loadCustomer()` or `refreshUserInfoBar()` when a Kustomer ticket loads. `clearCustomerCtx()` resets all customer-scoped state when the toolbar is removed from the DOM (agent navigated away). The mutation observer detects toolbar removal.

---

## Token Management

The Bearer token is stored in `localStorage('sb_crm_token')` and loaded on script init.

- **`validateToken()`** — fires once per toolbar session (`tokenValidatedForSession` guard). Sends a lightweight `userSearch` query; marks `tokenValid = true` on success.
- **`handle401()`** — called on any 401 response. Clears `BEARER_TOKEN`, removes localStorage entry, sets `tokenValid = false`, updates toolbar button to `⚠ Token Expired`.
- **`handle403()`** — called on 403 (CRM Cloudflare captcha). Updates toolbar button to `⚠ CRM Captcha`. Token is NOT cleared.
- Token Panel (`sb-token-panel`) — manual paste/save UI. Triggers `validateToken()` after save.

---

## GraphQL Queries

### `CRM_QUERY` — `userSearch`
Searches CRM by free-text filter. Returns up to 50 results.
```
fields: id, email, firstName, lastName, origin
        userAddress.shipping { street1, city, region, postcode, country }
        subscription { id, status }
```

### `ORDERS_QUERY` — `ordersByUserId`
Last 4 orders for a user.
```
fields: id, type, status, year, month, orderNumber, created, subscriptionId
        initialShippingAddress { street1, street2, city, region, postcode, country }
        tracking { carrier, number, items { status, date, location } }
        warehouseOrder.data { status }
        orderItems { id, starterSet, product { id, productInfo { name, brand } } }
```

### `CHARGES_QUERY` — `chargesByUserId`
All charges for a user.
```
fields: date, kind
        charge { id, paymentDate, category, status, totalPrice, refundAmount }
        charge.cashbirdDetails {
          invoice.lineItems { uuid, description, productCode, price, total, tax, discount, type, currency, refundedAmount }
          credits           { id, status, reason, price, amount, tax, total, currency }
          shippingCredits   { id, type, status, price, amount, tax, total, currency }
          paymentMethod     { methodName, type }
        }
        charge.orders { id, status, warehouseOrder.data.status }
```

`fetchCharges()` post-processes the raw list:
1. Finds the most recent `SUCCESS` charge → `refDate`
2. Keeps only charges within ±10 days of `refDate`
3. Returns `{ success, all, failed }` — `all` = SUCCESS entries, `failed` = FAIL entries

### `USER_DETAILS_QUERY` — `userDetailsById`
Extended user info fetched after initial CRM search.
```
fields: fraudInfo.status, gwebLink, gender
        userAddress.shipping.id
        subscriptionList { id, status, subscribed, nextBillingDate, planName, subscriptionDate, subscriptionEndDate
          cashbirdDetails.data { billingDay, nextBillingDate, isAwaitCancellation }
          credits { id, type, status, reason }
          coupons { id, code, type, amount, status, applyToNextBillingDates }
          addOnSettings { candleSubscription, caseSubscription, carScentSubscription,
                          samplesSubscription, homeDiffuserSubscription }
        }
```

### `COMMENTS_QUERY` — `userComments`
```
fields: id, created, author, comment, zendeskUrl, ownerId
```

### `PRODUCT_SEARCH_QUERY` — `productSuggestion`
Used in Replacement and Custom Shipment product search panels.
```
fields: id, sku, status, section, price, currency, volume, volumeUnit
        productInfo { id, name, brand, upchargePrice, productType }
```

### `QUEUE_QUERY` — `subscriptionQueue`
```
fields: blockList { yearMonth, status, processedSubscriptionOrder { id, status }
          productList { tradingItem { id, sku, productInfo { id, name, brand, upchargePrice, productType } }
                        source, subscriptionQueueItemId, upchargePrice }
        }
```

---

## GraphQL Mutations

All mutations go through `gqlMutate(operationName, query, variables, callback)` which handles 401/403, parses errors, and triggers `refreshSubscriptionBar()` after success (except comment operations).

| Mutation | Trigger | Key inputs |
|----------|---------|-----------|
| `createUserComment` | Post-action comment | `userId`, `comment`, `zendeskUrl` |
| `deleteUserComment` | Auto-comment removal | `id` |
| `replacementSave` | Replacement form confirm | `zendesk`, `reason`, `bom`, `customerId`, `addressConfirmed`, `picturesProvided`, `orderNumber`, `sourceOrder.orderItems` |
| `customShipmentSave` | Custom Shipment confirm | `zendesk`, `reason`, `bom`, `customerId`, `sourceOrder { year, month, subscriptionId, orderItems }` |
| `refundCredits` | Refund flow — subscription credits | `userId`, `credits[]`, `shippingCredits[]`, `forceRefundUsedCredits: true`, `forceRefundUsedShippingCredits: true`, `reason: 'CUSTOMER_CHANGED_THEIR_MIND'` |
| `refundInvoiceItems` | Refund flow — upcharge line items | `userId`, `items[{ lineItemUuid, amount, forceRefundUsedCredits: false }]`, `reason: 'CUSTOMER_CHANGED_THEIR_MIND'` |
| `cancelSubscription` | Cancel Sub panel | `subscriptionId` |
| `deleteAllPaymentMethods` | Cancel Sub panel | `userId` |
| `userChangeEmail` | Edit Customer panel | `userId`, `email` |
| `userSetFraudStatus` | Cancel Sub panel fraud toggle | `userId`, `status` |
| `addSupportCredits` | Custom Shipment (credit top-up) | `subscriptionId`, `credits`, `comment` |
| `charge` | Replacement upcharge billing | `userId`, `subscriptionId`, `amount`, `currency`, `description` |
| `cancelOrder` | Refund flow — order cancel before refund | `id`, `subscriptionHold: true`, `unlockItems: false`, `refundReason: 'CUSTOMER_CHANGED_THEIR_MIND'` |

---

## Comment System

`mutPostComment(userId, comment, callback)` is the standard comment writer. It:
1. Fetches recent comments for this Kustomer URL
2. If our own recent comment exists (within 3 days, matching `zendeskUrl`), merges new text into it by deleting and re-creating
3. Discovers `_myOwnerId` from the first successful comment to identify future "ours" comments

Auto-generated CRM comments (from mutations like `replacementSave` / `customShipmentSave`) are deleted before our custom comment is posted:
- Replacement: single `fetchComments` call — auto-comment references the pre-existing source order number (already present)
- Custom Shipment: polling loop up to 4 attempts × 1.5 s — auto-comment references the newly minted order number which may not exist yet

---

## Customer Resolution Flow

```
Kustomer cookie API → GET /v1/customers/{kustomerId}
  → email (primary) + alternate emails + phones

CRM GraphQL → userSearch(filter: email)
  → filter origin === 'SCENTBIRD'
  → prefer exact email match
  → cache in cachedCustomerCtx
```

`resolveEmail()` — returns primary email only (string or null)
`resolveKustomerAttributes()` — returns `{ emails[], phones[] }` for all Kustomer contact fields

**Social Network tickets:** Kustomer returns `{"error": "No customer email found"}`. `resolveEmail()` returns null. `refreshUserInfoBar()` caches a sentinel `__no_email:{kustomerId}` in `userInfoLastEmail` to prevent repeated fetches on subsequent mutation observer firings.

---

## User Info Bar

Rendered by `renderUserInfoBar(sub, addressMatch, orderAddrStr)`. Displayed below the toolbar after customer resolution. Rebuilt on each `refreshSubscriptionBar()` call.

**Tag groups (left to right):**

1. **Status** — `Active since DD MON 'YY` / `Cancelled on DD MON 'YY` + `Scheduled Cancellation` if awaiting
2. **Billing day** — `BD: Nth`
3. **Plan** — plan label + add-ons (candle, case, car scent, samples, diffuser)
4. **Gender** — appended async from `USER_DETAILS_QUERY`
5. **Location** — Canada / UK / Hawaii / Alaska / Puerto Rico (from `userAddress.shipping`)
6. **Warnings** — `⚠ Addr Updated` / `⚠ Fraud` / `⚠ Chargeback` (appended async)

**Address comparison (`_compareAddresses`):**
Compares profile `userAddress.shipping` (raw user-entered) against `order.initialShippingAddress` (USPS-cleansed). Concatenates `street1 + street2` on both sides before normalizing — handles USPS folding unit numbers into `street1`. Returns `null` if either side is missing.

**Chargeback tag:** guarded by `data-tag="chargeback"` attribute to prevent duplicate tags on repeated renders.

---

## Panels Reference

| Panel ID | Function | Opens from |
|----------|----------|-----------|
| `sb-search-panel` | CRM user search + Find Other Accounts | 🔍 CRM Search button |
| `sb-token-panel` | Bearer token paste/save | 🔑 Token button |
| `sb-order-panel` | Last 4 orders with replace/cancel/hold | 📦 Last Order button |
| `sb-charges-panel` | Recent charges with refund entry | 💳 Recent Charges button |
| `sb-queue-panel` | Subscription queue with delete | 📋 Queue button |
| `sb-edit-customer-panel` | Email, address update | ✎ Edit Customer button |
| `sb-cancel-panel` | Cancel sub, delete payments, delete account, fraud toggle | ❌ Cancel Sub button |
| `sb-refund-charges-panel` | Charges list with refund buttons | 💰 Refund in order panel |

---

## Refund Flow (detailed)

```
showRefundChargesPanel(user)
  └─ fetchCharges(userId)
       └─ filter to ±10 days of last SUCCESS charge
       └─ per charge:
            canRefund = (credits.length > 0 || shippingCredits.length > 0 || lineItems.length > 0)
                        && !NON_REFUNDABLE_STATUSES.includes(orderStatus)

NON_REFUNDABLE_STATUSES = ["SHIPPED", "DELIVERED", "PRINTED", "DONE", "PROCESSED"]

showRefundConfirmDialog(user, charge, credits, shippingCredits, currency, totalCents, refBtn, block, orderId)
  └─ onConfirm:
       1. if orderId → mutCancelOrder(orderId)          [refundReason: 'CUSTOMER_CHANGED_THEIR_MIND']
       2. doRefund():
            hasCreditsFlag = credits.length > 0 || shippingCredits.length > 0
            unrefundedItems = lineItems.filter(total > 0 && refundedAmount === null && type !== 'SHIPPING_PURCHASE')
                                       .map({ lineItemUuid, amount, forceRefundUsedCredits: false })

            both  → mutRefundCredits → mutRefundInvoiceItems → onResult
            credits only → mutRefundCredits → onResult
            items only   → mutRefundInvoiceItems → onResult
            neither      → onResult directly

       3. onResult → mutPostComment("Cancelled & Refunded {Month Year} order")
                     deduped by _refundCommentsPosted Set
```

---

## Find Other Accounts

Located in `showSearchPanel()`. Searches CRM for accounts linked to the current customer by alternate contact info.

**Search order & tiers:**

| Tier | Source | Terms |
|------|--------|-------|
| 1 | Kustomer API | Alternate emails (excluding primary) + phone numbers |
| 2 | CRM cache | `userAddress.shipping.street1` |
| 3 | CRM cache | `firstName + ' ' + lastName` |

Searches run sequentially (no parallel) to avoid rate-limiting. Results collected into a `Map<userId, { user, tier }>` — first match wins the tier. Current customer excluded. Final list sorted by tier ascending before `renderUsers()`.

---

## Variable Fill

`fillVariables(btn)` scans the Kustomer composer for `[TOKEN]` placeholders, resolves each against a selected order and customer data, replaces in-place, and copies the result to clipboard.

**Available tokens:**

| Token | Source |
|-------|--------|
| `[ADDRESS]` | `userAddress.shipping` |
| `[DATE]` | Last tracking event date |
| `[FRAGRANCE]` | Order product names |
| `[TRACKING]` | Tracking number |
| `[CARRIER]` | Carrier name |
| `[ORDER_NUMBER]` | Order number |
| `[MONTH]` | Order month/year |
| `[NAME]` | Customer first name |
| `[NEXT_BILLING]` | Next billing date |
| `[BILLING_DAY]` | Billing day of month |

Slash autocomplete (`/`) shows available tokens in a picker dropdown while composing.

---

## Fill Name Button

`registerFillName()` adds the **Fill Name** toolbar button. On click:

1. Opens Kustomer "Edit Customer Profile" modal
2. Resolves customer from CRM (or uses cache)
3. Priority:
   - CRM name found → set to proper case → submit
   - CRM name not found but field has content → title-case existing value (`✔ Cased`)
   - No CRM user AND no token AND field empty → dismiss modal, show `⚠ Token Expired`
   - Otherwise → dismiss modal, show `✘ Not Found`

---

## Customer API (api.scentbird.com)

Used for queue delete and address update. Requires a separate CSRF token obtained via:
1. `ensureCustomerAuth(callback)` — opens `gwebLink` in a 1×1 popup to establish session
2. `GM_xmlhttpRequest` to `api.scentbird.com/graphql?opname=queueProbe` — extracts `x-csrf-token` from response headers
3. Token stored in `_customerCsrfToken`, session flag in `_customerApiReady`

---

## UI Factory Functions

| Function | Returns | Purpose |
|----------|---------|---------|
| `createOverlay(id)` | `div` | Full-screen backdrop, closes on backdrop click |
| `createModal({ id, title, width })` | `{ overlay, box }` | Centered dialog |
| `createPanel({ id, title, width, right, maxHeight, onClose })` | `div` | Fixed floating panel |
| `makeTag(label, color, bg, border)` | `span` | Info bar / warning tag |
| `makeSelect(options, selected)` | `select` | Styled dropdown |
| `makeLabel(text)` | `div` | Form section label |
| `makeStatusEl()` | `div` | Status message footer |
| `makeDialogButtons({ confirmLabel, confirmColor, onCancel, onConfirm })` | `div` | Cancel + Confirm button pair |
| `makeStatusLog()` | `div` | Multi-line sequential status log |
| `makeCopyBtn(textOrFn, label)` | `button` | Clipboard copy button |
| `renderUserRow(user)` | `div` | Name · email identity row |

---

## Utility Functions

| Function | Purpose |
|----------|---------|
| `setReactValue(input, value)` | Sets React-controlled input value and fires synthetic events |
| `toProperCase(str)` | Title-cases a string |
| `fmtMoney(cents, currency)` | `$X.XX` / `£X.XX` |
| `fmtDateTag(isoStr)` | `15 MAR '24` format for info bar |
| `fmtDateLong(isoStr)` | `March 30th` format for variable fill |
| `detectCurrency(lineItems)` | Reads currency from first line item with one set |
| `chargeDateLabel(isoStr)` | `Mar 15, 2025` |
| `waitForField(selector, timeout)` | Promise that resolves when a DOM element appears |
| `_normalizeAddr(s)` | Lowercase, strip punctuation, collapse whitespace |
| `_compareAddresses(user, order)` | Compares profile vs order shipping address; returns `null` if data missing |

---

## Initialization

```
MutationObserver (childList, subtree on document.body)
  → ensureToolbar()         inject toolbar if missing
  → registerFillName()      once per toolbar lifetime
  → registerCrmButtons()    once per toolbar lifetime
  → refreshUserInfoBar()    debounced 800ms — resolves customer and renders info bar
  → validateToken()         once per session if BEARER_TOKEN present
```

Toolbar is injected inside the Kustomer header element (`.timeline-header` or equivalent). When the header is removed by Kustomer's SPA navigation, `toolbarEl` is detected as detached and `clearCustomerCtx()` fires.

---

## Reason Constants

**Replacement reasons** (`REPLACEMENT_REASONS`):
`DAMAGED_LEAKING`, `DAMAGED_VIAL`, `INCOMPLETE`, `WRONG_ITEM`, `NOT_RECEIVED`, `DELIVERY_DELAY`, `RETURNED_SHIPMENT`, `ADDRESS_CHANGE`, `COURTESY_DISLIKE`, `COURTESY_FORGET_UPDATE_QUEUE`, `COURTESY_GENDER_ISSUE`, `OTHER`

**Custom Shipment reasons** (`CS_REASONS`):
`SKIPPED_SHIPMENT`, `DELAY_ISSUE`, `DAMAGED_LEAKING`, `DAMAGED_VIAL`, `WRONG_ITEM`, `NOT_RECEIVED`, `OTHER`

**Cancel order reasons** (`CANCEL_ORDER_REASONS`):
Used in the cancel order dialog in the orders panel.

**Refund reason** (hardcoded): `CUSTOMER_CHANGED_THEIR_MIND` — used in all refund and cancel-order mutations.


**Non-refundable order statuses**: `SHIPPED`, `DELIVERED`, `PRINTED`, `DONE`, `PROCESSED`
