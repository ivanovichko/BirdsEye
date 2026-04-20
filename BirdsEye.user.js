// ==UserScript==
// @name         BirdsEye
// @namespace    scentbird-kustomer
// @version      8.7
// @description  Unified toolbar: Fill Name + CRM Search + Last Orders + Recent Charges
// @author       You
// @match        https://scentbird.kustomerapp.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setClipboard
// @connect      crm.scentbird.com
// @connect      www.scentbird.com
// @connect      api.scentbird.com
// @updateURL    https://raw.githubusercontent.com/ivanovichko/BirdsEye/main/BirdsEye.user.js
// @downloadURL  https://raw.githubusercontent.com/ivanovichko/BirdsEye/main/BirdsEye.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ══════════════════════════════════════════════════════════════════════════
  // CONSTANTS & STATE
  // ══════════════════════════════════════════════════════════════════════════

  const GRAPHQL_URL   = 'https://crm.scentbird.com/graphql';

  let toolbarEl           = null;
  let fillNameRegistered  = false;
  let stashedSelection    = '';    // captured on mousedown before Edge clears it

  // ── Customer context cache ────────────────────────────────────────────────
  // Populated once per customer page by loadCustomer or refreshUserInfoBar.
  // Cleared when toolbar is removed from DOM (agent navigated away).

  let cachedCustomerCtx   = null;   // { email, user }
  let userInfoLastEmail   = null;
  let _userInfoTimer      = null;
  let _refundCommentsPosted = new Set(); // tracks posted refund comments to avoid duplicates
  let _totalRefundedCents = 0;           // running total of refunds in this session
  let _totalRefundedCurrency = 'USD';    // currency of refunded charges
  let _myOwnerId = null;                 // our CRM comment ownerId, discovered on first comment create

  function clearCustomerCtx() {
    cachedCustomerCtx = null;
    userInfoLastEmail = null;
    _customerApiReady = false;
    _customerCsrfToken = null;
    _refundCommentsPosted.clear();
    _totalRefundedCents = 0;
    _totalRefundedCurrency = 'USD';
  }

  // ── Token storage ─────────────────────────────────────────────────────────

  let BEARER_TOKEN = localStorage.getItem('sb_crm_token') || '';

  function handle401() {
    BEARER_TOKEN = '';
    localStorage.removeItem('sb_crm_token');
    tokenValid = false;
    const tokenBtn = document.getElementById('sb-crm-token-btn');
    if (tokenBtn) {
      tokenBtn.textContent = '⚠ Token Expired';
      tokenBtn.style.color = '#fca5a5';
    }
  }

  function handle403() {
    const tokenBtn = document.getElementById('sb-crm-token-btn');
    if (tokenBtn) {
      tokenBtn.textContent = '⚠ CRM Captcha';
      tokenBtn.style.color = '#f59e0b';
    }
  }

  let tokenValid = false;

  /** Lightweight CRM call to validate the bearer token. Runs once per toolbar init. */
  function validateToken() {
    if (!BEARER_TOKEN) {
      tokenValid = false;
      console.warn('[BirdsEye] No token set');
      return;
    }
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({
        operationName: 'userSearch', query: CRM_QUERY,
        variables: { input: { filter: 'token-validation-check', statuses: [], page: { index: 1, size: 1 } } }
      }),
      onload(res) {
        console.log('[BirdsEye] Token validation response:', res.status, res.responseText?.substring(0, 200));
        if (res.status === 401) {
          handle401();
          return;
        }
        if (res.status === 403) {
          handle403();
          return;
        }
        try {
          const json = JSON.parse(res.responseText);
          if (json.errors) {
            console.warn('[BirdsEye] Token validation GraphQL error:', json.errors);
            handle401();
            return;
          }
          tokenValid = true;
          const tokenBtn = document.getElementById('sb-crm-token-btn');
          if (tokenBtn) {
            tokenBtn.textContent = '🔑 Token';
            tokenBtn.style.color = '';
          }
          console.log('[BirdsEye] Token valid');
        } catch(e) {
          console.warn('[BirdsEye] Token validation parse error:', e);
        }
      },
      onerror(err) {
        console.warn('[BirdsEye] Token validation network error:', err);
      }
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // UTILITY HELPERS
  // ══════════════════════════════════════════════════════════════════════════

  function setReactValue(input, value) {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(input, value);
    input.dispatchEvent(new Event('input',  { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function toProperCase(str) {
    if (!str) return str;
    return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  }

  function ordinal(n) {
    const s = ['th','st','nd','rd'];
    const v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
  }

  /** Format date as "15 MAR '24" for info bar tags. */
  function fmtDateTag(isoStr) {
    if (!isoStr) return null;
    const d = new Date(isoStr);
    if (isNaN(d)) return null;
    const day = d.getUTCDate();
    const mon = d.toLocaleDateString('en-US', { month: 'short', timeZone: 'UTC' }).toUpperCase();
    const yr  = String(d.getUTCFullYear()).slice(2);
    return `${day} ${mon} '${yr}`;
  }

  /** Format date as "March 30th" for fill variables. */
  function fmtDateLong(isoStr) {
    if (!isoStr) return null;
    const d = new Date(isoStr);
    if (isNaN(d)) return null;
    const month = d.toLocaleDateString('en-US', { month: 'long', timeZone: 'UTC' });
    const day = ordinal(d.getUTCDate());
    return `${month} ${day}`;
  }

  function fmtMoney(cents, currency) {
    const amount = (cents / 100).toFixed(2);
    return currency === 'GBP' ? '£' + amount : '$' + amount;
  }

  function detectCurrency(lineItems) {
    if (!lineItems?.length) return 'USD';
    return lineItems.find(i => i.currency)?.currency || 'USD';
  }

  function chargeDateLabel(isoStr) {
    if (!isoStr) return '';
    return new Date(isoStr).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function chargeMonthLabel(isoStr) {
    if (!isoStr) return '';
    return new Date(isoStr).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  }

  function waitForField(selector, timeout = 4000) {
    return new Promise((resolve) => {
      if (document.querySelector(selector)) return resolve(document.querySelector(selector));
      const obs = new MutationObserver(() => {
        const el = document.querySelector(selector);
        if (el) { obs.disconnect(); resolve(el); }
      });
      obs.observe(document.body, { childList: true, subtree: true });
      setTimeout(() => { obs.disconnect(); resolve(null); }, timeout);
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // UI FACTORIES — shared overlay / modal / panel / helpers
  // ══════════════════════════════════════════════════════════════════════════

  /** Full-screen overlay. Closes on backdrop click. Returns overlay element. */
  function createOverlay(id) {
    document.getElementById(id)?.remove();
    const overlay = document.createElement('div');
    overlay.id = id;
    overlay.style.cssText = `
      all:initial;position:fixed;inset:0;background:rgba(0,0,0,0.6);
      z-index:10000000;display:flex;align-items:center;justify-content:center;
      font-family:Arial,sans-serif;
    `;
    overlay.addEventListener('click', e => {
      if (e.target === overlay) {
        overlay.remove();
        // Clean up associated panels
        if (id === 'sb-replacement-form') document.getElementById('sb-product-search-panel')?.remove();
      }
    });
    document.body.appendChild(overlay);
    return overlay;
  }

  /** Modal dialog inside an overlay. Returns { overlay, box }. */
  function createModal({ id, title, width = 360 }) {
    const overlay = createOverlay(id);
    const box = document.createElement('div');
    box.style.cssText = `
      background:#1e1e2e;border-radius:10px;padding:20px 24px;width:${width}px;
      max-height:85vh;overflow-y:auto;
      border:1px solid rgba(255,255,255,0.1);box-shadow:0 20px 60px rgba(0,0,0,0.7);
      color:#e2e8f0;font-family:Arial,sans-serif;font-size:14px;box-sizing:border-box;
    `;
    if (title) {
      const hdr = document.createElement('div');
      hdr.style.cssText = 'display:flex;justify-content:space-between;align-items:center;font-weight:700;font-size:14px;margin-bottom:12px;';
      const span = document.createElement('span');
      span.textContent = title;
      hdr.appendChild(span);
      const xBtn = document.createElement('button');
      xBtn.textContent = 'Close \u2715';
      xBtn.style.cssText = 'background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.15);color:#94a3b8;font-size:11px;font-weight:600;cursor:pointer;padding:3px 10px;border-radius:20px;';
      xBtn.onmouseenter = () => xBtn.style.background = 'rgba(255,255,255,0.15)';
      xBtn.onmouseleave = () => xBtn.style.background = 'rgba(255,255,255,0.08)';
      xBtn.onclick = () => overlay.remove();
      hdr.appendChild(xBtn);
      box.appendChild(hdr);
    }
    overlay.appendChild(box);
    return { overlay, box };
  }

  /** Fixed-position floating panel. Returns panel element (already in DOM). */
  function createPanel({ id, title, width = 380, right = 20, maxHeight = 680, onClose }) {
    removePanel(id);
    const panel = document.createElement('div');
    panel.id = id;
    panel.style.cssText = `
      all:initial;position:fixed;top:60px;right:${right}px;width:${width}px;max-height:${maxHeight}px;
      overflow-y:auto;background:#1e1e2e;color:#e2e8f0;font-family:Arial,sans-serif;
      font-size:13px;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,0.5);
      padding:14px;z-index:999999;border:1px solid rgba(255,255,255,0.1);box-sizing:border-box;
    `;
    const hdr = document.createElement('div');
    hdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;';
    hdr.innerHTML = `<span style="font-weight:700;font-size:14px;">${title}</span>`;
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close \u2715';
    closeBtn.style.cssText = 'background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.15);color:#94a3b8;font-size:11px;font-weight:600;cursor:pointer;padding:3px 10px;border-radius:20px;';
    closeBtn.onmouseenter = () => closeBtn.style.background = 'rgba(255,255,255,0.15)';
    closeBtn.onmouseleave = () => closeBtn.style.background = 'rgba(255,255,255,0.08)';
    closeBtn.onclick = () => { removePanel(id); onClose?.(); };
    hdr.appendChild(closeBtn);
    panel.appendChild(hdr);
    document.body.appendChild(panel);
    return panel;
  }

  /** Small user identity row used in panels. */
  function renderUserRow(user) {
    const row = document.createElement('div');
    row.style.cssText = 'color:#94a3b8;font-size:11px;margin-bottom:12px;';
    row.textContent = `${user.firstName || ''} ${user.lastName || ''} · ${user.email}`.trim();
    return row;
  }

  /** Status log widget for sequential operations. Call log.log(msg, color). */
  function makeStatusLog() {
    const el = document.createElement('div');
    el.style.cssText = 'margin-top:10px;font-size:11px;line-height:1.8;';
    el.log = (msg, color = '#94a3b8') => {
      const row = document.createElement('div');
      row.style.color = color;
      row.textContent = msg;
      el.appendChild(row);
    };
    return el;
  }

  function makeCopyBtn(textOrFn, label = '📋') {
    const b = document.createElement('button');
    b.textContent = label;
    b.title = 'Copy';
    b.style.cssText = 'background:none;border:none;cursor:pointer;color:#818cf8;font-size:13px;padding:0 4px;flex-shrink:0;';
    b.onclick = () => {
      const text = typeof textOrFn === 'function' ? textOrFn() : textOrFn;
      navigator.clipboard.writeText(text);
      b.textContent = '✔';
      setTimeout(() => b.textContent = label, 1500);
    };
    return b;
  }

  /** Section label used in forms. */
  function makeLabel(text) {
    const l = document.createElement('div');
    l.style.cssText = 'font-size:12px;color:#64748b;font-weight:600;letter-spacing:0.05em;margin-bottom:5px;margin-top:12px;';
    l.textContent = text;
    return l;
  }

  /** Select dropdown for forms. */
  function makeSelect(options, selected) {
    const s = document.createElement('select');
    s.style.cssText = 'width:100%;background:#0f172a;color:#e2e8f0;border:1px solid rgba(255,255,255,0.1);border-radius:5px;padding:6px 8px;font-size:13px;box-sizing:border-box;cursor:pointer;';
    options.forEach(o => {
      const opt = document.createElement('option');
      opt.value = o.value;
      opt.textContent = o.label;
      if (o.value === selected) opt.selected = true;
      s.appendChild(opt);
    });
    return s;
  }

  /** Status element for panel footers. */
  function makeStatusEl() {
    const el = document.createElement('div');
    el.style.cssText = 'font-size:12px;min-height:16px;margin-top:10px;color:#94a3b8;';
    return el;
  }

  /** Styled tag (used in info bar and elsewhere). */
  function makeTag(label, color, bg, border) {
    const t = document.createElement('span');
    t.style.cssText = `display:inline-flex;align-items:center;padding:2px 7px;border-radius:4px;font-weight:600;font-size:10px;color:${color};background:${bg};border:1px solid ${border};white-space:nowrap;`;
    t.textContent = label;
    return t;
  }

  /** Standard action button pair (cancel + confirm) for dialogs. */
  function makeDialogButtons({ confirmLabel, confirmColor = '#dc2626', onCancel, onConfirm }) {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;';

    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.cssText = `
      padding:7px 16px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);
      background:transparent;color:#e2e8f0;font-size:13px;cursor:pointer;flex:1;
    `;
    cancelBtn.onclick = onCancel;

    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = confirmLabel;
    confirmBtn.style.cssText = `
      padding:7px 16px;border-radius:6px;border:none;
      background:${confirmColor};color:#fff;font-weight:700;font-size:13px;cursor:pointer;flex:1;
    `;
    confirmBtn.onmouseenter = () => confirmBtn.style.filter = 'brightness(0.85)';
    confirmBtn.onmouseleave = () => confirmBtn.style.filter = '';
    confirmBtn.onclick = () => onConfirm(confirmBtn, cancelBtn);

    row.appendChild(cancelBtn);
    row.appendChild(confirmBtn);
    row._confirmBtn = confirmBtn;
    row._cancelBtn  = cancelBtn;
    return row;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // TOOLBAR
  // ══════════════════════════════════════════════════════════════════════════

  function ensureToolbar() {
    if (toolbarEl && document.body.contains(toolbarEl)) return toolbarEl;

    const card = document.querySelector('[data-kt="card_card"]');
    if (!card) return null;

    document.getElementById('sb-toolbar')?.remove();

    toolbarEl = document.createElement('div');
    toolbarEl.id = 'sb-toolbar';
    toolbarEl.style.cssText = `
      display:flex;align-items:center;gap:6px;
      padding:3px 10px;
      background:#1a1a2e;
      border-bottom:1px solid rgba(255,255,255,0.08);
      min-height:28px;flex-wrap:wrap;
      box-sizing:border-box;width:100%;
    `;

    document.getElementById('sb-info-bar-slot')?.remove();
    const infoSlot = document.createElement('div');
    infoSlot.id = 'sb-info-bar-slot';
    infoSlot.style.cssText = 'box-sizing:border-box;width:100%;';

    card.insertAdjacentElement('afterend', infoSlot);
    infoSlot.insertAdjacentElement('beforebegin', toolbarEl);
    return toolbarEl;
  }

  function addToolbarButton(id, label, onClick) {
    const toolbar = ensureToolbar();
    if (!toolbar || toolbar.querySelector('#' + id)) return null;

    const btn = document.createElement('button');
    btn.id = id;
    btn.textContent = label;
    btn.style.cssText = `
      font-size:13px;font-weight:600;padding:5px 12px;border-radius:4px;
      border:1px solid rgba(255,255,255,0.15);background:rgba(255,255,255,0.07);
      color:#e2e8f0;cursor:pointer;white-space:nowrap;transition:background 0.15s;
    `;
    btn.addEventListener('mouseenter', () => btn.style.background = 'rgba(255,255,255,0.15)');
    btn.addEventListener('mouseleave', () => btn.style.background = 'rgba(255,255,255,0.07)');
    btn.addEventListener('click', () => onClick(btn));
    toolbar.appendChild(btn);
    return btn;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // EMAIL RESOLUTION (Kustomer API)
  // ══════════════════════════════════════════════════════════════════════════

  /** Extract Kustomer customer ID (24-char hex) from the current URL. */
  function getCustomerIdFromURL() {
    const m = window.location.href.match(/\/customers\/([0-9a-f]{24})/);
    return m ? m[1] : null;
  }

  /** Fetch customer email via Kustomer's own API. Cookie-authenticated. */
  async function resolveEmail() {
    const customerId = getCustomerIdFromURL();
    if (!customerId) return null;
    try {
      const res = await fetch(`https://scentbird.api.kustomerapp.com/v1/customers/${customerId}`, {
        credentials: 'include',
        headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
      });
      if (!res.ok) return null;
      const json = await res.json();
      const email = json?.data?.attributes?.emails?.[0]?.email;
      return email ? email.trim().toLowerCase() : null;
    } catch(e) {
      console.warn('[BirdsEye] Kustomer API fetch failed:', e);
      return null;
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // GRAPHQL QUERIES & MUTATIONS (single source of truth)
  // ══════════════════════════════════════════════════════════════════════════

  const CRM_QUERY = `
    query userSearch($input: SearchUsersInput) {
      userSearch(input: $input) {
        total
        data { id email firstName lastName origin
          userAddress { shipping { street1 city region postcode } }
          subscription { id status }
        }
      }
    }`;

  const ORDERS_QUERY = `
    query ordersByUserId($id: Long) {
      userById(id: $id) {
        data {
          orderList {
            id status type month year
            shipment { trackingNumber trackingUrl }
            tracking { status trackingNumber trackingUrl items { date status description } }
            warehouseOrder { data { number status } }
            tags orderNumber
            initialShippingAddress { street1 city region postcode }
            orderItems {
              id
              product { id status productInfo { name brand productType } }
            }
          }
        }
      }
    }`;

  const CHARGES_QUERY = `
    query chargesByUserId($id: Long) {
      userById(id: $id) {
        data {
          chargeList {
            date kind
            charge {
              id paymentDate category status totalPrice
              cashbirdDetails {
                invoice {
                  lineItems { uuid description productCode price total tax discount type currency refundedAmount }
                }
                credits { id status reason price amount tax total currency }
                shippingCredits { id type status price amount tax total currency }
                paymentMethod { methodName type }
              }
              refundAmount
              orders { id }
            }
          }
        }
      }
    }`;

  const USER_DETAILS_QUERY = `
    query userDetailsById($id: Long) {
      userById(id: $id) {
        data {
          fraudInfo { status }
          gwebLink
          gender
          userAddress { shipping { id } }
          subscriptionList {
            id status subscribed nextBillingDate planName subscriptionDate subscriptionEndDate
            cashbirdDetails { data { billingDay nextBillingDate isAwaitCancellation } }
            credits { id type status reason }
            coupons { id code type amount status applyToNextBillingDates }
            addOnSettings {
              candleSubscription    { enabled selected kind }
              caseSubscription      { enabled selected kind }
              carScentSubscription  { enabled selected kind }
              samplesSubscription   { enabled selected kind }
              homeDiffuserSubscription { enabled selected kind }
            }
          }
        }
      }
    }`;

  const COMMENTS_QUERY = `
    query userComments($id: Long) {
      userById(id: $id) {
        data {
          commentList { id created author comment zendeskUrl ownerId }
        }
      }
    }`;

  const CREATE_COMMENT_MUTATION = `
    mutation createUserComment($input: CreateCommentInput) {
      createUserComment(input: $input) {
        message
        error { message __typename }
        __typename
      }
    }`;

  const DELETE_COMMENT_MUTATION = `
    mutation deleteUserComment($id: Long) {
      deleteUserComment(id: $id) {
        message
        error { message __typename }
        __typename
      }
    }`;

  const REPLACEMENT_MUTATION = `
    mutation replacementSave($input: SaveReplacementTaskInput) {
      replacementSave(input: $input) {
        data { id orderNumber status __typename }
        error { message __typename }
        __typename
      }
    }`;

  const DELETE_PAYMENT_METHODS_MUTATION = `
    mutation deleteAllPaymentMethods($input: PaymentMethodDeleteAllInput!) {
      paymentMethodDeleteAll(input: $input) {
        error {
          ... on ServerError {
            serverErrorCode: errorCode
            serverErrorMessage: message
            __typename
          }
          ... on PaymentMethodDeleteError {
            errorCode
            message
            __typename
          }
          __typename
        }
        __typename
      }
    }`;

  const CHANGE_EMAIL_MUTATION = `
    mutation userChangeEmail($input: ChangeUserEmailInput) {
      userChangeEmail(input: $input) {
        data { id username }
        error {
          ... on UpdateUserError { message }
        }
      }
    }`;

  const SET_FRAUD_STATUS_MUTATION = `
    mutation userSetFraudStatus($input: UserSetFraudStatusInput!) {
      userSetFraudStatus(input: $input) {
        error { message }
        data { fraudInfo { status } }
      }
    }`;

  const CUSTOM_SHIPMENT_MUTATION = `
    mutation customShipmentSave($input: SaveCustomShipmentTaskInput) {
      customShipmentSave(input: $input) {
        data {
          id status type orderNumber year month
          orderItems {
            id
            product { id productInfo { name brand } }
          }
        }
        error { message }
      }
    }`;

  const ADD_CREDITS_MUTATION = `
    mutation addSupportCredits($input: AddSupportCreditsInput) {
      addSupportCredits(input: $input) {
        error {
          ... on AddSupportCreditsError {
            errorCode
            message
          }
        }
      }
    }`;

  const CHARGE_MUTATION = `
    mutation charge($input: ChargeInput!) {
      charge(input: $input) {
        data { id number paid total subtotal tax state paymentMethodType }
        error { message errorCode }
      }
    }`;

  const PRODUCT_SEARCH_QUERY = `
    query productSuggestion($input: FindProductInput) {
      productSuggestion(input: $input) {
        data {
          id sku status section price currency volume volumeUnit
          productInfo { id name brand upchargePrice productType __typename }
          __typename
        }
        error { message __typename }
        __typename
      }
    }`;

  const QUEUE_QUERY = `
    query subscriptionQueue($input: SubscriptionQueueInput!) {
      subscriptionQueue(input: $input) {
        data {
          blockList {
            yearMonth
            status
            processedSubscriptionOrder { id status }
            productList {
              tradingItem {
                id sku
                productInfo { id name brand upchargePrice productType }
              }
              source
              subscriptionQueueItemId
              upchargePrice
            }
          }
        }
        error { message }
      }
    }`;

  // ══════════════════════════════════════════════════════════════════════════
  // DATA FETCHERS
  // ══════════════════════════════════════════════════════════════════════════

  const COMMENT_OPS = ['createUserComment', 'deleteUserComment'];

  function gqlMutate(operationName, query, variables, callback) {
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName, query, variables }),
      onload(res) {
        console.log('[BirdsEye]', operationName, '→', res.status, res.responseText?.substring(0, 300));
        if (res.status === 401) { handle401(); return callback(null, 'Token expired'); }
        if (res.status === 403) { handle403(); return callback(null, 'CRM captcha required — open CRM in browser'); }
        try {
          const json = JSON.parse(res.responseText);
          const errMsg = json?.errors?.[0]?.message;
          if (errMsg) return callback(null, errMsg);
          // Auto-refresh info bar after data-changing mutations
          if (!COMMENT_OPS.includes(operationName)) refreshSubscriptionBar();
          callback(json.data);
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); }
    });
  }

  function searchCRM(q, callback) {
    if (!BEARER_TOKEN) {
      callback(null, 'No token set — click 🔑 Token to add one.');
      return;
    }
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({
        operationName: 'userSearch', query: CRM_QUERY,
        variables: { input: { filter: q, statuses: [], page: { index: 1, size: 50 } } }
      }),
      onload(res) {
        console.log('[BirdsEye] userSearch →', res.status, res.responseText?.substring(0, 300));
        if (res.status === 401) { handle401(); callback(null, 'Token expired'); return; }
        if (res.status === 403) { handle403(); callback(null, 'CRM captcha required'); return; }
        try {
          const data = JSON.parse(res.responseText);
          if (data.errors) return callback(null, data.errors[0]?.message || 'GraphQL error');
          callback(data?.data?.userSearch?.data || []);
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); }
    });
  }

  function fetchLastOrders(userId, callback) {
    if (!BEARER_TOKEN) { console.warn('[BirdsEye] fetchLastOrders blocked — no token'); return callback(null); }
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName: 'ordersByUserId', query: ORDERS_QUERY, variables: { id: userId } }),
      onload(res) {
        console.log('[BirdsEye] ordersByUserId →', res.status, res.responseText?.substring(0, 300));
        if (res.status === 401) { handle401(); return callback(null, 'Token expired'); }
        if (res.status === 403) { handle403(); return callback(null); }
        try {
          const json = JSON.parse(res.responseText);
          const list = json?.data?.userById?.data?.orderList;
          if (!list?.length) return callback([]);
          callback(list.slice(0, 4));
        } catch (e) { callback(null); }
      },
      onerror() { callback(null); },
    });
  }

  function fetchCharges(userId, callback) {
    if (!BEARER_TOKEN) { console.warn('[BirdsEye] fetchCharges blocked — no token'); return callback(null); }
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName: 'chargesByUserId', query: CHARGES_QUERY, variables: { id: userId } }),
      onload(res) {
        console.log('[BirdsEye] chargesByUserId →', res.status, res.responseText?.substring(0, 300));
        if (res.status === 401) { handle401(); return callback(null); }
        if (res.status === 403) { handle403(); return callback(null); }
        try {
          const json = JSON.parse(res.responseText);
          const list = json?.data?.userById?.data?.chargeList;
          if (!list?.length) return callback({ success: null, failed: [] });

          const lastSuccess = list.find(c => c.charge?.status === 'SUCCESS');
          if (!lastSuccess) return callback({ success: null, failed: [] });

          const refDate = new Date(lastSuccess.charge.paymentDate || lastSuccess.date);
          const PROXIMITY_MS = 10 * 24 * 60 * 60 * 1000;

          const monthCharges = list.filter(c => {
            const d = new Date(c.charge?.paymentDate || c.date);
            return Math.abs(d - refDate) <= PROXIMITY_MS;
          });

          const successEntries = monthCharges.filter(c => c.charge?.status === 'SUCCESS');
          const failedEntries  = monthCharges.filter(c => c.charge?.status === 'FAIL' || c.charge?.category === 'FAILED_CHARGE');
          callback({ success: lastSuccess, all: successEntries, failed: failedEntries });
        } catch(e) { callback(null); }
      },
      onerror() { callback(null); }
    });
  }

  function fetchActiveSubscription(userId, callback) {
    if (!BEARER_TOKEN) { console.warn('[BirdsEye] fetchActiveSubscription blocked — no token'); return callback(null, 'No token'); }
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName: 'userDetailsById', query: USER_DETAILS_QUERY, variables: { id: userId } }),
      onload(res) {
        console.log('[BirdsEye] subscriptionsByUserId →', res.status, res.responseText?.substring(0, 300));
        if (res.status === 401) { handle401(); return callback(null, 'Token expired'); }
        if (res.status === 403) { handle403(); return callback(null, 'CRM captcha required'); }
        try {
          const json = JSON.parse(res.responseText);
          const list = json?.data?.userById?.data?.subscriptionList || [];
          const active = list.find(s => s.status === 'Active' || s.subscribed);
          callback(active?.id || null, active ? null : 'No active subscription');
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); }
    });
  }

  function fetchComments(userId, callback) {
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName: 'userComments', query: COMMENTS_QUERY, variables: { id: userId } }),
      onload(res) {
        try {
          const json = JSON.parse(res.responseText);
          callback(json?.data?.userById?.data?.commentList || []);
        } catch(e) { callback([]); }
      },
      onerror() { callback([]); }
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // MUTATION HELPERS
  // ══════════════════════════════════════════════════════════════════════════

  function mutCancelSub(subscriptionId, callback) {
    gqlMutate('cancelSubscription',
      `mutation cancelSubscription($input: CancelSubscriptionInput) {
        cancelSubscription(input: $input) {
          data { id status __typename }
          error { message __typename }
          __typename
        }
      }`,
      { input: { id: subscriptionId, decision: 'RECEIVE', comment: null } },
      (data, err) => callback(err ? null : data?.cancelSubscription, err)
    );
  }

  function mutCancelOrder(orderId, callback) {
    gqlMutate('cancelOrderById',
      `mutation cancelOrderById($input: CancelOrderByIdInput!) {
        cancelOrderById(input: $input) {
          data { order { id status __typename } __typename }
          error { message __typename }
          __typename
        }
      }`,
      { input: { id: orderId, subscriptionHold: true, unlockItems: false, refundReason: 'CUSTOMER_CHANGED_THEIR_MIND' } },
      (data, err) => callback(err ? null : data?.cancelOrderById, err)
    );
  }

  function mutHoldOrder(orderId, callback) {
    gqlMutate('holdOrderById',
      `mutation holdOrderById($id: Long) {
        holdOrderById(id: $id) {
          data { id status __typename }
          error { message __typename }
          __typename
        }
      }`,
      { id: orderId },
      (data, err) => callback(err ? null : data?.holdOrderById, err)
    );
  }

  function mutUnholdOrder(orderId, callback) {
    gqlMutate('unholdOrderById',
      `mutation unholdOrderById($id: Long) {
        unholdOrderById(id: $id) {
          data { id status __typename }
          error { message __typename }
          __typename
        }
      }`,
      { id: orderId },
      (data, err) => callback(err ? null : data?.unholdOrderById, err)
    );
  }

  function mutRefundCredits(userId, creditIds, shippingCreditIds, callback) {
    gqlMutate('refundCredits',
      `mutation refundCredits($input: RefundCreditsInput) {
        refundCredits(input: $input) {
          message
          error { message __typename }
          __typename
        }
      }`,
      { input: {
        userId,
        credits: creditIds,
        forceRefundUsedCredits: true,
        shippingCredits: shippingCreditIds,
        forceRefundUsedShippingCredits: true,
        reason: 'CUSTOMER_CHANGED_THEIR_MIND'
      }},
      (data, err) => callback(err ? null : data?.refundCredits, err)
    );
  }

  function mutRefundInvoiceItems(userId, items, callback) {
    gqlMutate('refundInvoiceItems',
      `mutation refundInvoiceItems($input: RefundInvoiceItemListInput) {
        refundInvoiceItems(input: $input) {
          message
          error { message __typename }
          __typename
        }
      }`,
      { input: { userId, items, reason: 'CUSTOMER_CHANGED_THEIR_MIND' } },
      (data, err) => callback(err ? null : data?.refundInvoiceItems, err)
    );
  }

  /** Posts a comment, merging with our own recent comment for this ticket URL. */
  function mutPostComment(userId, comment, callback) {
    const zendeskUrl = window.location.href;
    const THREE_DAYS_MS = 3 * 24 * 60 * 60 * 1000;

    fetchComments(userId, (comments) => {
      // Find our own recent comment on this ticket
      const now = Date.now();
      const match = comments.find(c => {
        if (c.zendeskUrl !== zendeskUrl) return false;
        if (_myOwnerId && c.ownerId !== _myOwnerId) return false;
        if (!c.created) return false;
        const age = now - new Date(c.created).getTime();
        return age <= THREE_DAYS_MS;
      });

      const finalComment = match ? match.comment + '\n' + comment : comment;

      const doCreate = () => gqlMutate('createUserComment',
        CREATE_COMMENT_MUTATION,
        { input: { userId, comment: finalComment, zendeskUrl } },
        (data, err) => {
          // Discover our ownerId after first successful create
          if (!err && !_myOwnerId) {
            fetchComments(userId, (freshComments) => {
              const ours = freshComments
                .filter(c => c.zendeskUrl === zendeskUrl)
                .sort((a, b) => new Date(b.created) - new Date(a.created))[0];
              if (ours?.ownerId) _myOwnerId = ours.ownerId;
            });
          }
          callback(err ? null : data?.createUserComment, err);
        }
      );

      if (match) {
        gqlMutate('deleteUserComment', DELETE_COMMENT_MUTATION, { id: match.id },
          () => doCreate()   // proceed even if delete failed
        );
      } else {
        doCreate();
      }
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // CUSTOMER CONTEXT — resolve once, share across panels
  // ══════════════════════════════════════════════════════════════════════════

  /**
   * Resolves the current customer from the iframe email → CRM search.
   * Caches the result so subsequent calls within the same customer page
   * return instantly without extra API calls.
   */
  async function loadCustomer(btn, callback) {
    // Token gate — check before anything else, including cache
    if (!BEARER_TOKEN) {
      const orig = btn.textContent;
      btn.textContent = '\u2718 No token'; btn.style.color = '#fca5a5';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; }, 2000);
      console.warn('[BirdsEye] loadCustomer blocked — no token set');
      return;
    }

    // Fast path: already resolved for this customer page
    if (cachedCustomerCtx?.user) {
      return callback({ user: cachedCustomerCtx.user });
    }

    const orig = btn.textContent;
    btn.textContent = '\u23F3'; btn.disabled = true;
    const restore = () => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; };
    const fail = (msg) => {
      btn.textContent = '\u2718 ' + msg; btn.style.color = '#fca5a5';
      setTimeout(restore, 2000);
    };

    const startId = getCustomerIdFromURL();
    const email = await resolveEmail();
    if (getCustomerIdFromURL() !== startId) { restore(); return; } // stale
    if (!email) return fail('No email');
    searchCRM(email, (users, err) => {
      if (getCustomerIdFromURL() !== startId) { restore(); return; } // stale
      restore();
      if (err || !users?.length) return fail('Not found');
      const sbUsers = users.filter(u => !u.origin || u.origin === 'SCENTBIRD');
      if (!sbUsers.length) return fail('No Scentbird account found');
      // Prefer exact email match to avoid binding wrong account on partial matches
      const exactMatch = sbUsers.find(u => u.email?.toLowerCase() === email.toLowerCase());
      const bestUser = exactMatch || sbUsers[0];
      cachedCustomerCtx = { email, user: bestUser, _kustomerId: getCustomerIdFromURL() };
      callback({ user: bestUser });
    });
  }

  function loadCustomerOrders(userId, callback) {
    fetchLastOrders(userId, (orders) => callback(orders || []));
  }

  function loadCustomerCharges(userId, callback) {
    fetchCharges(userId, (result) => callback(result));
  }

  function fetchQueue(userId, callback) {
    if (!BEARER_TOKEN) { console.warn('[BirdsEye] fetchQueue blocked — no token'); return callback(null); }
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName: 'subscriptionQueue', query: QUEUE_QUERY, variables: { input: { userId } } }),
      onload(res) {
        console.log('[BirdsEye] subscriptionQueue →', res.status, res.responseText?.substring(0, 300));
        if (res.status === 401) { handle401(); return callback(null); }
        if (res.status === 403) { handle403(); return callback(null); }
        try {
          const json = JSON.parse(res.responseText);
          const blockList = json?.data?.subscriptionQueue?.data?.blockList;
          callback(blockList || []);
        } catch(e) { callback(null); }
      },
      onerror() { callback(null); }
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // ORDER RENDERING
  // ══════════════════════════════════════════════════════════════════════════

  function orderMonthLabel(order) {
    const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${m[(order.month || 1) - 1]} ${order.year}`;
  }

  function orderProductLines(order) {
    return (order.orderItems || [])
      .map(i => {
        const p = i?.product?.productInfo;
        if (!p) return null;
        return `${p.name} by ${p.brand}`;
      })
      .filter(Boolean);
  }

  function orderCopyText(order) {
    return orderProductLines(order).join('\n');
  }

  const CANCELLABLE_STATUSES = ["PENDING", "UNPROCESSED", "UPCHARGE_WAITING_CHARGE", "ITEM_NOT_AVAILABLE", "BACKORDERED", "NEW"];
  const NON_REFUNDABLE_STATUSES = ["SHIPPED", "DELIVERED", "PRINTED", "DONE", "PROCESSED"];

  function renderOrderBlock(order, compact = false, ctx = null, isReplacement = false) {
    const wrap = document.createElement('div');
    const borderColor = isReplacement ? '#7c3aed' : '#4f46e5';
    wrap.style.cssText = `
      margin-top:${compact ? 4 : 0}px;padding:${compact ? '6px 8px' : '10px 12px'};
      background:${isReplacement ? '#1a1530' : '#1a1a2e'};border-radius:6px;border-left:2px solid ${borderColor};
    `;

    const _ws = (order.warehouseOrder?.data?.status || order.status || '').toUpperCase();
    const statusColor = (_ws === 'SHIPPED' || _ws === 'DELIVERED') ? '#6ee7b7'
      : (_ws.includes('CANCEL') || _ws.includes('BACK')) ? '#fca5a5'
      : (_ws === 'PRINTED' || _ws.includes('PROCESS')) ? '#818cf8'
      : (_ws === 'PENDING') ? '#f59e0b'
      : '#94a3b8';

    // Header row
    const hdr = document.createElement('div');
    hdr.style.cssText = 'display:flex;align-items:center;gap:6px;margin-bottom:4px;flex-wrap:wrap;';
    const hasWelcomeKit = (order.tags || []).includes('REBRAND_CASE') || order.hasRebrandCase;
    const hasSamples = (order.orderItems || []).some(oi =>
      oi.product?.productInfo?.productType === 'UPCHARGE_SAMPLES' || oi.productType === 'UPCHARGE_SAMPLES'
    );
    const hasCase = (order.orderItems || []).some(oi =>
      oi.product?.productInfo?.productType === 'PERFUME_CASE_UPCHARGE' || oi.productType === 'PERFUME_CASE_UPCHARGE'
    );
    hdr.innerHTML = `
      <span style="color:#94a3b8;font-size:11px;">${orderMonthLabel(order)}</span>
      <span style="color:${statusColor};font-weight:600;font-size:11px;">${order.warehouseOrder?.data?.status || order.status}</span>
      ${(order.type && order.type !== 'SUBSCRIPTION') ? `<span style="color:#f59e0b;font-size:11px;font-weight:600;">${order.type.replace(/_/g,' ')}</span>` : ''}
      ${hasWelcomeKit ? '<span style="background:rgba(99,102,241,0.2);color:#a5b4fc;font-size:10px;font-weight:600;padding:1px 5px;border-radius:4px;border:1px solid rgba(99,102,241,0.4);">Welcome Kit</span>' : ''}
      ${hasSamples ? '<span style="background:rgba(245,158,11,0.15);color:#fcd34d;font-size:10px;font-weight:600;padding:1px 5px;border-radius:4px;border:1px solid rgba(245,158,11,0.3);">Samples</span>' : ''}
      ${hasCase ? '<span style="background:rgba(99,102,241,0.15);color:#a5b4fc;font-size:10px;font-weight:600;padding:1px 5px;border-radius:4px;border:1px solid rgba(99,102,241,0.3);">Case</span>' : ''}
    `;
    hdr.appendChild(makeCopyBtn(() => orderCopyText(order), '📋 Copy'));
    wrap.appendChild(hdr);

    // Product lines
    orderProductLines(order).forEach(line => {
      const d = document.createElement('div');
      d.style.cssText = 'color:#cbd5e1;font-size:12px;margin-bottom:2px;';
      d.textContent = line;
      wrap.appendChild(d);
    });

    // Tracking row
    const trackNo  = order.tracking?.trackingNumber || order.shipment?.trackingNumber || '';
    const trackUrl = order.tracking?.trackingUrl  || order.shipment?.trackingUrl  || '';

    if (trackNo || trackUrl) {
      const tRow = document.createElement('div');
      tRow.style.cssText = 'display:flex;align-items:center;gap:4px;margin-top:5px;flex-wrap:wrap;';

      if (trackUrl) {
        const link = document.createElement('a');
        link.href = trackUrl; link.target = '_blank';
        link.textContent = trackNo || 'Track';
        link.style.cssText = 'color:#818cf8;font-size:11px;text-decoration:none;';
        tRow.appendChild(link);
      } else {
        const span = document.createElement('span');
        span.style.cssText = 'color:#818cf8;font-size:11px;';
        span.textContent = trackNo;
        tRow.appendChild(span);
      }

      if (trackUrl) {
        const linkBtn = document.createElement('button');
        linkBtn.textContent = '🔗';
        linkBtn.title = 'Copy tracking link';
        linkBtn.style.cssText = 'background:none;border:none;cursor:pointer;color:#818cf8;font-size:13px;padding:0 4px;flex-shrink:0;';
        linkBtn.onclick = () => {
          navigator.clipboard.write([
            new ClipboardItem({
              'text/html':  new Blob([`<a href="${trackUrl}">${trackNo}</a>`], { type: 'text/html' }),
              'text/plain': new Blob([trackUrl], { type: 'text/plain' }),
            })
          ]);
          linkBtn.textContent = '✔';
          setTimeout(() => linkBtn.textContent = '🔗', 1500);
        };
        tRow.appendChild(linkBtn);
      }
      wrap.appendChild(tRow);
    }

    // Last 2 tracking events
    const items = order.tracking?.items;
    if (items?.length) {
      const statusEl = document.createElement('div');
      statusEl.style.cssText = 'margin-top:6px;';
      items.slice(-2).forEach(item => {
        const d = document.createElement('div');
        d.style.cssText = 'color:#64748b;font-size:11px;margin-top:2px;';
        const date = new Date(item.date).toLocaleDateString('en-US', { month:'short', day:'numeric' });
        d.textContent = `${date} · ${item.description}`;
        statusEl.appendChild(d);
      });
      wrap.appendChild(statusEl);
    }

    // Replace + Cancel buttons — only in full panel
    if (!compact && ctx) {
      const btnRow = document.createElement('div');
      btnRow.style.cssText = 'display:flex;gap:6px;margin-top:8px;flex-wrap:wrap;';

      const replBtn = document.createElement('button');
      replBtn.textContent = '\uD83D\uDD04 Replace';
      replBtn.style.cssText = 'padding:5px 10px;border-radius:5px;border:1px solid #6366f1;background:transparent;color:#a5b4fc;font-weight:600;font-size:11px;cursor:pointer;';
      replBtn.onmouseenter = () => replBtn.style.background = 'rgba(99,102,241,0.15)';
      replBtn.onmouseleave = () => replBtn.style.background = 'transparent';
      replBtn.onclick = () => showReplacementForm(order, ctx.user, ctx.totalOrders);
      btnRow.appendChild(replBtn);

      const ws = (order.warehouseOrder?.data?.status || order.status || '').toUpperCase();
      if (CANCELLABLE_STATUSES.includes(ws)) {
        const cancelOrderBtn = document.createElement('button');
        cancelOrderBtn.textContent = '\u274C Cancel Order';
        cancelOrderBtn.style.cssText = 'padding:5px 10px;border-radius:5px;border:1px solid #dc2626;background:transparent;color:#fca5a5;font-weight:600;font-size:11px;cursor:pointer;';
        cancelOrderBtn.onmouseenter = () => cancelOrderBtn.style.background = 'rgba(220,38,38,0.15)';
        cancelOrderBtn.onmouseleave = () => cancelOrderBtn.style.background = 'transparent';
        cancelOrderBtn.onclick = () => showCancelOrderDialog(order, ctx.user, cancelOrderBtn);
        btnRow.appendChild(cancelOrderBtn);
      }

      // Hold / Unhold button
      const isOnHold = ws === 'ON_HOLD';
      const canHold = ws === 'UNPROCESSED' || isOnHold;
      if (canHold) {
        const holdBtn = document.createElement('button');
        holdBtn.textContent = isOnHold ? '▶ Unhold' : '⏸ Hold';
        holdBtn.style.cssText = `padding:5px 10px;border-radius:5px;border:1px solid ${isOnHold ? '#6ee7b7' : '#f59e0b'};background:transparent;color:${isOnHold ? '#6ee7b7' : '#f59e0b'};font-weight:600;font-size:11px;cursor:pointer;`;
        holdBtn.onmouseenter = () => holdBtn.style.background = isOnHold ? 'rgba(110,231,183,0.15)' : 'rgba(245,158,11,0.15)';
        holdBtn.onmouseleave = () => holdBtn.style.background = 'transparent';
        holdBtn.onclick = () => {
          holdBtn.disabled = true;
          holdBtn.textContent = isOnHold ? 'Unholding...' : 'Holding...';
          const mut = isOnHold ? mutUnholdOrder : mutHoldOrder;
          mut(order.id, (data, err) => {
            if (err) {
              holdBtn.textContent = '✘ ' + err;
              holdBtn.style.color = '#fca5a5';
              setTimeout(() => { holdBtn.disabled = false; holdBtn.textContent = isOnHold ? '▶ Unhold' : '⏸ Hold'; holdBtn.style.color = isOnHold ? '#6ee7b7' : '#f59e0b'; }, 2000);
              return;
            }
            holdBtn.textContent = '✔ Done';
            holdBtn.style.color = '#6ee7b7';
            holdBtn.style.borderColor = '#6ee7b7';
            // Post comment on hold (not on unhold)
            if (!isOnHold && ctx?.user) {
              const monthLabel = orderMonthLabel(order);
              mutPostComment(ctx.user.id, `Order ${monthLabel} on hold`, () => {});
            }
          });
        };
        btnRow.appendChild(holdBtn);
      }
      wrap.appendChild(btnRow);
    }

    return wrap;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // CHARGE EXPLANATION
  // ══════════════════════════════════════════════════════════════════════════

  function lineItemExplanation(item, billingDate, currency) {
    const t     = item.type;
    const tax   = item.tax;
    const disc  = item.discount;
    const total = item.total;
    const month = chargeMonthLabel(billingDate);
    const taxStr  = tax  ? ` (after ${fmtMoney(tax, currency)} for sales tax)` : '';
    const discStr = disc ? `, including a ${fmtMoney(disc, currency)} discount` : '';

    switch (t) {
      case 'FIRST_SUBSCRIPTION':
        return `* ${fmtMoney(total, currency)} was your initial subscription payment for ${month} (${fmtMoney(item.price, currency)}${discStr} + ${fmtMoney(tax, currency)} tax);`;
      case 'RECURRENT_CHARGE':
        return `* ${fmtMoney(total, currency)} was your monthly subscription payment for ${month} (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax);`;
      case 'RESUBSCRIBE':
        return `* ${fmtMoney(total, currency)} was your resubscription payment for ${month} (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax);`;
      case 'UPGRADE_PARTIAL':
        return `* ${fmtMoney(total, currency)} was a charge for upgrading your account (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax).\nAs soon as you upgraded your plan, the difference in prices between the two plans was billed. Since all our subscriptions start out as a basic 1 fragrance/month and can be upgraded at any point, this transaction was applied as a separate charge;`;
      case 'UPGRADE_FULL':
        return `* ${fmtMoney(total, currency)} was a charge for a full plan upgrade (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax).\nWhen upgrading your subscription plan, the full amount for the new plan is billed immediately as a separate charge;`;
      case 'PREMIUM_PRODUCT_UPCHARGE': {
        const name = (item.description && item.description.toLowerCase() !== 'premium product')
          ? item.description : 'a premium fragrance';
        return `* ${fmtMoney(total, currency)} was a premium fragrance upcharge for including ${name} in your order (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax).\nSome fragrances on our website are part of a high-end lux collection and carry an upcharge of up to $25. They are marked by a tag in the left corner, and adding them to your queue triggers a pop-up with information about the extra cost;`;
      }
      case 'PERFUME_CASE_UPCHARGE':
        return `* ${fmtMoney(total, currency)} was a charge for including a Case Subscription add-on in your plan (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax). With this, you get a different Fragrance Case sent to you together with each shipment;`;
      case 'SHIPPING_CREDIT':
      case 'SHIPPING_UPCHARGE':
        return `* ${fmtMoney(total, currency)} was a shipping surcharge (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax);`;
      case 'SAMPLES_UPCHARGE':
        return `* ${fmtMoney(total, currency)} was a charge for including a Samples subscription add-on in your order (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax). With this add-on you get two 1.5ml samples with each order;`;
      case 'SHIPPING_PURCHASE':
        return `* ${fmtMoney(total, currency)} was a charge for the shipping of your order (${fmtMoney(item.price, currency)} + ${fmtMoney(tax, currency)} tax);`;
      case 'ECOMMERCE_PURCHASE':
        return `* ${fmtMoney(total, currency)} was a charge for the separate Online Shop order you placed with us (${fmtMoney(item.price, currency)}${discStr}${taxStr});`;
      case 'FAILED_CHARGE':
        return null;
      default:
        return `* ${item.description || t} — ${fmtMoney(total, currency)}${taxStr}; [unknown type: ${t}]`;
    }
  }

  function buildBillExplanation(successEntry) {
    const charge = successEntry?.charge;
    if (!charge) return '';
    const items = charge.cashbirdDetails?.invoice?.lineItems || [];
    const billingDate = charge.paymentDate || successEntry.date;
    const currency = detectCurrency(items);
    const lines = items.filter(i => (i.total || 0) !== 0).map(i => lineItemExplanation(i, billingDate, currency)).filter(Boolean);
    const hasShipping = (charge.cashbirdDetails?.shippingCredits?.length > 0)
      || items.some(i => i.type === 'SHIPPING_CREDIT' || i.type === 'SHIPPING_UPCHARGE');
    if (hasShipping) {
      lines.push('\nPlease note that all charges are processed in USD.');
    }
    return lines.join('\n');
  }

  // ══════════════════════════════════════════════════════════════════════════
  // APPLY USER (Kustomer modal)
  // ══════════════════════════════════════════════════════════════════════════

  async function applyUser(user) {
    if (user.origin && user.origin !== 'SCENTBIRD') return;

    if (!document.querySelector('input[data-kt="customerModalEmailField"]')) {
      document.querySelector('button[data-kt="customerTimelineEditCustomerProfileButton"]')?.click();
    }
    await waitForField('input[data-kt="customerModalEmailField"]');
    await new Promise(r => setTimeout(r, 200));

    const allFields = document.querySelectorAll('input[data-kt="customerModalEmailField"]');
    const existingValues = Array.from(allFields).map(f => f.value.trim().toLowerCase());
    const targetEmail = user.email.trim().toLowerCase();
    const matchIndex = existingValues.indexOf(targetEmail);

    if (matchIndex !== -1) {
      document.querySelector(`button[data-kt="customerModalEmailField_star_${matchIndex}"]`)?.click();
      await new Promise(r => setTimeout(r, 200));
    } else {
      const newIndex = allFields.length;
      document.querySelector('button[data-kt="customerModalEmailField_addRow"]')?.click();
      await new Promise(r => setTimeout(r, 300));

      const updatedFields = document.querySelectorAll('input[data-kt="customerModalEmailField"]');
      const newField = updatedFields[updatedFields.length - 1];
      if (newField) { setReactValue(newField, user.email); await new Promise(r => setTimeout(r, 200)); }

      document.querySelector(`button[data-kt="customerModalEmailField_star_${newIndex}"]`)?.click();
      await new Promise(r => setTimeout(r, 200));
    }

    const nameField = document.querySelector('input[data-kt="customerModalNameField"]');
    if (nameField && !nameField.value.trim()) {
      const full = toProperCase([user.firstName, user.lastName].filter(Boolean).join(' '));
      if (full) { setReactValue(nameField, full); await new Promise(r => setTimeout(r, 200)); }
    }

    document.querySelector('button[data-kt="modalFooterBasic_buttonPrimary"]')?.click();
  }

  // ══════════════════════════════════════════════════════════════════════════
  // PANELS
  // ══════════════════════════════════════════════════════════════════════════

  function removePanel(id) {
    document.getElementById(id)?.remove();
  }

  // ── Search Panel ──────────────────────────────────────────────────────────

  function showSearchPanel(prefill = '') {
    const panel = createPanel({ id: 'sb-search-panel', title: '🔍 CRM User Search', width: 400 });

    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:6px;margin-bottom:10px;';

    const input = document.createElement('input');
    input.type = 'text'; input.value = prefill;
    input.placeholder = 'Name, email, address…';
    input.style.cssText = `
      flex:1;padding:7px 10px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);
      background:#2a2a3e;color:#e2e8f0;font-size:13px;outline:none;
    `;

    const searchBtn = document.createElement('button');
    searchBtn.textContent = 'Search';
    searchBtn.style.cssText = `
      padding:7px 14px;border-radius:6px;border:none;background:#4f46e5;
      color:#fff;font-weight:700;font-size:13px;cursor:pointer;
    `;

    const filterBtn = document.createElement('button');
    filterBtn.textContent = 'Filter Active';
    filterBtn.style.cssText = `
      padding:7px 10px;border-radius:6px;border:1px solid rgba(110,231,183,0.3);background:transparent;
      color:#6ee7b7;font-weight:600;font-size:11px;cursor:pointer;white-space:nowrap;
    `;
    let filterActive = false;
    filterBtn.onclick = () => {
      filterActive = !filterActive;
      filterBtn.style.background = filterActive ? 'rgba(110,231,183,0.15)' : 'transparent';
      filterBtn.textContent = filterActive ? '✔ Active Only' : 'Filter Active';
      // Re-filter displayed results
      const cards = results.querySelectorAll('[data-sub-status]');
      cards.forEach(card => {
        if (filterActive && card.dataset.subStatus !== 'Active' && card.dataset.subStatus !== 'Unpaid' && card.dataset.subStatus !== 'OnHold') {
          card.style.display = 'none';
        } else {
          card.style.display = '';
        }
      });
    };

    const results = document.createElement('div');

    function doSearch() {
      const q = input.value.trim(); if (!q) return;
      results.innerHTML = `<div style="color:#94a3b8;padding:8px 0;">Searching…</div>`;
      searchCRM(q, (users, err) => {
        results.innerHTML = '';
        if (err) { results.innerHTML = `<div style="color:#fca5a5;padding:8px 0;">${err}</div>`; return; }
        if (!users.length) { results.innerHTML = `<div style="color:#94a3b8;padding:8px 0;">No users found.</div>`; return; }

        const count = document.createElement('div');
        count.style.cssText = 'color:#94a3b8;margin-bottom:8px;font-size:12px;';
        count.textContent = `${users.length} result(s)`;
        results.appendChild(count);

        users.forEach(user => {
          const card = document.createElement('div');
          card.dataset.subStatus = user.subscription?.status || '';
          card.style.cssText = `
            padding:10px;margin-bottom:8px;border-radius:8px;background:#2a2a3e;
            border:1px solid transparent;transition:0.15s;
          `;
          if (filterActive && card.dataset.subStatus !== 'Active' && card.dataset.subStatus !== 'Unpaid' && card.dataset.subStatus !== 'OnHold') card.style.display = 'none';
          const s = user.userAddress?.shipping;
          const addr = s ? [s.street1,s.city,s.region,s.postcode].filter(Boolean).join(', ') : '';
          const subSt = user.subscription?.status || '';
          const subColor = subSt === 'Active' ? '#6ee7b7' : subSt ? '#fca5a5' : '#64748b';
          const subLabel = subSt || 'No subscription';
          const isDrift = user.origin && user.origin !== 'SCENTBIRD';
          if (isDrift) card.style.opacity = '0.5';
          card.innerHTML = `
            <div style="display:flex;justify-content:space-between;align-items:baseline;">
              <div style="font-weight:700;">${user.firstName||''} ${user.lastName||''} ${isDrift ? '<span style="font-size:10px;color:#f59e0b;font-weight:600;background:rgba(245,158,11,0.15);padding:1px 5px;border-radius:3px;margin-left:4px;">' + user.origin + '</span>' : ''}</div>
              <div style="font-size:11px;font-weight:600;color:${subColor};">${subLabel}</div>
            </div>
            <div style="color:#94a3b8;margin-top:2px;">${user.email}</div>
            <div style="color:#64748b;font-size:11px;margin-top:2px;">ID: ${user.id}</div>
            ${addr ? `<div style="color:#64748b;font-size:11px;margin-top:2px;">${addr}</div>` : ''}
            <div style="margin-top:8px;display:flex;gap:6px;">
              <button class="sb-apply" style="padding:4px 12px;border-radius:5px;border:none;background:#4f46e5;color:#fff;font-size:12px;font-weight:600;cursor:pointer;">Apply to Kustomer</button>
              <a class="sb-profile" href="https://crm.scentbird.com/user/${user.id}/profile/subscription" target="_blank" style="padding:4px 12px;border-radius:5px;border:1px solid rgba(255,255,255,0.15);background:transparent;color:#e2e8f0;font-size:12px;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;">Open Profile</a>
            </div>
          `;
          card.onmouseenter = () => card.style.borderColor = 'rgba(79,70,229,0.5)';
          card.onmouseleave = () => card.style.borderColor = 'transparent';

          const applyBtn = card.querySelector('.sb-apply');
          if (isDrift) {
            applyBtn.disabled = true;
            applyBtn.style.opacity = '0.4';
            applyBtn.style.cursor = 'not-allowed';
            applyBtn.title = 'Drift account — cannot apply';
          } else {
            applyBtn.onclick = async (e) => {
              const btn = e.target; btn.textContent = '⏳ Applying…'; btn.disabled = true;
              await applyUser(user);
              await new Promise(r => setTimeout(r, 500));
              cachedCustomerCtx = { email: user.email.trim().toLowerCase(), user, _kustomerId: getCustomerIdFromURL() };
              btn.textContent = '✔ Done'; btn.style.background = '#059669';
              setTimeout(() => removePanel('sb-search-panel'), 1500);
              // Auto-trigger Fill Name
              const fillBtn = document.getElementById('sb-fill-name-btn');
              if (fillBtn && !fillBtn.disabled) fillBtn.click();
            };
          }

          // Last order preview (skip Drift)
          const orderEl = document.createElement('div');
          orderEl.style.cssText = 'margin-top:8px;font-size:11px;color:#64748b;';
          card.appendChild(orderEl);

          if (!isDrift) {
            orderEl.textContent = '⏳ Loading orders…';
            fetchLastOrders(user.id, (orders) => {
              orderEl.innerHTML = '';
              if (!orders?.length) return;
              const mainOrder = orders.find(o => o.type !== 'REPLACEMENT' && o.type !== 'REPLACEMENT_ORDER') || orders[0];
              orderEl.appendChild(renderOrderBlock(mainOrder, true));
            });
          }

          results.appendChild(card);
        });
      });
    }

    searchBtn.onclick = doSearch;
    input.addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });
    row.appendChild(input); row.appendChild(searchBtn); row.appendChild(filterBtn);
    panel.appendChild(row); panel.appendChild(results);
    if (prefill) doSearch();
    setTimeout(() => input.focus(), 50);
  }

  // ── Token Panel ───────────────────────────────────────────────────────────

  function showTokenPanel() {
    const panel = createPanel({ id: 'sb-token-panel', title: '🔑 CRM Bearer Token', width: 380 });
    // Remove the default maxHeight since this is a short form
    panel.style.maxHeight = 'none';

    const hint = document.createElement('div');
    hint.style.cssText = 'color:#94a3b8;font-size:12px;margin-bottom:8px;';
    hint.textContent = 'Saved to localStorage — never hardcoded.';
    panel.appendChild(hint);

    const textarea = document.createElement('textarea');
    textarea.rows = 5;
    textarea.placeholder = BEARER_TOKEN ? '(token set — paste new to replace)' : 'eyJra…';
    textarea.style.cssText = `
      width:100%;padding:8px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);
      background:#2a2a3e;color:#e2e8f0;font-size:11px;font-family:monospace;
      resize:vertical;box-sizing:border-box;outline:none;
    `;
    panel.appendChild(textarea);

    const btnRow = document.createElement('div');
    btnRow.style.cssText = 'display:flex;gap:8px;margin-top:8px;';

    const saveBtn = document.createElement('button');
    saveBtn.textContent = 'Save';
    saveBtn.style.cssText = 'flex:1;padding:7px;border-radius:6px;border:none;background:#4f46e5;color:#fff;font-weight:700;font-size:13px;cursor:pointer;';

    const clearBtn = document.createElement('button');
    clearBtn.textContent = 'Clear';
    clearBtn.style.cssText = 'padding:7px 12px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);background:transparent;color:#fca5a5;font-size:13px;cursor:pointer;';

    const statusEl = makeStatusEl();
    statusEl.style.color = '#6ee7b7';

    saveBtn.onclick = () => {
      const val = textarea.value.trim();
      if (!val) { statusEl.style.color = '#fca5a5'; statusEl.textContent = 'Paste a token first.'; return; }
      BEARER_TOKEN = val; localStorage.setItem('sb_crm_token', val);
      statusEl.style.color = '#6ee7b7';
      statusEl.textContent = '✔ Saved.';
      setTimeout(() => removePanel('sb-token-panel'), 1000);
    };
    clearBtn.onclick = () => {
      BEARER_TOKEN = ''; localStorage.removeItem('sb_crm_token');
      textarea.value = '';
      statusEl.style.color = '#fca5a5';
      statusEl.textContent = 'Cleared.';
    };

    btnRow.appendChild(saveBtn); btnRow.appendChild(clearBtn);
    panel.appendChild(btnRow);
    panel.appendChild(statusEl);
  }

  // ── Last Orders Panel ─────────────────────────────────────────────────────

  function showLastOrderPanel(btn) {
    removePanel('sb-order-panel');

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;
      loadCustomerOrders(user.id, (orders) => {
        const panel = createPanel({ id: 'sb-order-panel', title: '📦 Last Orders', width: 380 });
        panel.appendChild(renderUserRow(user));

        if (!orders?.length) {
          const empty = document.createElement('div');
          empty.style.color = '#94a3b8';
          empty.textContent = 'No orders found.';
          panel.appendChild(empty);
          return;
        }

        const subOrders = orders.filter(o => o.type === 'SUBSCRIPTION' || !o.type || o.type === 'MANUAL');
        const orderCtx = { user, totalOrders: subOrders.length };

        // Group by month+year
        const groups = [];
        const groupIndex = {};
        orders.forEach(o => {
          const key = `${o.year}-${String(o.month).padStart(2,'0')}`;
          if (!groupIndex[key]) {
            groupIndex[key] = { key, year: o.year, month: o.month, mains: [], replacements: [] };
            groups.push(groupIndex[key]);
          }
          const g = groupIndex[key];
          if (o.type === 'REPLACEMENT' || o.type === 'REPLACEMENT_ORDER') {
            g.replacements.push(o);
          } else {
            g.mains.push(o);
          }
        });

        groups.sort((a, b) => (b.year - a.year) || (b.month - a.month));

        groups.forEach((g, gi) => {
          if (gi > 0) {
            const sep = document.createElement('div');
            sep.style.cssText = 'height:1px;background:rgba(255,255,255,0.06);margin:10px 0;';
            panel.appendChild(sep);
          }

          g.mains.forEach((o, oi) => {
            if (oi > 0) {
              const s = document.createElement('div');
              s.style.cssText = 'height:1px;background:rgba(255,255,255,0.04);margin:8px 0;';
              panel.appendChild(s);
            }
            panel.appendChild(renderOrderBlock(o, false, orderCtx, false));
          });

          g.replacements.forEach(r => {
            const nest = document.createElement('div');
            nest.style.cssText = 'margin-left:12px;margin-top:6px;padding-left:8px;border-left:1px solid rgba(124,58,237,0.3);';
            const replLabel = document.createElement('div');
            replLabel.style.cssText = 'font-size:10px;color:#7c3aed;font-weight:600;letter-spacing:0.05em;margin-bottom:4px;';
            replLabel.textContent = 'REPLACEMENT';
            nest.appendChild(replLabel);
            nest.appendChild(renderOrderBlock(r, false, orderCtx, true));
            panel.appendChild(nest);
          });
        });

        // Custom Shipment button
        const csSep = document.createElement('div');
        csSep.style.cssText = 'height:1px;background:rgba(255,255,255,0.07);margin:12px 0;';
        panel.appendChild(csSep);

        const csBtn = document.createElement('button');
        csBtn.textContent = '📦 Custom Shipment';
        csBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid #4f46e5;background:transparent;
          color:#a5b4fc;font-weight:600;font-size:12px;cursor:pointer;box-sizing:border-box;`;
        csBtn.onmouseenter = () => csBtn.style.background = 'rgba(79,70,229,0.1)';
        csBtn.onmouseleave = () => csBtn.style.background = 'transparent';
        csBtn.onclick = () => showCustomShipmentForm(user);
        panel.appendChild(csBtn);
      });
    });
  }

  // ── Charges Panel ─────────────────────────────────────────────────────────

  function showChargesPanel(btn) {
    removePanel('sb-charges-panel');

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;
      loadCustomerCharges(user.id, (result) => {
        const panel = createPanel({ id: 'sb-charges-panel', title: '💳 Recent Charges', width: 420 });
        panel.appendChild(renderUserRow(user));

        if (!result) {
          const errEl = document.createElement('div');
          errEl.style.color = '#fca5a5';
          errEl.textContent = 'Failed to load charges.';
          panel.appendChild(errEl);
          return;
        }

        const { success, failed } = result;

        // Failed charges
        if (failed?.length) {
          const failHdr = document.createElement('div');
          failHdr.style.cssText = 'color:#fca5a5;font-size:11px;font-weight:600;margin-bottom:4px;';
          failHdr.textContent = `\u274C ${failed.length} failed charge attempt${failed.length > 1 ? 's' : ''}`;
          panel.appendChild(failHdr);

          const failWrap = document.createElement('div');
          failWrap.style.cssText = 'background:#2a1a1a;border-radius:6px;padding:6px 10px;margin-bottom:12px;';
          failed.forEach(entry => {
            const c = entry.charge;
            const desc = c?.cashbirdDetails?.invoice?.lineItems?.[0]?.description || c?.category || '\u2014';
            const row = document.createElement('div');
            row.style.cssText = 'color:#94a3b8;font-size:11px;margin-bottom:3px;display:flex;gap:8px;align-items:baseline;';
            row.innerHTML = `
              <span style="color:#64748b;white-space:nowrap;">${chargeDateLabel(c?.paymentDate || entry.date)}</span>
              <span style="color:#fca5a5;white-space:nowrap;">${fmtMoney((c?.totalPrice || 0) * 100)}</span>
              <span style="color:#94a3b8;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;">${desc}</span>
            `;
            failWrap.appendChild(row);
          });
          panel.appendChild(failWrap);
        }

        // Successful charges
        const allSuccess = result.all || (result.success ? [result.success] : []);
        if (!allSuccess.length) {
          const noEl = document.createElement('div');
          noEl.style.cssText = 'color:#94a3b8;font-size:12px;';
          noEl.textContent = 'No successful charges found.';
          panel.appendChild(noEl);
          return;
        }

        const TYPE_LABELS = {
          FIRST_SUBSCRIPTION:       'SUB (initial)',
          RECURRENT_CHARGE:         'SUB (renewal)',
          RESUBSCRIBE:              'SUB (resubscribe)',
          UPGRADE_PARTIAL:          'Partial upgrade',
          UPGRADE_FULL:             'Full upgrade',
          PREMIUM_PRODUCT_UPCHARGE: 'Premium upcharge',
          PERFUME_CASE_UPCHARGE:    'Case',
          SHIPPING_CREDIT:          'Shipping',
          SHIPPING_UPCHARGE:        'Shipping',
          SAMPLES_UPCHARGE:         'Samples',
          SHIPPING_PURCHASE:        'Shipping',
          ECOMMERCE_PURCHASE:       'Ecommerce',
        };

        const orderedSuccess = [...allSuccess].reverse();
        const allExplanations = [];

        orderedSuccess.forEach((entry, idx) => {
          const c = entry.charge;
          const pm = c?.cashbirdDetails?.paymentMethod;
          const items = c?.cashbirdDetails?.invoice?.lineItems || [];
          const billingDate = c?.paymentDate || entry.date;
          const currency = detectCurrency(items);

          if (idx > 0) {
            const sep = document.createElement('div');
            sep.style.cssText = 'height:1px;background:rgba(255,255,255,0.06);margin:10px 0;';
            panel.appendChild(sep);
          }

          const cHdr = document.createElement('div');
          cHdr.style.cssText = 'display:flex;align-items:baseline;justify-content:space-between;margin-bottom:4px;';
          const lineItemsForRefundCalc = c?.cashbirdDetails?.invoice?.lineItems || [];
          const refundAmt = lineItemsForRefundCalc.reduce((sum, li) => sum + (li.refundedAmount || 0), 0);
          const totalRefunded = refundAmt >= Math.round(c.totalPrice * 100);
          const refundTag = refundAmt > 0
            ? ` <span style="font-size:11px;font-weight:600;color:${totalRefunded ? '#6ee7b7' : '#f59e0b'};">&#8617; ${fmtMoney(refundAmt, currency)} refunded</span>`
            : '';
          cHdr.innerHTML = `
            <span style="font-weight:700;font-size:13px;color:#6ee7b7;">\u2714 ${fmtMoney(c.totalPrice * 100, currency)}</span>${refundTag}
            <span style="color:#64748b;font-size:11px;">${chargeDateLabel(billingDate)}</span>
          `;
          panel.appendChild(cHdr);

          if (pm) {
            const pmRow = document.createElement('div');
            pmRow.style.cssText = 'color:#64748b;font-size:11px;margin-bottom:8px;';
            pmRow.textContent = `${pm.type} \u00b7 ${pm.methodName}`;
            panel.appendChild(pmRow);
          }

          const lineWrap = document.createElement('div');
          lineWrap.style.cssText = 'background:#2a2a3e;border-radius:6px;padding:8px 10px;margin-bottom:8px;';
          const visibleItems = items.filter(i => i.type !== 'FAILED_CHARGE');
          if (visibleItems.length) {
            visibleItems.forEach(item => {
              const label = TYPE_LABELS[item.type] || item.type;
              const nameStr = (item.type === 'PREMIUM_PRODUCT_UPCHARGE' && item.description)
                ? `${label} \u2014 ${item.description}` : label;
              const discPart = item.discount
                ? `<span style="color:#6ee7b7;margin-right:4px;">-${fmtMoney(item.discount, currency)}</span>` : '';
              const row = document.createElement('div');
              row.style.cssText = 'display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px;font-size:12px;';
              row.innerHTML = `
                <span style="color:#cbd5e1;">${nameStr}</span>
                <span style="color:#94a3b8;white-space:nowrap;margin-left:8px;">${discPart}${fmtMoney(item.price, currency)} + ${fmtMoney(item.tax, currency)} tax = <strong>${fmtMoney(item.total, currency)}</strong></span>
              `;
              lineWrap.appendChild(row);
            });
          } else {
            const cat = c?.category || 'CHARGE';
            const fallback = document.createElement('div');
            fallback.style.cssText = 'color:#64748b;font-size:11px;font-style:italic;';
            fallback.textContent = cat.replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, l => l.toUpperCase());
            lineWrap.appendChild(fallback);
          }
          panel.appendChild(lineWrap);

          const explanation = buildBillExplanation(entry);
          if (explanation) allExplanations.push(explanation);
        });

        // Merged bill explanation
        if (allExplanations.length) {
          const mergedText = allExplanations.join('\n');
          const expHdr = document.createElement('div');
          expHdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-top:4px;margin-bottom:4px;';
          expHdr.innerHTML = '<span style="font-size:11px;color:#64748b;font-weight:600;letter-spacing:0.05em;">BILL EXPLANATION</span>';
          expHdr.appendChild(makeCopyBtn(mergedText, '\uD83D\uDCCB Copy'));
          panel.appendChild(expHdr);

          const expBox = document.createElement('div');
          expBox.style.cssText = `
            background:#0f172a;border-radius:6px;padding:10px 12px;
            font-size:12px;color:#cbd5e1;white-space:pre-wrap;line-height:1.7;
            border:1px solid rgba(255,255,255,0.06);
          `;
          expBox.textContent = mergedText;
          panel.appendChild(expBox);
        }
      });
    });
  }

  // ── Queue Panel ────────────────────────────────────────────────────────────

  function showQueuePanel(btn) {
    if (document.getElementById('sb-queue-panel')) {
      removePanel('sb-queue-panel'); return;
    }

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;

      const panel = createPanel({ id: 'sb-queue-panel', title: '📋 Queue', width: 360 });
      panel.appendChild(renderUserRow(user));

      const loadingEl = document.createElement('div');
      loadingEl.style.cssText = 'color:#94a3b8;font-size:12px;';
      loadingEl.textContent = 'Loading queue...';
      panel.appendChild(loadingEl);

      fetchQueue(user.id, (blocks) => {
        loadingEl.remove();

        if (!blocks || !blocks.length) {
          const empty = document.createElement('div');
          empty.style.cssText = 'color:#94a3b8;font-size:12px;';
          empty.textContent = blocks ? 'Queue is empty.' : 'Failed to load queue.';
          panel.appendChild(empty);
          return;
        }

        // Filter to blocks with products only
        const filledBlocks = blocks.filter(b => b.productList?.length > 0);

        if (!filledBlocks.length) {
          const empty = document.createElement('div');
          empty.style.cssText = 'color:#94a3b8;font-size:12px;';
          empty.textContent = 'No products in queue.';
          panel.appendChild(empty);
          return;
        }

        // Collect all product names for copy-all
        const allProducts = [];

        // Compute flatIndex for each product across unprocessed blocks
        // flatIndex is sequential across ALL unprocessed products in the full blockList, not just filledBlocks
        let flatIdx = 0;
        blocks.forEach(block => {
          if (!block._flatIndices) block._flatIndices = [];
          (block.productList || []).forEach(() => {
            if (block.processedSubscriptionOrder) {
              block._flatIndices.push(-1);
            } else {
              block._flatIndices.push(flatIdx++);
            }
          });
        });

        filledBlocks.forEach((block, idx) => {
          if (idx > 0) {
            const sep = document.createElement('div');
            sep.style.cssText = 'height:1px;background:rgba(255,255,255,0.06);margin:8px 0;';
            panel.appendChild(sep);
          }

          const blockEl = document.createElement('div');
          blockEl.style.cssText = 'background:#0f172a;border-radius:8px;padding:10px 12px;';

          // Month header
          const [yearStr, monthStr] = (block.yearMonth || '').split('-');
          const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
          const monthLabel = monthNames[parseInt(monthStr, 10) - 1] + ' ' + yearStr;

          const hdr = document.createElement('div');
          hdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;';

          const monthEl = document.createElement('span');
          monthEl.style.cssText = 'color:#94a3b8;font-size:11px;font-weight:600;';
          monthEl.textContent = monthLabel;
          hdr.appendChild(monthEl);

          // Order status if processed
          const orderStatus = block.processedSubscriptionOrder?.status;
          if (orderStatus) {
            const statusColor = (orderStatus === 'SHIPPED' || orderStatus === 'DELIVERED') ? '#6ee7b7'
              : orderStatus === 'PENDING' ? '#f59e0b' : '#94a3b8';
            const statusEl = document.createElement('span');
            statusEl.style.cssText = `font-size:10px;font-weight:600;color:${statusColor};`;
            statusEl.textContent = orderStatus;
            hdr.appendChild(statusEl);
          }

          blockEl.appendChild(hdr);

          // Products
          block.productList.forEach((prod, pIdx) => {
            const pi = prod.tradingItem?.productInfo;
            if (!pi) return;
            const prodName = `${pi.name} by ${pi.brand}`;
            allProducts.push(prodName);

            const row = document.createElement('div');
            row.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:3px;';

            const nameEl = document.createElement('span');
            nameEl.style.cssText = 'color:#cbd5e1;font-size:12px;flex:1;margin-right:6px;';
            nameEl.textContent = prodName;
            row.appendChild(nameEl);

            // Upcharge indicator
            const upcharge = prod.upchargePrice || pi.upchargePrice;
            if (upcharge && upcharge > 0) {
              const upEl = document.createElement('span');
              upEl.style.cssText = 'color:#f59e0b;font-size:10px;font-weight:600;white-space:nowrap;margin-right:4px;';
              upEl.textContent = '+$' + (upcharge / 100).toFixed(0);
              row.appendChild(upEl);
            }

            row.appendChild(makeCopyBtn(prodName, '📋'));

            // Delete button (only for unprocessed items)
            const prodFlatIdx = block._flatIndices[pIdx];
            if (prodFlatIdx >= 0) {
              const delBtn = document.createElement('button');
              delBtn.textContent = '✕';
              delBtn.title = 'Remove from queue';
              delBtn.style.cssText = 'background:none;border:none;cursor:pointer;color:#64748b;font-size:13px;padding:0 3px;flex-shrink:0;';
              delBtn.onmouseenter = () => delBtn.style.color = '#fca5a5';
              delBtn.onmouseleave = () => delBtn.style.color = '#64748b';
              delBtn.onclick = () => {
                delBtn.disabled = true;
                delBtn.textContent = '⏳';
                delBtn.style.color = '#94a3b8';

                const doDelete = () => {
                  customerApiDelete(prodFlatIdx, (data, err) => {
                    if (err) {
                      delBtn.textContent = '✕';
                      delBtn.style.color = '#fca5a5';
                      delBtn.disabled = false;
                      return;
                    }
                    // Refresh the panel with fresh data
                    removePanel('sb-queue-panel');
                    showQueuePanel(btn);
                  });
                };

                // Ensure customer API auth first
                if (_customerApiReady) {
                  doDelete();
                } else {
                  ensureCustomerAuth((ok) => {
                    if (!ok) {
                      delBtn.textContent = '✕';
                      delBtn.style.color = '#fca5a5';
                      delBtn.disabled = false;
                      return;
                    }
                    doDelete();
                  });
                }
              };
              row.appendChild(delBtn);
            }

            blockEl.appendChild(row);
          });

          panel.appendChild(blockEl);
        });

        // Copy all button at bottom
        if (allProducts.length > 1) {
          const copyAllRow = document.createElement('div');
          copyAllRow.style.cssText = 'display:flex;justify-content:flex-end;margin-top:8px;';
          const copyAllBtn = document.createElement('button');
          copyAllBtn.textContent = '📋 Copy All';
          copyAllBtn.style.cssText = 'background:none;border:1px solid rgba(255,255,255,0.1);border-radius:5px;color:#818cf8;font-size:11px;padding:4px 10px;cursor:pointer;';
          copyAllBtn.onmouseenter = () => copyAllBtn.style.background = 'rgba(99,102,241,0.1)';
          copyAllBtn.onmouseleave = () => copyAllBtn.style.background = 'none';
          copyAllBtn.onclick = () => {
            navigator.clipboard.writeText(allProducts.join('\n'));
            copyAllBtn.textContent = '✔ Copied';
            setTimeout(() => copyAllBtn.textContent = '📋 Copy All', 1500);
          };
          copyAllRow.appendChild(copyAllBtn);
          panel.appendChild(copyAllRow);
        }
      });
    });
  }

  // ── Custom Shipment Form ────────────────────────────────────────────────

  const CS_REASONS = [
    { value: 'SKIPPED_SHIPMENT', label: 'Skipped Shipment' },
    { value: 'BACKORDER',        label: 'Backorder' },
    { value: 'MULTIPLE_CHARGE',  label: 'Multiple Charge' },
    { value: 'PROMO_OFFER',      label: 'Promo Offer' },
  ];

  function showCustomShipmentForm(user) {
    const existingForm = document.getElementById('sb-cs-form');
    if (existingForm) { existingForm.remove(); document.getElementById('sb-product-search-panel')?.remove(); return; }

    const { overlay, box: form } = createModal({
      id: 'sb-cs-form',
      title: '📦 Custom Shipment',
      width: 400,
    });

    form.appendChild(renderUserRow(user));

    // Reason
    form.appendChild(makeLabel('REASON'));
    const reasonSel = makeSelect(CS_REASONS, 'SKIPPED_SHIPMENT');
    form.appendChild(reasonSel);

    // BOM
    form.appendChild(makeLabel('ORDER TYPE (BOM)'));
    const bomSel = makeSelect(REPLACEMENT_BOMS, 'FEMALE_RECURRENT_MONTH_SET');
    form.appendChild(bomSel);

    // Month/Year picker
    form.appendChild(makeLabel('SHIPMENT MONTH'));
    const monthYearRow = document.createElement('div');
    monthYearRow.style.cssText = 'display:flex;gap:6px;margin-bottom:8px;';

    const now = new Date();
    const monthOpts = ['January','February','March','April','May','June','July','August','September','October','November','December']
      .map((m, i) => ({ value: String(i + 1), label: m }));
    const monthSel = makeSelect(monthOpts, String(now.getMonth() + 1));
    monthSel.style.flex = '1';

    const yearOpts = [now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1]
      .map(y => ({ value: String(y), label: String(y) }));
    const yearSel = makeSelect(yearOpts, String(now.getFullYear()));
    yearSel.style.width = '90px';

    monthYearRow.appendChild(monthSel);
    monthYearRow.appendChild(yearSel);
    form.appendChild(monthYearRow);

    // Products
    form.appendChild(makeLabel('PRODUCTS'));
    const itemsWrap = document.createElement('div');
    itemsWrap.style.cssText = 'background:#0f172a;border-radius:6px;padding:8px 10px;min-height:30px;';
    const itemChecks = [];

    const noItems = document.createElement('div');
    noItems.style.cssText = 'color:#64748b;font-size:11px;font-style:italic;';
    noItems.textContent = 'No products added. Use Search Product below.';
    itemsWrap.appendChild(noItems);
    form.appendChild(itemsWrap);

    // Product search button
    const searchRow = document.createElement('div');
    searchRow.style.cssText = 'margin-top:8px;';
    const searchProductBtn = document.createElement('button');
    searchProductBtn.textContent = '🔍 Search Product';
    searchProductBtn.style.cssText = 'padding:5px 10px;border-radius:5px;border:1px solid #818cf8;background:transparent;color:#818cf8;font-weight:600;font-size:11px;cursor:pointer;';
    searchProductBtn.onmouseenter = () => searchProductBtn.style.background = 'rgba(129,140,248,0.15)';
    searchProductBtn.onmouseleave = () => searchProductBtn.style.background = 'transparent';
    searchProductBtn.onclick = () => {
      if (noItems.parentNode) noItems.remove();
      showProductSearchPanel(itemChecks, itemsWrap, () => csUpchargeUI.refresh());
    };
    searchRow.appendChild(searchProductBtn);
    form.appendChild(searchRow);

    // Additional comment
    form.appendChild(makeLabel('ADDITIONAL COMMENT (OPTIONAL)'));
    const commentInput = document.createElement('textarea');
    commentInput.placeholder = 'Add a note...';
    commentInput.rows = 2;
    commentInput.style.cssText = 'width:100%;background:#0f172a;color:#e2e8f0;border:1px solid rgba(255,255,255,0.1);border-radius:5px;padding:6px 8px;font-size:12px;box-sizing:border-box;resize:vertical;font-family:Arial,sans-serif;';
    form.appendChild(commentInput);

    // Credit info (loaded async)
    const creditInfo = document.createElement('div');
    creditInfo.style.cssText = 'margin-top:12px;font-size:11px;color:#94a3b8;';
    creditInfo.textContent = 'Loading credit info...';
    form.appendChild(creditInfo);

    let _subscriptionId = null;
    let _availableCredits = 0;

    // Fetch subscription for credits
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName: 'userDetailsById', query: USER_DETAILS_QUERY, variables: { id: user.id } }),
      onload(res) {
        try {
          const json = JSON.parse(res.responseText);
          const list = json?.data?.userById?.data?.subscriptionList || [];
          const sub = list.find(s => s.status === 'Active' || s.subscribed) || list[0];
          if (sub) {
            _subscriptionId = sub.id;
            const credits = sub.credits || [];
            _availableCredits = credits.filter(c => c.status === 'NEW').length;
            creditInfo.textContent = _availableCredits + ' NEW credit(s) available';
            creditInfo.style.color = _availableCredits > 0 ? '#6ee7b7' : '#f59e0b';
          } else {
            creditInfo.textContent = 'No subscription found';
            creditInfo.style.color = '#fca5a5';
          }
        } catch(e) {
          creditInfo.textContent = 'Failed to load credits';
          creditInfo.style.color = '#fca5a5';
        }
      }
    });

    // Status
    const statusEl = makeStatusEl();
    form.appendChild(statusEl);

    // Confirm button
    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'Create Custom Shipment';
    confirmBtn.style.cssText = `
      margin-top:14px;width:100%;padding:9px;border-radius:6px;
      border:none;background:#4f46e5;color:#fff;font-weight:700;
      font-size:13px;cursor:pointer;box-sizing:border-box;
    `;
    confirmBtn.onmouseenter = () => confirmBtn.style.background = '#4338ca';
    confirmBtn.onmouseleave = () => confirmBtn.style.background = '#4f46e5';

    // Upcharge manager
    const csUpchargeUI = createUpchargeManager(user.id, itemChecks, confirmBtn);
    form.appendChild(csUpchargeUI.el);

    confirmBtn.onclick = () => {
      const selectedItems = itemChecks
        .filter(({ chk }) => chk.checked)
        .map(({ item }) => ({
          id: null,
          starterSet: false,
          product: { id: item.product?.id || item.id },
        }));

      if (!selectedItems.length) {
        statusEl.style.color = '#fca5a5';
        statusEl.textContent = 'Add at least one product.';
        return;
      }

      const csMonth = parseInt(monthSel.value, 10);
      const csYear = parseInt(yearSel.value, 10);
      const creditsNeeded = selectedItems.length;
      const creditsToAdd = Math.max(0, creditsNeeded - _availableCredits);

      // Show confirmation
      const { overlay: confOverlay, box: confBox } = createModal({
        id: 'sb-cs-confirm',
        title: '📦 Confirm Custom Shipment',
        width: 340,
      });

      const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      const reasonLabel = CS_REASONS.find(r => r.value === reasonSel.value)?.label || reasonSel.value;

      const detailsEl = document.createElement('div');
      detailsEl.style.cssText = 'background:#0f172a;border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:12px;line-height:1.8;color:#cbd5e1;';

      let detailsHtml = `
        <div><span style="color:#64748b;">Month</span>&nbsp;&nbsp;&nbsp;&nbsp; ${monthNames[csMonth - 1]} ${csYear}</div>
        <div><span style="color:#64748b;">Reason</span>&nbsp;&nbsp;&nbsp; ${reasonLabel}</div>
        <div><span style="color:#64748b;">Products</span>&nbsp; ${selectedItems.length}</div>
      `;
      if (creditsToAdd > 0) {
        detailsHtml += `<div style="margin-top:6px;padding-top:6px;border-top:1px solid rgba(255,255,255,0.07);color:#f59e0b;font-weight:600;">${creditsToAdd} credit(s) will be added</div>`;
      }
      detailsEl.innerHTML = detailsHtml;
      confBox.appendChild(detailsEl);

      // Product list
      const prodList = document.createElement('div');
      prodList.style.cssText = 'margin-bottom:12px;font-size:11px;color:#94a3b8;';
      itemChecks.filter(({ chk }) => chk.checked).forEach(({ item }) => {
        const pi = item.product?.productInfo || item.productInfo;
        const row = document.createElement('div');
        row.style.cssText = 'margin-bottom:2px;';
        row.textContent = '• ' + (pi ? pi.name + ' by ' + pi.brand : 'Unknown product');
        prodList.appendChild(row);
      });
      confBox.appendChild(prodList);

      const confStatus = makeStatusEl();
      confBox.appendChild(confStatus);

      confBox.appendChild(makeDialogButtons({
        confirmLabel: 'Confirm',
        confirmColor: '#4f46e5',
        onCancel: () => confOverlay.remove(),
        onConfirm: (cBtn, cancelBtn) => {
          cBtn.disabled = true;
          cancelBtn.disabled = true;
          cBtn.textContent = 'Processing...';

          const doShipment = () => {
            confStatus.textContent = 'Creating custom shipment...';

            const variables = {
              input: {
                zendesk: window.location.href,
                reason: reasonSel.value,
                bom: bomSel.value,
                customerId: user.id,
                addressConfirmed: false,
                picturesProvided: false,
                sourceOrder: {
                  id: null,
                  year: csYear,
                  month: csMonth,
                  subscriptionId: _subscriptionId,
                  orderItems: selectedItems,
                },
              },
            };

            gqlMutate('customShipmentSave', CUSTOM_SHIPMENT_MUTATION, variables,
              (data, err) => {
                const respErr = data?.customShipmentSave?.error;
                const errMsg = err || respErr?.message;
                if (errMsg) {
                  confStatus.style.color = '#fca5a5';
                  confStatus.textContent = '\u2718 ' + errMsg;
                  cBtn.disabled = false;
                  cancelBtn.disabled = false;
                  cBtn.textContent = 'Confirm';
                  return;
                }

                confStatus.style.color = '#94a3b8';
                confStatus.textContent = 'Finding automated comment...';

                const csOrderNumber = data?.customShipmentSave?.data?.orderNumber;
                const monthLabel = monthNames[csMonth - 1] + ' ' + csYear;
                const commentParts = ['Custom Shipment - ' + reasonLabel + ' - ' + monthLabel];
                if (creditsToAdd > 0) commentParts.push(creditsToAdd + ' Credit(s) Added for CS to ship');
                const extra = commentInput.value.trim();
                if (extra) commentParts.push(extra);
                const commentText = commentParts.join('\n');

                fetchComments(user.id, (comments) => {
                  const autoComment = comments.find(c =>
                    c.comment && csOrderNumber && c.comment.includes(csOrderNumber)
                  );

                  const postComment = () => {
                    confStatus.textContent = 'Posting comment...';
                    gqlMutate('createUserComment', CREATE_COMMENT_MUTATION,
                      { input: { userId: user.id, comment: commentText, zendeskUrl: window.location.href } },
                      () => {
                        confStatus.style.color = '#6ee7b7';
                        confStatus.textContent = '\u2714 Custom shipment created!';
                        cBtn.textContent = '\u2714 Done';
                        confirmBtn.disabled = true;
                        confirmBtn.textContent = '\u2714 Created';
                        confirmBtn.style.background = 'rgba(110,231,183,0.15)';
                        confirmBtn.style.color = '#6ee7b7';
                        setTimeout(() => {
                          confOverlay.remove();
                          overlay.remove();
                          document.getElementById('sb-product-search-panel')?.remove();
                        }, 1800);
                      }
                    );
                  };

                  if (autoComment) {
                    confStatus.textContent = 'Removing automated comment...';
                    gqlMutate('deleteUserComment', DELETE_COMMENT_MUTATION, { id: autoComment.id },
                      () => postComment()
                    );
                  } else {
                    postComment();
                  }
                });
              }
            );
          };

          // Step 1: Add credits if needed
          if (creditsToAdd > 0 && _subscriptionId) {
            confStatus.textContent = 'Adding ' + creditsToAdd + ' credit(s)...';
            gqlMutate('addSupportCredits', ADD_CREDITS_MUTATION,
              { input: { subscriptionId: _subscriptionId, credits: creditsToAdd, comment: 'Added for Custom Shipment' } },
              (data, err) => {
                const respErr = data?.addSupportCredits?.error;
                const errMsg = err || respErr?.message;
                if (errMsg) {
                  confStatus.style.color = '#fca5a5';
                  confStatus.textContent = '\u2718 Credit add failed: ' + errMsg;
                  cBtn.disabled = false;
                  cancelBtn.disabled = false;
                  cBtn.textContent = 'Confirm';
                  return;
                }
                doShipment();
              }
            );
          } else {
            doShipment();
          }
        },
      }));
    };

    form.appendChild(confirmBtn);
  }

  // ── Edit Customer Panel ──────────────────────────────────────────────────

  function showEditCustomerPanel(btn) {
    if (document.getElementById('sb-edit-customer-panel')) {
      removePanel('sb-edit-customer-panel'); return;
    }

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;

      const panel = createPanel({ id: 'sb-edit-customer-panel', title: '✎ Edit Customer', width: 340 });
      panel.appendChild(renderUserRow(user));

      // Email section
      panel.appendChild(makeLabel('EMAIL'));
      const emailRow = document.createElement('div');
      emailRow.style.cssText = 'display:flex;gap:6px;margin-bottom:12px;';

      const emailInput = document.createElement('input');
      emailInput.type = 'email';
      emailInput.value = user.email || '';
      emailInput.style.cssText = `
        flex:1;padding:7px 10px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);
        background:#0f172a;color:#e2e8f0;font-size:12px;outline:none;box-sizing:border-box;
      `;

      const updateBtn = document.createElement('button');
      updateBtn.textContent = 'Update';
      updateBtn.style.cssText = 'padding:7px 14px;border-radius:6px;border:none;background:#4f46e5;color:#fff;font-weight:700;font-size:12px;cursor:pointer;';
      updateBtn.onmouseenter = () => updateBtn.style.background = '#4338ca';
      updateBtn.onmouseleave = () => updateBtn.style.background = '#4f46e5';

      const emailStatus = makeStatusEl();

      updateBtn.onclick = () => {
        const newEmail = emailInput.value.trim();
        if (!newEmail || !newEmail.includes('@')) {
          emailStatus.style.color = '#fca5a5';
          emailStatus.textContent = 'Enter a valid email.';
          return;
        }
        if (newEmail === user.email) {
          emailStatus.style.color = '#f59e0b';
          emailStatus.textContent = 'Email is unchanged.';
          return;
        }
        updateBtn.disabled = true;
        updateBtn.textContent = '⏳';
        emailStatus.textContent = '';

        gqlMutate('userChangeEmail', CHANGE_EMAIL_MUTATION,
          { input: { userId: user.id, email: newEmail } },
          (data, err) => {
            const respErr = data?.userChangeEmail?.error;
            const errMsg = err || respErr?.message;
            if (errMsg) {
              emailStatus.style.color = '#fca5a5';
              emailStatus.textContent = '\u2718 ' + errMsg;
              updateBtn.disabled = false;
              updateBtn.textContent = 'Update';
              return;
            }
            emailStatus.style.color = '#6ee7b7';
            emailStatus.textContent = '\u2714 Email updated.';
            updateBtn.textContent = '✔';
            user.email = newEmail;
            mutPostComment(user.id, 'Email updated to ' + newEmail, () => {});
            setTimeout(() => {
              updateBtn.textContent = 'Update';
              updateBtn.disabled = false;
            }, 2000);
          }
        );
      };

      emailRow.appendChild(emailInput);
      emailRow.appendChild(updateBtn);
      panel.appendChild(emailRow);
      panel.appendChild(emailStatus);

      // ── Gender section ──────────────────────────────────────────────────
      panel.appendChild(makeLabel('GENDER'));
      const genderRow = document.createElement('div');
      genderRow.style.cssText = 'display:flex;gap:6px;margin-bottom:12px;';

      const currentGender = cachedCustomerCtx?._gender || '';
      const genderSel = makeSelect([
        { value: 'MALE', label: '♂ Male (Colognes)' },
        { value: 'FEMALE', label: '♀ Female (Perfumes)' },
      ], currentGender || 'MALE');
      genderSel.style.flex = '1';

      const genderBtn = document.createElement('button');
      genderBtn.textContent = 'Update';
      genderBtn.style.cssText = 'padding:7px 14px;border-radius:6px;border:none;background:#4f46e5;color:#fff;font-weight:700;font-size:12px;cursor:pointer;';
      genderBtn.onmouseenter = () => genderBtn.style.background = '#4338ca';
      genderBtn.onmouseleave = () => genderBtn.style.background = '#4f46e5';

      const genderStatus = makeStatusEl();

      genderBtn.onclick = () => {
        const newGender = genderSel.value;
        if (newGender === currentGender) {
          genderStatus.style.color = '#f59e0b';
          genderStatus.textContent = 'Gender is unchanged.';
          return;
        }
        genderBtn.disabled = true;
        genderBtn.textContent = '⏳';
        genderStatus.textContent = 'Authenticating...';

        ensureCustomerAuth((ok) => {
          if (!ok) {
            genderStatus.style.color = '#fca5a5';
            genderStatus.textContent = '\u2718 Auth failed';
            genderBtn.disabled = false;
            genderBtn.textContent = 'Update';
            return;
          }
          genderStatus.textContent = 'Updating gender...';
          customerApiCall('UserPersonalInfoUpdate',
            `mutation UserPersonalInfoUpdate($input: UserPersonalInfoUpdateInput!) {
              userPersonalInfoUpdate(input: $input) {
                data { id gender }
                error {
                  ... on SecurityError { message }
                  ... on ServerError { message }
                  ... on ValidationError { message }
                }
              }
            }`,
            { input: { gender: newGender } },
            (data, err) => {
              const respErr = data?.userPersonalInfoUpdate?.error;
              const errMsg = err || respErr?.message;
              if (errMsg) {
                genderStatus.style.color = '#fca5a5';
                genderStatus.textContent = '\u2718 ' + errMsg;
                genderBtn.disabled = false;
                genderBtn.textContent = 'Update';
                return;
              }
              genderStatus.style.color = '#6ee7b7';
              genderStatus.textContent = '\u2714 Gender updated.';
              genderBtn.textContent = '✔';
              if (cachedCustomerCtx) cachedCustomerCtx._gender = newGender;
              mutPostComment(user.id, 'Gender updated to ' + (newGender === 'MALE' ? 'Male' : 'Female'), () => {});
              setTimeout(() => { genderBtn.textContent = 'Update'; genderBtn.disabled = false; }, 2000);
            }
          );
        });
      };

      genderRow.appendChild(genderSel);
      genderRow.appendChild(genderBtn);
      panel.appendChild(genderRow);
      panel.appendChild(genderStatus);

      // ── Address section ─────────────────────────────────────────────────
      panel.appendChild(makeLabel('SHIPPING ADDRESS'));

      // Current address display
      const shipping = user.userAddress?.shipping;
      if (shipping) {
        const addrDisplay = document.createElement('div');
        addrDisplay.style.cssText = 'font-size:11px;color:#94a3b8;margin-bottom:8px;line-height:1.5;';
        addrDisplay.textContent = [shipping.street1, shipping.city, shipping.region, shipping.postcode].filter(Boolean).join(', ');
        panel.appendChild(addrDisplay);
      }

      // Container for address edit UI (hidden until auth)
      const addrEditWrap = document.createElement('div');
      addrEditWrap.style.cssText = 'display:none;';
      panel.appendChild(addrEditWrap);

      const addrAuthStatus = makeStatusEl();
      panel.appendChild(addrAuthStatus);

      // Update Address button — triggers auth, then reveals search
      const updateAddrBtn = document.createElement('button');
      updateAddrBtn.textContent = '📍 Update Address';
      updateAddrBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid #4f46e5;background:transparent;
        color:#a5b4fc;font-weight:600;font-size:12px;cursor:pointer;box-sizing:border-box;`;
      updateAddrBtn.onmouseenter = () => updateAddrBtn.style.background = 'rgba(79,70,229,0.1)';
      updateAddrBtn.onmouseleave = () => updateAddrBtn.style.background = 'transparent';
      updateAddrBtn.onclick = () => {
        updateAddrBtn.disabled = true;
        updateAddrBtn.textContent = '⏳ Authenticating...';
        addrAuthStatus.textContent = '';

        ensureCustomerAuth((ok) => {
          if (!ok) {
            addrAuthStatus.style.color = '#fca5a5';
            addrAuthStatus.textContent = '\u2718 Authentication failed';
            updateAddrBtn.disabled = false;
            updateAddrBtn.textContent = '📍 Update Address';
            return;
          }
          // Hide button, show search UI
          updateAddrBtn.style.display = 'none';
          addrAuthStatus.textContent = '';
          addrEditWrap.style.display = '';
        });
      };
      panel.appendChild(updateAddrBtn);

      // Build address edit UI inside the hidden container
      const addrInput = document.createElement('input');
      addrInput.type = 'text';
      addrInput.placeholder = 'Start typing new address...';
      addrInput.style.cssText = `
        width:100%;padding:7px 10px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);
        background:#0f172a;color:#e2e8f0;font-size:12px;outline:none;box-sizing:border-box;
      `;
      addrEditWrap.appendChild(addrInput);

      const addrResults = document.createElement('div');
      addrResults.style.cssText = 'max-height:200px;overflow-y:auto;';
      addrEditWrap.appendChild(addrResults);

      const selectedAddr = document.createElement('div');
      selectedAddr.style.cssText = 'display:none;background:#0f172a;border-radius:6px;padding:8px 10px;margin-top:8px;font-size:11px;color:#cbd5e1;line-height:1.6;';
      addrEditWrap.appendChild(selectedAddr);

      const addrStatus = makeStatusEl();
      addrEditWrap.appendChild(addrStatus);

      const addrSaveBtn = document.createElement('button');
      addrSaveBtn.textContent = 'Save Address';
      addrSaveBtn.style.cssText = 'display:none;width:100%;padding:9px;border-radius:6px;border:none;background:#4f46e5;color:#fff;font-weight:700;font-size:12px;cursor:pointer;box-sizing:border-box;margin-top:8px;';
      addrSaveBtn.onmouseenter = () => addrSaveBtn.style.background = '#4338ca';
      addrSaveBtn.onmouseleave = () => addrSaveBtn.style.background = '#4f46e5';
      addrEditWrap.appendChild(addrSaveBtn);

      let _selectedAddrData = null;
      const userCountry = shipping?.country || 'US';

      // Debounced autocomplete
      let _addrTimer = null;
      addrInput.addEventListener('input', () => {
        clearTimeout(_addrTimer);
        const q = addrInput.value.trim();
        if (q.length < 3) { addrResults.innerHTML = ''; return; }

        _addrTimer = setTimeout(() => {
          customerApiCall('AddressAutocomplete',
            `query AddressAutocomplete($input: AddressAutocompleteInput!) {
              addressAutocomplete(input: $input) {
                data { type placeId mainText secondaryText }
              }
            }`,
            { input: { query: q, country: userCountry, apiVersion: 'V2' } },
            (data, err) => {
              addrResults.innerHTML = '';
              if (err || !data?.addressAutocomplete?.data?.length) return;

              data.addressAutocomplete.data.forEach(suggestion => {
                const row = document.createElement('div');
                row.style.cssText = 'padding:6px 10px;cursor:pointer;font-size:11px;color:#cbd5e1;border-bottom:1px solid rgba(255,255,255,0.05);';
                row.onmouseenter = () => row.style.background = 'rgba(99,102,241,0.15)';
                row.onmouseleave = () => row.style.background = '';

                const main = document.createElement('div');
                main.textContent = suggestion.mainText;
                main.style.fontWeight = '600';
                row.appendChild(main);

                if (suggestion.secondaryText) {
                  const sec = document.createElement('div');
                  sec.style.cssText = 'color:#64748b;font-size:10px;';
                  sec.textContent = suggestion.secondaryText;
                  row.appendChild(sec);
                }

                row.onclick = () => {
                  if (suggestion.type === 'CONTAINER') {
                    addrInput.value = suggestion.mainText;
                    addrInput.dispatchEvent(new Event('input', { bubbles: true }));
                    return;
                  }

                  addrResults.innerHTML = '<div style="color:#94a3b8;font-size:11px;padding:6px 10px;">Loading details...</div>';
                  customerApiCall('AddressDetails',
                    `query AddressDetails($input: AddressAutocompleteDetailsInput!) {
                      addressAutocompleteDetails(input: $input) {
                        data { city region postalCode street1 street2 }
                      }
                    }`,
                    { input: { placeId: suggestion.placeId, apiVersion: 'V2' } },
                    (detData, detErr) => {
                      addrResults.innerHTML = '';
                      if (detErr || !detData?.addressAutocompleteDetails?.data) {
                        addrStatus.style.color = '#fca5a5';
                        addrStatus.textContent = '\u2718 Failed to load address details';
                        return;
                      }
                      const addr = detData.addressAutocompleteDetails.data;
                      _selectedAddrData = addr;
                      addrInput.value = '';
                      selectedAddr.style.display = '';
                      selectedAddr.innerHTML = `
                        <div style="font-weight:600;color:#e2e8f0;">${addr.street1 || ''}${addr.street2 ? ', ' + addr.street2 : ''}</div>
                        <div>${addr.city || ''}, ${addr.region || ''} ${addr.postalCode || ''}</div>
                      `;
                      addrSaveBtn.style.display = '';
                    }
                  );
                };

                addrResults.appendChild(row);
              });
            }
          );
        }, 300);
      });

      // Save address
      addrSaveBtn.onclick = () => {
        if (!_selectedAddrData) return;
        addrSaveBtn.disabled = true;
        addrSaveBtn.textContent = 'Validating...';
        addrStatus.textContent = '';

        const firstName = (user.firstName || '').toUpperCase();
        const lastName = (user.lastName || '').toUpperCase();

        // Step 1: Validate
        customerApiCall('AddressValidate',
          `mutation AddressValidate($inputNew: AddressValidateInput!) {
            addressValidate(inputNew: $inputNew) {
              data {
                address { id primary firstName lastName country city region postalCode street1 street2 phone }
                isPrimaryAddress
              }
              error {
                ... on ValidationError { message }
                ... on AddressValidationError { message }
                ... on SecurityError { message }
                ... on ServerError { message }
                ... on AmbiguousAddressError { message suggestions { id street1 street2 city region postalCode country } }
              }
            }
          }`,
          { inputNew: {
            address: {
              firstName, lastName, country: userCountry,
              postalCode: _selectedAddrData.postalCode || '',
              region: _selectedAddrData.region || '',
              city: _selectedAddrData.city || '',
              street1: _selectedAddrData.street1 || '',
              street2: _selectedAddrData.street2 || '',
              phone: '', receiveAlertsByPhone: false, receiveOffersByPhone: false,
            },
            apiVersion: 'V2',
            strict: false,
          }},
          (valData, valErr) => {
            const valResult = valData?.addressValidate;
            const valErrMsg = valErr || valResult?.error?.message;
            if (valErrMsg) {
              addrStatus.style.color = '#fca5a5';
              addrStatus.textContent = '\u2718 ' + valErrMsg;
              addrSaveBtn.disabled = false;
              addrSaveBtn.textContent = 'Save Address';
              return;
            }

            const validatedAddr = valResult?.data?.address;
            if (!validatedAddr) {
              addrStatus.style.color = '#fca5a5';
              addrStatus.textContent = '\u2718 Validation returned no address data';
              addrSaveBtn.disabled = false;
              addrSaveBtn.textContent = 'Save Address';
              return;
            }

            const addrId = validatedAddr.id || cachedCustomerCtx?._shippingAddrId;
            if (!addrId) {
              addrStatus.style.color = '#fca5a5';
              addrStatus.textContent = '\u2718 No address ID available — try refreshing';
              addrSaveBtn.disabled = false;
              addrSaveBtn.textContent = 'Save Address';
              return;
            }

            // Step 2: Update
            addrSaveBtn.textContent = 'Saving...';
            customerApiCall('AddressUpdate',
              `mutation AddressUpdate($input: AddressUpdateInput!) {
                addressUpdate(input: $input) {
                  data {
                    personalInfo { addressInfo { shipping { id street1 street2 city region postalCode country } } }
                  }
                  error {
                    ... on SecurityError { message }
                    ... on ServerError { message }
                    ... on AddressError { message }
                    ... on ValidationError { message }
                  }
                }
              }`,
              { input: {
                addressId: addrId,
                primary: true,
                address: {
                  firstName: validatedAddr.firstName || firstName,
                  lastName: validatedAddr.lastName || lastName,
                  country: validatedAddr.country || userCountry,
                  postalCode: validatedAddr.postalCode || '',
                  region: validatedAddr.region || '',
                  city: validatedAddr.city || '',
                  street1: validatedAddr.street1 || '',
                  street2: validatedAddr.street2 || '',
                  phone: '', receiveAlertsByPhone: false, receiveOffersByPhone: false,
                },
              }},
              (updData, updErr) => {
                const updErrMsg = updErr || updData?.addressUpdate?.error?.message;
                if (updErrMsg) {
                  addrStatus.style.color = '#fca5a5';
                  addrStatus.textContent = '\u2718 ' + updErrMsg;
                  addrSaveBtn.disabled = false;
                  addrSaveBtn.textContent = 'Save Address';
                  return;
                }
                addrStatus.style.color = '#6ee7b7';
                addrStatus.textContent = '\u2714 Address updated.';
                addrSaveBtn.textContent = '✔ Saved';
                const addrStr = [validatedAddr.street1, validatedAddr.city, validatedAddr.region, validatedAddr.postalCode].filter(Boolean).join(', ');
                mutPostComment(user.id, 'Address updated to ' + addrStr, () => {});
                setTimeout(() => { addrSaveBtn.disabled = false; addrSaveBtn.textContent = 'Save Address'; }, 2000);
              }
            );
          }
        );
      };
    });
  }

  // ── Cancel Subscription Panel ─────────────────────────────────────────────

  function showCancelPanel(btn) {
    if (document.getElementById('sb-cancel-panel')) {
      removePanel('sb-cancel-panel'); return;
    }

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;
      loadCustomerOrders(user.id, (orders) => {
        const panel = createPanel({ id: 'sb-cancel-panel', title: '❌ Cancel Subscription', width: 320 });

        const subStatus = user.subscription?.status || '';
        const hasRecentSubOrder = (orders || []).some(o => o.type === 'SUBSCRIPTION' || !o.type);
        const isAlreadyCancelled = subStatus.toLowerCase() === 'cancelled';

        // Customer info
        const name = [user.firstName, user.lastName].filter(Boolean).join(' ') || '\u2014';
        const infoEl = document.createElement('div');
        infoEl.style.cssText = 'font-size:12px;color:#94a3b8;margin-bottom:10px;line-height:1.6;';
        infoEl.innerHTML = `<span style="color:#e2e8f0;font-weight:600;">${name}</span> &middot; ${user.email}`;
        panel.appendChild(infoEl);

        // Last subscription order summary
        const lastOrder = orders?.find(o => o.type === 'SUBSCRIPTION');
        if (lastOrder) {
          const orderEl = document.createElement('div');
          orderEl.style.cssText = 'background:#1a1a2e;border-radius:6px;padding:8px 10px;margin-bottom:12px;border-left:2px solid #4f46e5;';
          const monthLbl = orderMonthLabel(lastOrder);
          const ws = lastOrder.warehouseOrder?.data?.status || lastOrder.status || '';
          const wsUC = ws.toUpperCase();
          const statusColor = wsUC === 'DONE' ? '#6ee7b7' : '#fca5a5';
          const lines = orderProductLines(lastOrder);

          const orderTop = document.createElement('div');
          orderTop.style.cssText = 'display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px;';

          const orderInfo = document.createElement('div');
          orderInfo.innerHTML = `<div style="display:flex;gap:6px;align-items:center;margin-bottom:3px;">
            <span style="color:#94a3b8;font-size:11px;">${monthLbl}</span>
            <span style="color:${statusColor};font-size:11px;font-weight:600;">${ws}</span>
          </div>` + lines.map(l => `<div style="color:#cbd5e1;font-size:11px;">${l}</div>`).join('');

          // Tracking
          const trackNo  = lastOrder.tracking?.trackingNumber || lastOrder.shipment?.trackingNumber || '';
          const trackUrl = lastOrder.tracking?.trackingUrl  || lastOrder.shipment?.trackingUrl  || '';
          const trackItems = lastOrder.tracking?.items || [];
          if (trackNo) {
            const tRow = document.createElement('div');
            tRow.style.cssText = 'display:flex;align-items:center;gap:4px;margin-top:4px;flex-wrap:wrap;';
            if (trackUrl) {
              const a = document.createElement('a');
              a.href = trackUrl; a.target = '_blank';
              a.textContent = trackNo;
              a.style.cssText = 'color:#818cf8;font-size:11px;text-decoration:none;';
              tRow.appendChild(a);
              const copyBtn = document.createElement('button');
              copyBtn.textContent = '\uD83D\uDD17';
              copyBtn.title = 'Copy tracking link';
              copyBtn.style.cssText = 'background:none;border:none;cursor:pointer;color:#818cf8;font-size:12px;padding:0 2px;';
              copyBtn.onclick = () => {
                navigator.clipboard.write([
                  new ClipboardItem({
                    'text/html':  new Blob([`<a href="${trackUrl}">${trackNo}</a>`], { type: 'text/html' }),
                    'text/plain': new Blob([trackUrl], { type: 'text/plain' }),
                  })
                ]);
                copyBtn.textContent = '\u2714'; setTimeout(() => copyBtn.textContent = '\uD83D\uDD17', 1500);
              };
              tRow.appendChild(copyBtn);
            } else {
              const sp = document.createElement('span');
              sp.style.cssText = 'color:#818cf8;font-size:11px;';
              sp.textContent = trackNo;
              tRow.appendChild(sp);
            }
            orderInfo.appendChild(tRow);
          }
          if (trackItems.length) {
            const evEl = document.createElement('div');
            evEl.style.cssText = 'margin-top:3px;';
            trackItems.slice(-2).forEach(item => {
              const d = document.createElement('div');
              d.style.cssText = 'color:#64748b;font-size:11px;margin-top:1px;';
              const dt = new Date(item.date).toLocaleDateString('en-US', { month:'short', day:'numeric' });
              d.textContent = dt + ' · ' + item.description;
              evEl.appendChild(d);
            });
            orderInfo.appendChild(evEl);
          }
          orderTop.appendChild(orderInfo);

          // Refund button
          const wsUpper = ws.toUpperCase();
          const refundable = !NON_REFUNDABLE_STATUSES.includes(wsUpper);
          const orderRefBtn = document.createElement('button');
          orderRefBtn.textContent = '💰 Refund';
          orderRefBtn.disabled = !refundable;
          orderRefBtn.title = refundable ? 'Open refund charges' : 'Order cannot be refunded at this stage (' + ws + ')';
          orderRefBtn.style.cssText = `
            margin-left:8px;padding:3px 8px;border-radius:5px;border:none;font-size:11px;font-weight:600;
            white-space:nowrap;flex-shrink:0;
            cursor:${refundable ? 'pointer' : 'not-allowed'};
            background:${refundable ? '#7c3aed' : 'rgba(124,58,237,0.2)'};
            color:${refundable ? '#fff' : '#64748b'};
          `;
          if (refundable) {
            orderRefBtn.onmouseenter = () => orderRefBtn.style.background = '#6d28d9';
            orderRefBtn.onmouseleave = () => orderRefBtn.style.background = '#7c3aed';
            orderRefBtn.onclick = () => showRefundChargesPanel(user);
          }
          orderTop.appendChild(orderRefBtn);
          orderEl.appendChild(orderTop);
          panel.appendChild(orderEl);
        }

        const statusEl = makeStatusEl();
        panel.appendChild(statusEl);

        if (isAlreadyCancelled) {
          const subStatusEl = document.createElement('div');
          subStatusEl.style.cssText = 'font-size:11px;color:#f59e0b;margin-bottom:10px;';
          subStatusEl.textContent = '⚠ Subscription status: ' + subStatus;
          panel.appendChild(subStatusEl);
        }

        const confirmBtn = document.createElement('button');
        confirmBtn.textContent = 'Cancel Subscription';
        confirmBtn.disabled = isAlreadyCancelled;
        confirmBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:none;font-weight:700;font-size:13px;box-sizing:border-box;
          background:${isAlreadyCancelled ? '#374151' : '#dc2626'};
          color:${isAlreadyCancelled ? '#6b7280' : '#fff'};
          cursor:${isAlreadyCancelled ? 'not-allowed' : 'pointer'};`;
        if (!isAlreadyCancelled) {
          confirmBtn.onmouseenter = () => confirmBtn.style.background = '#b91c1c';
          confirmBtn.onmouseleave = () => confirmBtn.style.background = '#dc2626';
        }

        confirmBtn.onclick = () => {
          confirmBtn.disabled = true;
          confirmBtn.textContent = 'Fetching subscription...';
          statusEl.textContent = '';

          fetchActiveSubscription(user.id, (subscriptionId, fetchErr) => {
            if (fetchErr || !subscriptionId) {
              statusEl.style.color = '#fca5a5';
              statusEl.textContent = '\u2718 ' + (fetchErr || 'No active subscription found.');
              confirmBtn.disabled = false;
              confirmBtn.textContent = 'Confirm Cancellation';
              return;
            }
            confirmBtn.textContent = 'Cancelling...';
            mutCancelSub(subscriptionId, (data, cancelErr) => {
              if (cancelErr) {
                statusEl.style.color = '#fca5a5';
                statusEl.textContent = '\u2718 ' + cancelErr;
                confirmBtn.disabled = false;
                confirmBtn.textContent = 'Confirm Cancellation';
                return;
              }
              statusEl.style.color = '#94a3b8';
              statusEl.textContent = 'Posting comment...';
              mutPostComment(user.id, 'Subscription Cancelled', () => {
                statusEl.style.color = '#6ee7b7';
                statusEl.textContent = '\u2714 Subscription cancelled.';
                confirmBtn.textContent = '\u2714 Done';
                setTimeout(() => removePanel('sb-cancel-panel'), 2000);
              });
            });
          });
        };

        panel.appendChild(confirmBtn);

        // ── Delete Payment Methods ────────────────────────────────────────
        const pmSep = document.createElement('div');
        pmSep.style.cssText = 'height:1px;background:rgba(255,255,255,0.07);margin:12px 0;';
        panel.appendChild(pmSep);

        const delPmBtn = document.createElement('button');
        delPmBtn.textContent = '🗑 Delete Payment Methods';
        delPmBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid #dc2626;background:transparent;
          color:#fca5a5;font-weight:600;font-size:12px;cursor:pointer;box-sizing:border-box;`;
        delPmBtn.onmouseenter = () => delPmBtn.style.background = 'rgba(220,38,38,0.1)';
        delPmBtn.onmouseleave = () => delPmBtn.style.background = 'transparent';
        delPmBtn.onclick = () => {
          const { overlay, box } = createModal({ id: 'sb-del-pm-overlay', title: '🗑 Delete All Payment Methods', width: 320 });

          const warnEl = document.createElement('div');
          warnEl.style.cssText = 'background:rgba(220,38,38,0.1);border:1px solid rgba(220,38,38,0.3);border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:12px;color:#fca5a5;line-height:1.6;';
          warnEl.innerHTML = `This will permanently remove <strong>all stored payment methods</strong> for:<br><br>
            <span style="color:#e2e8f0;font-weight:600;">${name}</span><br>
            <span style="color:#94a3b8;">${user.email}</span><br><br>
            This action cannot be undone.`;
          box.appendChild(warnEl);

          const pmStatus = makeStatusEl();
          box.appendChild(pmStatus);

          box.appendChild(makeDialogButtons({
            confirmLabel: 'Delete All',
            confirmColor: '#dc2626',
            onCancel: () => overlay.remove(),
            onConfirm: (cBtn) => {
              cBtn.disabled = true;
              cBtn.textContent = 'Deleting...';
              gqlMutate('deleteAllPaymentMethods', DELETE_PAYMENT_METHODS_MUTATION,
                { input: { userId: user.id } },
                (data, err) => {
                  const respErr = data?.paymentMethodDeleteAll?.error;
                  const errMsg = err || respErr?.message || respErr?.serverErrorMessage;
                  if (errMsg) {
                    pmStatus.style.color = '#fca5a5';
                    pmStatus.textContent = '\u2718 ' + errMsg;
                    cBtn.disabled = false;
                    cBtn.textContent = 'Delete All';
                    return;
                  }
                  pmStatus.style.color = '#6ee7b7';
                  pmStatus.textContent = '\u2714 All payment methods deleted.';
                  cBtn.textContent = '\u2714 Done';
                  delPmBtn.disabled = true;
                  delPmBtn.textContent = '\u2714 Payment Methods Deleted';
                  delPmBtn.style.color = '#6ee7b7';
                  delPmBtn.style.borderColor = '#6ee7b7';
                  delPmBtn.style.cursor = 'default';
                  mutPostComment(user.id, 'Billing method removed', () => {});
                  setTimeout(() => overlay.remove(), 1800);
                }
              );
            },
          }));
        };
        panel.appendChild(delPmBtn);

        // ── Delete Account ────────────────────────────────────────────────
        const delAccBtn = document.createElement('button');
        delAccBtn.textContent = '⛔ Delete Account';
        delAccBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid #991b1b;background:transparent;
          color:#fca5a5;font-weight:600;font-size:12px;cursor:pointer;box-sizing:border-box;margin-top:6px;`;
        delAccBtn.onmouseenter = () => delAccBtn.style.background = 'rgba(153,27,27,0.15)';
        delAccBtn.onmouseleave = () => delAccBtn.style.background = 'transparent';
        delAccBtn.onclick = () => {
          const { overlay, box } = createModal({ id: 'sb-del-acc-overlay', title: '⛔ Delete Account', width: 320 });

          const warnEl = document.createElement('div');
          warnEl.style.cssText = 'background:rgba(153,27,27,0.15);border:1px solid rgba(153,27,27,0.4);border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:12px;color:#fca5a5;line-height:1.6;';
          warnEl.innerHTML = `This will:<br>
            <span style="color:#e2e8f0;">1.</span> Remove all payment methods<br>
            <span style="color:#e2e8f0;">2.</span> Disable the account email<br><br>
            <span style="color:#e2e8f0;font-weight:600;">${name}</span><br>
            <span style="color:#94a3b8;">${user.email}</span><br><br>
            This action cannot be undone.`;
          box.appendChild(warnEl);

          const accStatus = makeStatusEl();
          box.appendChild(accStatus);

          box.appendChild(makeDialogButtons({
            confirmLabel: 'Delete Account',
            confirmColor: '#991b1b',
            onCancel: () => overlay.remove(),
            onConfirm: (cBtn) => {
              cBtn.disabled = true;
              cBtn.textContent = 'Deleting...';
              accStatus.textContent = 'Removing payment methods...';

              // Step 1: Delete payment methods (skip if fails)
              gqlMutate('deleteAllPaymentMethods', DELETE_PAYMENT_METHODS_MUTATION,
                { input: { userId: user.id } },
                (pmData, pmErr) => {
                  // Step 2: Disable email
                  accStatus.textContent = 'Disabling account email...';
                  const [localPart, domain] = user.email.split('@');
                  const rand = Math.floor(Math.random() * 900000000) + 100000000;
                  const disabledEmail = localPart + '_disabled' + rand + '@' + domain;

                  gqlMutate('userChangeEmail', CHANGE_EMAIL_MUTATION,
                    { input: { userId: user.id, email: disabledEmail } },
                    (emailData, emailErr) => {
                      const respErr = emailData?.userChangeEmail?.error;
                      const errMsg = emailErr || respErr?.message;
                      if (errMsg) {
                        accStatus.style.color = '#fca5a5';
                        accStatus.textContent = '\u2718 ' + errMsg;
                        cBtn.disabled = false;
                        cBtn.textContent = 'Delete Account';
                        return;
                      }
                      accStatus.style.color = '#6ee7b7';
                      accStatus.textContent = '\u2714 Account deleted.';
                      cBtn.textContent = '\u2714 Done';
                      delAccBtn.disabled = true;
                      delAccBtn.textContent = '\u2714 Account Deleted';
                      delAccBtn.style.color = '#6ee7b7';
                      delAccBtn.style.borderColor = '#6ee7b7';
                      delAccBtn.style.cursor = 'default';
                      mutPostComment(user.id, 'Account deleted', () => {});
                      setTimeout(() => overlay.remove(), 1800);
                    }
                  );
                }
              );
            },
          }));
        };
        panel.appendChild(delAccBtn);

        // ── Fraud Blocklist Toggle ────────────────────────────────────────
        const fraudBtn = document.createElement('button');
        fraudBtn.textContent = '🚫 Blocklist (Fraud)';
        fraudBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid #7c3aed;background:transparent;
          color:#c4b5fd;font-weight:600;font-size:12px;cursor:pointer;box-sizing:border-box;margin-top:6px;`;
        fraudBtn.onmouseenter = () => fraudBtn.style.background = 'rgba(124,58,237,0.1)';
        fraudBtn.onmouseleave = () => fraudBtn.style.background = 'transparent';

        // Use cached fraud status from initial user details query
        const cachedFraudStatus = cachedCustomerCtx?._fraudStatus || 'DEFAULT';
        fraudBtn._currentStatus = cachedFraudStatus;
        if (cachedFraudStatus === 'DECLINE') {
          fraudBtn.textContent = '✔ Blocklisted (Fraud)';
          fraudBtn.style.color = '#ef4444';
          fraudBtn.style.borderColor = '#ef4444';
        }

        fraudBtn.onclick = () => {
          const newStatus = (fraudBtn._currentStatus === 'DECLINE') ? 'DEFAULT' : 'DECLINE';
          const action = newStatus === 'DECLINE' ? 'Blocklisting...' : 'Removing blocklist...';
          fraudBtn.disabled = true;
          fraudBtn.textContent = action;

          gqlMutate('userSetFraudStatus', SET_FRAUD_STATUS_MUTATION,
            { input: { userId: user.id, status: newStatus } },
            (data, err) => {
              const respErr = data?.userSetFraudStatus?.error;
              const errMsg = err || respErr?.message;
              if (errMsg) {
                fraudBtn.style.color = '#fca5a5';
                fraudBtn.textContent = '\u2718 ' + errMsg;
                fraudBtn.disabled = false;
                setTimeout(() => {
                  fraudBtn.textContent = fraudBtn._currentStatus === 'DECLINE' ? '✔ Blocklisted (Fraud)' : '🚫 Blocklist (Fraud)';
                  fraudBtn.style.color = fraudBtn._currentStatus === 'DECLINE' ? '#ef4444' : '#c4b5fd';
                }, 2000);
                return;
              }
              fraudBtn._currentStatus = newStatus;
              if (newStatus === 'DECLINE') {
                fraudBtn.textContent = '✔ Blocklisted (Fraud)';
                fraudBtn.style.color = '#ef4444';
                fraudBtn.style.borderColor = '#ef4444';
                mutPostComment(user.id, 'Fraud blocklist applied', () => {});
              } else {
                fraudBtn.textContent = '🚫 Blocklist (Fraud)';
                fraudBtn.style.color = '#c4b5fd';
                fraudBtn.style.borderColor = '#7c3aed';
                mutPostComment(user.id, 'Fraud blocklist removed', () => {});
              }
              fraudBtn.disabled = false;
            }
          );
        };
        panel.appendChild(fraudBtn);
      });
    });
  }

  // ── Refund Charges Panel ──────────────────────────────────────────────────

  function showRefundChargesPanel(user) {
    if (document.getElementById('sb-refund-charges-panel')) {
      document.getElementById('sb-refund-charges-panel').remove(); return;
    }

    const panel = createPanel({ id: 'sb-refund-charges-panel', title: '💳 Charges — Refund', width: 360, right: 350 });

    const loadingEl = document.createElement('div');
    loadingEl.style.cssText = 'color:#94a3b8;font-size:12px;';
    loadingEl.textContent = 'Loading charges...';
    panel.appendChild(loadingEl);

    fetchCharges(user.id, (result) => {
      loadingEl.remove();

      if (!result || !result.all?.length) {
        const empty = document.createElement('div');
        empty.style.cssText = 'color:#94a3b8;font-size:12px;';
        empty.textContent = result ? 'No charges found for this month.' : 'Failed to load charges.';
        panel.appendChild(empty);
        return;
      }

      result.all.forEach((entry, idx) => {
        const charge = entry.charge;
        const invoice = charge?.cashbirdDetails?.invoice;
        const credits = charge?.cashbirdDetails?.credits || [];
        const shippingCredits = charge?.cashbirdDetails?.shippingCredits || [];
        const currency = detectCurrency(invoice?.lineItems);
        const totalCents = invoice?.paid || Math.round((charge.totalPrice || 0) * 100);
        const alreadyRefunded = (charge.refundAmount || 0) > 0;
        const hasRefunds = charge?.cashbirdDetails?.refunds?.length > 0;

        if (idx > 0) {
          const sep = document.createElement('div');
          sep.style.cssText = 'height:1px;background:rgba(255,255,255,0.07);margin:10px 0;';
          panel.appendChild(sep);
        }

        const block = document.createElement('div');
        block.style.cssText = 'background:#0f172a;border-radius:8px;padding:10px 12px;';

        // Charge header
        const chgHdr = document.createElement('div');
        chgHdr.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;';
        const chgLeft = document.createElement('div');
        chgLeft.style.cssText = 'font-size:12px;color:#94a3b8;';
        chgLeft.textContent = chargeDateLabel(charge.paymentDate);
        const chgTotal = document.createElement('div');
        chgTotal.style.cssText = 'display:flex;align-items:center;gap:2px;font-weight:700;font-size:13px;color:#e2e8f0;';
        chgTotal.textContent = fmtMoney(totalCents, currency);
        chgTotal.appendChild(makeCopyBtn(fmtMoney(totalCents, currency), '📋'));
        chgHdr.appendChild(chgLeft);
        chgHdr.appendChild(chgTotal);
        block.appendChild(chgHdr);

        // Payment method
        const pmMethod = charge?.cashbirdDetails?.paymentMethod?.methodName || '';
        if (pmMethod) {
          const pmEl = document.createElement('div');
          pmEl.style.cssText = 'font-size:11px;color:#64748b;margin-bottom:6px;';
          pmEl.textContent = pmMethod;
          block.appendChild(pmEl);
        }

        // Line items breakdown
        const lineItems = invoice?.lineItems || [];
        if (lineItems.length) {
          const breakdown = document.createElement('div');
          breakdown.style.cssText = 'margin-bottom:8px;';
          lineItems.forEach(li => {
            const liEl = document.createElement('div');
            liEl.style.cssText = 'display:flex;justify-content:space-between;font-size:11px;color:#94a3b8;padding:2px 0;';
            const liName = document.createElement('span');
            liName.style.cssText = 'flex:1;margin-right:8px;';
            liName.textContent = li.description || li.productCode || li.type || '—';
            const liRight = document.createElement('span');
            liRight.style.cssText = 'display:flex;align-items:center;gap:6px;white-space:nowrap;';
            const liAmt = document.createElement('span');
            liAmt.style.color = '#cbd5e1';
            liAmt.textContent = fmtMoney(li.total || li.price || 0, currency);
            liRight.appendChild(liAmt);
            if (li.refundedAmount !== null && li.refundedAmount !== undefined) {
              const liRefTag = document.createElement('span');
              const fullyRefunded = li.refundedAmount >= (li.total || li.price || 0);
              liRefTag.style.cssText = 'font-size:10px;font-weight:600;color:' + (fullyRefunded ? '#6ee7b7' : '#f59e0b') + ';';
              liRefTag.textContent = fullyRefunded ? '\u21A9 refunded' : '\u21A9 partial';
              liRight.appendChild(liRefTag);
            }
            liEl.appendChild(liName);
            liEl.appendChild(liRight);
            breakdown.appendChild(liEl);
          });
          if (invoice.tax) {
            const taxEl = document.createElement('div');
            taxEl.style.cssText = 'display:flex;justify-content:space-between;font-size:11px;color:#64748b;padding:2px 0;border-top:1px solid rgba(255,255,255,0.05);margin-top:3px;';
            taxEl.innerHTML = `<span>Tax</span><span>${fmtMoney(invoice.tax, currency)}</span>`;
            breakdown.appendChild(taxEl);
          }
          block.appendChild(breakdown);
        }

        // Credit summary
        if (credits.length) {
          const creditEl = document.createElement('div');
          creditEl.style.cssText = 'font-size:11px;color:#94a3b8;margin-bottom:8px;';
          const perCredit = credits[0]?.total || credits[0]?.amount || 0;
          creditEl.textContent = credits.length === 1
            ? fmtMoney(perCredit, currency) + ' credit'
            : credits.length + ' × ' + fmtMoney(perCredit, currency) + ' credits';
          if (shippingCredits.length) {
            const sc = shippingCredits[0]?.total || shippingCredits[0]?.amount || 0;
            creditEl.textContent += ' + ' + fmtMoney(sc, currency) + ' shipping';
          }
          block.appendChild(creditEl);
        }

        // Already refunded warning
        if (alreadyRefunded || hasRefunds) {
          const warnEl = document.createElement('div');
          warnEl.style.cssText = 'font-size:11px;color:#f59e0b;margin-bottom:8px;';
          warnEl.textContent = alreadyRefunded
            ? '⚠ Partially refunded: ' + fmtMoney(Math.round(charge.refundAmount * 100), currency) + ' already returned'
            : '⚠ Refund records exist for this charge';
          block.appendChild(warnEl);
        }

        // Refund button
        const hasCreditsFlag = credits.length > 0 || shippingCredits.length > 0;
        const lineItemsForRefund = invoice?.lineItems || [];
        const canRefund = hasCreditsFlag || lineItemsForRefund.length > 0;
        const refBtn = document.createElement('button');
        refBtn.textContent = 'Refund ' + fmtMoney(totalCents, currency);
        refBtn.disabled = !canRefund;
        refBtn.style.cssText = `
          width:100%;padding:7px;border-radius:6px;border:none;margin-top:4px;
          font-weight:700;font-size:12px;cursor:${canRefund ? 'pointer' : 'not-allowed'};
          background:${canRefund ? '#7c3aed' : 'rgba(124,58,237,0.2)'};
          color:${canRefund ? '#fff' : '#64748b'};
          box-sizing:border-box;
        `;
        if (canRefund) {
          refBtn.onmouseenter = () => refBtn.style.background = '#6d28d9';
          refBtn.onmouseleave = () => refBtn.style.background = '#7c3aed';
          const orderId = charge?.orders?.[0]?.id || null;
          refBtn.onclick = () => showRefundConfirmDialog(user, charge, credits, shippingCredits, currency, totalCents, refBtn, block, orderId);
        }
        block.appendChild(refBtn);
        panel.appendChild(block);
      });

      // Refund total row — updates dynamically after each refund
      const refundTotalRow = document.createElement('div');
      refundTotalRow.id = 'sb-refund-total-row';
      refundTotalRow.style.cssText = 'display:none;margin-top:10px;padding:8px 12px;background:#0f172a;border-radius:8px;border:1px solid rgba(110,231,183,0.2);';
      const refundTotalInner = document.createElement('div');
      refundTotalInner.style.cssText = 'display:flex;align-items:center;justify-content:space-between;';
      const refundTotalLabel = document.createElement('span');
      refundTotalLabel.style.cssText = 'font-size:12px;color:#6ee7b7;font-weight:700;';
      const refundTotalCopy = makeCopyBtn(() => {
        const cur = result.all?.[0]?.charge?.cashbirdDetails?.invoice?.lineItems?.[0]?.currency || 'USD';
        return fmtMoney(_totalRefundedCents, cur);
      }, '📋');
      refundTotalInner.appendChild(refundTotalLabel);
      refundTotalInner.appendChild(refundTotalCopy);
      refundTotalRow.appendChild(refundTotalInner);
      panel.appendChild(refundTotalRow);

      // Observe _totalRefundedCents changes via polling (simple approach)
      const refundTotalTimer = setInterval(() => {
        if (!document.getElementById('sb-refund-charges-panel')) { clearInterval(refundTotalTimer); return; }
        if (_totalRefundedCents > 0) {
          const cur = result.all?.[0]?.charge?.cashbirdDetails?.invoice?.lineItems?.[0]?.currency || 'USD';
          refundTotalLabel.textContent = 'Total Refunded: ' + fmtMoney(_totalRefundedCents, cur);
          refundTotalRow.style.display = '';
        }
      }, 500);
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // MODALS / DIALOGS (using createModal factory)
  // ══════════════════════════════════════════════════════════════════════════

  // ── Confirm Dialog (generic) ──────────────────────────────────────────────

  function showConfirmDialog({ title, lines, confirmLabel, confirmColor = '#dc2626', onConfirm }) {
    const { overlay, box } = createModal({ id: 'sb-confirm-dialog', title, width: 360 });

    lines.forEach(line => {
      const el = document.createElement('div');
      el.style.cssText = 'font-size:12px;color:#94a3b8;margin-bottom:5px;padding:4px 8px;background:#2a2a3e;border-radius:4px;';
      el.textContent = line;
      box.appendChild(el);
    });

    const warn = document.createElement('div');
    warn.style.cssText = 'font-size:12px;color:#fca5a5;margin-top:10px;margin-bottom:14px;';
    warn.textContent = 'This cannot be undone.';
    box.appendChild(warn);

    const btns = makeDialogButtons({
      confirmLabel, confirmColor,
      onCancel: () => overlay.remove(),
      onConfirm: (confirmBtn) => { overlay.remove(); onConfirm(confirmBtn); },
    });
    box.appendChild(btns);
  }

  // ── Cancel Order Dialog ───────────────────────────────────────────────────

  const CANCEL_ORDER_REASONS = [
    { value: 'charge clarification', label: 'Charge clarification' },
    { value: 'order change needed',  label: 'Order change needed' },
    { value: 'clarification needed', label: 'Clarification needed' },
  ];

  function showCancelOrderDialog(order, user, triggerBtn) {
    const { overlay, box } = createModal({ id: 'sb-cancel-order-dialog', title: '❌ Cancel Order', width: 320 });

    // Order info
    const info = document.createElement('div');
    info.style.cssText = 'font-size:12px;color:#94a3b8;margin-bottom:14px;';
    info.textContent = orderMonthLabel(order) + ' order';
    box.appendChild(info);

    // Reason
    box.appendChild(makeLabel('Reason'));
    const select = makeSelect(CANCEL_ORDER_REASONS, 'charge clarification');
    select.style.marginBottom = '14px';
    box.appendChild(select);

    // Status
    const statusEl = makeStatusEl();
    box.appendChild(statusEl);

    // Buttons
    const btns = makeDialogButtons({
      confirmLabel: 'Cancel Order',
      confirmColor: '#dc2626',
      onCancel: () => overlay.remove(),
      onConfirm: (confirmBtn, cancelBtn) => {
        confirmBtn.disabled = true; cancelBtn.disabled = true;
        confirmBtn.textContent = 'Cancelling...';
        statusEl.textContent = 'Cancelling order...';

        mutCancelOrder(order.id, (data, err) => {
          if (err) {
            statusEl.style.color = '#fca5a5';
            statusEl.textContent = '✘ ' + err;
            confirmBtn.disabled = false; cancelBtn.disabled = false;
            confirmBtn.textContent = 'Cancel Order';
            return;
          }

          statusEl.style.color = '#94a3b8';
          statusEl.textContent = 'Posting comment...';

          const d = new Date(order.year + '-' + String(order.month).padStart(2, '0') + '-01');
          const mLabel = d.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }).replace(' ', "'");
          const comment = 'Cancelled ' + mLabel + ' order, ' + select.value;

          mutPostComment(user.id, comment, () => {
            statusEl.style.color = '#6ee7b7';
            statusEl.textContent = '✔ Done!';
            confirmBtn.textContent = '✔ Done';
            triggerBtn.disabled = true;
            triggerBtn.textContent = '✔ Cancelled';
            triggerBtn.style.color = '#6ee7b7';
            triggerBtn.style.borderColor = '#6ee7b7';
            setTimeout(() => overlay.remove(), 1800);
          });
        });
      },
    });
    box.appendChild(btns);
  }

  // ── Refund Confirm Dialog ─────────────────────────────────────────────────

  function showRefundConfirmDialog(user, charge, credits, shippingCredits, currency, totalCents, refBtn, parentBlock, orderId) {
    const { overlay, box } = createModal({ id: 'sb-refund-confirm-overlay', title: 'Confirm Refund', width: 300 });

    const invoice = charge?.cashbirdDetails?.invoice;
    const pmMethod = charge?.cashbirdDetails?.paymentMethod?.methodName || 'Unknown';
    const dateStr = chargeDateLabel(charge.paymentDate);

    // Details
    const details = document.createElement('div');
    details.style.cssText = 'background:#0f172a;border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:12px;line-height:1.8;color:#cbd5e1;';
    const perCredit = credits[0]?.total || credits[0]?.amount || 0;
    const creditLine = credits.length === 1
      ? fmtMoney(perCredit, currency) + ' credit'
      : credits.length + ' × ' + fmtMoney(perCredit, currency) + ' credits';
    let shippingLine = '';
    if (shippingCredits.length) {
      const sc = shippingCredits[0]?.total || shippingCredits[0]?.amount || 0;
      shippingLine = '<br>' + fmtMoney(sc, currency) + ' shipping credit';
    }
    details.innerHTML = `
      <div><span style="color:#64748b;">Date</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ${dateStr}</div>
      <div><span style="color:#64748b;">Method</span>&nbsp;&nbsp;&nbsp; ${pmMethod}</div>
      <div><span style="color:#64748b;">Credits</span>&nbsp;&nbsp;&nbsp; ${creditLine}${shippingLine}</div>
      <div style="margin-top:6px;padding-top:6px;border-top:1px solid rgba(255,255,255,0.07);font-weight:700;font-size:13px;color:#e2e8f0;">
        Total refund: ${fmtMoney(totalCents, currency)}
      </div>`;
    box.appendChild(details);

    // Existing refund warning
    const existingRefunds = charge?.cashbirdDetails?.invoice?.refunds || [];
    if (existingRefunds.length) {
      const refWarn = document.createElement('div');
      refWarn.style.cssText = 'background:rgba(245,158,11,0.1);border:1px solid rgba(245,158,11,0.3);border-radius:6px;padding:8px 10px;margin-bottom:10px;font-size:11px;color:#f59e0b;line-height:1.6;';
      const refundLines = existingRefunds.map(r => {
        const amt = fmtMoney(r.amount, currency);
        const status = r.status === 'TRACKING_REFUND_IN_PROCESS' ? 'in progress' : (r.status || '').toLowerCase().replace(/_/g,' ');
        const date = new Date(r.requestDate).toLocaleDateString('en-US', { month:'short', day:'numeric' });
        return `${amt} — ${status} (${date})`;
      }).join('<br>');
      refWarn.innerHTML = '\u26A0 Existing refund(s):<br>' + refundLines;
      box.appendChild(refWarn);
    }

    // Per-line-item refund status
    const dlgLineItems = charge?.cashbirdDetails?.invoice?.lineItems || [];
    const refundedItems = dlgLineItems.filter(li => li.refundedAmount !== null && li.refundedAmount !== undefined && li.total > 0);
    if (refundedItems.length) {
      const liWarn = document.createElement('div');
      liWarn.style.cssText = 'font-size:11px;color:#94a3b8;margin-bottom:10px;line-height:1.8;';
      refundedItems.forEach(li => {
        const row = document.createElement('div');
        const fullyRefunded = li.refundedAmount >= li.total;
        row.style.color = fullyRefunded ? '#6ee7b7' : '#f59e0b';
        row.textContent = (fullyRefunded ? '\u2714' : '\u21A9') + ' ' + (li.description || li.type) + ': ' + fmtMoney(li.refundedAmount, currency) + ' refunded';
        liWarn.appendChild(row);
      });
      box.appendChild(liWarn);
    }

    const warn = document.createElement('div');
    warn.style.cssText = 'font-size:12px;color:#fca5a5;margin-bottom:14px;line-height:1.5;';
    warn.textContent = 'This will refund the payment to the customer. This cannot be undone.';
    box.appendChild(warn);

    const statusEl = makeStatusEl();
    box.appendChild(statusEl);

    const btns = makeDialogButtons({
      confirmLabel: 'Refund',
      confirmColor: '#7c3aed',
      onCancel: () => overlay.remove(),
      onConfirm: (confirmBtn, cancelBtn) => {
        confirmBtn.disabled = true;
        cancelBtn.disabled = true;
        confirmBtn.textContent = 'Refunding...';
        statusEl.textContent = '';

        const creditIds = credits.map(c => c.id);
        const shippingCreditIds = shippingCredits.map(c => c.id);

        const doRefund = () => {
          const onResult = (data, err) => {
            if (err) {
              statusEl.style.color = '#fca5a5';
              statusEl.textContent = '\u2718 ' + err;
              confirmBtn.disabled = false;
              cancelBtn.disabled = false;
              confirmBtn.textContent = 'Refund';
              return;
            }

            statusEl.style.color = '#94a3b8';
            statusEl.textContent = 'Posting comment...';

            // Track total refunded
            _totalRefundedCents += totalCents;
            _totalRefundedCurrency = currency;

            const refDate = new Date(charge.paymentDate);
            const refMonthLabel = refDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            const commentText = 'Cancelled & Refunded ' + refMonthLabel + ' order';

            const afterComment = () => {
              statusEl.style.color = '#6ee7b7';
              statusEl.textContent = '\u2714 Refunded!';
              confirmBtn.textContent = '\u2714 Done';
              refBtn.disabled = true;
              refBtn.textContent = '\u2714 Refunded';
              refBtn.style.background = 'rgba(110,231,183,0.15)';
              refBtn.style.color = '#6ee7b7';
              refBtn.style.cursor = 'not-allowed';
              setTimeout(() => overlay.remove(), 1800);
            };

            // Skip duplicate comment for same month
            if (_refundCommentsPosted.has(commentText)) {
              afterComment();
            } else {
              _refundCommentsPosted.add(commentText);
              mutPostComment(user.id, commentText, afterComment);
            }
          };

          const unrefundedItems = (charge?.cashbirdDetails?.invoice?.lineItems || [])
            .filter(li => li.total > 0 && li.refundedAmount === null && li.type !== 'SHIPPING_PURCHASE')
            .map(li => ({ lineItemUuid: li.uuid, amount: li.total, forceRefundUsedCredits: false }));

          const hasCreditsFlag = creditIds.length > 0 || shippingCreditIds.length > 0;
          const hasItems = unrefundedItems.length > 0;

          if (hasCreditsFlag && hasItems) {
            statusEl.textContent = 'Refunding subscription...';
            mutRefundCredits(user.id, creditIds, shippingCreditIds, (data, err) => {
              if (err) return onResult(null, err);
              statusEl.textContent = 'Refunding upcharge...';
              mutRefundInvoiceItems(user.id, unrefundedItems, onResult);
            });
          } else if (hasCreditsFlag) {
            mutRefundCredits(user.id, creditIds, shippingCreditIds, onResult);
          } else if (hasItems) {
            mutRefundInvoiceItems(user.id, unrefundedItems, onResult);
          } else {
            onResult({}, null);
          }
        };

        if (orderId) {
          statusEl.style.color = '#94a3b8';
          statusEl.textContent = 'Cancelling order...';
          mutCancelOrder(orderId, (data, cancelErr) => {
            if (cancelErr) console.warn('Order cancel failed:', cancelErr);
            doRefund();
          });
        } else {
          doRefund();
        }
      },
    });
    box.appendChild(btns);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // REPLACEMENT FORM
  // ══════════════════════════════════════════════════════════════════════════

  const REPLACEMENT_REASONS = [
    { value: 'DAMAGED_LEAKING',              label: 'Leaked' },
    { value: 'DAMAGED_VIAL',                 label: 'Broken / Not Spraying' },
    { value: 'INCOMPLETE',                   label: 'Incomplete' },
    { value: 'WRONG_ITEM',                   label: 'Wrong' },
    { value: 'NOT_RECEIVED',                 label: 'Not Received' },
    { value: 'DELIVERY_DELAY',               label: 'Lost in Transit / Delayed' },
    { value: 'RETURNED_SHIPMENT',            label: 'Returned' },
    { value: 'ADDRESS_CHANGE',               label: 'Address Change' },
    { value: 'COURTESY_DISLIKE',             label: 'Disliked' },
    { value: 'COURTESY_FORGET_UPDATE_QUEUE', label: 'Forgot to update queue' },
    { value: 'COURTESY_GENDER_ISSUE',        label: 'Gender issue' },
    { value: 'OTHER',                        label: 'Other' },
  ];

  const REPLACEMENT_BOMS = [
    { value: 'FEMALE_FIRST_MONTH_SET',     label: 'Welcome Kit' },
    { value: 'FEMALE_RECURRENT_MONTH_SET', label: 'Recurring Order' },
    { value: 'ECOMMERCE_MALE',             label: 'Ecommerce' },
  ];

  /**
   * Creates upcharge billing UI. Attach to a form with a confirm button.
   * Returns an object with { el, refresh() } — call refresh() when items change.
   */
  function createUpchargeManager(userId, itemChecks, confirmBtn) {
    const wrap = document.createElement('div');
    wrap.style.cssText = 'margin-top:10px;';

    const infoEl = document.createElement('div');
    infoEl.style.cssText = 'display:none;background:rgba(245,158,11,0.1);border:1px solid rgba(245,158,11,0.3);border-radius:6px;padding:8px 10px;margin-bottom:8px;font-size:12px;color:#f59e0b;line-height:1.6;';
    wrap.appendChild(infoEl);

    const btnRow = document.createElement('div');
    btnRow.style.cssText = 'display:none;gap:6px;';
    wrap.appendChild(btnRow);

    const billBtn = document.createElement('button');
    billBtn.textContent = '💳 Bill Upcharge';
    billBtn.style.cssText = 'flex:1;padding:7px;border-radius:6px;border:none;background:#f59e0b;color:#000;font-weight:700;font-size:12px;cursor:pointer;';
    billBtn.onmouseenter = () => billBtn.style.background = '#d97706';
    billBtn.onmouseleave = () => billBtn.style.background = '#f59e0b';
    btnRow.appendChild(billBtn);

    const skipBtn = document.createElement('button');
    skipBtn.textContent = 'Billed Already';
    skipBtn.style.cssText = 'padding:7px 12px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);background:transparent;color:#94a3b8;font-size:12px;cursor:pointer;';
    btnRow.appendChild(skipBtn);

    const statusEl = makeStatusEl();
    wrap.appendChild(statusEl);

    let _upchargeCleared = false;

    function getUpchargeTotal() {
      return itemChecks
        .filter(({ chk }) => chk.checked)
        .reduce((sum, { item }) => {
          const pi = item.product?.productInfo || item.productInfo;
          const up = pi?.upchargePrice || 0;
          return sum + up;
        }, 0);
    }

    function refresh() {
      const total = getUpchargeTotal();
      if (total > 0 && !_upchargeCleared) {
        infoEl.style.display = '';
        infoEl.textContent = '⚠ Upcharge required: $' + total + ' — bill before confirming.';
        btnRow.style.display = 'flex';
        confirmBtn.disabled = true;
        confirmBtn.style.opacity = '0.4';
        confirmBtn.style.cursor = 'not-allowed';
      } else {
        infoEl.style.display = 'none';
        btnRow.style.display = 'none';
        // Only re-enable if upcharge is cleared or no upcharge
        if (_upchargeCleared || total === 0) {
          confirmBtn.disabled = false;
          confirmBtn.style.opacity = '1';
          confirmBtn.style.cursor = 'pointer';
        }
      }
    }

    billBtn.onclick = () => {
      const total = getUpchargeTotal();
      if (total <= 0) return;

      const { overlay, box } = createModal({ id: 'sb-upcharge-modal', title: '💳 Bill Upcharge', width: 320 });

      const detailsEl = document.createElement('div');
      detailsEl.style.cssText = 'background:#0f172a;border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:12px;color:#cbd5e1;line-height:1.8;';

      // List upcharge products
      let detailsHtml = '';
      itemChecks.filter(({ chk }) => chk.checked).forEach(({ item }) => {
        const pi = item.product?.productInfo || item.productInfo;
        const up = pi?.upchargePrice || 0;
        if (up > 0) {
          detailsHtml += `<div>${pi.name} by ${pi.brand} — <span style="color:#f59e0b;font-weight:600;">+$${up}</span></div>`;
        }
      });
      detailsHtml += `<div style="margin-top:6px;padding-top:6px;border-top:1px solid rgba(255,255,255,0.07);font-weight:700;font-size:13px;color:#e2e8f0;">Total: $${total}</div>`;
      detailsEl.innerHTML = detailsHtml;
      box.appendChild(detailsEl);

      const chargeStatus = makeStatusEl();
      box.appendChild(chargeStatus);

      box.appendChild(makeDialogButtons({
        confirmLabel: 'Charge $' + total,
        confirmColor: '#f59e0b',
        onCancel: () => overlay.remove(),
        onConfirm: (cBtn, cancelBtn) => {
          cBtn.disabled = true;
          cancelBtn.disabled = true;
          cBtn.textContent = 'Charging...';

          const amountCents = total * 100;
          gqlMutate('charge', CHARGE_MUTATION,
            { input: { amount: String(amountCents), comment: 'Upcharge billing', paymentMethodId: null, userId } },
            (data, err) => {
              const respErr = data?.charge?.error;
              const errMsg = err || respErr?.message;
              if (errMsg) {
                chargeStatus.style.color = '#fca5a5';
                chargeStatus.textContent = '\u2718 ' + errMsg;
                cBtn.disabled = false;
                cancelBtn.disabled = false;
                cBtn.textContent = 'Charge $' + total;
                return;
              }

              const state = data?.charge?.data?.state || 'UNKNOWN';
              const paid = data?.charge?.data?.paid;
              if (state === 'PAID') {
                chargeStatus.style.color = '#6ee7b7';
                chargeStatus.textContent = '\u2714 Charged successfully' + (paid ? ' — $' + (paid / 100).toFixed(2) + ' (incl. tax)' : '');
                cBtn.textContent = '\u2714 Done';
                _upchargeCleared = true;
                statusEl.style.color = '#6ee7b7';
                statusEl.textContent = '\u2714 Upcharge billed — $' + total;
                refresh();
                setTimeout(() => overlay.remove(), 1800);
              } else {
                chargeStatus.style.color = '#fca5a5';
                chargeStatus.textContent = '\u2718 Charge state: ' + state;
                cBtn.disabled = false;
                cancelBtn.disabled = false;
                cBtn.textContent = 'Charge $' + total;
              }
            }
          );
        },
      }));
    };

    skipBtn.onclick = () => {
      _upchargeCleared = true;
      statusEl.style.color = '#94a3b8';
      statusEl.textContent = 'Upcharge skipped — billed separately.';
      refresh();
    };

    return { el: wrap, refresh };
  }

  function showProductSearchPanel(itemChecks, itemsWrap, onItemAdded) {
    const existingPanel = document.getElementById('sb-product-search-panel');
    if (existingPanel) { existingPanel.remove(); return; }

    // Position to the left of the replacement form overlay
    const panel = document.createElement('div');
    panel.id = 'sb-product-search-panel';
    panel.style.cssText = `
      all:initial;position:fixed;top:50%;left:50%;transform:translate(-110%, -50%);
      width:520px;max-height:80vh;overflow-y:auto;
      background:#1e1e2e;color:#e2e8f0;font-family:Arial,sans-serif;font-size:13px;
      border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,0.5);padding:14px;
      z-index:10000001;border:1px solid rgba(255,255,255,0.1);box-sizing:border-box;
    `;

    // Header
    const hdr = document.createElement('div');
    hdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;';
    hdr.innerHTML = '<span style="font-weight:700;font-size:14px;">🔍 Product Search</span>';
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close \u2715';
    closeBtn.style.cssText = 'background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.15);color:#94a3b8;font-size:11px;font-weight:600;cursor:pointer;padding:3px 10px;border-radius:20px;';
    closeBtn.onmouseenter = () => closeBtn.style.background = 'rgba(255,255,255,0.15)';
    closeBtn.onmouseleave = () => closeBtn.style.background = 'rgba(255,255,255,0.08)';
    closeBtn.onclick = () => panel.remove();
    hdr.appendChild(closeBtn);
    panel.appendChild(hdr);

    // Search input row
    const inputRow = document.createElement('div');
    inputRow.style.cssText = 'display:flex;gap:6px;margin-bottom:10px;';
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search by name...';
    searchInput.style.cssText = 'flex:1;padding:7px 10px;border-radius:6px;border:1px solid rgba(255,255,255,0.15);background:#2a2a3e;color:#e2e8f0;font-size:13px;outline:none;box-sizing:border-box;';
    const searchBtn = document.createElement('button');
    searchBtn.textContent = 'Search';
    searchBtn.style.cssText = 'padding:7px 14px;border-radius:6px;border:none;background:#4f46e5;color:#fff;font-weight:700;font-size:13px;cursor:pointer;';
    inputRow.appendChild(searchInput);
    inputRow.appendChild(searchBtn);
    panel.appendChild(inputRow);

    const resultsEl = document.createElement('div');
    panel.appendChild(resultsEl);

    function doSearch() {
      const q = searchInput.value.trim();
      if (!q || q.length < 2) { resultsEl.innerHTML = '<div style="color:#94a3b8;font-size:12px;">Type at least 2 characters.</div>'; return; }
      resultsEl.innerHTML = '<div style="color:#94a3b8;font-size:12px;">Searching…</div>';

      GM_xmlhttpRequest({
        method: 'POST', url: GRAPHQL_URL,
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
        data: JSON.stringify({
          operationName: 'productSuggestion',
          query: PRODUCT_SEARCH_QUERY,
          variables: { input: { name: q, sections: ['Subscription', 'Extras', 'AddonSubscription'], statuses: ['LIVE', 'OUT_OF_STOCK', 'NOT_AVAILABLE_FOR_NEW_ORDERS'] } },
        }),
        onload(res) {
          console.log('[BirdsEye] Product search response:', res.status, res.responseText?.substring(0, 500));
          if (res.status === 401) { handle401(); resultsEl.innerHTML = '<div style="color:#fca5a5;">Token expired.</div>'; return; }
          if (res.status === 403) { handle403(); resultsEl.innerHTML = '<div style="color:#fca5a5;">CRM captcha required.</div>'; return; }
          try {
            const json = JSON.parse(res.responseText);
            console.log('[BirdsEye] Product search parsed:', json?.data?.productSuggestion?.data?.length, 'results, error:', json?.data?.productSuggestion?.error);
            const products = json?.data?.productSuggestion?.data || [];
            const err = json?.data?.productSuggestion?.error?.message;
            if (err) { resultsEl.innerHTML = `<div style="color:#fca5a5;">${err}</div>`; return; }
            if (!products.length) { resultsEl.innerHTML = '<div style="color:#94a3b8;font-size:12px;">No products found.</div>'; return; }
            renderProductResults(products);
          } catch(e) {
            resultsEl.innerHTML = '<div style="color:#fca5a5;">Parse error.</div>';
          }
        },
        onerror() { resultsEl.innerHTML = '<div style="color:#fca5a5;">Network error.</div>'; },
      });
    }

    function renderProductResults(products) {
      resultsEl.innerHTML = '';
      const count = document.createElement('div');
      count.style.cssText = 'color:#94a3b8;margin-bottom:8px;font-size:11px;';
      count.textContent = products.length + ' result(s)';
      resultsEl.appendChild(count);

      // Split into sections
      const subProducts = products.filter(p => p.section === 'Subscription' || p.section === 'Extras');
      const sampleProducts = products.filter(p => p.section === 'AddonSubscription');

      const columnsWrap = document.createElement('div');

      function renderColumn(title, items, container) {
        const col = document.createElement('div');
        col.style.cssText = 'flex:1;min-width:0;';
        const hdr = document.createElement('div');
        hdr.style.cssText = 'font-size:10px;font-weight:600;color:#64748b;letter-spacing:0.05em;margin-bottom:6px;';
        hdr.textContent = title;
        col.appendChild(hdr);

        if (!items.length) {
          const empty = document.createElement('div');
          empty.style.cssText = 'font-size:11px;color:#475569;';
          empty.textContent = 'No results';
          col.appendChild(empty);
        }

        items.forEach(product => {
          const info = product.productInfo;
          if (!info) return;

          const card = document.createElement('div');
          card.style.cssText = `
            padding:6px 8px;margin-bottom:4px;border-radius:5px;background:#2a2a3e;
            border:1px solid transparent;transition:0.15s;display:flex;justify-content:space-between;align-items:center;gap:6px;
          `;
          card.onmouseenter = () => card.style.borderColor = 'rgba(99,102,241,0.4)';
          card.onmouseleave = () => card.style.borderColor = 'transparent';

          const leftCol = document.createElement('div');
          leftCol.style.cssText = 'flex:1;min-width:0;';

          const nameEl = document.createElement('div');
          nameEl.style.cssText = 'font-size:11px;font-weight:600;color:#e2e8f0;';
          nameEl.textContent = info.name + ' by ' + info.brand;
          leftCol.appendChild(nameEl);

          const tagsRow = document.createElement('div');
          tagsRow.style.cssText = 'display:flex;gap:3px;margin-top:2px;flex-wrap:wrap;align-items:center;';

          // Add-on tag for Extras
          if (product.section === 'Extras') {
            const addonBadge = document.createElement('span');
            addonBadge.style.cssText = 'font-size:9px;padding:1px 4px;border-radius:3px;background:rgba(167,139,250,0.15);color:#c4b5fd;border:1px solid rgba(167,139,250,0.3);';
            addonBadge.textContent = 'add-on';
            tagsRow.appendChild(addonBadge);
          }

          // Upcharge badge
          if (info.upchargePrice && info.upchargePrice > 0) {
            const upBadge = document.createElement('span');
            upBadge.style.cssText = 'font-size:9px;padding:1px 4px;border-radius:3px;background:rgba(245,158,11,0.15);color:#f59e0b;border:1px solid rgba(245,158,11,0.3);';
            upBadge.textContent = '+$' + info.upchargePrice;
            tagsRow.appendChild(upBadge);
          }

          // Stock badge
          if (product.status && product.status !== 'LIVE') {
            const stockBadge = document.createElement('span');
            const stockLabel = product.status.replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
            stockBadge.style.cssText = 'font-size:9px;padding:1px 4px;border-radius:3px;background:rgba(239,68,68,0.15);color:#ef4444;border:1px solid rgba(239,68,68,0.3);';
            stockBadge.textContent = stockLabel;
            tagsRow.appendChild(stockBadge);
          }

          leftCol.appendChild(tagsRow);
          card.appendChild(leftCol);

          // Add button
          const vol = product.volume;
          const unit = product.volumeUnit || '';
          const addBtn = document.createElement('button');
          addBtn.textContent = '＋';
          addBtn.title = 'Add to replacement';
          addBtn.style.cssText = 'padding:3px 8px;border-radius:4px;border:1px solid #6366f1;background:transparent;color:#a5b4fc;font-weight:700;font-size:12px;cursor:pointer;flex-shrink:0;';
          addBtn.onmouseenter = () => addBtn.style.background = 'rgba(99,102,241,0.15)';
          addBtn.onmouseleave = () => addBtn.style.background = 'transparent';
          addBtn.onclick = () => {
            const fakeItem = {
              id: product.id,
              product: { id: product.id, status: product.status, productInfo: info },
            };

            const chk = document.createElement('input');
            chk.type = 'checkbox'; chk.checked = true;
            chk.style.cssText = 'cursor:pointer;accent-color:#818cf8;flex-shrink:0;';
            itemChecks.push({ item: fakeItem, chk, starterChk: null });
            chk.addEventListener('change', () => { if (onItemAdded) onItemAdded(); });

            const row = document.createElement('div');
            row.style.cssText = 'margin-bottom:6px;border-left:2px solid #6366f1;padding-left:6px;';
            const lbl = document.createElement('label');
            lbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:12px;color:#a5b4fc;flex-wrap:wrap;';
            lbl.appendChild(chk);
            lbl.appendChild(document.createTextNode(info.name + ' by ' + info.brand));

            if (vol) {
              const volTag = document.createElement('span');
              const volText = unit === 'oz' ? vol.toFixed(2) + 'oz' : vol + unit;
              volTag.style.cssText = 'font-size:10px;color:#94a3b8;';
              volTag.textContent = '(' + volText + ')';
              lbl.appendChild(volTag);
            }

            const addedTag = document.createElement('span');
            addedTag.style.cssText = 'font-size:10px;padding:1px 5px;border-radius:3px;background:rgba(99,102,241,0.2);color:#a5b4fc;border:1px solid rgba(99,102,241,0.3);';
            addedTag.textContent = 'Added';
            lbl.appendChild(addedTag);

            row.appendChild(lbl);
            itemsWrap.appendChild(row);

            addBtn.textContent = '✔';
            addBtn.style.color = '#6ee7b7';
            addBtn.style.borderColor = '#6ee7b7';
            addBtn.disabled = true;
            if (onItemAdded) onItemAdded();
          };

          card.appendChild(addBtn);
          col.appendChild(card);
        });

        container.appendChild(col);
      }

      if (sampleProducts.length) {
        columnsWrap.style.cssText = 'display:flex;gap:8px;';
        renderColumn('SUBSCRIPTION (0.27oz)', subProducts, columnsWrap);
        renderColumn('SAMPLES (1.5ml)', sampleProducts, columnsWrap);
      } else {
        columnsWrap.style.cssText = '';
        renderColumn('SUBSCRIPTION (0.27oz)', subProducts, columnsWrap);
      }
      resultsEl.appendChild(columnsWrap);
    }

    searchBtn.onclick = doSearch;
    searchInput.addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });

    document.body.appendChild(panel);
    setTimeout(() => searchInput.focus(), 50);
  }

  function showReplacementForm(order, user, totalOrders) {
    const existingForm = document.getElementById('sb-replacement-form');
    if (existingForm) { existingForm.remove(); document.getElementById('sb-product-search-panel')?.remove(); return; }

    const isEcomm = order.type && order.type !== 'SUBSCRIPTION' && order.type !== 'MANUAL';
    const monthLabel = order.month
      ? new Date(order.year, order.month - 1).toLocaleString('en-US', { month:'short', year:'numeric' })
      : 'Order';

    const { overlay, box: form } = createModal({
      id: 'sb-replacement-form',
      title: '\uD83D\uDD04 Replacement — ' + monthLabel,
      width: 400,
    });

    // Reason
    form.appendChild(makeLabel('REASON'));
    const reasonSel = makeSelect(REPLACEMENT_REASONS, 'DAMAGED_VIAL');
    form.appendChild(reasonSel);

    // BOM
    const availableBoms = isEcomm
      ? REPLACEMENT_BOMS
      : REPLACEMENT_BOMS.filter(b => b.value !== 'ECOMMERCE');
    const defaultBom = isEcomm ? 'ECOMMERCE_MALE'
      : (totalOrders <= 1 ? 'FEMALE_FIRST_MONTH_SET' : 'FEMALE_RECURRENT_MONTH_SET');
    form.appendChild(makeLabel('ORDER TYPE (BOM)'));
    const bomSel = makeSelect(availableBoms, defaultBom);
    form.appendChild(bomSel);

    // Items
    form.appendChild(makeLabel('ITEMS TO REPLACE'));
    const itemsWrap = document.createElement('div');
    itemsWrap.style.cssText = 'background:#0f172a;border-radius:6px;padding:8px 10px;';
    const itemChecks = [];

    (order.orderItems || []).forEach(item => {
      const info = item?.product?.productInfo;
      const name = info ? `${info.name} by ${info.brand}` : `Item #${item.id}`;
      const isDrift = /drift|wood car freshener/i.test(name);
      const stockStatus = item?.product?.status;
      const stockBadge = {
        'OUT_OF_STOCK':              { text: 'Out of Stock',    color: '#ef4444' },
        'NOT_AVAILABLE_FOR_NEW_ORDERS': { text: 'Not Available', color: '#f97316' },
        'INTERNAL_USE':              { text: 'Internal Use',    color: '#eab308' },
      }[stockStatus];

      const row = document.createElement('div');
      row.style.cssText = 'margin-bottom:6px;';

      const lbl = document.createElement('label');
      lbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:12px;color:#cbd5e1;flex-wrap:wrap;';
      const chk = document.createElement('input');
      chk.type = 'checkbox'; chk.checked = true;
      chk.style.cssText = 'cursor:pointer;accent-color:#818cf8;flex-shrink:0;';
      lbl.appendChild(chk);
      lbl.appendChild(document.createTextNode(name));
      if (stockBadge) {
        const badge = document.createElement('span');
        badge.style.cssText = `font-size:10px;font-weight:600;color:${stockBadge.color};white-space:nowrap;`;
        badge.textContent = '● ' + stockBadge.text;
        lbl.appendChild(badge);
      }
      row.appendChild(lbl);

      let starterChk = null;
      if (isDrift) {
        const starterRow = document.createElement('div');
        starterRow.style.cssText = 'margin-left:22px;margin-top:3px;';
        const sLbl = document.createElement('label');
        sLbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:11px;color:#94a3b8;';
        starterChk = document.createElement('input');
        starterChk.type = 'checkbox'; starterChk.checked = true;
        starterChk.style.cssText = 'cursor:pointer;accent-color:#818cf8;';
        sLbl.appendChild(starterChk);
        sLbl.appendChild(document.createTextNode('Starter Set'));
        starterRow.appendChild(sLbl);
        row.appendChild(starterRow);
      }

      itemsWrap.appendChild(row);
      itemChecks.push({ item, chk, starterChk });
      chk.addEventListener('change', () => { if (typeof upchargeUI !== 'undefined') upchargeUI.refresh(); });
    });
    form.appendChild(itemsWrap);

    // Product search button
    const searchRow = document.createElement('div');
    searchRow.style.cssText = 'margin-top:8px;';
    const searchProductBtn = document.createElement('button');
    searchProductBtn.textContent = '🔍 Search Product';
    searchProductBtn.style.cssText = 'padding:5px 10px;border-radius:5px;border:1px solid #818cf8;background:transparent;color:#818cf8;font-weight:600;font-size:11px;cursor:pointer;';
    searchProductBtn.onmouseenter = () => searchProductBtn.style.background = 'rgba(129,140,248,0.15)';
    searchProductBtn.onmouseleave = () => searchProductBtn.style.background = 'transparent';
    searchProductBtn.onclick = () => showProductSearchPanel(itemChecks, itemsWrap, () => upchargeUI.refresh());
    searchRow.appendChild(searchProductBtn);
    form.appendChild(searchRow);

    // Comment
    form.appendChild(makeLabel('ADDITIONAL COMMENT (OPTIONAL)'));
    const commentInput = document.createElement('textarea');
    commentInput.placeholder = 'Add a note...';
    commentInput.rows = 2;
    commentInput.style.cssText = 'width:100%;background:#0f172a;color:#e2e8f0;border:1px solid rgba(255,255,255,0.1);border-radius:5px;padding:6px 8px;font-size:12px;box-sizing:border-box;resize:vertical;font-family:Arial,sans-serif;';
    form.appendChild(commentInput);

    // Options
    form.appendChild(makeLabel('OPTIONS'));
    const optsWrap = document.createElement('div');
    optsWrap.style.cssText = 'background:#0f172a;border-radius:6px;padding:8px 10px;display:flex;flex-direction:column;gap:6px;';

    function makeCheckRow(labelText, defaultChecked = false) {
      const lbl = document.createElement('label');
      lbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:12px;color:#cbd5e1;';
      const chk = document.createElement('input');
      chk.type = 'checkbox'; chk.checked = defaultChecked;
      chk.style.cssText = 'cursor:pointer;accent-color:#818cf8;';
      const warn = document.createElement('span');
      warn.textContent = '⚠';
      warn.style.cssText = 'color:#f59e0b;font-size:14px;display:none;margin-left:2px;';
      lbl.appendChild(chk);
      lbl.appendChild(document.createTextNode(labelText));
      lbl.appendChild(warn);
      optsWrap.appendChild(lbl);
      chk._warn = warn;
      return chk;
    }

    const picturesChk = makeCheckRow('Pictures provided');
    const addressChk  = makeCheckRow('Address confirmed');
    form.appendChild(optsWrap);

    // Live validation warnings
    const ADDR_REASONS = ['NOT_RECEIVED', 'DELIVERY_DELAY', 'RETURNED_SHIPMENT', 'ADDRESS_CHANGE'];
    const PIC_REASONS  = ['DAMAGED_LEAKING', 'DAMAGED_VIAL', 'WRONG_ITEM'];

    function updateWarnings() {
      const reason = reasonSel.value;
      addressChk._warn.style.display = (ADDR_REASONS.includes(reason) && !addressChk.checked) ? 'inline' : 'none';
      picturesChk._warn.style.display = (PIC_REASONS.includes(reason) && !picturesChk.checked) ? 'inline' : 'none';
    }

    reasonSel.addEventListener('change', updateWarnings);
    addressChk.addEventListener('change', updateWarnings);
    picturesChk.addEventListener('change', updateWarnings);
    updateWarnings(); // initial check

    // Status
    const statusEl = makeStatusEl();
    form.appendChild(statusEl);

    // Confirm button
    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'Confirm Replacement';
    confirmBtn.style.cssText = `
      margin-top:14px;width:100%;padding:9px;border-radius:6px;
      border:none;background:#6366f1;color:#fff;font-weight:700;
      font-size:13px;cursor:pointer;box-sizing:border-box;
    `;
    confirmBtn.onmouseenter = () => confirmBtn.style.background = '#4f46e5';
    confirmBtn.onmouseleave = () => confirmBtn.style.background = '#6366f1';

    // Upcharge manager (inserted before confirm button)
    const upchargeUI = createUpchargeManager(user.id, itemChecks, confirmBtn);
    form.appendChild(upchargeUI.el);
    upchargeUI.refresh();

    confirmBtn.onclick = () => {
      const selectedItems = itemChecks
        .filter(({ chk }) => chk.checked)
        .map(({ item, starterChk }) => ({
          id: item.id,
          starterSet: starterChk ? starterChk.checked : false,
          product: { id: item.product?.id },
        }));

      if (!selectedItems.length) {
        statusEl.style.color = '#fca5a5';
        statusEl.textContent = 'Select at least one item.';
        return;
      }

      confirmBtn.disabled = true;
      confirmBtn.textContent = 'Submitting...';
      statusEl.textContent = '';

      const variables = {
        input: {
          zendesk: window.location.href,
          reason: reasonSel.value,
          bom: bomSel.value,
          customerId: user.id,
          addressConfirmed: addressChk.checked,
          picturesProvided: picturesChk.checked,
          orderNumber: order.orderNumber,
          sourceOrder: {
            id: order.id,
            orderItems: selectedItems.map(i => ({
              id: i.id, starterSet: i.starterSet, product: { id: i.product.id },
            })),
          },
        },
      };

      GM_xmlhttpRequest({
        method: 'POST', url: GRAPHQL_URL,
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
        data: JSON.stringify({ operationName: 'replacementSave', query: REPLACEMENT_MUTATION, variables }),
        onload(res) {
          if (res.status === 401) {
            handle401();
            statusEl.style.color = '#fca5a5';
            statusEl.textContent = '\u2718 Token expired — click 🔑 Token to update.';
            confirmBtn.disabled = false;
            confirmBtn.textContent = 'Confirm Replacement';
            return;
          }
          if (res.status === 403) {
            handle403();
            statusEl.style.color = '#fca5a5';
            statusEl.textContent = '\u2718 CRM captcha required — open CRM in browser.';
            confirmBtn.disabled = false;
            confirmBtn.textContent = 'Confirm Replacement';
            return;
          }
          try {
            const json = JSON.parse(res.responseText);
            const err = json?.data?.replacementSave?.error?.message;
            if (err) {
              statusEl.style.color = '#fca5a5';
              statusEl.textContent = '\u2718 ' + err;
              confirmBtn.disabled = false;
              confirmBtn.textContent = 'Confirm Replacement';
              return;
            }

            // Build comment text
            const reasonLabel = REPLACEMENT_REASONS.find(r => r.value === reasonSel.value)?.label || reasonSel.value;
            const orderDate = new Date(order.year, (order.month || 1) - 1);
            const monthStr = isEcomm
              ? `Ecommerce ${order.orderNumber || ''}`
              : orderDate.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            const picStr  = picturesChk.checked ? 'Yes' : 'No';
            const addrStr = addressChk.checked  ? 'Yes' : 'No';
            const extra   = commentInput.value.trim();
            const commentText = [
              `Replacement - ${reasonLabel} - ${monthStr}`,
              `Pictures provided: ${picStr} | Address confirmed: ${addrStr}`,
              extra || null,
            ].filter(Boolean).join('\n');

            statusEl.style.color = '#94a3b8';
            statusEl.textContent = 'Finding automated comment...';

            // Find and delete auto-generated comment, then post ours
            fetchComments(user.id, (comments) => {
              const autoComment = comments.find(c =>
                c.comment && order.orderNumber && c.comment.includes(order.orderNumber)
              );

              const postComment = () => {
                statusEl.textContent = 'Posting comment...';
                gqlMutate('createUserComment', CREATE_COMMENT_MUTATION,
                  { input: { userId: user.id, comment: commentText, zendeskUrl: window.location.href } },
                  (data, err) => {
                    if (err) {
                      statusEl.style.color = '#f59e0b';
                      statusEl.textContent = '\u2714 Replacement created, but comment failed.';
                    } else {
                      statusEl.style.color = '#6ee7b7';
                      statusEl.textContent = '\u2714 Done!';
                    }
                    confirmBtn.textContent = '\u2714 Done';
                    setTimeout(() => { overlay.remove(); document.getElementById('sb-product-search-panel')?.remove(); }, 2000);
                  }
                );
              };

              if (autoComment) {
                statusEl.textContent = 'Removing automated comment...';
                gqlMutate('deleteUserComment', DELETE_COMMENT_MUTATION, { id: autoComment.id },
                  () => postComment()
                );
              } else {
                postComment();
              }
            });
          } catch(e) {
            statusEl.style.color = '#fca5a5';
            statusEl.textContent = '\u2718 Unexpected error.';
            confirmBtn.disabled = false;
            confirmBtn.textContent = 'Confirm Replacement';
          }
        },
        onerror() {
          statusEl.style.color = '#fca5a5';
          statusEl.textContent = '\u2718 Network error.';
          confirmBtn.disabled = false;
          confirmBtn.textContent = 'Confirm Replacement';
        },
      });
    };

    form.appendChild(confirmBtn);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // FILL NAME (Button 1)
  // ══════════════════════════════════════════════════════════════════════════

  function registerFillName() {
    if (fillNameRegistered) return;
    fillNameRegistered = true;

    addToolbarButton('sb-fill-name-btn', 'Fill Name', async (btn) => {
      const orig = btn.textContent;
      btn.textContent = '⏳'; btn.disabled = true;

      // Open modal first — we need it for the name field anyway
      if (!document.querySelector('input[data-kt="customerModalNameField"]')) {
        document.querySelector('button[data-kt="customerTimelineEditCustomerProfileButton"]')?.click();
      }
      const nameField = await waitForField('input[data-kt="customerModalNameField"]');

      // Resolve customer from CRM if not cached
      const startId = getCustomerIdFromURL();
      const getUser = () => new Promise(async (resolve) => {
        if (cachedCustomerCtx?.user) return resolve(cachedCustomerCtx.user);
        const email = await resolveEmail();
        if (getCustomerIdFromURL() !== startId) return resolve(null); // stale
        if (!email) return resolve(null);
        searchCRM(email, (users, err) => {
          if (getCustomerIdFromURL() !== startId) return resolve(null); // stale
          if (err || !users?.length) return resolve(null);
          const sbUsers = users.filter(u => !u.origin || u.origin === 'SCENTBIRD');
          const exactMatch = sbUsers.find(u => u.email?.toLowerCase() === email.toLowerCase());
          if (exactMatch) cachedCustomerCtx = { email, user: exactMatch, _kustomerId: getCustomerIdFromURL() };
          resolve(exactMatch || null);
        });
      });

      const user = await getUser();
      const name = user ? toProperCase([user.firstName, user.lastName].filter(Boolean).join(' ')) : null;

      if (name && nameField) {
        setReactValue(nameField, name);
        await new Promise(r => setTimeout(r, 300));
        document.querySelector('button[data-kt="modalFooterBasic_buttonPrimary"]')?.click();
        btn.textContent = '✔ Done'; btn.style.color = '#6ee7b7';
      } else if (nameField && nameField.value.trim()) {
        // Fallback: title-case the existing name
        setReactValue(nameField, toProperCase(nameField.value.trim()));
        await new Promise(r => setTimeout(r, 300));
        document.querySelector('button[data-kt="modalFooterBasic_buttonPrimary"]')?.click();
        btn.textContent = '✔ Cased'; btn.style.color = '#fcd34d';
      } else {
        btn.textContent = '✘ Not Found'; btn.style.color = '#fca5a5';
      }
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; }, 2000);
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // VARIABLE FILL
  // ══════════════════════════════════════════════════════════════════════════

  const FILL_VARIABLES = [
    {
      key: 'ADDRESS', label: 'Address',
      resolve({ user }) {
        const s = user?.userAddress?.shipping;
        if (!s) return null;
        return [s.street1, s.city, s.region, s.postcode].filter(Boolean).join(', ') || null;
      },
    },
    {
      key: 'DATE', label: 'Tracking date',
      resolve({ order }) {
        if (!order) return null;
        const items = order.tracking?.items;
        if (!items?.length) return null;
        const last = items[items.length - 1];
        return last?.date ? new Date(last.date).toLocaleDateString('en-US', { month: 'long', day: 'numeric' }) : null;
      },
    },
    {
      key: 'FRAGRANCE', label: 'Fragrances',
      resolve({ order }) {
        if (!order) return null;
        const lines = orderProductLines(order);
        if (!lines.length) return null;
        // Append case line for Welcome Kit orders
        const isWelcomeKit = (order.tags || []).includes('REBRAND_CASE') || order.hasRebrandCase;
        if (isWelcomeKit) lines.push('Signature Grey Fragrance Case (welcome gift)');
        return lines.join('\n');
      },
    },
    {
      key: 'MONTH', label: 'Order month',
      resolve({ order }) {
        if (!order) return null;
        const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
        return order.month ? months[order.month - 1] : null;
      },
    },
    {
      key: 'TRACKING', label: 'Tracking number',
      resolve({ order }) {
        if (!order) return null;
        const trackNo  = order.tracking?.trackingNumber || order.shipment?.trackingNumber || null;
        if (!trackNo) return null;
        const trackUrl = order.tracking?.trackingUrl || order.shipment?.trackingUrl || null;
        return trackUrl ? `<a href="${trackUrl}">${trackNo}</a>` : trackNo;
      },
    },
    {
      key: 'BD', label: 'Billing date (ordinal)',
      resolve({ subscription }) {
        if (!subscription) return null;
        const day = subscription.nextBillingDate
          ? new Date(subscription.nextBillingDate).getUTCDate()
          : subscription.cashbirdDetails?.data?.billingDay || null;
        if (!day) return null;
        return ordinal(day);
      },
    },
    {
      key: 'BD-1', label: 'Billing date minus 1 day',
      resolve({ subscription }) {
        if (!subscription) return null;
        const day = subscription.nextBillingDate
          ? new Date(subscription.nextBillingDate).getUTCDate()
          : subscription.cashbirdDetails?.data?.billingDay || null;
        if (!day) return null;
        return ordinal(day === 1 ? 30 : day - 1);
      },
    },
    {
      key: 'DATE_SUBSCRIBED', label: 'Subscription start date',
      resolve({ subscription }) {
        if (!subscription?.subscriptionDate) return null;
        return fmtDateLong(subscription.subscriptionDate);
      },
    },
    {
      key: 'DATE_CANCELLED', label: 'Subscription cancel date',
      resolve({ subscription }) {
        if (!subscription?.subscriptionEndDate) return null;
        return fmtDateLong(subscription.subscriptionEndDate);
      },
    },
    {
      key: 'TOTAL_REFUNDED', label: 'Total refunded amount',
      resolve() {
        if (_totalRefundedCents <= 0) return null;
        return fmtMoney(_totalRefundedCents, _totalRefundedCurrency);
      },
    },
    {
      key: 'BILL_EXPLANATION', label: 'Bill explanation',
      needsCharges: true,
      resolve({ charges }) {
        if (!charges?.length) return null;
        const successEntries = charges.filter(e => e.success);
        if (!successEntries.length) return null;
        return successEntries.map(e => buildBillExplanation(e)).filter(Boolean).join('\n\n') || null;
      },
    },
  ];

  function getComposer() {
    return document.querySelector('.public-DraftEditor-content[contenteditable="true"]');
  }

  function fillComposer(replacements) {
    const composer = getComposer();
    if (!composer) return 0;

    let html = composer.innerHTML;
    let count = 0;

    replacements.forEach(({ key, value }) => {
      const token = '[' + key + ']';
      const str = (value && typeof value === 'object') ? value.text : value;
      if (html.includes(token)) {
        const htmlStr = str.split('\n').map((line, i) =>
          i === 0 ? line : '<div>' + (line || '<br>') + '</div>'
        ).join('');
        html = html.split(token).join(htmlStr);
        count++;
      }
    });

    if (count === 0) return 0;
    GM_setClipboard(html, 'html');
    return count;
  }

  function fillVariables(btn) {
    // Remove existing dropdown if open
    const existingDrop = document.getElementById('sb-fill-dropdown');
    if (existingDrop) { existingDrop.remove(); return; }

    const composer = getComposer();
    if (!composer) {
      const orig = btn.textContent;
      btn.textContent = '✘ No editor'; btn.style.color = '#fca5a5';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; }, 2000);
      return;
    }

    const text = composer.innerText;
    const needed = FILL_VARIABLES.filter(v => text.includes('[' + v.key + ']'));
    if (!needed.length) {
      const orig = btn.textContent;
      btn.textContent = '✘ No tokens'; btn.style.color = '#fca5a5';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; }, 2000);
      return;
    }

    // Fetch orders to build dropdown with real month labels
    const origText = btn.textContent;
    btn.textContent = '⏳'; btn.disabled = true;

    loadCustomer(btn, (result) => {
      if (!result) { btn.textContent = origText; btn.disabled = false; return; }
      const { user } = result;

      loadCustomerOrders(user.id, (orders) => {
        btn.textContent = origText; btn.disabled = false;

        const mainOrders = (orders || []).filter(o => o.type !== 'REPLACEMENT' && o.type !== 'REPLACEMENT_ORDER');
        const monthNames = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

        const options = mainOrders.slice(0, 3).map((o, i) => {
          const label = o.month && o.year
            ? monthNames[o.month - 1] + '-' + String(o.year).slice(2)
            : 'Order ' + (i + 1);
          return { label, index: i };
        });

        if (!options.length) {
          options.push({ label: 'No orders found', index: 0 });
        }

        // Show dropdown
        const rect = btn.getBoundingClientRect();
        const drop = document.createElement('div');
        drop.id = 'sb-fill-dropdown';
        drop.style.cssText = `
          all:initial;position:fixed;z-index:9999999;
          top:${rect.bottom + 4}px;left:${rect.left}px;
          background:#1e1e2e;border:1px solid rgba(255,255,255,0.15);border-radius:7px;
          box-shadow:0 8px 24px rgba(0,0,0,0.5);overflow:hidden;
          font-family:Arial,sans-serif;font-size:12px;min-width:120px;
        `;

        options.forEach(opt => {
          const row = document.createElement('div');
          row.style.cssText = 'padding:8px 14px;cursor:pointer;color:#e2e8f0;';
          row.textContent = opt.label;
          row.onmouseenter = () => row.style.background = 'rgba(99,102,241,0.25)';
          row.onmouseleave = () => row.style.background = '';
          row.onclick = () => {
            drop.remove();
            doFill(btn, needed, opt.index);
          };
          drop.appendChild(row);
        });

        document.body.appendChild(drop);

        const closeHandler = (e) => {
          if (!drop.contains(e.target) && e.target !== btn) {
            drop.remove();
            document.removeEventListener('mousedown', closeHandler, true);
          }
        };
        setTimeout(() => document.addEventListener('mousedown', closeHandler, true), 0);
      });
    });
  }

  function doFill(btn, needed, orderIndex) {
    const orig = btn.textContent;
    btn.textContent = '⏳'; btn.disabled = true;
    const restore = (label, color) => {
      btn.textContent = label; btn.style.color = color || '';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; }, 2000);
    };

    const needsOrder        = needed.some(v => ['MONTH','DATE','FRAGRANCE','TRACKING'].includes(v.key));
    const needsSubscription = needed.some(v => ['BD','BD-1','DATE_SUBSCRIBED','DATE_CANCELLED'].includes(v.key));
    const needsCharges      = needed.some(v => v.needsCharges);

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;

      const afterOrders = (order) => {
        const afterSubscription = (subscription) => {
          const afterCharges = (charges) => {
            const data = { user, order, subscription, charges };
            const replacements = needed
              .map(v => ({ key: v.key, value: v.resolve(data) }))
              .filter(r => r.value !== null);

            const count = fillComposer(replacements);
            const skipped = needed.length - replacements.length;
            if (count === 0 && skipped > 0) return restore('✘ No data', '#fca5a5');
            const label = (skipped > 0 ? '⚠ Paste (missing)' : '📋 Paste now');
            restore(label, '#6ee7b7');
          };

          if (needsCharges) {
            loadCustomerCharges(user.id, (chargeResult) => {
              afterCharges(chargeResult?.entries || []);
            });
          } else {
            afterCharges(null);
          }
        };

        if (needsSubscription) {
          GM_xmlhttpRequest({
            method: 'POST', url: GRAPHQL_URL,
            headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
            data: JSON.stringify({ operationName: 'userDetailsById', query: USER_DETAILS_QUERY, variables: { id: user.id } }),
            onload(res) {
              try {
                const json = JSON.parse(res.responseText);
                const list = json?.data?.userById?.data?.subscriptionList || [];
                const active = list.find(s => s.status === 'Active' || s.subscribed) || list[0] || null;
                afterSubscription(active);
              } catch(e) { afterSubscription(null); }
            },
            onerror() { afterSubscription(null); },
          });
        } else {
          afterSubscription(null);
        }
      };

      if (needsOrder) {
        loadCustomerOrders(user.id, (orders) => {
          const mainOrders = (orders || []).filter(o => o.type !== 'REPLACEMENT' && o.type !== 'REPLACEMENT_ORDER');
          const order = mainOrders[orderIndex] || mainOrders[0] || null;
          afterOrders(order);
        });
      } else {
        afterOrders(null);
      }
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SLASH AUTOCOMPLETE (disabled — buggy in Draft.js composer)
  // ══════════════════════════════════════════════════════════════════════════

  let _slashPicker = null;

  function removeSlashPicker() {
    _slashPicker?.remove();
    _slashPicker = null;
  }

  function showSlashPicker(composer, caretRect, filter) {
    removeSlashPicker();

    const matches = FILL_VARIABLES.filter(v =>
      filter === '' || v.key.toLowerCase().startsWith(filter.toLowerCase()) || v.label.toLowerCase().startsWith(filter.toLowerCase())
    );
    if (!matches.length) return;

    const picker = document.createElement('div');
    _slashPicker = picker;
    picker.style.cssText = `
      all:initial;position:fixed;z-index:9999999;
      background:#1e1e2e;border:1px solid rgba(255,255,255,0.15);border-radius:7px;
      box-shadow:0 8px 24px rgba(0,0,0,0.5);overflow:hidden;
      font-family:Arial,sans-serif;font-size:12px;min-width:200px;
    `;
    picker.style.left = caretRect.left + 'px';
    picker.style.top  = (caretRect.bottom + 4) + 'px';

    let selected = 0;

    const rows = matches.map((v, i) => {
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:6px 12px;cursor:pointer;gap:16px;';
      row.innerHTML = `<span style="color:#e2e8f0;font-weight:600;">[${v.key}]</span><span style="color:#64748b;font-size:11px;">${v.label}</span>`;
      row.addEventListener('mouseenter', () => { selected = i; highlight(); });
      row.addEventListener('mousedown', (e) => { e.preventDefault(); insertVariable(v.key); });
      return row;
    });

    function highlight() {
      rows.forEach((r, i) => r.style.background = i === selected ? 'rgba(99,102,241,0.25)' : '');
    }
    rows.forEach(r => picker.appendChild(r));
    highlight();
    document.body.appendChild(picker);

    const keyHandler = (e) => {
      if (!_slashPicker) return;
      if (e.key === 'ArrowDown') { e.preventDefault(); selected = (selected + 1) % rows.length; highlight(); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); selected = (selected - 1 + rows.length) % rows.length; highlight(); }
      else if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); insertVariable(matches[selected].key); }
      else if (e.key === 'Escape') { removeSlashPicker(); }
    };
    document.addEventListener('keydown', keyHandler, true);
    picker._keyHandler = keyHandler;

    function insertVariable(key) {
      removeSlashPicker();
      document.removeEventListener('keydown', keyHandler, true);
      const sel = window.getSelection();
      if (sel?.rangeCount) {
        const range = sel.getRangeAt(0);
        const textNode = range.startContainer;
        if (textNode.nodeType === Node.TEXT_NODE) {
          const text = textNode.nodeValue;
          const slashIdx = text.lastIndexOf('/', range.startOffset);
          if (slashIdx !== -1) {
            const deleteRange = document.createRange();
            deleteRange.setStart(textNode, slashIdx);
            deleteRange.setEnd(textNode, range.startOffset);
            deleteRange.deleteContents();
            sel.collapseToStart();
          }
        }
      }
      document.execCommand('insertText', false, '[' + key + ']');
      composer.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }

  function initSlashAutocomplete() {
    document.addEventListener('input', (e) => {
      const composer = getComposer();
      if (!composer || !composer.contains(e.target)) { removeSlashPicker(); return; }

      const sel = window.getSelection();
      if (!sel?.rangeCount) return;
      const range = sel.getRangeAt(0);
      const textNode = range.startContainer;
      if (textNode.nodeType !== Node.TEXT_NODE) return;

      const textBefore = textNode.nodeValue.slice(0, range.startOffset);
      const slashIdx = textBefore.lastIndexOf('/');
      if (slashIdx === -1) { removeSlashPicker(); return; }

      const charBefore = textBefore[slashIdx - 1];
      if (slashIdx > 0 && charBefore && !/\s/.test(charBefore)) { removeSlashPicker(); return; }

      const filter = textBefore.slice(slashIdx + 1);
      if (/\s/.test(filter)) { removeSlashPicker(); return; }

      const caretRect = range.getBoundingClientRect();
      showSlashPicker(composer, caretRect, filter);
    }, true);

    document.addEventListener('mousedown', (e) => {
      if (_slashPicker && !_slashPicker.contains(e.target)) removeSlashPicker();
    }, true);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // BUTTON REGISTRATION
  // ══════════════════════════════════════════════════════════════════════════

  function registerCrmButtons() {
    addToolbarButton('sb-crm-search-btn', '🔍 CRM Search', () => {
      const existing = document.getElementById('sb-search-panel');
      if (existing) {
        if (stashedSelection) {
          const input = existing.querySelector('input[type="text"]');
          if (input && input.value.trim() !== stashedSelection) {
            input.value = stashedSelection;
            input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
            return;
          }
        }
        removePanel('sb-search-panel'); return;
      }
      showSearchPanel(stashedSelection);
    });
    addToolbarButton('sb-last-order-btn', '📦 Last Order', (btn) => {
      if (document.getElementById('sb-order-panel')) { removePanel('sb-order-panel'); return; }
      showLastOrderPanel(btn);
    });
    addToolbarButton('sb-charges-btn', '💳 Recent Charges', (btn) => {
      if (document.getElementById('sb-charges-panel')) { removePanel('sb-charges-panel'); return; }
      showChargesPanel(btn);
    });
    addToolbarButton('sb-queue-btn', '📋 Queue', (btn) => {
      showQueuePanel(btn);
    });
    addToolbarButton('sb-fill-vars-btn', '✏️ Fill', (btn) => {
      fillVariables(btn);
    });
    addToolbarButton('sb-edit-customer-btn', '✎ Edit Customer', (btn) => {
      showEditCustomerPanel(btn);
    });
    addToolbarButton('sb-cancel-btn', '❌ Cancel Sub', (btn) => {
      showCancelPanel(btn);
    });

    // Spacer to push right-side buttons
    const toolbar = ensureToolbar();
    if (toolbar && !toolbar.querySelector('#sb-toolbar-spacer')) {
      const spacer = document.createElement('div');
      spacer.id = 'sb-toolbar-spacer';
      spacer.style.cssText = 'flex:1;';
      toolbar.appendChild(spacer);
    }

    addToolbarButton('sb-open-account-btn', '🔗 Open Account', (btn) => {
      const user = cachedCustomerCtx?.user;
      if (!user?.id) {
        btn.textContent = '✘ No customer';
        btn.style.color = '#fca5a5';
        setTimeout(() => { btn.textContent = '🔗 Open Account'; btn.style.color = ''; }, 2000);
        return;
      }
      window.open(`https://crm.scentbird.com/user/${user.id}/profile/subscription`, '_blank');
    });
    addToolbarButton('sb-crm-token-btn', '🔑 Token', () => {
      if (document.getElementById('sb-token-panel')) { removePanel('sb-token-panel'); return; }
      showTokenPanel();
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // USER INFO BAR
  // ══════════════════════════════════════════════════════════════════════════

  const PLAN_LABELS = {
    'MONTHLY':         '1 fragrance',
    'MONTHLY_1PC':     '1 fragrance',
    'MONTHLY_2PCS':    '2 fragrances',
    'MONTHLY_3PCS':    '3 fragrances',
    'MONTHLY_4PCS':    '4 fragrances',
    'MONTHLY_5PCS':    '5 fragrances',
    'MONTHLY_6PCS':    '6 fragrances',
    'BIMONTHLY':       'Bimonthly · 1 fragrance',
    'BIMONTHLY_1PC':   'Bimonthly · 1 fragrance',
    'BIMONTHLY_2PCS':  'Bimonthly · 2 fragrances',
    'BIMONTHLY_3PCS':  'Bimonthly · 3 fragrances',
    'TRIMONTHLY':      'Trimonthly · 1 fragrance',
    'TRIMONTHLY_1PC':  'Trimonthly · 1 fragrance',
    'TRIMONTHLY_2PCS': 'Trimonthly · 2 fragrances',
    'TRIMONTHLY_3PCS': 'Trimonthly · 3 fragrances',
  };

  function renderUserInfoBar(sub, addressMatch, orderAddrStr) {
    const existing = document.getElementById('sb-user-info-bar');
    if (existing) existing.remove();

    const bar = document.createElement('div');
    bar.id = 'sb-user-info-bar';
    bar.style.cssText = 'display:flex;align-items:center;flex-wrap:wrap;gap:6px;padding:4px 10px;background:#0f172a;border-top:1px solid rgba(255,255,255,0.06);font-family:Arial,sans-serif;font-size:11px;';

    function makeGroup() {
      const g = document.createElement('div');
      g.style.cssText = 'display:inline-flex;align-items:center;gap:4px;border:1px solid rgba(255,255,255,0.08);border-radius:6px;padding:2px 4px;';
      return g;
    }

    // Group 1: Sub status + BD + Gender (gender appended async later)
    const group1 = makeGroup();
    group1.id = 'sb-info-group-status';

    const status = sub.status || 'Unknown';
    const isActive = status === 'Active';
    const isCancelled = status === 'Cancelled';
    let statusLabel = status;
    if (isActive && sub.subscriptionDate) {
      statusLabel = 'Active since ' + fmtDateTag(sub.subscriptionDate);
    } else if (isCancelled && sub.subscriptionEndDate) {
      statusLabel = 'Cancelled on ' + fmtDateTag(sub.subscriptionEndDate);
    }
    group1.appendChild(makeTag(statusLabel,
      isActive ? '#6ee7b7' : '#fca5a5',
      isActive ? 'rgba(110,231,183,0.1)' : 'rgba(252,165,165,0.1)',
      isActive ? 'rgba(110,231,183,0.3)' : 'rgba(252,165,165,0.3)'
    ));

    if (sub.cashbirdDetails?.data?.isAwaitCancellation) {
      group1.appendChild(makeTag('Scheduled Cancellation', '#ef4444', 'rgba(239,68,68,0.15)', 'rgba(239,68,68,0.4)'));
    }

    const bdDate = sub.nextBillingDate || sub.cashbirdDetails?.data?.nextBillingDate;
    const bdDay  = bdDate
      ? new Date(bdDate).getUTCDate()
      : sub.cashbirdDetails?.data?.billingDay || null;
    if (bdDay) {
      let bdLabel = 'BD: ' + ordinal(bdDay);
      if (bdDate) {
        const mon = new Date(bdDate).toLocaleDateString('en-US', { month: 'short', timeZone: 'UTC' }).toUpperCase();
        bdLabel = 'BD: ' + mon + ' ' + bdDay;
      }
      group1.appendChild(makeTag(bdLabel, '#94a3b8', 'rgba(148,163,184,0.08)', 'rgba(148,163,184,0.2)'));
    }

    // Gender tag from cache
    const cachedGender = cachedCustomerCtx?._gender;
    if (cachedGender) {
      const isMale = cachedGender === 'MALE';
      const gTag = makeTag(
        isMale ? '♂ Colognes' : '♀ Perfumes',
        isMale ? '#60a5fa' : '#f472b6',
        isMale ? 'rgba(96,165,250,0.1)' : 'rgba(244,114,182,0.1)',
        isMale ? 'rgba(96,165,250,0.3)' : 'rgba(244,114,182,0.3)'
      );
      gTag.dataset.tag = 'gender';
      group1.appendChild(gTag);
    }

    bar.appendChild(group1);

    // Group 2: Plan + add-ons
    const group2 = makeGroup();

    if (sub.planName) {
      const planLabel = PLAN_LABELS[sub.planName] || sub.planName.replace(/_/g, ' ').toLowerCase();
      group2.appendChild(makeTag(planLabel, '#c4b5fd', 'rgba(167,139,250,0.1)', 'rgba(167,139,250,0.25)'));
    }

    const addOnLabels = {
      caseSubscription:         'Case',
      samplesSubscription:      'Samples',
      candleSubscription:       'Candle',
      carScentSubscription:     'Car Scent',
      homeDiffuserSubscription: 'Diffuser',
    };
    const addOns = sub.addOnSettings || {};
    Object.entries(addOnLabels).forEach(([key, label]) => {
      if (addOns[key]?.selected) {
        group2.appendChild(makeTag(label, '#fcd34d', 'rgba(245,158,11,0.1)', 'rgba(245,158,11,0.25)'));
      }
    });

    if (group2.children.length) bar.appendChild(group2);

    // Group 3: Warnings (addr, fraud, chargeback — fraud/chargeback appended async)
    const group3 = makeGroup();
    group3.id = 'sb-info-group-warnings';
    group3.style.display = 'none'; // hidden until a warning is added

    if (addressMatch === false) {
      const addrTag = makeTag('⚠ Addr Updated', '#f59e0b', 'rgba(245,158,11,0.15)', 'rgba(245,158,11,0.3)');
      if (orderAddrStr) addrTag.title = 'Last order shipped to: ' + orderAddrStr;
      group3.appendChild(addrTag);
      group3.style.display = '';
    }

    bar.appendChild(group3);

    const slot = document.getElementById('sb-info-bar-slot');
    if (slot) {
      slot.innerHTML = '';
      slot.appendChild(bar);
    } else {
      const toolbar = document.getElementById('sb-toolbar');
      if (toolbar && toolbar.parentNode) toolbar.insertAdjacentElement('afterend', bar);
    }
  }

  /** Normalize address string for comparison. */
  function _normalizeAddr(s) {
    return (s || '').trim().toLowerCase().replace(/[.,#\-]+/g, ' ').replace(/\s+/g, ' ');
  }

  /** Compare user's current address against last order's shipping address. */
  function _compareAddresses(user, order) {
    const current = user?.userAddress?.shipping;
    const orderAddr = order?.initialShippingAddress;
    if (!current?.street1 || !orderAddr?.street1) return null; // can't compare
    return _normalizeAddr(current.street1) === _normalizeAddr(orderAddr.street1);
  }

  /** Fetch user details (subscription + fraud + gender + gweb) + last order and render the info bar. */
  function _fetchAndRenderSubBar(userId) {
    const user = cachedCustomerCtx?.user;
    const startId = getCustomerIdFromURL();

    // Single query: subscriptions + fraud + gweb + gender + shipping addr
    GM_xmlhttpRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
      data: JSON.stringify({ operationName: 'userDetailsById', query: USER_DETAILS_QUERY, variables: { id: userId } }),
      onload(res) {
        if (getCustomerIdFromURL() !== startId) return; // stale
        try {
          const data = JSON.parse(res.responseText);
          const userData = data?.data?.userById?.data;
          if (!userData) return;

          // Cache fraud, gweb, gender, shipping addr
          if (cachedCustomerCtx) {
            cachedCustomerCtx._fraudStatus = userData.fraudInfo?.status || 'DEFAULT';
            if (userData.gwebLink) cachedCustomerCtx._gwebLink = userData.gwebLink;
            if (userData.gender) cachedCustomerCtx._gender = userData.gender;
            if (userData.userAddress?.shipping?.id) cachedCustomerCtx._shippingAddrId = userData.userAddress.shipping.id;
          }

          // Subscription
          const subs = userData.subscriptionList || [];
          const active = subs.find(s => s.status === 'Active') || subs[0];
          if (!active) return;

          // Fetch last order for address comparison + chargeback check
          fetchLastOrders(userId, (orders) => {
            if (getCustomerIdFromURL() !== startId) return; // stale
            const lastSub = (orders || []).find(o => o.type === 'SUBSCRIPTION' || !o.type);
            const addrMatch = lastSub ? _compareAddresses(user, lastSub) : null;
            const orderAddr = lastSub?.initialShippingAddress;
            const orderAddrStr = orderAddr ? [orderAddr.street1, orderAddr.city, orderAddr.region, orderAddr.postcode].filter(Boolean).join(', ') : '';
            renderUserInfoBar(active, addrMatch, orderAddrStr);

            // Async: fraud tag
            if (userData.fraudInfo?.status === 'DECLINE') {
              const warnGroup = document.getElementById('sb-info-group-warnings');
              if (warnGroup) {
                warnGroup.appendChild(makeTag('⚠ Fraud', '#ef4444', 'rgba(239,68,68,0.15)', 'rgba(239,68,68,0.4)'));
                warnGroup.style.display = '';
              }
            }

            // Async: chargeback tag
            GM_xmlhttpRequest({
              method: 'POST', url: GRAPHQL_URL,
              headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + BEARER_TOKEN },
              data: JSON.stringify({ operationName: 'chargesByUserId', query: CHARGES_QUERY, variables: { id: userId } }),
              onload(cRes) {
                if (getCustomerIdFromURL() !== startId) return; // stale
                try {
                  const cData = JSON.parse(cRes.responseText);
                  const chargeList = cData?.data?.userById?.data?.chargeList || [];
                  const hasChargeback = chargeList.some(c => c.charge?.status === 'CHARGEBACK');
                  if (hasChargeback) {
                    const warnGroup = document.getElementById('sb-info-group-warnings');
                    if (warnGroup) {
                      warnGroup.appendChild(makeTag('⚠ Chargeback', '#ef4444', 'rgba(239,68,68,0.15)', 'rgba(239,68,68,0.4)'));
                      warnGroup.style.display = '';
                    }
                  }
                } catch(e) {}
              }
            });
          });
        } catch(e) {}
      }
    });
  }

  /** Quick refresh — skips CRM search, goes straight to subscription fetch. */
  function refreshSubscriptionBar() {
    if (cachedCustomerCtx?.user) {
      _fetchAndRenderSubBar(cachedCustomerCtx.user.id);
    }
  }

  /**
   * Full refresh — resolves customer if needed, then fetches subscription.
   * Called by the MutationObserver on page changes. Debounced.
   */
  function refreshUserInfoBar() {
    clearTimeout(_userInfoTimer);
    _userInfoTimer = setTimeout(async () => {
      const startId = getCustomerIdFromURL();

      // Fast path: cache exists, just refresh subscription data
      if (cachedCustomerCtx?.user) {
        // Verify we're still on the same customer (URL may have changed)
        const currentId = startId;
        if (currentId && cachedCustomerCtx._kustomerId && currentId !== cachedCustomerCtx._kustomerId) {
          clearCustomerCtx();
        } else {
          if (document.getElementById('sb-user-info-bar')) return;
          _fetchAndRenderSubBar(cachedCustomerCtx.user.id);
          return;
        }
      }

      const email = await resolveEmail();
      if (getCustomerIdFromURL() !== startId) return; // stale
      if (!email) return;

      // Full resolve
      if (email === userInfoLastEmail && document.getElementById('sb-user-info-bar')) return;
      userInfoLastEmail = email;
      document.getElementById('sb-user-info-bar')?.remove();

      searchCRM(email, (users, err) => {
        if (getCustomerIdFromURL() !== startId) return; // stale
        if (err || !users) return;
        const sbUsers = users.filter(u => !u.origin || u.origin === 'SCENTBIRD');
        const exactMatch = sbUsers.find(u => u.email?.toLowerCase() === email.toLowerCase());
        const sbUser = exactMatch || sbUsers[0];
        if (!sbUser) return;
        cachedCustomerCtx = { email, user: sbUser, _kustomerId: getCustomerIdFromURL() };
        _fetchAndRenderSubBar(sbUser.id);
      });
    }, 800);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // MUTATION OBSERVER — wire everything on header appear
  // ══════════════════════════════════════════════════════════════════════════

  let tokenValidatedForSession = false;

  const observer = new MutationObserver(() => {
    if (toolbarEl && !document.body.contains(toolbarEl)) {
      toolbarEl = null;
      fillNameRegistered = false;
      clearCustomerCtx();
      ['sb-order-panel','sb-charges-panel','sb-cancel-panel','sb-search-panel','sb-queue-panel','sb-edit-customer-panel'].forEach(id => document.getElementById(id)?.remove());
    }
    if (!ensureToolbar()) return;
    registerFillName();
    registerCrmButtons();
    refreshUserInfoBar();
    if (!tokenValidatedForSession && BEARER_TOKEN) {
      tokenValidatedForSession = true;
      validateToken();
    }
  });

  observer.observe(document.body, { childList: true, subtree: true });

  // ── Stash selection on mousedown (Edge clears it before click fires) ────

  document.addEventListener('mousedown', () => {
    stashedSelection = (window.getSelection()?.toString() || '').trim();
  }, true);

  // ── Keyboard shortcut ─────────────────────────────────────────────────────

  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.shiftKey && e.key === 'Q') {
      if (document.getElementById('sb-search-panel')) { removePanel('sb-search-panel'); return; }
      showSearchPanel(window.getSelection().toString().trim());
    }
  });

  // ── Customer API auth via hidden iframe ──

  let _customerApiReady = false;
  let _customerCsrfToken = null;

  function ensureCustomerAuth(callback) {
    if (_customerApiReady) return callback(true);

    const gwebLink = cachedCustomerCtx?._gwebLink;
    if (!gwebLink) { console.warn('[BirdsEye] No gwebLink'); return callback(false); }

    const popup = window.open(gwebLink, 'sb_auth', 'width=1,height=1,left=-9999,top=-9999,menubar=no,toolbar=no,location=no,status=no');
    window.focus();

    setTimeout(() => {
      try { popup?.close(); } catch(e) {}

      GM_xmlhttpRequest({
        method: 'POST',
        url: 'https://api.scentbird.com/graphql?opname=queueProbe',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Origin': 'https://www.scentbird.com',
          'Referer': 'https://www.scentbird.com/',
        },
        data: JSON.stringify({
          operationName: 'queueProbe',
          variables: {},
          query: `query queueProbe { __schema { queryType { fields { name } } } }`
        }),
        onload(res) {
          const csrfMatch = res.responseHeaders.match(/x-csrf-token:\s*(.+)/i);
          if (csrfMatch) _customerCsrfToken = csrfMatch[1].trim();

          try {
            const body = JSON.parse(res.responseText);
            if (body?.data) {
              _customerApiReady = true;
              callback(true);
            } else {
              callback(false);
            }
          } catch(e) {
            callback(false);
          }
        },
        onerror() {
          callback(false);
        }
      });
    }, 5000);
  }

  function customerApiDelete(flatIndex, callback) {
    const headers = {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Origin': 'https://www.scentbird.com',
      'Referer': 'https://www.scentbird.com/',
    };
    if (_customerCsrfToken) headers['x-csrf-token'] = _customerCsrfToken;

    GM_xmlhttpRequest({
      method: 'POST',
      url: 'https://api.scentbird.com/graphql?opname=QueueDeleteItem',
      headers,
      data: JSON.stringify({
        operationName: 'QueueDeleteItem',
        variables: {
          mutationInput: { index: flatIndex, metadata: { page: 'Queue', pageType: 'Queue', placement: 'queue' } },
          queueInput: { limit: 24 }
        },
        query: `mutation QueueDeleteItem($mutationInput: QueueDeleteItemInput!, $queueInput: QueueInput) {
          queueDeleteItem(input: $mutationInput) {
            queue(input: $queueInput) {
              queueItems { month year products { flatIndex queueItemId tradingItem { productInfo { name } } } }
            }
            error {
              ... on QueueDeleteItemError { queueDeleteItemErrorCode: errorCode message }
              ... on SecurityError { securityErrorCode: errorCode message }
              ... on ServerError { serverErrorCode: errorCode message }
              ... on ValidationError { validationErrorCode: errorCode message }
            }
          }
        }`
      }),
      onload(res) {
        try {
          const body = JSON.parse(res.responseText);
          const err = body?.data?.queueDeleteItem?.error;
          const errMsg = err?.message;
          if (errMsg) return callback(null, errMsg);
          if (body?.errors?.length) return callback(null, body.errors[0].message);
          callback(body?.data?.queueDeleteItem, null);
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); }
    });
  }

  /** Generic customer-side API call (api.scentbird.com). Requires ensureCustomerAuth first. */
  function customerApiCall(operationName, query, variables, callback) {
    const headers = {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Origin': 'https://www.scentbird.com',
      'Referer': 'https://www.scentbird.com/',
    };
    if (_customerCsrfToken) headers['x-csrf-token'] = _customerCsrfToken;

    GM_xmlhttpRequest({
      method: 'POST',
      url: 'https://api.scentbird.com/graphql?opname=' + operationName,
      headers,
      data: JSON.stringify({ operationName, variables, query }),
      onload(res) {
        try {
          const body = JSON.parse(res.responseText);
          if (body?.errors?.length) return callback(null, body.errors[0].message);
          callback(body?.data, null);
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); }
    });
  }

})();