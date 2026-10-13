// ==UserScript==
// @name         BirdsEye
// @namespace    scentbird-kustomer
// @version      9.0.0
// @description  Unified toolbar for Scentbird support: CRM search, orders, charges, queue, edit / downgrade / cancel, replacements, refunds, Specials (OT, account deletion)
// @author       You
// @match        https://scentbird.kustomerapp.com/*
// @match        https://crm.scentbird.com/login/callback*
// @grant        GM_xmlhttpRequest
// @grant        GM_setClipboard
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @connect      crm.scentbird.com
// @connect      www.scentbird.com
// @connect      api.scentbird.com
// @connect      scentbird.okta.com
// @connect      a1.asendiausa.com
// @connect      a1reportapi.asendiaprod.com
// Install-only: no @updateURL/@downloadURL. Cowork remints the artifact URL on
// every upload and the git host serves clone/push only (no raw files), so there
// is no stable URL for Tampermonkey autoupdate. Reinstall manually per release.
// ==/UserScript==

(function () {
  'use strict';
  try { document.documentElement.dataset.sbBeta = '8.35.0-beta'; } catch (_) {}

  // ══════════════════════════════════════════════════════════════════════════
  // OKTA / OAUTH CONSTANTS
  // ══════════════════════════════════════════════════════════════════════════

  const OKTA_CLIENT_ID    = '0oayh4u7fglCFfDbL697';
  const OKTA_AUTHORIZE    = 'https://scentbird.okta.com/oauth2/v1/authorize';
  const OKTA_TOKEN        = 'https://scentbird.okta.com/oauth2/v1/token';
  const OKTA_REDIRECT_URI = 'https://crm.scentbird.com/login/callback';
  const OKTA_SCOPE        = 'openid email groups offline_access';
  const TOKENS_KEY        = 'birdseye_okta_tokens';   // GM_setValue bundle: {access_token, refresh_token, id_token}
  const PENDING_AUTH_KEY  = 'birdseye_okta_pending';  // GM_setValue: {code_verifier, state} during a login flow

  // ══════════════════════════════════════════════════════════════════════════
  // PKCE + TOKEN STORAGE
  // ══════════════════════════════════════════════════════════════════════════

  function b64url(bytes) {
    let s = '';
    for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  function randomB64Url(byteLen) {
    const buf = new Uint8Array(byteLen);
    crypto.getRandomValues(buf);
    return b64url(buf);
  }

  /** Returns a Promise<string> resolving to the SHA-256 base64url hash of `s`. */
  async function sha256B64Url(s) {
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
    return b64url(new Uint8Array(digest));
  }

  function getTokens() {
    try { return JSON.parse(GM_getValue(TOKENS_KEY, '') || 'null'); } catch(e) { return null; }
  }
  function setTokens(bundle) { GM_setValue(TOKENS_KEY, JSON.stringify(bundle)); }
  function clearTokens()     { GM_deleteValue(TOKENS_KEY); }
  // CRM's backend validates the ID token as its bearer (not the access token).
  function getAccessToken()  { return getTokens()?.id_token || ''; }

  function getPendingAuth() {
    try { return JSON.parse(GM_getValue(PENDING_AUTH_KEY, '') || 'null'); } catch(e) { return null; }
  }
  function setPendingAuth(p) { GM_setValue(PENDING_AUTH_KEY, JSON.stringify(p)); }
  function clearPendingAuth() { GM_deleteValue(PENDING_AUTH_KEY); }

  // POST to Okta's /token endpoint. Centralized so the code exchange and the refresh
  // path share identical transport hardening:
  //  - anonymous: no cookies. The public-client PKCE grant needs none, and cookie
  //    mirroring is one signal Okta uses to classify a call as a browser request.
  //  - Origin/Referer pinned to the CRM origin: Okta's /token browser guard ("Browser
  //    requests to the token endpoint must be part of at least one whitelisted
  //    redirect_uri") keys on the Origin VALUE, not merely its presence — it must match
  //    the origin of a whitelisted redirect_uri. Tampermonkey strips Origin on
  //    cross-origin GM_xmlhttpRequest, so its call reads as a native request and passes.
  //    Violentmonkey (Edge/MV3) leaves a non-matching origin on the request (e.g.
  //    chrome-extension://…), which Okta rejects. Pinning Origin to the CRM origin — the
  //    same one crmRequest sends, and the origin of our whitelisted redirect_uri —
  //    satisfies the guard on managers that forward Origin, and is a whitelisted value
  //    for any that strip-then-set it. Fallback if a manager still fails: register
  //    crm.scentbird.com as an Okta Trusted Origin.
  function oktaTokenRequest(body, handlers) {
    GM_xmlhttpRequest({
      method: 'POST', url: OKTA_TOKEN,
      anonymous: true,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept':       'application/json',
        'Origin':       'https://crm.scentbird.com',
        'Referer':      'https://crm.scentbird.com/',
      },
      data: body,
      onload:  handlers.onload,
      onerror: handlers.onerror,
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // OAUTH CALLBACK HANDLER (runs in popup at crm.scentbird.com/login/callback)
  // ══════════════════════════════════════════════════════════════════════════

  function handleOktaCallback() {
    // Block CRM from consuming the auth code before we do.
    const isOktaToken = (u) => typeof u === 'string' && u.indexOf('scentbird.okta.com/oauth2/v1/token') !== -1;
    const origFetch = window.fetch;
    window.fetch = function(input, init) {
      const url = typeof input === 'string' ? input : input?.url;
      if (isOktaToken(url)) return Promise.reject(new Error('blocked by BirdsEye'));
      return origFetch.apply(this, arguments);
    };
    const origOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function(method, url) {
      if (isOktaToken(url)) { this._birdsEyeBlocked = true; return; }
      return origOpen.apply(this, arguments);
    };
    const origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function() {
      if (this._birdsEyeBlocked) return;
      return origSend.apply(this, arguments);
    };

    // Render an overlay status message; do not clobber CRM's mounted DOM.
    let _statusEl = null;
    const showStatus = (text, color) => {
      try {
        if (!_statusEl) {
          _statusEl = document.createElement('div');
          _statusEl.id = 'sb-okta-status';
          _statusEl.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;z-index:2147483647;background:var(--sb-bg-deep);color:var(--sb-text);font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;font-size:var(--sb-fs-14);';
          (document.body || document.documentElement).appendChild(_statusEl);
        }
        _statusEl.style.color = color || 'var(--sb-text-muted)';
        _statusEl.textContent = text;
      } catch(e) {}
    };
    const showStatusWhenReady = (text, color) => {
      if (document.body) showStatus(text, color);
      else document.addEventListener('DOMContentLoaded', () => showStatus(text, color), { once: true });
    };
    showStatusWhenReady('Completing sign-in…');

    const params = new URLSearchParams(location.search);
    const code  = params.get('code');
    const state = params.get('state');
    const err   = params.get('error');

    if (err) {
      clearPendingAuth();
      showStatus('Sign-in failed: ' + err, 'var(--sb-danger)');
      setTimeout(() => { try { window.close(); } catch(e) {} }, 2500);
      return;
    }
    const pending = getPendingAuth();
    if (!code || !state || !pending || pending.state !== state) {
      clearPendingAuth();
      showStatus('Sign-in failed: invalid state', 'var(--sb-danger)');
      setTimeout(() => { try { window.close(); } catch(e) {} }, 2500);
      return;
    }

    const body = new URLSearchParams({
      grant_type:    'authorization_code',
      client_id:     OKTA_CLIENT_ID,
      code,
      code_verifier: pending.code_verifier,
      redirect_uri:  OKTA_REDIRECT_URI,
    }).toString();

    oktaTokenRequest(body, {
      onload(res) {
        try {
          const json = JSON.parse(res.responseText);
          if (res.status >= 200 && res.status < 300 && json.access_token) {
            setTokens({
              access_token:  json.access_token,
              refresh_token: json.refresh_token || null,
              id_token:      json.id_token || null,
            });
            clearPendingAuth();
            // Warm up DataDome by navigating to CRM root. Opener watches for tokens
            // landing in storage and closes this popup after the warm-up window.
            showStatus('Warming up CRM session…', 'var(--sb-warning)');
            try { location.replace('https://crm.scentbird.com/'); } catch(e) {
              // Fallback: if navigation fails, close ourselves.
              setTimeout(() => { try { window.close(); } catch(_) {} }, 800);
            }
          } else {
            clearPendingAuth();
            showStatus('Token exchange failed: ' + (json.error_description || json.error || res.status), 'var(--sb-danger)');
          }
        } catch(e) {
          clearPendingAuth();
          showStatus('Token exchange parse error', 'var(--sb-danger)');
        }
      },
      onerror() {
        clearPendingAuth();
        showStatus('Network error during token exchange', 'var(--sb-danger)');
      },
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // OAUTH CALLBACK BRANCH — popup landed on crm.scentbird.com/login/callback
  // ══════════════════════════════════════════════════════════════════════════
  // Only intercept when WE started a flow; otherwise let CRM's own login proceed.
  if (location.hostname === 'crm.scentbird.com' && location.pathname.startsWith('/login/callback')) {
    const pending = getPendingAuth();
    const urlState = new URLSearchParams(location.search).get('state');
    if (pending && urlState && pending.state === urlState) {
      handleOktaCallback();
    }
    return;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SETTINGS — font scale, window scale, theme. Persisted in GM_setValue.
  // ══════════════════════════════════════════════════════════════════════════
  const SETTINGS_KEY = 'birdseye_settings';
  const SETTINGS_DEFAULTS = {
    fontScale: 1.15, modalScale: 1.0, theme: 'dark', hideTax: false, closeDelay: true, aiMode: false,
    // AI panel collapsed to its title bar (the caret in its header).
    aiCollapsed: false,
    // Whether the model's prose summary is expanded in the AI panel.
    aiSummaryOpen: true,
    // Where the agent last dropped the AI panel, {left, top}. null = centred.
    aiPanelPos: null,
    // CCPA cancellation macro — paste a Kustomer shortcut "edit" link; we read the
    // shortcut id from it. Defaults to the known Cancel-My-Subscription macro.
    ccpaShortcut: 'https://scentbird.kustomerapp.com/app/settings/shortcuts/6980de563330d0a608e17abf/edit',
    // CCPA-loc — the reply used when the email matches no CRM account, so there is no
    // subscription to cancel and we have to ask which account they mean.
    ccpaLocShortcut: '',
    // AI reply macros — per cancellation/refund flow, a Kustomer shortcut link the
    // "Approve" action applies as the reply. Keyed by flow name (see AI_ACTIONS).
    aiFlowShortcuts: {},
  };

  /** Extract a 24-hex Kustomer id from a shortcut link (or a bare id). '' if none. */
  function parseShortcutId(v) {
    const m = String(v || '').match(/[0-9a-f]{24}/i);
    return m ? m[0] : '';
  }
  /** CCPA cancellation-macro shortcut id (from settings). */
  function getCcpaShortcutId() { return parseShortcutId(getSettings().ccpaShortcut); }
  /** CCPA-loc shortcut id — the no-account-found reply. */
  function getCcpaLocShortcutId() { return parseShortcutId(getSettings().ccpaLocShortcut); }

  let MODAL_SCALE = SETTINGS_DEFAULTS.modalScale;   // mirror; createPanel/Modal read from this

  function getSettings() {
    try {
      const raw = JSON.parse(GM_getValue(SETTINGS_KEY, '') || 'null');
      return Object.assign({}, SETTINGS_DEFAULTS, raw || {});
    } catch(e) { return Object.assign({}, SETTINGS_DEFAULTS); }
  }
  function setSettings(patch) {
    const merged = Object.assign(getSettings(), patch);
    GM_setValue(SETTINGS_KEY, JSON.stringify(merged));
    applySettings(merged);
    return merged;
  }
  function applySettings(s) {
    const root = document.documentElement;
    root.style.setProperty('--sb-fs-scale',    String(s.fontScale));
    root.style.setProperty('--sb-modal-scale', String(s.modalScale));
    root.dataset.theme = s.theme === 'light' ? 'light' : 'dark';
    MODAL_SCALE = s.modalScale;
  }
  applySettings(getSettings());

  const BILL_TEMPLATES_KEY = 'birdseye_bill_templates';
  const BILL_TEMPLATES_DEFAULTS = {
    FIRST_SUBSCRIPTION:       '* {total} was your initial subscription payment for {month} ({price}{discStr}{taxClause});',
    RECURRENT_CHARGE:         '* {total} was your monthly subscription payment for {month} ({price}{taxClause});',
    RESUBSCRIBE:              '* {total} was your resubscription payment for {month} ({price}{taxClause});',
    UPGRADE_PARTIAL:          '* {total} was a charge for upgrading your account ({price}{taxClause}).\nAs soon as you upgraded your plan, the difference in prices between the two plans was billed. Since all our subscriptions start out as a basic 1 fragrance/month and can be upgraded at any point, this transaction was applied as a separate charge;',
    UPGRADE_FULL:             '* {total} was a charge for a full plan upgrade ({price}{taxClause}).\nWhen upgrading your subscription plan, the full amount for the new plan is billed immediately as a separate charge;',
    PREMIUM_PRODUCT_UPCHARGE: '* {total} was a premium fragrance upcharge for including {name} in your order ({price}{taxClause}).\nSome fragrances on our website are part of a high-end lux collection and carry an upcharge of up to $25. They are marked by a tag in the left corner, and adding them to your queue triggers a pop-up with information about the extra cost;',
    PERFUME_CASE_UPCHARGE:    '* {total} was a charge for including a Case Subscription add-on in your plan ({price}{taxClause}). With this, you get a different Fragrance Case sent to you together with each shipment;',
    SHIPPING_UPCHARGE:        '* {total} was a shipping surcharge ({price}{taxClause});',
    SAMPLES_UPCHARGE:         '* {total} was a charge for including a Samples subscription add-on in your order ({price}{taxClause}). With this add-on you get two 1.5ml samples with each order;',
    CAR_SCENT_UPCHARGE:       '* {total} was a charge for your Drift Car scent add-on ({price}{taxClause}). With this add-on, you get a car fragrance in every order;',
    CANDLE_UPCHARGE:          '* {total} was a charge for your Candle Club add-on ({price}{taxClause}). You get a Scented Candle included in every order with this add-on;',
    HSD_UPCHARGE:             '* {total} was a charge for your Home Diffuser add-on ({price}{taxClause}). With this add-on you get a diffuser included in every order;',
    SHIPPING_PURCHASE:        '* {total} was a charge for the shipping of your order ({price}{taxClause});',
    ECOMMERCE_PURCHASE:       '* {total} was a charge for the separate Online Shop order you placed with us ({price}{discStr}{taxStr});',
  };
  const BILL_TEMPLATE_LABELS = {
    FIRST_SUBSCRIPTION:       'First subscription',
    RECURRENT_CHARGE:         'Monthly charge',
    RESUBSCRIBE:              'Resubscription',
    UPGRADE_PARTIAL:          'Partial plan upgrade',
    UPGRADE_FULL:             'Full plan upgrade',
    PREMIUM_PRODUCT_UPCHARGE: 'Premium product upcharge',
    PERFUME_CASE_UPCHARGE:    'Case add-on',
    SHIPPING_UPCHARGE:        'Shipping surcharge',
    SAMPLES_UPCHARGE:         'Samples add-on',
    CAR_SCENT_UPCHARGE:       'Car scent add-on',
    CANDLE_UPCHARGE:          'Candle Club add-on',
    HSD_UPCHARGE:             'Home Diffuser add-on',
    SHIPPING_PURCHASE:        'Shipping purchase',
    ECOMMERCE_PURCHASE:       'Online Shop order',
  };

  function getBillTemplates() {
    try { return JSON.parse(GM_getValue(BILL_TEMPLATES_KEY) || 'null') || {}; }
    catch(e) { return {}; }
  }
  function applyBillTemplate(template, vars) {
    return template.replace(/\{(\w+)\}/g, (_, k) => (k in vars ? vars[k] : '{' + k + '}'));
  }
  function autoClose(fn) {
    if (getSettings().closeDelay) setTimeout(fn, 1800);
    else fn();
  }

  if (!GM_getValue(BILL_TEMPLATES_KEY)) {
    GM_setValue(BILL_TEMPLATES_KEY, JSON.stringify(BILL_TEMPLATES_DEFAULTS));
  }

  // ══════════════════════════════════════════════════════════════════════════
  // GLOBAL CSS — font-size, modal-size, and theme color variables.
  // Custom properties bypass `all:initial` and inherit into panels.
  // ══════════════════════════════════════════════════════════════════════════
  (function injectStyles() {
    if (document.getElementById('sb-style')) return;
    const style = document.createElement('style');
    style.id = 'sb-style';
    style.textContent = `
      :root {
        --sb-fs-scale: 1.15;
        --sb-fs-9:  calc(9px  * var(--sb-fs-scale));
        --sb-fs-10: calc(10px * var(--sb-fs-scale));
        --sb-fs-11: calc(11px * var(--sb-fs-scale));
        --sb-fs-12: calc(12px * var(--sb-fs-scale));
        --sb-fs-13: calc(13px * var(--sb-fs-scale));
        --sb-fs-14: calc(14px * var(--sb-fs-scale));

        --sb-modal-scale: 1.0;

        /* Surfaces */
        --sb-bg-panel:           #1e1e2e;
        --sb-bg-deep:            #0f172a;
        --sb-bg-row:             #1a1a2e;
        --sb-bg-row-replacement: #1a1530;
        --sb-bg-input:           #2a2a3e;

        /* Text */
        --sb-text:           #e2e8f0;
        --sb-text-secondary: #cbd5e1;
        --sb-text-muted:     #94a3b8;
        --sb-text-dim:       #64748b;

        /* Borders */
        --sb-border:      rgba(255,255,255,0.1);
        --sb-border-soft: rgba(255,255,255,0.06);

        /* Overlays — used for button bg, hover, dividers */
        --sb-control-bg:  rgba(255,255,255,0.08);
        --sb-hover:       rgba(255,255,255,0.15);

        /* Shadow */
        --sb-shadow: 0 12px 40px rgba(0,0,0,0.5);

        /* Accents — used for buttons, status badges, plaques. Flipped in light theme. */
        --sb-accent:          #a5b4fc;
        --sb-accent-strong:   #818cf8;
        --sb-success:         #86efac;
        --sb-success-soft:    #6ee7b7;
        --sb-warning:         #fcd34d;
        --sb-warning-strong:  #f59e0b;
        --sb-danger:          #fca5a5;
        --sb-danger-strong:   #ef4444;
        --sb-danger-border:   #dc2626;
        --sb-info:            #60a5fa;

        /* Tags in the user info bar — paired text/bg/border per flavor, theme-aware */
        --sb-tag-plan-fg:        #c4b5fd;
        --sb-tag-plan-bg:        rgba(167,139,250,0.1);
        --sb-tag-plan-border:    rgba(167,139,250,0.25);
        --sb-tag-success-bg:     rgba(110,231,183,0.1);
        --sb-tag-success-border: rgba(110,231,183,0.3);
        --sb-tag-warning-bg:     rgba(245,158,11,0.1);
        --sb-tag-warning-border: rgba(245,158,11,0.25);
        --sb-tag-danger-bg:      rgba(239,68,68,0.15);
        --sb-tag-danger-border:  rgba(239,68,68,0.4);
        --sb-tag-info-bg:        rgba(96,165,250,0.1);
        --sb-tag-info-border:    rgba(96,165,250,0.25);
        --sb-tag-muted-bg:       rgba(148,163,184,0.08);
        --sb-tag-muted-border:   rgba(148,163,184,0.2);
        --sb-tag-pink-fg:        #f472b6;
        --sb-tag-pink-bg:        rgba(244,114,182,0.1);
        --sb-tag-pink-border:    rgba(244,114,182,0.25);
        --sb-tag-emerald-fg:     #34d399;
        --sb-tag-emerald-bg:     rgba(52,211,153,0.1);
        --sb-tag-emerald-border: rgba(52,211,153,0.25);

        /* Solid button fills — white text on a colored ground. Distinct from the
           accent tokens above, which are FOREGROUND colors on a dark panel. A solid
           has to stay dark enough for white text in BOTH themes, so most of these
           are not flipped below. The -hi suffix is a base's hover/armed partner. */
        --sb-solid-primary:     #4f46e5;
        --sb-solid-primary-hi:  #4338ca;
        --sb-solid-primary-lo:  #6366f1;
        --sb-solid-violet:      #7c3aed;
        --sb-solid-violet-hi:   #6d28d9;
        --sb-solid-violet-dim:  rgba(124,58,237,0.2);
        --sb-solid-danger-hi:   #b91c1c;
        --sb-solid-warn:        #a16207;
        --sb-solid-warn-hi:     #b45309;
        --sb-solid-success:     #059669;
        --sb-solid-disabled:    #374151;
        --sb-warning-hi:        #d97706;
        --sb-plaque-orange:     #f97316;
        --sb-plaque-yellow:     #eab308;
        --sb-on-solid:          #ffffff;
        /* Amber is too light for white text — the one solid that inverts. */
        --sb-on-warning:        #000000;

        /* Note/alert surfaces — the four bespoke rgba() boxes, unified. */
        --sb-note-warn-bg:      rgba(245,158,11,0.10);
        --sb-note-warn-border:  rgba(245,158,11,0.35);
        --sb-note-danger-bg:    rgba(239,68,68,0.10);
        --sb-note-danger-border:rgba(239,68,68,0.35);
        --sb-note-info-bg:      rgba(96,165,250,0.10);
        --sb-note-info-border:  rgba(96,165,250,0.30);
        --sb-badge-ai-bg:       rgba(99,102,241,0.20);
        --sb-badge-ai-border:   rgba(99,102,241,0.30);
      }

      /* Light theme — Kustomer-soft palette */
      :root[data-theme="light"] {
        --sb-bg-panel:           #ffffff;
        --sb-bg-deep:            #f7f7fa;
        --sb-bg-row:             #fafbfc;
        --sb-bg-row-replacement: #f5f0ff;
        --sb-bg-input:           #f7f7fa;

        --sb-text:           #111827;
        --sb-text-secondary: #1f2937;
        --sb-text-muted:     #374151;
        --sb-text-dim:       #4b5563;

        --sb-border:      #cbd5e1;
        --sb-border-soft: #e5e7eb;

        --sb-control-bg:  rgba(0,0,0,0.06);
        --sb-hover:       rgba(0,0,0,0.10);

        --sb-shadow: 0 12px 40px rgba(0,0,0,0.10);

        --sb-accent:          #4f46e5;
        --sb-accent-strong:   #4338ca;
        --sb-success:         #15803d;
        --sb-success-soft:    #16a34a;
        --sb-warning:         #b45309;
        --sb-warning-strong:  #92400e;
        --sb-danger:          #b91c1c;
        --sb-danger-strong:   #991b1b;
        --sb-danger-border:   #b91c1c;
        --sb-info:            #1d4ed8;

        --sb-tag-plan-fg:        #4338ca;
        --sb-tag-plan-bg:        rgba(99,102,241,0.14);
        --sb-tag-plan-border:    rgba(99,102,241,0.4);
        --sb-tag-success-bg:     rgba(22,163,74,0.14);
        --sb-tag-success-border: rgba(22,163,74,0.4);
        --sb-tag-warning-bg:     rgba(180,83,9,0.14);
        --sb-tag-warning-border: rgba(180,83,9,0.4);
        --sb-tag-danger-bg:      rgba(185,28,28,0.14);
        --sb-tag-danger-border:  rgba(185,28,28,0.4);
        --sb-tag-info-bg:        rgba(29,78,216,0.14);
        --sb-tag-info-border:    rgba(29,78,216,0.4);
        --sb-tag-muted-bg:       rgba(71,85,105,0.10);
        --sb-tag-muted-border:   rgba(71,85,105,0.30);
        --sb-tag-pink-fg:        #9d174d;
        --sb-tag-pink-bg:        rgba(157,23,77,0.14);
        --sb-tag-pink-border:    rgba(157,23,77,0.4);
        --sb-tag-emerald-fg:     #047857;
        --sb-tag-emerald-bg:     rgba(4,120,87,0.14);
        --sb-tag-emerald-border: rgba(4,120,87,0.4);

        /* Solids mostly carry over — they are already dark enough for white text on
           a white panel. The exception is the armed danger fill: light's
           --sb-danger-border is already #b91c1c, so reusing it would make arming
           invisible. Push it darker instead. */
        --sb-solid-danger-hi:   #7f1d1d;
        --sb-solid-disabled:    #9ca3af;
        --sb-warning-hi:        #92400e;
        --sb-plaque-orange:     #c2410c;
        --sb-plaque-yellow:     #a16207;

        --sb-note-warn-bg:      rgba(180,83,9,0.10);
        --sb-note-warn-border:  rgba(180,83,9,0.35);
        --sb-note-danger-bg:    rgba(185,28,28,0.10);
        --sb-note-danger-border:rgba(185,28,28,0.35);
        --sb-note-info-bg:      rgba(29,78,216,0.10);
        --sb-note-info-border:  rgba(29,78,216,0.30);
        --sb-badge-ai-bg:       rgba(79,70,229,0.14);
        --sb-badge-ai-border:   rgba(79,70,229,0.35);
      }

      /* Collapsed panel — header only. Done in CSS rather than by walking children
         in JS because a panel's body is appended AFTER createPanel returns, so a
         one-shot JS pass would miss everything and re-collapsing on each append is
         a footgun. The rule applies to whatever is in the panel at any moment.
         Bottom padding/margin collapse too, so the result is a tight title bar. */
      [data-sb-collapsed="1"] > *:not([data-sb-panel-header]) { display: none !important; }
      [data-sb-collapsed="1"] { padding-bottom: 0 !important; overflow: hidden !important; }
      [data-sb-collapsed="1"] > [data-sb-panel-header] { margin-bottom: 0 !important; border-bottom: none !important; }

      /* ══ COMPONENT LAYER ═══════════════════════════════════════════════════
         Everything is scoped under .sb-ui — the class every panel/modal/overlay
         root carries. That gives 0,2,0 specificity, which beats Kustomer's own
         element and single-class rules; a bare .sb-btn would lose to them.

         Panel roots set all:initial inline to stop CSS bleed, but that only
         resets the root itself — descendants still inherit Kustomer's font and
         are still reachable by its selectors. So every interactive base rule
         re-states font/box-sizing/margin/appearance rather than trusting them.

         Hover lives here, not in JS. The ~39 onmouseenter/onmouseleave pairs
         this replaces all did the same thing: swap between two constants. */

      .sb-ui .sb-btn {
        font: inherit; box-sizing: border-box; margin: 0; appearance: none;
        display: inline-flex; align-items: center; justify-content: center; gap: 6px;
        padding: 7px 12px; border-radius: 6px;
        border: 1px solid var(--sb-hover);
        background: var(--sb-control-bg);
        color: var(--sb-text);
        font-size: var(--sb-fs-12); font-weight: 600; line-height: 1.2;
        white-space: nowrap; cursor: pointer; text-align: center;
        transition: background 0.15s, border-color 0.15s, opacity 0.15s;
      }
      .sb-ui .sb-btn:hover:not(:disabled) { background: var(--sb-hover); }
      .sb-ui .sb-btn:disabled { opacity: 0.55; cursor: not-allowed; }

      /* Sizes */
      .sb-ui .sb-btn--xs { padding: 3px 9px;  font-size: var(--sb-fs-10); }
      .sb-ui .sb-btn--sm { padding: 5px 10px; font-size: var(--sb-fs-11); }
      .sb-ui .sb-btn--lg { padding: 9px 14px; font-size: var(--sb-fs-13); }

      /* Shapes */
      .sb-ui .sb-btn--pill  { border-radius: 20px; }
      .sb-ui .sb-btn--block { display: flex; width: 100%; }
      .sb-ui .sb-btn--grow  { flex: 1; min-width: 0; }

      /* Solid fills. Each pins its own hover so the generic rule above can't
         wash it out with --sb-hover. */
      .sb-ui .sb-btn--primary { background: var(--sb-solid-primary); border-color: transparent; color: var(--sb-on-solid); font-weight: 700; }
      .sb-ui .sb-btn--primary:hover:not(:disabled) { background: var(--sb-solid-primary-hi); }
      .sb-ui .sb-btn--violet  { background: var(--sb-solid-violet);  border-color: transparent; color: var(--sb-on-solid); font-weight: 700; }
      .sb-ui .sb-btn--violet:hover:not(:disabled)  { background: var(--sb-solid-violet-hi); }
      .sb-ui .sb-btn--danger  { background: var(--sb-danger-border); border-color: transparent; color: var(--sb-on-solid); font-weight: 700; }
      .sb-ui .sb-btn--danger:hover:not(:disabled)  { background: var(--sb-solid-danger-hi); }
      .sb-ui .sb-btn--success { background: var(--sb-solid-success); border-color: transparent; color: var(--sb-on-solid); font-weight: 700; }
      .sb-ui .sb-btn--success:hover:not(:disabled) { background: var(--sb-solid-success); }
      .sb-ui .sb-btn--warn    { background: var(--sb-solid-warn);    border-color: transparent; color: var(--sb-on-solid); font-weight: 700; }
      .sb-ui .sb-btn--warn:hover:not(:disabled)    { background: var(--sb-solid-warn-hi); }

      /* Armed — the second click of a two-stage destructive control. Set as a
         state class so the button's resting variant survives underneath it. */
      .sb-ui .sb-btn--sent, .sb-ui .sb-btn--sent:disabled {
        background: var(--sb-solid-success); border-color: transparent;
        color: var(--sb-on-solid); opacity: 1; cursor: default;
      }

      .sb-ui .sb-btn--armed         { background: var(--sb-solid-danger-hi) !important; color: var(--sb-on-solid) !important; border-color: transparent !important; }
      .sb-ui .sb-btn--armed-warn    { background: var(--sb-solid-warn-hi)   !important; color: var(--sb-on-solid) !important; border-color: transparent !important; }

      /* Outline — a solid's colour on a transparent ground. */
      /* For buttons whose fill is a runtime parameter and therefore inline: an
         inline background beats any rule here, so darken it with a filter instead. */
      .sb-ui .sb-btn--tinted:hover:not(:disabled) { filter: brightness(0.85); }

      .sb-ui .sb-btn--ghost    { background: transparent; border-color: var(--sb-border); color: var(--sb-text-muted); }
      .sb-ui .sb-btn--ghost:hover:not(:disabled) { background: var(--sb-control-bg); }
      .sb-ui .sb-btn--outline  { background: transparent; border-color: var(--sb-accent-strong); color: var(--sb-accent); }
      .sb-ui .sb-btn--outline:hover:not(:disabled) { background: var(--sb-control-bg); }

      /* ── Notes / alerts ─────────────────────────────────────────────────── */
      .sb-ui .sb-note {
        box-sizing: border-box; padding: 8px 10px; border-radius: 6px;
        font-size: var(--sb-fs-11); line-height: 1.45;
        border: 1px solid transparent;
        background: var(--sb-bg-row); color: var(--sb-text-muted);
      }
      .sb-ui .sb-note--warn   { background: var(--sb-note-warn-bg);   border-color: var(--sb-note-warn-border);   color: var(--sb-warning-strong); }
      .sb-ui .sb-note--danger { background: var(--sb-note-danger-bg); border-color: var(--sb-note-danger-border); color: var(--sb-danger); }
      .sb-ui .sb-note--info   { background: var(--sb-note-info-bg);   border-color: var(--sb-note-info-border);   color: var(--sb-info); }

      /* ── Text ───────────────────────────────────────────────────────────── */
      .sb-ui .sb-label {
        font-size: var(--sb-fs-10); color: var(--sb-text-muted); font-weight: 600;
        text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 4px;
      }
      .sb-ui .sb-hint  { font-size: var(--sb-fs-10); color: var(--sb-text-dim); }
      .sb-ui .sb-hint--warn { color: var(--sb-warning-strong); }
      .sb-ui .sb-badge {
        font-size: var(--sb-fs-10); padding: 1px 5px; border-radius: 3px;
        background: var(--sb-badge-ai-bg); color: var(--sb-accent);
        border: 1px solid var(--sb-badge-ai-border); white-space: nowrap;
      }
      .sb-ui .sb-divider { height: 1px; background: var(--sb-border-soft); margin: 10px 0; border: 0; }
      .sb-ui .sb-card    { background: var(--sb-bg-row); border-radius: 6px; padding: 8px 10px; }

      /* ── Inputs ─────────────────────────────────────────────────────────── */
      .sb-ui .sb-input, .sb-ui .sb-select, .sb-ui .sb-textarea {
        font: inherit; box-sizing: border-box; margin: 0;
        width: 100%; padding: 6px 9px; border-radius: 6px;
        border: 1px solid var(--sb-border);
        background: var(--sb-bg-input); color: var(--sb-text);
        font-size: var(--sb-fs-12);
      }
      .sb-ui .sb-input:focus, .sb-ui .sb-select:focus, .sb-ui .sb-textarea:focus {
        outline: none; border-color: var(--sb-accent-strong);
      }
      .sb-ui .sb-select   { cursor: pointer; }
      .sb-ui .sb-textarea { resize: vertical; font-family: Arial, sans-serif; }
      .sb-ui .sb-check    { cursor: pointer; accent-color: var(--sb-accent-strong); flex-shrink: 0; margin: 0; }

      /* ── Selectable rows (order picker, item picker, product results) ───── */
      .sb-ui .sb-choice {
        display: flex; align-items: center; gap: 8px; box-sizing: border-box;
        width: 100%; text-align: left;
        padding: 7px 9px; border-radius: 6px; margin-bottom: 4px;
        background: transparent; border: 1px solid transparent;
        font-size: var(--sb-fs-11); color: var(--sb-text); cursor: pointer;
        transition: background 0.15s, border-color 0.15s;
      }
      .sb-ui .sb-choice:hover { border-color: var(--sb-border); }
      .sb-ui .sb-choice[data-checked="1"] { background: var(--sb-bg-row); border-color: var(--sb-border); }
      .sb-ui .sb-choice--top { align-items: flex-start; }
      .sb-ui .sb-choice--top .sb-check { margin-top: 2px; }

      /* Two-line choice body: bold identity on top, dim detail under it. */
      .sb-ui .sb-choice-body { min-width: 0; flex: 1; }
      .sb-ui .sb-choice-main { font-size: var(--sb-fs-12); font-weight: 700; color: var(--sb-text); }
      .sb-ui .sb-choice-sub  { font-size: var(--sb-fs-10); color: var(--sb-text-dim); }
      .sb-ui .sb-truncate    { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

      /* ── Layout ─────────────────────────────────────────────────────────── */
      /* ── make* family ───────────────────────────────────────────────────── */
      .sb-ui .sb-copy {
        background: none; border: none; margin: 0; font: inherit;
        padding: 0 4px; color: var(--sb-accent-strong);
        font-size: var(--sb-fs-13); cursor: pointer; flex-shrink: 0;
      }
      .sb-ui .sb-form-label {
        font-size: var(--sb-fs-12); color: var(--sb-text-dim);
        font-weight: 600; letter-spacing: 0.05em; margin: 12px 0 5px;
      }
      /* The forms' select predates .sb-select and sits a size larger on a deeper
         ground. Kept as a modifier rather than normalised away, so Phase 1 stays
         a pure re-plumbing with no visual diff. */
      .sb-ui .sb-select--form { background: var(--sb-bg-deep); border-radius: 5px; padding: 6px 8px; font-size: var(--sb-fs-13); }
      .sb-ui .sb-status    { font-size: var(--sb-fs-12); min-height: 16px; margin-top: 10px; color: var(--sb-text-muted); }
      .sb-ui .sb-statuslog { margin-top: 10px; font-size: var(--sb-fs-11); line-height: 1.8; }
      .sb-ui .sb-tag {
        display: inline-flex; align-items: center; padding: 2px 7px;
        border-radius: 4px; font-weight: 600; font-size: var(--sb-fs-10); white-space: nowrap;
      }
      .sb-ui .sb-toolbar-btn { padding: 5px 12px; border-radius: 4px; font-size: var(--sb-fs-13); }

      /* ── Layout ─────────────────────────────────────────────────────────── */
      .sb-ui .sb-row       { display: flex; align-items: center; gap: 8px; }
      .sb-ui .sb-row--gap6 { gap: 6px; }
      .sb-ui .sb-spread    { justify-content: space-between; }
      .sb-ui .sb-fill      { flex: 1; min-width: 0; }
    `;
    (document.head || document.documentElement).appendChild(style);
  })();

  // ══════════════════════════════════════════════════════════════════════════
  // OKTA REFRESH + CRM REQUEST WRAPPER
  // ══════════════════════════════════════════════════════════════════════════

  let _refreshInFlight = null;     // coalesces concurrent refreshes
  let _reauthRequired  = false;    // set after a hard refresh failure; UI reads this

  /** Refresh the access token using the stored refresh token. Returns Promise<string>. */
  function refreshTokens() {
    if (_refreshInFlight) return _refreshInFlight;
    const tokens = getTokens();
    if (!tokens?.refresh_token) {
      _reauthRequired = true;
      return Promise.reject(new Error('reauth_required'));
    }
    const body = new URLSearchParams({
      grant_type:    'refresh_token',
      client_id:     OKTA_CLIENT_ID,
      refresh_token: tokens.refresh_token,
      scope:         OKTA_SCOPE,
    }).toString();

    _refreshInFlight = new Promise((resolve, reject) => {
      oktaTokenRequest(body, {
        onload(res) {
          try {
            const json = JSON.parse(res.responseText);
            if (res.status >= 200 && res.status < 300 && json.access_token) {
              const newId = json.id_token || tokens.id_token || null;
              setTokens({
                access_token:  json.access_token,
                refresh_token: json.refresh_token || tokens.refresh_token,
                id_token:      newId,
              });
              _reauthRequired = false;
              resolve(newId);
            } else if (res.status >= 400 && res.status < 500) {
              clearTokens();
              _reauthRequired = true;
              reject(new Error('reauth_required'));
            } else {
              reject(new Error('refresh_failed_' + res.status));
            }
          } catch(e) {
            reject(new Error('refresh_parse_error'));
          }
        },
        onerror() { reject(new Error('refresh_network_error')); },
      });
    }).finally(() => { _refreshInFlight = null; });

    return _refreshInFlight;
  }

  /**
   * Wraps GM_xmlhttpRequest. Auto-injects Authorization, refreshes-and-retries on 401.
   * `opts` = { method, url, headers, data, onload(res), onerror() }
   * onload/onerror semantics match GM_xmlhttpRequest's.
   */
  function crmRequest(opts) {
    const send = (token, isRetry) => {
      // Mimic a real same-origin CRM request so DataDome doesn't gate the cross-origin
      // call. Origin + Sec-Fetch-* must all match what a real CRM tab sends.
      const headers = Object.assign({
        'Origin':          'https://crm.scentbird.com',
        'Referer':         'https://crm.scentbird.com/',
        'Sec-Fetch-Site':  'same-origin',
        'Sec-Fetch-Mode':  'cors',
        'Sec-Fetch-Dest':  'empty',
      }, opts.headers || {});
      if (token) headers['Authorization'] = 'Bearer ' + token;
      GM_xmlhttpRequest({
        method: opts.method || 'POST',
        url:    opts.url,
        headers,
        data:   opts.data,
        onload(res) {
          if (res.status === 403) {
            console.log('[BirdsEye] 403', opts.url, res.responseText?.substring(0, 200));
          }
          if (res.status === 401 && !isRetry) {
            refreshTokens().then(
              (newToken) => send(newToken, true),
              ()         => opts.onload && opts.onload(res)
            );
            return;
          }
          opts.onload && opts.onload(res);
        },
        onerror: opts.onerror,
      });
    };
    send(getAccessToken(), false);
  }

  /** Kicks off Okta login flow: generates PKCE pair, opens popup to authorize URL. */
  async function startOktaLogin() {
    const code_verifier  = randomB64Url(48);
    const code_challenge = await sha256B64Url(code_verifier);
    const state          = randomB64Url(24);
    const nonce          = randomB64Url(24);
    setPendingAuth({ code_verifier, state });

    const authUrl = OKTA_AUTHORIZE + '?' + new URLSearchParams({
      client_id:             OKTA_CLIENT_ID,
      response_type:         'code',
      scope:                 OKTA_SCOPE,
      redirect_uri:          OKTA_REDIRECT_URI,
      state,
      nonce,
      code_challenge,
      code_challenge_method: 'S256',
    }).toString();

    const w = 500, h = 700;
    const left = (screen.width  - w) / 2;
    const top  = (screen.height - h) / 2;
    const popup = window.open(authUrl, 'sb_okta_login', `width=${w},height=${h},left=${left},top=${top}`);

    // Watch GM_setValue for the token bundle to land. Once it does, the popup
    // is on crm.scentbird.com warming up DataDome — give it 5s, then force-close.
    const beforeId = getTokens()?.id_token || null;
    const startedAt = Date.now();
    const watcher = setInterval(() => {
      if (!popup || popup.closed) { clearInterval(watcher); return; }
      const cur = getTokens();
      if (cur?.id_token && cur.id_token !== beforeId) {
        clearInterval(watcher);
        setTimeout(() => { try { popup.close(); } catch(e) {} refreshOktaButtonState(); }, 5000);
      } else if (Date.now() - startedAt > 180000) {
        // 3-minute safety: if user hasn't completed auth, stop polling.
        clearInterval(watcher);
      }
    }, 500);
  }

  function refreshOktaButtonState() {
    const btn = document.getElementById('sb-crm-token-btn');
    if (!btn) return;
    const haveToken = !!getAccessToken();
    if (haveToken && _reauthRequired) _reauthRequired = false;   // popup login completed
    const state = _captchaSolving  ? 'solving'
                : _captchaRequired ? 'captcha'
                : !haveToken       ? 'signin'
                : _reauthRequired  ? 'reauth'
                                   : 'in';
    if (btn.dataset.authState === state) return;
    btn.dataset.authState = state;
    if (state === 'in') {
      btn.style.display = 'none';
      return;
    }
    btn.style.display = '';
    if (state === 'solving') {
      btn.textContent = '⏳ Solving challenge…';
      btn.style.color = 'var(--sb-warning)';
    } else if (state === 'captcha') {
      btn.textContent = '⚠ CRM Captcha';
      btn.style.color = 'var(--sb-warning-strong)';
    } else {
      btn.textContent = '🔐 Sign in with Okta';
      btn.style.color = state === 'reauth' ? 'var(--sb-danger)' : 'var(--sb-text)';
    }
  }

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
  let _myAuthorEmail = null;             // our CRM comment author email, discovered on first comment create

  function clearCustomerCtx() {
    cachedCustomerCtx = null;
    userInfoLastEmail = null;
    _customerApiReady = false;
    _customerCsrfToken = null;
    _refundCommentsPosted.clear();
    _totalRefundedCents = 0;
    _totalRefundedCurrency = 'USD';
  }

  // ── Auth state ─────────────────────────────────────────────────────────────
  // Token storage is in GM_setValue (TOKENS_KEY); see PKCE section above.

  /** Called when a request gets 401 even after a refresh attempt — refresh chain dead. */
  function handle401() {
    _reauthRequired = true;
    refreshOktaButtonState();
  }

  let _captchaRequired = false;
  function handle403() {
    _captchaRequired = true;
    refreshOktaButtonState();
  }

  /** Opens crm.scentbird.com in a popup and auto-closes once a CRM probe stops 403'ing.
   *  Gives DataDome up to ~30s to complete its challenge before giving up. */
  let _captchaSolving = false;
  function openCaptchaPopup() {
    const w = 500, h = 700;
    const left = (screen.width  - w) / 2;
    const top  = (screen.height - h) / 2;
    const popup = window.open('https://crm.scentbird.com/', 'sb_crm_captcha', `width=${w},height=${h},left=${left},top=${top}`);

    _captchaSolving = true;
    refreshOktaButtonState();

    let timedOut = false;
    const cleanup = (resolved) => {
      _captchaSolving = false;
      if (resolved) _captchaRequired = false;
      refreshOktaButtonState();
    };

    const closeWatch = setInterval(() => {
      if (!popup || popup.closed) {
        clearInterval(closeWatch);
        clearInterval(probe);
        clearTimeout(giveUp);
        cleanup(false);   // user closed — leave _captchaRequired set so they can retry
      }
    }, 500);

    // Wait 5s before the first probe (DataDome JS challenge takes a beat to run).
    let probe = null;
    setTimeout(() => {
      probe = setInterval(() => {
        if (timedOut) return;
        GM_xmlhttpRequest({
          method: 'POST', url: GRAPHQL_URL,
          headers: {
            'Content-Type':  'application/json',
            'Authorization': 'Bearer ' + getAccessToken(),
            'Origin':        'https://crm.scentbird.com',
            'Referer':       'https://crm.scentbird.com/',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Dest': 'empty',
          },
          data: JSON.stringify({ query: '{ __typename }' }),
          onload(res) {
            if (res.status !== 403) {
              clearInterval(probe);
              clearInterval(closeWatch);
              clearTimeout(giveUp);
              try { popup?.close(); } catch(e) {}
              cleanup(true);
            }
          },
        });
      }, 3000);
    }, 5000);

    // 30s timeout: stop polling and revert toolbar; popup stays open for the user to interact.
    const giveUp = setTimeout(() => {
      timedOut = true;
      if (probe) clearInterval(probe);
      clearInterval(closeWatch);
      cleanup(false);
    }, 30000);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // UTILITY HELPERS
  // ══════════════════════════════════════════════════════════════════════════

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

  // ══════════════════════════════════════════════════════════════════════════
  // UI FACTORIES — shared overlay / modal / panel / helpers
  // ══════════════════════════════════════════════════════════════════════════

  /** Build an element in one call. `cls` names classes from the component layer
   *  in the stylesheet above; `style` is for the genuinely one-off bits that do
   *  not deserve a class (a computed width, a colour from trackingStatusColor).
   *  Prefer cls — an inline style beats every rule in the sheet, including :hover.
   *
   *    el('button', { cls: 'sb-btn sb-btn--primary', text: 'Approve', on: { click } })
   *    el('div',    { cls: 'sb-row' }, [left, right])
   */
  function el(tag, opts = {}, children = []) {
    const node = document.createElement(tag);
    if (opts.cls)   node.className = opts.cls;
    if (opts.id)    node.id = opts.id;
    if (opts.text  != null) node.textContent = opts.text;
    if (opts.html  != null) node.innerHTML = opts.html;
    if (opts.title != null) node.title = opts.title;
    if (opts.style) node.style.cssText = opts.style;
    if (opts.attrs) for (const k of Object.keys(opts.attrs)) {
      const v = opts.attrs[k];
      if (v === false || v == null) continue;
      node.setAttribute(k, v === true ? '' : v);
    }
    if (opts.data) for (const k of Object.keys(opts.data)) node.dataset[k] = opts.data[k];
    if (opts.props) Object.assign(node, opts.props);
    if (opts.on) for (const k of Object.keys(opts.on)) node.addEventListener(k, opts.on[k]);
    for (const c of [].concat(children)) if (c) node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    return node;
  }

  /** Two-stage destructive button: first click arms, second fires, and it
   *  disarms itself after `armMs`. Replaces four hand-rolled copies of the same
   *  timer + colour dance (Approve, Send, merge-on-duplicate, cancel-sub), which
   *  had drifted to three different arm colours and two different timeouts. */
  function armedButton(btn, { idle, armed = '⚠ Click again', armMs = 4000, armedCls = 'sb-btn--armed', onArm, onFire }) {
    let isArmed = false, timer = null;
    const rest = () => {
      isArmed = false;
      btn.textContent = idle;
      btn.classList.remove(armedCls);
      clearTimeout(timer);
    };
    btn.textContent = idle;
    btn.addEventListener('click', async () => {
      if (!isArmed) {
        isArmed = true;
        btn.textContent = armed;
        btn.classList.add(armedCls);
        timer = setTimeout(rest, armMs);
        try { onArm?.(); } catch (_) {}
        return;
      }
      clearTimeout(timer);
      btn.classList.remove(armedCls);
      isArmed = false;
      await onFire(btn);
    });
    // Callers need to re-arm after a failed fire (the agent already confirmed
    // once, so a retry should go on one click) and to reset after a disarm.
    btn._sbRest = rest;
    btn._sbPreArm = () => { isArmed = true; };
    return btn;
  }

  /** Full-screen overlay. Closes on backdrop click. Returns overlay element. */
  function createOverlay(id) {
    document.getElementById(id)?.remove();
    const overlay = document.createElement('div');
    overlay.id = id;
    overlay.className = 'sb-ui';
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
  function createModal({ id, title, width = 360, noScale = false }) {
    const overlay = createOverlay(id);
    const scale = noScale ? 1 : MODAL_SCALE;
    width = Math.round(width * scale);
    const box = document.createElement('div');
    box.style.cssText = `
      background:var(--sb-bg-panel);border-radius:10px;padding:20px 24px;width:${width}px;
      max-height:85vh;overflow-y:auto;
      border:1px solid var(--sb-border);box-shadow:var(--sb-shadow);
      color:var(--sb-text);font-family:Arial,sans-serif;font-size:var(--sb-fs-14);box-sizing:border-box;
    `;
    if (title) {
      const hdr = document.createElement('div');
      hdr.style.cssText = `
        position:sticky;top:-20px;z-index:5;
        margin:-20px -24px 12px;padding:14px 24px;
        background:var(--sb-bg-panel);
        display:flex;justify-content:space-between;align-items:center;
        font-weight:700;font-size:var(--sb-fs-14);
        border-bottom:1px solid var(--sb-border-soft);
      `;
      const span = document.createElement('span');
      span.textContent = title;
      hdr.appendChild(span);
      const xBtn = el('button', {
        cls: 'sb-btn sb-btn--sm sb-btn--pill',
        text: 'Close \u2715',
        style: 'color:var(--sb-text-muted);',
      });
      xBtn.onclick = () => overlay.remove();
      hdr.appendChild(xBtn);
      box.appendChild(hdr);
    }
    overlay.appendChild(box);
    return { overlay, box };
  }

  /** Fixed-position floating panel. Returns panel element (already in DOM).
   *  Options: centered (override right-anchored placement), draggable (grab header to move),
   *  noScale (bypass MODAL_SCALE — for the settings panel), collapsible (header caret
   *  that hides the body; `collapsed` sets the initial state and `onCollapse(bool)`
   *  fires on every toggle, so callers can persist it), position + onMove (restore a
   *  remembered {left,top} and report every drop, so a panel can stay where it was put). */
  function createPanel({ id, title, width = 380, right = 20, maxHeight = 680, centered = false, draggable = false, noScale = false, collapsible = false, collapsed = false, position = null, onCollapse, onMove, onClose }) {
    removePanel(id);
    const scale = noScale ? 1 : MODAL_SCALE;
    width     = Math.round(width     * scale);
    maxHeight = Math.round(maxHeight * scale);
    const panel = document.createElement('div');
    panel.id = id;
    panel.className = 'sb-ui';

    // A remembered position outranks both placement modes. Clamped on restore: the
    // window may be smaller than it was when the panel was dropped, and a panel
    // restored off-screen is indistinguishable from a broken one.
    const saved = (position && Number.isFinite(position.left) && Number.isFinite(position.top))
      ? {
          left: Math.min(window.innerWidth  - 80, Math.max(-width + 80, position.left)),
          top:  Math.min(window.innerHeight - 40, Math.max(0,           position.top)),
        }
      : null;

    const initLeft = centered ? Math.max(20, Math.floor((window.innerWidth  - width)         / 2)) : null;
    const initTop  = centered ? Math.max(40, Math.floor((window.innerHeight - maxHeight) / 3)) : 60;
    const placement = saved
      ? `left:${saved.left}px;top:${saved.top}px;`
      : centered
        ? `left:${initLeft}px;top:${initTop}px;`
        : `top:${initTop}px;right:${right}px;`;
    panel.style.cssText = `
      all:initial;position:fixed;${placement}width:${width}px;max-height:${maxHeight}px;
      overflow-y:auto;background:var(--sb-bg-panel);color:var(--sb-text);font-family:Arial,sans-serif;
      font-size:var(--sb-fs-13);border-radius:10px;box-shadow:var(--sb-shadow);
      padding:14px;z-index:999999;border:1px solid var(--sb-border);box-sizing:border-box;
    `;
    const hdr = document.createElement('div');
    hdr.style.cssText = `
      position:sticky;top:-14px;z-index:5;
      margin:-14px -14px 10px;padding:10px 14px;
      background:var(--sb-bg-panel);
      display:flex;align-items:center;justify-content:space-between;
      border-bottom:1px solid var(--sb-border-soft);
      ${draggable ? 'cursor:grab;user-select:none;' : ''}
    `;
    hdr.dataset.sbPanelHeader = '1';
    hdr.innerHTML = `<span style="font-weight:700;font-size:var(--sb-fs-14);">${title}</span>`;

    // Right-hand control cluster: [caret] [Close ✕]. Built even when there's only
    // one button so the header layout doesn't shift between panel kinds.
    const ctrls = document.createElement('div');
    ctrls.style.cssText = 'display:flex;align-items:center;gap:6px;flex-shrink:0;';

    let collapseBtn = null;
    if (collapsible) {
      collapseBtn = el('button', { cls: 'sb-btn sb-btn--sm sb-btn--pill', style: 'color:var(--sb-text-muted);line-height:1;' });
      const paint = () => {
        const isCollapsed = panel.dataset.sbCollapsed === '1';
        collapseBtn.textContent = isCollapsed ? '▾' : '▴';
        collapseBtn.title = isCollapsed ? 'Expand' : 'Collapse';
      };
      /** Single entry point for the state — the header click and the caret share it,
       *  and callers can drive it via panel._sbSetCollapsed(). */
      panel._sbSetCollapsed = (next, notify = true) => {
        panel.dataset.sbCollapsed = next ? '1' : '0';
        // max-height is inline on the panel; while collapsed it must not reserve the
        // full body height or the title bar floats in a tall empty box.
        panel.style.maxHeight = next ? 'none' : maxHeight + 'px';
        paint();
        if (notify) { try { onCollapse?.(!!next); } catch (_) {} }
      };
      collapseBtn.onclick = (e) => {
        e.stopPropagation();   // don't also trip the header's own toggle
        panel._sbSetCollapsed(panel.dataset.sbCollapsed !== '1');
      };
      ctrls.appendChild(collapseBtn);

      // Clicking the title bar itself toggles too. On a draggable panel that would
      // fire at the end of every drag, so a click that moved the panel is ignored.
      let downAt = null;
      hdr.addEventListener('mousedown', (e) => { downAt = { x: e.clientX, y: e.clientY }; });
      hdr.addEventListener('click', (e) => {
        if (ctrls.contains(e.target)) return;
        if (downAt && Math.hypot(e.clientX - downAt.x, e.clientY - downAt.y) > 4) return;
        panel._sbSetCollapsed(panel.dataset.sbCollapsed !== '1');
      });
    }

    const closeBtn = el('button', {
      cls: 'sb-btn sb-btn--sm sb-btn--pill',
      text: 'Close \u2715',
      style: 'color:var(--sb-text-muted);',
    });
    closeBtn.onclick = () => { removePanel(id); onClose?.(); };
    ctrls.appendChild(closeBtn);
    hdr.appendChild(ctrls);
    panel.appendChild(hdr);
    document.body.appendChild(panel);

    // Applied after mount so the initial state goes through the same path as a click.
    // Silent: restoring a remembered state isn't the agent changing it.
    if (collapsible && collapsed) panel._sbSetCollapsed(true, false);

    if (draggable) attachPanelDrag(panel, hdr, ctrls, onMove);
    return panel;
  }

  /** Wire mousedown-on-header drag for a panel. Switches the panel to absolute left/top.
   *  `ctrls` is the header's button cluster — dragging never starts on it. */
  function attachPanelDrag(panel, hdr, ctrls, onDrop) {
    let dragging = false, startX = 0, startY = 0, originX = 0, originY = 0;
    const onDown = (e) => {
      if (e.target === ctrls || ctrls.contains(e.target)) return;
      dragging = true;
      const rect = panel.getBoundingClientRect();
      // Switch to left/top so we can move freely (drop right anchor if any).
      panel.style.left  = rect.left + 'px';
      panel.style.top   = rect.top  + 'px';
      panel.style.right = 'auto';
      originX = rect.left;
      originY = rect.top;
      startX = e.clientX;
      startY = e.clientY;
      hdr.style.cursor = 'grabbing';
      e.preventDefault();
    };
    const onMove = (e) => {
      if (!dragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const rect = panel.getBoundingClientRect();
      // Clamp so the title bar always stays grabbable inside the viewport.
      const minLeft = -rect.width  + 80;
      const maxLeft = window.innerWidth  - 80;
      const minTop  = 0;
      const maxTop  = window.innerHeight - 40;
      const nx = Math.min(maxLeft, Math.max(minLeft, originX + dx));
      const ny = Math.min(maxTop,  Math.max(minTop,  originY + dy));
      panel.style.left = nx + 'px';
      panel.style.top  = ny + 'px';
    };
    const onUp = () => {
      if (!dragging) return;
      dragging = false;
      hdr.style.cursor = 'grab';
      // Report the drop, not every mousemove — persisting mid-drag would write to
      // storage a few hundred times per drag.
      if (onDrop) {
        const r = panel.getBoundingClientRect();
        try { onDrop({ left: Math.round(r.left), top: Math.round(r.top) }); } catch (_) {}
      }
    };
    hdr.addEventListener('mousedown', onDown);
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  /** Small user identity row used in panels. */
  function renderUserRow(user) {
    const row = document.createElement('div');
    row.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-11);margin-bottom:12px;';
    row.textContent = `${user.firstName || ''} ${user.lastName || ''} · ${user.email}`.trim();
    return row;
  }

  /** Returns a single-line formatted shipping address, or '' if shipping is missing. */
  function formatShippingAddress(shipping) {
    if (!shipping) return '';
    const street = [shipping.street1, shipping.street2].filter(Boolean).join(', ');
    const city   = [shipping.city, shipping.region, shipping.postcode].filter(Boolean).join(', ');
    return [street, city].filter(Boolean).join(' · ');
  }

  /** High-contrast color for a tracking-event status. */
  function trackingStatusColor(status) {
    const s = (status || '').toUpperCase();
    if (s === 'DELIVERED')                                                       return 'var(--sb-success)'; // bright lime
    if (s === 'OUT_FOR_DELIVERY' || s === 'IN_TRANSIT' || s === 'SHIPPED')       return 'var(--sb-warning)'; // warm yellow
    if (s === 'RETURNED' || s === 'LOST' || s === 'FAILED' || s === 'EXCEPTION') return 'var(--sb-danger)'; // red
    return 'var(--sb-text)'; // default bright
  }

  /** Returns a small inline pill for non-LIVE product statuses, or null if status is LIVE/missing. */
  function makeProductStatusPlaque(status) {
    if (!status || status === 'LIVE') return null;
    const palette = {
      OUT_OF_STOCK:                  { bg: 'rgba(245,158,11,0.15)', fg: 'var(--sb-warning)', border: 'rgba(245,158,11,0.4)' },
      NOT_AVAILABLE_FOR_NEW_ORDERS:  { bg: 'rgba(245,158,11,0.15)', fg: 'var(--sb-warning)', border: 'rgba(245,158,11,0.4)' },
      INTERNAL_USE:                  { bg: 'rgba(239,68,68,0.15)',  fg: 'var(--sb-danger)', border: 'rgba(239,68,68,0.4)'  },
      DISCONTINUED:                  { bg: 'rgba(239,68,68,0.15)',  fg: 'var(--sb-danger)', border: 'rgba(239,68,68,0.4)'  },
    };
    const c = palette[status] || { bg: 'rgba(148,163,184,0.15)', fg: 'var(--sb-text-secondary)', border: 'rgba(148,163,184,0.4)' };
    const label = status.replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, ch => ch.toUpperCase());
    const span = document.createElement('span');
    span.textContent = label;
    span.style.cssText = `background:${c.bg};color:${c.fg};border:1px solid ${c.border};font-size:var(--sb-fs-10);font-weight:600;padding:1px 5px;border-radius:4px;`;
    return span;
  }

  function makeCopyBtn(textOrFn, label = '📋') {
    const b = el('button', { cls: 'sb-copy', text: label, title: 'Copy' });
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
    return el('div', { cls: 'sb-form-label', text });
  }

  /** Select dropdown for forms. */
  function makeSelect(options, selected) {
    const s = el('select', { cls: 'sb-select sb-select--form' });
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
    return el('div', { cls: 'sb-status' });
  }

  /** Styled tag (used in info bar and elsewhere). */
  function makeTag(label, color, bg, border) {
    // Colours are per-call (one of the --sb-tag-* triples), so they stay inline;
    // the class carries only the shape.
    return el('span', {
      cls: 'sb-tag', text: label,
      style: `color:${color};background:${bg};border:1px solid ${border};`,
    });
  }

  /** Standard action button pair (cancel + confirm) for dialogs. */
  function makeDialogButtons({ confirmLabel, confirmColor = 'var(--sb-danger-border)', onCancel, onConfirm }) {
    const row = el('div', { cls: 'sb-row', style: 'justify-content:flex-end;' });

    const cancelBtn = el('button', {
      cls: 'sb-btn sb-btn--lg sb-btn--grow',
      text: 'Cancel',
      style: 'background:transparent;color:var(--sb-text);padding:7px 16px;',
    });
    cancelBtn.onclick = onCancel;

    // confirmColor is a per-dialog parameter (danger / violet / success), so the
    // fill stays inline; --hover-fill drives the class's :hover without JS.
    const confirmBtn = el('button', {
      cls: 'sb-btn sb-btn--lg sb-btn--grow sb-btn--tinted',
      text: confirmLabel,
      style: `background:${confirmColor};border-color:transparent;color:var(--sb-on-solid);font-weight:700;padding:7px 16px;`,
    });
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
    toolbarEl.className = 'sb-ui';
    toolbarEl.style.cssText = `
      display:flex;align-items:center;gap:6px;
      padding:3px 10px;
      background:var(--sb-bg-row);
      border-bottom:1px solid var(--sb-control-bg);
      min-height:28px;flex-wrap:wrap;
      box-sizing:border-box;width:100%;
    `;

    document.getElementById('sb-info-bar-slot')?.remove();
    const infoSlot = document.createElement('div');
    infoSlot.id = 'sb-info-bar-slot';
    infoSlot.className = 'sb-ui';
    infoSlot.style.cssText = 'box-sizing:border-box;width:100%;';

    card.insertAdjacentElement('afterend', infoSlot);
    infoSlot.insertAdjacentElement('beforebegin', toolbarEl);
    return toolbarEl;
  }

  /** Toolbar button with a dropdown of item buttons. Items run their handler with the
   *  item element (for in-place ⏳/✔ feedback), so the menu stays open until a click
   *  lands outside it. Idempotent by id, like addToolbarButton. Returns the menu body
   *  (the element items go into), or null if it already exists. */
  function addToolbarMenu(id, label) {
    const toolbar = ensureToolbar();
    if (!toolbar || toolbar.querySelector('#' + id)) return null;
    const wrap = el('div', { id, style: 'position:relative;display:inline-flex;' });
    const btn = el('button', { cls: 'sb-btn sb-toolbar-btn', text: label + ' ▾' });
    const menu = el('div', { style: 'display:none;position:absolute;top:calc(100% + 4px);left:0;z-index:10000;min-width:200px;flex-direction:column;gap:4px;padding:6px;background:var(--sb-bg-panel);border:1px solid var(--sb-control-bg);border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,0.35);' });
    const close = () => { menu.style.display = 'none'; document.removeEventListener('mousedown', outside, true); };
    const outside = (e) => { if (!wrap.contains(e.target)) close(); };
    btn.addEventListener('click', () => {
      if (menu.style.display === 'none') { menu.style.display = 'flex'; document.addEventListener('mousedown', outside, true); }
      else close();
    });
    wrap.appendChild(btn); wrap.appendChild(menu);
    // Left group, where CCPA / OT used to sit — appending would land it past the
    // spacer on the right, after ⚙.
    const spacer = toolbar.querySelector('#sb-toolbar-spacer');
    spacer ? spacer.before(wrap) : toolbar.appendChild(wrap);
    return menu;
  }

  function addMenuItem(menu, id, label, onClick) {
    const btn = el('button', { id, cls: 'sb-btn sb-toolbar-btn', text: label, style: 'width:100%;text-align:left;' });
    btn.addEventListener('click', () => onClick(btn));
    menu.appendChild(btn);
    return btn;
  }

  function addToolbarButton(id, label, onClick) {
    const toolbar = ensureToolbar();
    if (!toolbar || toolbar.querySelector('#' + id)) return null;

    const btn = el('button', { id, cls: 'sb-btn sb-toolbar-btn', text: label });
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

  /** Fetch all emails + phone numbers from Kustomer for the current customer. */
  async function resolveKustomerAttributes() {
    const customerId = cachedCustomerCtx?._kustomerId || getCustomerIdFromURL();
    if (!customerId) return null;
    try {
      const res = await fetch(`https://scentbird.api.kustomerapp.com/v1/customers/${customerId}`, {
        credentials: 'include',
        headers: { 'Accept': 'application/json' },
      });
      if (!res.ok) return null;
      const attrs = (await res.json())?.data?.attributes;
      return {
        emails: (attrs?.emails || []).map(e => e.email).filter(Boolean),
        phones: (attrs?.phones || []).map(p => p.phone).filter(Boolean),
      };
    } catch(e) { return null; }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // GRAPHQL QUERIES & MUTATIONS (single source of truth)
  // ══════════════════════════════════════════════════════════════════════════

  const CRM_QUERY = `
    query userSearch($input: SearchUsersInput) {
      userSearch(input: $input) {
        total
        data { id email firstName lastName origin
          userAddress { shipping { street1 street2 city region postcode country } }
          subscription { id status subscriptionEndDate }
          subscriptionList { id status }
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
              orders { id status warehouseOrder { data { status } } }
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
          userAddress { shipping { id street1 street2 city region postcode country } }
          subscriptionList {
            id status subscribed nextBillingDate planName subscriptionDate subscriptionEndDate
            cashbirdDetails { data { cashbirdId billingDay nextBillingDate isAwaitCancellation } }
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
          commentList { id created author { firstName lastName email } comment kustomerUrl }
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
                id sku status
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
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ operationName, query, variables }),
      onload(res) {
        if (res.status === 401) { handle401(); return callback(null, 'Token expired'); }
        if (res.status === 403) { handle403(); return callback(null, 'CRM captcha required — open CRM in browser'); }
        try {
          const json = JSON.parse(res.responseText);
          const errMsg = json?.errors?.[0]?.message;
          if (errMsg) return callback(null, errMsg);
          if (!COMMENT_OPS.includes(operationName)) refreshSubscriptionBar();
          callback(json.data);
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); }
    });
  }

  function searchCRM(q, callback) {
    if (!getAccessToken()) { callback(null, 'Not signed in — click Sign in with Okta.'); return; }
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({
        operationName: 'userSearch', query: CRM_QUERY,
        variables: { input: { filter: q, statuses: [], page: { index: 1, size: 50 } } }
      }),
      onload(res) {
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

  // Run searchCRM across multiple terms grouped by tier, dedupe by user.id keeping
  // the lowest (best) tier seen. Used by Find Other Accounts (search panel) and AI
  // customer discovery (AI panel).
  async function searchCRMTiered({ tier1, tier2, tier3, excludeId } = {}) {
    const terms = [
      ...(tier1 || []).map(t => ({ term: t, tier: 1 })),
      ...(tier2 || []).map(t => ({ term: t, tier: 2 })),
      ...(tier3 || []).map(t => ({ term: t, tier: 3 })),
    ];
    const seen = new Map();
    const errors = [];
    const trace = [];
    for (const { term, tier } of terms) {
      if (!term) continue;
      await new Promise(resolve => {
        searchCRM(term, (users, err) => {
          if (err) {
            // Errors used to be dropped here, so a dead token or a captcha produced
            // an empty result set indistinguishable from "no such customer".
            errors.push({ term, tier, error: err });
            trace.push(`  tier${tier} "${term}" → ✖ ${err}`);
          } else {
            const hits = (users || []).filter(u => !(excludeId && u.id === excludeId));
            hits.forEach(u => { if (!seen.has(u.id)) seen.set(u.id, { user: u, tier }); });
            trace.push(`  tier${tier} "${term}" → ${hits.length} hit(s)`);
          }
          resolve();
        });
      });
    }
    const byTier = { 1: [], 2: [], 3: [] };
    seen.forEach(({ user, tier }) => byTier[tier].push(user));
    console.log(`[BirdsEye:search] ${terms.filter(t => t.term).length} term(s), ${seen.size} unique match(es)\n` + (trace.join('\n') || '  (no terms)'));
    if (errors.length) console.warn('[BirdsEye:search] ✖ some searches errored — an empty result may be a failure, not a miss', errors);
    byTier.errors = errors;
    return byTier;
  }

  function fetchLastOrders(userId, callback) {
    if (!getAccessToken()) { return callback(null); }
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ operationName: 'ordersByUserId', query: ORDERS_QUERY, variables: { id: userId } }),
      onload(res) {
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
    if (!getAccessToken()) { return callback(null); }
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ operationName: 'chargesByUserId', query: CHARGES_QUERY, variables: { id: userId } }),
      onload(res) {
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

  /** Fetch the full `userById.data` object via USER_DETAILS_QUERY. Returns the
   *  raw object (with full subscriptionList, addOnSettings, fraud info, etc.) or
   *  null on failure. Used to enrich the lean searchCRM result. */
  function fetchUserDetails(userId, callback) {
    if (!getAccessToken()) return callback(null, 'No token');
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ operationName: 'userDetailsById', query: USER_DETAILS_QUERY, variables: { id: userId } }),
      onload(res) {
        if (res.status === 401) { handle401(); return callback(null, 'Token expired'); }
        if (res.status === 403) { handle403(); return callback(null, 'CRM captcha required'); }
        try {
          const json = JSON.parse(res.responseText);
          callback(json?.data?.userById?.data || null);
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); },
    });
  }

  // Subscription statuses that can still be cancelled. 'Cancelled' is terminal;
  // everything else here is a live sub the agent can cancel (Unpaid / OnHold
  // subscriptions previously slipped through because we only matched 'Active').
  const CANCELLABLE_SUB_STATUSES = ['Active', 'Unpaid', 'OnHold'];

  function fetchActiveSubscription(userId, callback) {
    if (!getAccessToken()) { return callback(null, 'No token'); }
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ operationName: 'userDetailsById', query: USER_DETAILS_QUERY, variables: { id: userId } }),
      onload(res) {
        if (res.status === 401) { handle401(); return callback(null, 'Token expired'); }
        if (res.status === 403) { handle403(); return callback(null, 'CRM captcha required'); }
        try {
          const json = JSON.parse(res.responseText);
          const list = json?.data?.userById?.data?.subscriptionList || [];
          // Prefer a truly Active sub, then fall back to any cancellable one
          // (Unpaid / OnHold) so those statuses can be cancelled too.
          const active = list.find(s => s.status === 'Active' || s.subscribed)
                      || list.find(s => CANCELLABLE_SUB_STATUSES.includes(s.status));
          callback(active?.id || null, active ? null : 'No active subscription');
        } catch(e) { callback(null, 'Parse error'); }
      },
      onerror() { callback(null, 'Network error'); }
    });
  }

  function fetchComments(userId, callback) {
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
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

  /** Resolve active subscription, cancel it, post the CRM-side comment, update cache.
   *  Shared by showCancelPanel and the AI action layer. */
  function executeCancelSubscription(user, onProgress) {
    const progress = onProgress || (() => {});
    return new Promise((resolve, reject) => {
      progress('Fetching subscription…');
      fetchActiveSubscription(user.id, (subscriptionId, fetchErr) => {
        if (fetchErr || !subscriptionId) return reject(new Error(fetchErr || 'No active subscription found.'));
        progress('Cancelling subscription…');
        mutCancelSub(subscriptionId, (data, cancelErr) => {
          if (cancelErr) return reject(new Error(cancelErr));
          progress('Posting CRM comment…');
          mutPostComment(user.id, 'Subscription Cancelled', () => {
            if (cachedCustomerCtx?.user) {
              if (!cachedCustomerCtx.user.subscription) cachedCustomerCtx.user.subscription = {};
              cachedCustomerCtx.user.subscription.status = 'Cancelled';
            }
            refreshOpenPanels('sb-cancel-panel');
            resolve();
          });
        });
      });
    });
  }

  function mutCancelOrder(orderId, reasonOrCallback, maybeCallback) {
    const callback = typeof reasonOrCallback === 'function' ? reasonOrCallback : maybeCallback;
    const reason = typeof reasonOrCallback === 'string' ? reasonOrCallback : 'CUSTOMER_CHANGED_THEIR_MIND';
    gqlMutate('cancelOrderById',
      `mutation cancelOrderById($input: CancelOrderByIdInput!) {
        cancelOrderById(input: $input) {
          data { order { id status __typename } __typename }
          error { message __typename }
          __typename
        }
      }`,
      { input: { id: orderId, subscriptionHold: true, unlockItems: false, refundReason: reason } },
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

  function mutRefundCredits(userId, creditIds, shippingCreditIds, reasonOrCallback, maybeCallback) {
    const callback = typeof reasonOrCallback === 'function' ? reasonOrCallback : maybeCallback;
    const reason = typeof reasonOrCallback === 'string' ? reasonOrCallback : 'CUSTOMER_CHANGED_THEIR_MIND';
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
        reason
      }},
      (data, err) => callback(err ? null : data?.refundCredits, err)
    );
  }

  function mutRefundInvoiceItems(userId, items, reasonOrCallback, maybeCallback) {
    const callback = typeof reasonOrCallback === 'function' ? reasonOrCallback : maybeCallback;
    const reason = typeof reasonOrCallback === 'string' ? reasonOrCallback : 'CUSTOMER_CHANGED_THEIR_MIND';
    gqlMutate('refundInvoiceItems',
      `mutation refundInvoiceItems($input: RefundInvoiceItemListInput) {
        refundInvoiceItems(input: $input) {
          message
          error { message __typename }
          __typename
        }
      }`,
      { input: { userId, items, reason } },
      (data, err) => callback(err ? null : data?.refundInvoiceItems, err)
    );
  }

  // CRM's "unsubscribe from emails" (captured 2026-10-13). Keyed by email, not user id.
  const EMAIL_UNSUBSCRIBE_MUTATION = `
    mutation userEmailUnsubscribe($input: UserEmailUnsubscribeInput!) {
      userEmailUnsubscribe(input: $input) { data { email } error { message errorCode } }
    }`;
  async function crmUnsubscribeEmail(email) {
    const data = await crmQueryP('userEmailUnsubscribe', EMAIL_UNSUBSCRIBE_MUTATION, { input: { email } });
    const r = data.userEmailUnsubscribe;
    if (r?.error) throw new Error(r.error.message || r.error.errorCode || 'Unsubscribe failed');
    if (!r?.data) throw new Error('Unsubscribe: no result');
    return r.data.email;
  }

  /** Posts a comment, merging with our own recent comment for this ticket URL. */
  function mutPostComment(userId, comment, callback) {
    const kustomerUrl = window.location.href;
    const THREE_DAYS_MS = 3 * 24 * 60 * 60 * 1000;
    const tag = '[BirdsEye:comment]';
    console.log(`${tag} ▶ post requested`, { userId, comment, kustomerUrl, _myAuthorEmail });

    fetchComments(userId, (comments) => {
      console.log(`${tag} fetchComments returned ${comments.length} comment(s)`, comments);
      // Find our own recent comment on this ticket
      const now = Date.now();
      const match = comments.find(c => {
        if (c.kustomerUrl !== kustomerUrl) return false;
        if (_myAuthorEmail && c.author?.email !== _myAuthorEmail) return false;
        if (!c.created) return false;
        const age = now - new Date(c.created).getTime();
        return age <= THREE_DAYS_MS;
      });
      console.log(`${tag} merge match:`, match || '(none — will create fresh)');

      const finalComment = match ? match.comment + '\n' + comment : comment;

      const doCreate = () => {
        console.log(`${tag} → createUserComment`, { input: { userId, comment: finalComment, kustomerUrl } });
        gqlMutate('createUserComment',
          CREATE_COMMENT_MUTATION,
          { input: { userId, comment: finalComment, kustomerUrl } },
          (data, err) => {
            const payload = data?.createUserComment;
            console.log(`${tag} ◀ createUserComment response`, { err, payload, appError: payload?.error });
            if (err) console.error(`${tag} ✖ top-level error:`, err);
            else if (payload?.error) console.error(`${tag} ✖ app-level error:`, payload.error);
            else console.log(`${tag} ✔ server reports success:`, payload?.message);
            // Discover our author email after first successful create
            if (!err && !_myAuthorEmail) {
              fetchComments(userId, (freshComments) => {
                const ours = freshComments
                  .filter(c => c.kustomerUrl === kustomerUrl)
                  .sort((a, b) => new Date(b.created) - new Date(a.created))[0];
                console.log(`${tag} post-create readback (filtered by kustomerUrl):`, ours || '(NONE — comment not visible under this URL!)', `total now=${freshComments.length}`);
                if (ours?.author?.email) _myAuthorEmail = ours.author.email;
              });
            }
            callback(err ? null : data?.createUserComment, err);
          }
        );
      };

      if (match) {
        console.log(`${tag} deleting prior comment id=${match.id} before merge-recreate`);
        gqlMutate('deleteUserComment', DELETE_COMMENT_MUTATION, { id: match.id },
          (d, e) => { console.log(`${tag} delete response`, { err: e, payload: d?.deleteUserComment }); doCreate(); }   // proceed even if delete failed
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
    if (!getAccessToken()) {
      const orig = btn.textContent;
      btn.textContent = '\u2718 Sign in'; btn.style.color = 'var(--sb-danger)';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; }, 2000);
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
      btn.textContent = '\u2718 ' + msg; btn.style.color = 'var(--sb-danger)';
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
      // Require exact email match — never bind to a partial/related account
      const exactMatch = sbUsers.find(u => u.email?.toLowerCase() === email.toLowerCase());
      if (!exactMatch) return fail('No exact email match');
      cachedCustomerCtx = { email, user: exactMatch, _kustomerId: getCustomerIdFromURL() };
      callback({ user: exactMatch });
    });
  }

  function loadCustomerOrders(userId, callback) {
    fetchLastOrders(userId, (orders) => callback(orders || []));
  }

  function loadCustomerCharges(userId, callback) {
    fetchCharges(userId, (result) => callback(result));
  }

  function fetchQueue(userId, callback) {
    if (!getAccessToken()) { return callback(null); }
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ operationName: 'subscriptionQueue', query: QUEUE_QUERY, variables: { input: { userId } } }),
      onload(res) {
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
    const lines = orderProductLines(order);
    const isWelcomeKit = (order.tags || []).includes('REBRAND_CASE') || order.hasRebrandCase;
    if (isWelcomeKit) lines.push('Signature Grey Fragrance Case (welcome gift)');
    return lines.join('\n');
  }

  const NON_REFUNDABLE_STATUSES = ["NEW", "SHIPPED", "DELIVERED", "PRINTED", "DONE", "PROCESSED"];

  /**
   * Refund gate for a charge, judged across ALL of its linked orders.
   *
   * A charge carries `orders` as a LIST, and an upcharge puts more than one in it:
   * the original is superseded (status UPGRADED — a shell that carries no fulfilment
   * state) and the real order sits alongside it. Reading only `orders[0]` can
   * therefore land on the shell, and UPGRADED is not in NON_REFUNDABLE_STATUSES, so
   * the refund is offered on goods that already PRINTED or shipped.
   *
   * Skip the shells, judge every real order, block if ANY is past the point of no
   * return. Returns the blocking status so the button can name it.
   */
  function chargeRefundBlock(charge) {
    for (const o of (charge?.orders || [])) {
      const own = (o?.status || '').toUpperCase();
      const wh  = (o?.warehouseOrder?.data?.status || '').toUpperCase();
      // A shell's own status (UPGRADED) says nothing about fulfilment, but its
      // warehouse record still can — so read that rather than dropping the order
      // outright, which is how a printed upcharge order used to slip the gate.
      const st = wh || (own === 'UPGRADED' ? '' : own);
      if (st && NON_REFUNDABLE_STATUSES.includes(st)) return { blocked: true, status: st };
    }
    return { blocked: false, status: '' };
  }

  /** Live colour for an Asendia event, by its free-text description. */
  function asendiaEventColor(desc) {
    const d = desc || '';
    if (/deliver/i.test(d) && !/out for|attempt|unable|failed|not delivered/i.test(d)) return 'var(--sb-success)';
    if (/return|exception|unable|failed|lost/i.test(d)) return 'var(--sb-danger)';
    return 'var(--sb-warning)';
  }

  /** "ASENDIA LIVE" block for an order card: placeholder now, last 3 scans when the
   *  lookup lands. Oldest → newest like the CRM lines; the newest is bold + coloured. */
  function renderAsendiaLive(trackNo) {
    const box = document.createElement('div');
    box.style.cssText = 'margin-top:6px;padding-top:4px;border-top:1px dashed var(--sb-control-bg);';
    const head = document.createElement('div');
    head.style.cssText = 'color:var(--sb-text-dim);font-size:var(--sb-fs-9);font-weight:700;letter-spacing:0.05em;';
    head.textContent = 'ASENDIA LIVE ⏳';
    box.appendChild(head);

    asendiaTracking(trackNo).then((live) => {
      if (!live.found) { head.textContent = 'ASENDIA LIVE · no record'; return; }
      const events = live.events.slice(-3);
      head.textContent = 'ASENDIA LIVE' + (events.length ? '' : ' · no scans yet');
      head.title = [live.service, live.lastMileMessage, live.finalMile?.carrier && `Final mile: ${live.finalMile.carrier} ${live.finalMile.trackingNumber || ''}`]
        .filter(Boolean).join('\n');
      events.forEach((e, i) => {
        const isLast = i === events.length - 1;
        const d = document.createElement('div');
        d.style.cssText = `color:${isLast ? asendiaEventColor(e.description) : 'var(--sb-text-dim)'};font-size:var(--sb-fs-11);margin-top:2px;${isLast ? 'font-weight:600;' : ''}`;
        const date = new Date(e.at).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        d.textContent = `${date} · ${e.description}${e.location ? ' — ' + e.location : ''}`;
        box.appendChild(d);
      });
      if (live.finalMile?.link) {
        const a = document.createElement('a');
        a.href = live.finalMile.link; a.target = '_blank'; a.rel = 'noopener';
        a.textContent = `${live.finalMile.carrier || 'Final mile'} ${live.finalMile.trackingNumber || 'tracking'} ↗`;
        a.style.cssText = 'display:inline-block;margin-top:3px;color:var(--sb-accent-strong);font-size:var(--sb-fs-10);text-decoration:none;';
        box.appendChild(a);
      }
    }).catch((e) => {
      head.textContent = 'ASENDIA LIVE · ' + (e?.message || 'lookup failed');
      head.style.color = 'var(--sb-danger)';
    });
    return box;
  }

  function renderOrderBlock(order, compact = false, ctx = null, isReplacement = false) {
    const wrap = document.createElement('div');
    const borderColor = isReplacement ? 'var(--sb-solid-violet)' : 'var(--sb-solid-primary)';
    wrap.style.cssText = `
      margin-top:${compact ? 4 : 0}px;padding:${compact ? '6px 8px' : '10px 12px'};
      background:${isReplacement ? 'var(--sb-bg-row-replacement)' : 'var(--sb-bg-row)'};border-radius:6px;border-left:2px solid ${borderColor};
    `;

    const _ws = (order.warehouseOrder?.data?.status || order.status || '').toUpperCase();
    const statusColor = (_ws === 'SHIPPED' || _ws === 'DELIVERED') ? 'var(--sb-success-soft)'
      : (_ws.includes('CANCEL') || _ws.includes('BACK')) ? 'var(--sb-danger)'
      : (_ws === 'PRINTED' || _ws.includes('PROCESS')) ? 'var(--sb-accent-strong)'
      : (_ws === 'PENDING') ? 'var(--sb-warning-strong)'
      : 'var(--sb-text-muted)';

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
      <span style="color:var(--sb-text-muted);font-size:var(--sb-fs-11);">${orderMonthLabel(order)}</span>
      <span style="color:${statusColor};font-weight:600;font-size:var(--sb-fs-11);">${order.warehouseOrder?.data?.status || order.status}</span>
      ${(order.type && order.type !== 'SUBSCRIPTION') ? `<span style="color:var(--sb-warning-strong);font-size:var(--sb-fs-11);font-weight:600;">${order.type.replace(/_/g,' ')}</span>` : ''}
      ${hasWelcomeKit ? '<span style="background:rgba(99,102,241,0.2);color:var(--sb-accent);font-size:var(--sb-fs-10);font-weight:600;padding:1px 5px;border-radius:4px;border:1px solid rgba(99,102,241,0.4);">Welcome Kit</span>' : ''}
      ${hasSamples ? '<span style="background:rgba(245,158,11,0.15);color:var(--sb-warning);font-size:var(--sb-fs-10);font-weight:600;padding:1px 5px;border-radius:4px;border:1px solid rgba(245,158,11,0.3);">Samples</span>' : ''}
      ${hasCase ? '<span style="background:rgba(99,102,241,0.15);color:var(--sb-accent);font-size:var(--sb-fs-10);font-weight:600;padding:1px 5px;border-radius:4px;border:1px solid rgba(99,102,241,0.3);">Case</span>' : ''}
    `;
    hdr.appendChild(makeCopyBtn(() => orderCopyText(order), '📋 Copy'));
    wrap.appendChild(hdr);

    // Product lines — append a plaque if the product status is anything but LIVE.
    (order.orderItems || []).forEach(i => {
      const p = i?.product?.productInfo;
      if (!p) return;
      const d = document.createElement('div');
      d.style.cssText = 'color:var(--sb-text-secondary);font-size:var(--sb-fs-12);margin-bottom:2px;display:flex;align-items:center;gap:6px;flex-wrap:wrap;';
      const text = document.createElement('span');
      text.textContent = `${p.name} by ${p.brand}`;
      d.appendChild(text);
      const plaque = makeProductStatusPlaque(i.product?.status);
      if (plaque) d.appendChild(plaque);
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
        link.style.cssText = 'color:var(--sb-accent-strong);font-size:var(--sb-fs-11);text-decoration:none;';
        tRow.appendChild(link);
      } else {
        const span = document.createElement('span');
        span.style.cssText = 'color:var(--sb-accent-strong);font-size:var(--sb-fs-11);';
        span.textContent = trackNo;
        tRow.appendChild(span);
      }

      if (trackUrl) {
        const linkBtn = document.createElement('button');
        linkBtn.textContent = '🔗';
        linkBtn.title = 'Copy tracking link';
        linkBtn.style.cssText = 'background:none;border:none;cursor:pointer;color:var(--sb-accent-strong);font-size:var(--sb-fs-13);padding:0 4px;flex-shrink:0;';
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

    // Last 2 tracking events — most recent line uses a status-bright color.
    const items = order.tracking?.items;
    if (items?.length) {
      const statusEl = document.createElement('div');
      statusEl.style.cssText = 'margin-top:6px;';
      const slice = items.slice(-2);
      slice.forEach((item, i) => {
        const isLast = i === slice.length - 1;
        const color = isLast ? trackingStatusColor(item.status) : 'var(--sb-text-dim)';
        const d = document.createElement('div');
        d.style.cssText = `color:${color};font-size:var(--sb-fs-11);margin-top:2px;${isLast ? 'font-weight:600;' : ''}`;
        const date = new Date(item.date).toLocaleDateString('en-US', { month:'short', day:'numeric' });
        d.textContent = `${date} · ${item.description}`;
        statusEl.appendChild(d);
      });
      wrap.appendChild(statusEl);
    }

    // Live Asendia scans (last 3) under the CRM events — CRM's copy of Asendia
    // tracking lags. Full panel only; the lookup is cached for 10 min per number.
    if (!compact && trackNo && /^asendia/.test(_orderCourier(order) || '')) {
      wrap.appendChild(renderAsendiaLive(trackNo));
    }

    // Replace + Cancel buttons — only in full panel
    if (!compact && ctx) {
      const btnRow = document.createElement('div');
      btnRow.style.cssText = 'display:flex;gap:6px;margin-top:8px;flex-wrap:wrap;';

      const isReplacementOrder = order.type === 'REPLACEMENT' || order.type === 'REPLACEMENT_ORDER';
      const replBtn = document.createElement('button');
      replBtn.textContent = '\uD83D\uDD04 Replace';
      replBtn.disabled = isReplacementOrder;
      replBtn.title = isReplacementOrder ? 'Cannot replace a replacement order' : '';
      replBtn.style.cssText = `padding:5px 10px;border-radius:5px;border:1px solid ${isReplacementOrder ? 'rgba(99,102,241,0.3)' : 'var(--sb-solid-primary-lo)'};background:transparent;color:${isReplacementOrder ? 'rgba(165,180,252,0.4)' : 'var(--sb-accent)'};font-weight:600;font-size:var(--sb-fs-11);cursor:${isReplacementOrder ? 'not-allowed' : 'pointer'};`;
      if (!isReplacementOrder) {
        replBtn.onmouseenter = () => replBtn.style.background = 'rgba(99,102,241,0.15)';
        replBtn.onmouseleave = () => replBtn.style.background = 'transparent';
        replBtn.onclick = () => showReplacementForm(order, ctx.user, ctx.totalOrders);
      }
      btnRow.appendChild(replBtn);

      const ws = (order.warehouseOrder?.data?.status || order.status || '').toUpperCase();
      const cancellable = !NON_REFUNDABLE_STATUSES.includes(ws);
      if (cancellable) {
      const cancelOrderBtn = document.createElement('button');
      cancelOrderBtn.textContent = '\u274C Cancel Order';
      cancelOrderBtn.style.cssText = 'padding:5px 10px;border-radius:5px;border:1px solid var(--sb-danger-border);background:transparent;color:var(--sb-danger);font-weight:600;font-size:var(--sb-fs-11);cursor:pointer;';
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
        holdBtn.style.cssText = `padding:5px 10px;border-radius:5px;border:1px solid ${isOnHold ? 'var(--sb-success-soft)' : 'var(--sb-warning-strong)'};background:transparent;color:${isOnHold ? 'var(--sb-success-soft)' : 'var(--sb-warning-strong)'};font-weight:600;font-size:var(--sb-fs-11);cursor:pointer;`;
        holdBtn.onmouseenter = () => holdBtn.style.background = isOnHold ? 'rgba(110,231,183,0.15)' : 'rgba(245,158,11,0.15)';
        holdBtn.onmouseleave = () => holdBtn.style.background = 'transparent';
        holdBtn.onclick = () => {
          holdBtn.disabled = true;
          holdBtn.textContent = isOnHold ? 'Unholding...' : 'Holding...';
          const mut = isOnHold ? mutUnholdOrder : mutHoldOrder;
          mut(order.id, (data, err) => {
            if (err) {
              holdBtn.textContent = '✘ ' + err;
              holdBtn.style.color = 'var(--sb-danger)';
              setTimeout(() => { holdBtn.disabled = false; holdBtn.textContent = isOnHold ? '▶ Unhold' : '⏸ Hold'; holdBtn.style.color = isOnHold ? 'var(--sb-success-soft)' : 'var(--sb-warning-strong)'; }, 2000);
              return;
            }
            holdBtn.textContent = '✔ Done';
            holdBtn.style.color = 'var(--sb-success-soft)';
            holdBtn.style.borderColor = 'var(--sb-success-soft)';
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

  function lineItemExplanation(item, billingDate, currency, hideTax) {
    const t = item.type;
    if (t === 'FAILED_CHARGE') return null;

    const tax  = item.tax;
    const disc = item.discount;
    const vars = {
      total:     fmtMoney(item.total, currency),
      price:     fmtMoney(item.price, currency),
      month:     chargeMonthLabel(billingDate),
      taxClause: (!hideTax && tax  > 0) ? ` + ${fmtMoney(tax, currency)} tax` : '',
      discStr:   disc ?    `, including a ${fmtMoney(disc, currency)} discount` : '',
      taxStr:    (!hideTax && tax  > 0) ? ` (after ${fmtMoney(tax, currency)} for sales tax)` : '',
      name:      (item.description && item.description.toLowerCase() !== 'premium product')
                   ? item.description : 'a premium fragrance',
    };

    const key = t === 'SHIPPING_CREDIT' ? 'SHIPPING_UPCHARGE' : t;
    const template = getBillTemplates()[key];
    if (!template) return `* ${item.description || t} — ${vars.total}${vars.taxStr}; [unknown type: ${t}]`;
    return applyBillTemplate(template, vars);
  }

  function buildBillExplanation(successEntry) {
    const charge = successEntry?.charge;
    if (!charge) return '';
    const items = charge.cashbirdDetails?.invoice?.lineItems || [];
    const billingDate = charge.paymentDate || successEntry.date;
    const currency = detectCurrency(items);

    const country = (cachedCustomerCtx?.user?.userAddress?.shipping?.country || '').trim().toUpperCase();
    const autoHide = country === 'CA' || country === 'CAN' || country === 'CANADA' ||
                     country === 'GB' || country === 'UK'  || country === 'UNITED KINGDOM';
    const hideTax = getSettings().hideTax || autoHide;

    // Fold standalone TAX line items (e.g. Colorado state tax) into the first
    // subscription-payment item so the bill reads as one charge with one tax figure.
    const extraTax = items
      .filter(i => i.type === 'TAX')
      .reduce((s, i) => s + (i.total || 0), 0);
    const TAX_TARGETS = new Set(['FIRST_SUBSCRIPTION', 'RESUBSCRIBE', 'RECURRENT_CHARGE']);
    let absorbed = false;
    const lines = items
      .filter(i => i.type !== 'TAX' && (i.total || 0) !== 0)
      .map(i => {
        if (!absorbed && TAX_TARGETS.has(i.type)) {
          absorbed = true;
          return lineItemExplanation({ ...i, tax: (i.tax || 0) + extraTax }, billingDate, currency, hideTax);
        }
        return lineItemExplanation(i, billingDate, currency, hideTax);
      })
      .filter(Boolean);

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

  // Apply a CRM user to the current Kustomer customer: prepend the user's email
  // (position 0 = "primary" / starred), keep other emails, fill name only if blank.
  // Mirrors the Kustomer modal's Save payload — see Apply-to-Kustomer HAR recon.
  async function applyUser(user) {
    if (user.origin && user.origin !== 'SCENTBIRD') return;
    const kustomerId = getCustomerIdFromURL();
    if (!kustomerId) throw new Error('No Kustomer customer in URL');
    const target = (user.email || '').trim();
    if (!target) throw new Error('Target user has no email');

    const attrs = await kustomerGetCustomer(kustomerId);
    const targetLc = target.toLowerCase();
    const others = (attrs.emails || []).filter(e => (e.email || '').toLowerCase() !== targetLc);
    attrs.emails = [{ type: 'home', email: target }, ...others];

    if (!(attrs.name || '').trim()) {
      const full = toProperCase([user.firstName, user.lastName].filter(Boolean).join(' '));
      if (full) attrs.name = full;
    }

    await kustomerPatchCustomer(kustomerId, attrs);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // PANELS
  // ══════════════════════════════════════════════════════════════════════════

  function removePanel(id) {
    document.getElementById(id)?.remove();
  }

  /** Refresh any open panel that exposes a `_refresh()` method. No-op if the
   *  panel isn't currently in the DOM. Callers pass the ids of panels whose
   *  data was just mutated by an action (cancel, refund, shipment, replacement). */
  function refreshOpenPanels(...ids) {
    ids.forEach(id => document.getElementById(id)?._refresh?.());
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
      flex:1;padding:7px 10px;border-radius:6px;border:1px solid var(--sb-hover);
      background:var(--sb-bg-input);color:var(--sb-text);font-size:var(--sb-fs-13);outline:none;
    `;

    const searchBtn = document.createElement('button');
    searchBtn.textContent = 'Search';
    searchBtn.style.cssText = `
      padding:7px 14px;border-radius:6px;border:none;background:var(--sb-solid-primary);
      color:var(--sb-on-solid);font-weight:700;font-size:var(--sb-fs-13);cursor:pointer;
    `;

    const filterBtn = document.createElement('button');
    filterBtn.textContent = 'Filter Active';
    filterBtn.style.cssText = `
      padding:7px 10px;border-radius:6px;border:1px solid rgba(110,231,183,0.3);background:transparent;
      color:var(--sb-success-soft);font-weight:600;font-size:var(--sb-fs-11);cursor:pointer;white-space:nowrap;
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

    function renderUsers(users, headerText, target, highlightTerm) {
      const container = target || results;
      if (!target) container.innerHTML = '';
      if (headerText) {
        const hdr = document.createElement('div');
        hdr.style.cssText = 'color:var(--sb-text-muted);margin-bottom:8px;font-size:var(--sb-fs-12);';
        hdr.textContent = headerText;
        container.appendChild(hdr);
      }
      users.forEach(user => {
        const card = document.createElement('div');
        card.dataset.subStatus = user.subscription?.status || '';
        card.style.cssText = `
          padding:10px;margin-bottom:8px;border-radius:8px;background:var(--sb-bg-input);
          border:1px solid transparent;transition:0.15s;
        `;
        if (filterActive && card.dataset.subStatus !== 'Active' && card.dataset.subStatus !== 'Unpaid' && card.dataset.subStatus !== 'OnHold') card.style.display = 'none';
        const s = user.userAddress?.shipping;
        const addrRaw = s ? [s.street1,s.street2,s.city,s.region,s.postcode].filter(Boolean).join(', ') : '';
        const addr = (highlightTerm && addrRaw)
          ? addrRaw.replace(new RegExp('(' + highlightTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'i'),
              '<span style="color:var(--sb-accent);font-weight:700;">$1</span>')
          : addrRaw;
        const subSt = user.subscription?.status || '';
        const subColor = subSt === 'Active' ? 'var(--sb-success-soft)' : subSt ? 'var(--sb-danger)' : 'var(--sb-text-dim)';
        let subLabel = subSt || 'No subscription';
        if (subSt === 'Cancelled' && user.subscription?.subscriptionEndDate) {
          subLabel = 'Cancelled ' + fmtDateTag(user.subscription.subscriptionEndDate);
        }
        const isDrift = user.origin && user.origin !== 'SCENTBIRD';
        if (isDrift) card.style.opacity = '0.5';
        card.innerHTML = `
          <div style="display:flex;justify-content:space-between;align-items:baseline;">
            <div style="font-weight:700;">${user.firstName||''} ${user.lastName||''} ${isDrift ? '<span style="font-size:var(--sb-fs-10);color:var(--sb-warning-strong);font-weight:600;background:rgba(245,158,11,0.15);padding:1px 5px;border-radius:3px;margin-left:4px;">' + user.origin + '</span>' : ''}</div>
            <div style="font-size:var(--sb-fs-11);font-weight:600;color:${subColor};">${subLabel}</div>
          </div>
          <div style="color:var(--sb-text-muted);margin-top:2px;">${user.email}</div>
          <div style="color:var(--sb-text-dim);font-size:var(--sb-fs-11);margin-top:2px;">ID: ${user.id}</div>
          ${addr ? `<div style="color:var(--sb-text-dim);font-size:var(--sb-fs-11);margin-top:2px;">${addr}</div>` : ''}
          <div style="margin-top:8px;display:flex;gap:6px;">
            <button class="sb-apply" style="padding:4px 12px;border-radius:5px;border:none;background:var(--sb-solid-primary);color:var(--sb-on-solid);font-size:var(--sb-fs-12);font-weight:600;cursor:pointer;">Apply to Kustomer</button>
            <a class="sb-profile" href="https://crm.scentbird.com/user/${user.id}/profile/subscription" target="_blank" style="padding:4px 12px;border-radius:5px;border:1px solid var(--sb-hover);background:transparent;color:var(--sb-text);font-size:var(--sb-fs-12);cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;">Open Profile</a>
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
            try {
              await applyUser(user);
              cachedCustomerCtx = { email: user.email.trim().toLowerCase(), user, _kustomerId: getCustomerIdFromURL() };
              btn.textContent = '✔ Done'; btn.style.background = 'var(--sb-solid-success)';
              autoClose(() => removePanel('sb-search-panel'));
              const fillBtn = document.getElementById('sb-fill-name-btn');
              if (fillBtn && !fillBtn.disabled) fillBtn.click();
            } catch (err) {
              const msg = err.message || 'Apply failed';
              if (/Duplicate email/i.test(msg) && await offerMergeOnDuplicateError(btn, user)) return;
              btn.textContent = '✘ ' + msg;
              btn.style.background = 'var(--sb-danger-border)';
              btn.disabled = false;
            }
          };
        }

        const orderEl = document.createElement('div');
        orderEl.style.cssText = 'margin-top:8px;font-size:var(--sb-fs-11);color:var(--sb-text-dim);';
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

        container.appendChild(card);
      });
    }

    function doSearch() {
      const q = input.value.trim(); if (!q) return;
      results.innerHTML = `<div style="color:var(--sb-text-muted);padding:8px 0;">Searching…</div>`;
      searchCRM(q, (users, err) => {
        if (err) { results.innerHTML = `<div style="color:var(--sb-danger);padding:8px 0;">${err}</div>`; return; }
        if (!users.length) { results.innerHTML = `<div style="color:var(--sb-text-muted);padding:8px 0;">No users found.</div>`; return; }
        renderUsers(users, `${users.length} result(s)`);
      });
    }

    // Find Other Accounts button
    const findOtherBtn = document.createElement('button');
    findOtherBtn.textContent = '🔎 Find Other Accounts';
    findOtherBtn.style.cssText = `
      width:100%;padding:7px;border-radius:6px;border:1px solid rgba(148,163,184,0.25);
      background:transparent;color:var(--sb-text-muted);font-weight:600;font-size:var(--sb-fs-12);
      cursor:pointer;margin-bottom:10px;
    `;
    findOtherBtn.onmouseenter = () => findOtherBtn.style.background = 'rgba(148,163,184,0.08)';
    findOtherBtn.onmouseleave = () => findOtherBtn.style.background = 'transparent';
    findOtherBtn.onclick = async () => {
      findOtherBtn.disabled = true;
      findOtherBtn.textContent = '⏳ Searching…';
      results.innerHTML = '';

      const currentUser = cachedCustomerCtx?.user;
      const currentId = currentUser?.id;
      const primaryEmail = (cachedCustomerCtx?.email || '').toLowerCase();

      // Collect terms by tier: 1=email/phone, 2=address, 3=name
      const tier1Email = [], tier1Phone = [], tier2 = [], tier3 = [];

      // Populate Kustomer-sourced emails/phones once on cachedCustomerCtx, then read from cache
      if (cachedCustomerCtx && (cachedCustomerCtx._allEmails === undefined || cachedCustomerCtx._phones === undefined)) {
        const kAttrs = await resolveKustomerAttributes();
        cachedCustomerCtx._allEmails = kAttrs?.emails || [];
        cachedCustomerCtx._phones    = kAttrs?.phones || [];
      }
      const cachedEmails = cachedCustomerCtx?._allEmails || [];
      const cachedPhones = cachedCustomerCtx?._phones || [];

      cachedEmails
        .filter(e => e.toLowerCase() !== primaryEmail)
        .forEach(e => tier1Email.push(e));
      cachedPhones.forEach(p => tier1Phone.push(p));

      const street1 = currentUser?.userAddress?.shipping?.street1;
      if (street1) tier2.push(street1);

      const fullName = [currentUser?.firstName, currentUser?.lastName].filter(Boolean).join(' ');
      if (fullName) tier3.push(fullName);

      const tier1 = [...tier1Email, ...tier1Phone];

      if (!tier1.length && !tier2.length && !tier3.length) {
        results.innerHTML = `<div style="color:var(--sb-text-muted);padding:8px 0;">No search terms available — load a customer first.</div>`;
        findOtherBtn.disabled = false; findOtherBtn.textContent = '🔎 Find Other Accounts';
        return;
      }

      const byTier = await searchCRMTiered({ tier1, tier2, tier3, excludeId: currentId });

      findOtherBtn.disabled = false; findOtherBtn.textContent = '🔎 Find Other Accounts';

      const total = byTier[1].length + byTier[2].length + byTier[3].length;
      if (!total) {
        results.innerHTML = `<div style="color:var(--sb-text-muted);padding:8px 0;">No other accounts found.</div>`;
        return;
      }

      results.innerHTML = '';
      const summary = document.createElement('div');
      summary.style.cssText = 'color:var(--sb-text-muted);margin-bottom:8px;font-size:var(--sb-fs-12);';
      summary.textContent = `Found ${total} other account(s)`;
      results.appendChild(summary);

      const sections = [
        { tier: 1, label: 'By email / phone' },
        { tier: 2, label: 'By address' },
        { tier: 3, label: 'By name' },
      ];
      sections.forEach(({ tier, label }) => {
        if (!byTier[tier].length) return;
        const sectionHdr = document.createElement('div');
        sectionHdr.style.cssText = 'font-weight:700;font-size:var(--sb-fs-12);color:var(--sb-text-secondary);margin:12px 0 6px;';
        sectionHdr.textContent = `${label} (${byTier[tier].length})`;
        results.appendChild(sectionHdr);
        renderUsers(byTier[tier], null, results, tier === 2 ? tier2[0] : '');
      });
    };

    searchBtn.onclick = doSearch;
    input.addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });
    row.appendChild(input); row.appendChild(searchBtn); row.appendChild(filterBtn);
    panel.appendChild(row); panel.appendChild(findOtherBtn); panel.appendChild(results);
    if (prefill) doSearch();
    setTimeout(() => input.focus(), 50);
  }

  // ── Last Orders Panel ─────────────────────────────────────────────────────

  function showLastOrderPanel(btn) {
    removePanel('sb-order-panel');

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;
      const panel = createPanel({ id: 'sb-order-panel', title: '📦 Last Orders', width: 380 });
      function render() {
        [...panel.children].slice(1).forEach(c => c.remove());
        loadCustomerOrders(user.id, (rawOrders) => {
        if (!document.getElementById('sb-order-panel')) return;
        // Hide upgraded orders — they are superseded shells that just clutter the list.
        const orders = (rawOrders || []).filter(o => (o.status || '').toUpperCase() !== 'UPGRADED');
        // Identity + shipping address now live in the user info bar; no duplicate header here.

        if (!orders?.length) {
          const empty = document.createElement('div');
          empty.style.color = 'var(--sb-text-muted)';
          empty.textContent = 'No orders found.';
          panel.appendChild(empty);
          return;
        }

        // Pick mode: prompt the agent to click an order to complete a pending Fill.
        if (_fillPickState) {
          const hint = document.createElement('div');
          hint.style.cssText = 'display:flex;align-items:center;justify-content:space-between;gap:8px;background:rgba(129,140,248,0.15);border:1px solid var(--sb-accent);border-radius:7px;padding:7px 10px;margin-bottom:10px;font-size:var(--sb-fs-12);color:var(--sb-text);';
          const label = document.createElement('span');
          label.textContent = '✏️ Pick an order to fill';
          const cancel = document.createElement('span');
          cancel.textContent = 'Cancel ✕';
          cancel.style.cssText = 'cursor:pointer;color:var(--sb-text-muted);font-weight:600;';
          cancel.onclick = cancelFillPick;
          hint.appendChild(label); hint.appendChild(cancel);
          panel.appendChild(hint);
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
            sep.style.cssText = 'height:1px;background:var(--sb-border-soft);margin:10px 0;';
            panel.appendChild(sep);
          }

          g.mains.forEach((o, oi) => {
            if (oi > 0) {
              const s = document.createElement('div');
              s.style.cssText = 'height:1px;background:var(--sb-border-soft);margin:8px 0;';
              panel.appendChild(s);
            }
            panel.appendChild(pickWrap(renderOrderBlock(o, false, orderCtx, false), o));
          });

          g.replacements.forEach(r => {
            const nest = document.createElement('div');
            nest.style.cssText = 'margin-left:12px;margin-top:6px;padding-left:8px;border-left:1px solid rgba(124,58,237,0.3);';
            const replLabel = document.createElement('div');
            replLabel.style.cssText = 'font-size:var(--sb-fs-10);color:var(--sb-solid-violet);font-weight:600;letter-spacing:0.05em;margin-bottom:4px;';
            replLabel.textContent = 'REPLACEMENT';
            nest.appendChild(replLabel);
            nest.appendChild(pickWrap(renderOrderBlock(r, false, orderCtx, true), r));
            panel.appendChild(nest);
          });
        });

        // Custom Shipment button
        const csSep = document.createElement('div');
        csSep.style.cssText = 'height:1px;background:var(--sb-control-bg);margin:12px 0;';
        panel.appendChild(csSep);

        const csBtn = document.createElement('button');
        csBtn.textContent = '📦 Custom Shipment';
        csBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid var(--sb-solid-primary);background:transparent;
          color:var(--sb-accent);font-weight:600;font-size:var(--sb-fs-12);cursor:pointer;box-sizing:border-box;`;
        csBtn.onmouseenter = () => csBtn.style.background = 'rgba(79,70,229,0.1)';
        csBtn.onmouseleave = () => csBtn.style.background = 'transparent';
        csBtn.onclick = () => showCustomShipmentForm(user);
        panel.appendChild(csBtn);
      });
      }
      render();
      panel._refresh = render;
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
          errEl.style.color = 'var(--sb-danger)';
          errEl.textContent = 'Failed to load charges.';
          panel.appendChild(errEl);
          return;
        }

        const { success, failed } = result;

        // Failed charges
        if (failed?.length) {
          const failHdr = document.createElement('div');
          failHdr.style.cssText = 'color:var(--sb-danger);font-size:var(--sb-fs-11);font-weight:600;margin-bottom:4px;';
          failHdr.textContent = `\u274C ${failed.length} failed charge attempt${failed.length > 1 ? 's' : ''}`;
          panel.appendChild(failHdr);

          const failWrap = document.createElement('div');
          failWrap.style.cssText = 'background:var(--sb-note-danger-bg);border-radius:6px;padding:6px 10px;margin-bottom:12px;';
          failed.forEach(entry => {
            const c = entry.charge;
            const desc = c?.cashbirdDetails?.invoice?.lineItems?.[0]?.description || c?.category || '\u2014';
            const row = document.createElement('div');
            row.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-11);margin-bottom:3px;display:flex;gap:8px;align-items:baseline;';
            row.innerHTML = `
              <span style="color:var(--sb-text-dim);white-space:nowrap;">${chargeDateLabel(c?.paymentDate || entry.date)}</span>
              <span style="color:var(--sb-danger);white-space:nowrap;">${fmtMoney((c?.totalPrice || 0) * 100)}</span>
              <span style="color:var(--sb-text-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;">${desc}</span>
            `;
            failWrap.appendChild(row);
          });
          panel.appendChild(failWrap);
        }

        // Successful charges
        const allSuccess = result.all || (result.success ? [result.success] : []);
        if (!allSuccess.length) {
          const noEl = document.createElement('div');
          noEl.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-12);';
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
          CAR_SCENT_UPCHARGE:       'Drift Car',
          CANDLE_UPCHARGE:          'Candle Club',
          HSD_UPCHARGE:             'Home Diffuser',
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
            sep.style.cssText = 'height:1px;background:var(--sb-border-soft);margin:10px 0;';
            panel.appendChild(sep);
          }

          const cHdr = document.createElement('div');
          cHdr.style.cssText = 'display:flex;align-items:baseline;justify-content:space-between;margin-bottom:4px;';
          const lineItemsForRefundCalc = c?.cashbirdDetails?.invoice?.lineItems || [];
          const refundAmt = lineItemsForRefundCalc.reduce((sum, li) => sum + (li.refundedAmount || 0), 0);
          const totalRefunded = refundAmt >= Math.round(c.totalPrice * 100);
          const refundTag = refundAmt > 0
            ? ` <span style="font-size:var(--sb-fs-11);font-weight:600;color:${totalRefunded ? 'var(--sb-success-soft)' : 'var(--sb-warning-strong)'};">&#8617; ${fmtMoney(refundAmt, currency)} refunded</span>`
            : '';
          cHdr.innerHTML = `
            <span style="font-weight:700;font-size:var(--sb-fs-13);color:var(--sb-success-soft);">\u2714 ${fmtMoney(c.totalPrice * 100, currency)}</span>${refundTag}
            <span style="color:var(--sb-text-dim);font-size:var(--sb-fs-11);">${chargeDateLabel(billingDate)}</span>
          `;
          panel.appendChild(cHdr);

          if (pm) {
            const pmRow = document.createElement('div');
            pmRow.style.cssText = 'color:var(--sb-text-dim);font-size:var(--sb-fs-11);margin-bottom:8px;';
            pmRow.textContent = `${pm.type} \u00b7 ${pm.methodName}`;
            panel.appendChild(pmRow);
          }

          const lineWrap = document.createElement('div');
          lineWrap.style.cssText = 'background:var(--sb-bg-input);border-radius:6px;padding:8px 10px;margin-bottom:8px;';
          const visibleItems = items.filter(i => i.type !== 'FAILED_CHARGE');
          if (visibleItems.length) {
            visibleItems.forEach(item => {
              const label = TYPE_LABELS[item.type] || item.type;
              const nameStr = (item.type === 'PREMIUM_PRODUCT_UPCHARGE' && item.description)
                ? `${label} \u2014 ${item.description}` : label;
              const discPart = item.discount
                ? `<span style="color:var(--sb-success-soft);margin-right:4px;">-${fmtMoney(item.discount, currency)}</span>` : '';
              const row = document.createElement('div');
              row.style.cssText = 'display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px;font-size:var(--sb-fs-12);';
              row.innerHTML = `
                <span style="color:var(--sb-text-secondary);">${nameStr}</span>
                <span style="color:var(--sb-text-muted);white-space:nowrap;margin-left:8px;">${discPart}${fmtMoney(item.price, currency)} + ${fmtMoney(item.tax, currency)} tax = <strong>${fmtMoney(item.total, currency)}</strong></span>
              `;
              lineWrap.appendChild(row);
            });
          } else {
            const cat = c?.category || 'CHARGE';
            const fallback = document.createElement('div');
            fallback.style.cssText = 'color:var(--sb-text-dim);font-size:var(--sb-fs-11);font-style:italic;';
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
          expHdr.innerHTML = '<span style="font-size:var(--sb-fs-11);color:var(--sb-text-dim);font-weight:600;letter-spacing:0.05em;">BILL EXPLANATION</span>';
          expHdr.appendChild(makeCopyBtn(mergedText, '\uD83D\uDCCB Copy'));
          panel.appendChild(expHdr);

          const expBox = document.createElement('div');
          expBox.style.cssText = `
            background:var(--sb-bg-deep);border-radius:6px;padding:10px 12px;
            font-size:var(--sb-fs-12);color:var(--sb-text-secondary);white-space:pre-wrap;line-height:1.7;
            border:1px solid var(--sb-border-soft);
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
      loadingEl.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-12);';
      loadingEl.textContent = 'Loading queue...';
      panel.appendChild(loadingEl);

      fetchQueue(user.id, (blocks) => {
        loadingEl.remove();

        if (!blocks || !blocks.length) {
          const empty = document.createElement('div');
          empty.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-12);';
          empty.textContent = blocks ? 'Queue is empty.' : 'Failed to load queue.';
          panel.appendChild(empty);
          return;
        }

        // Filter to blocks with products only
        const filledBlocks = blocks.filter(b => b.productList?.length > 0);

        if (!filledBlocks.length) {
          const empty = document.createElement('div');
          empty.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-12);';
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
            sep.style.cssText = 'height:1px;background:var(--sb-border-soft);margin:8px 0;';
            panel.appendChild(sep);
          }

          const blockEl = document.createElement('div');
          blockEl.style.cssText = 'background:var(--sb-bg-deep);border-radius:8px;padding:10px 12px;';

          // Month header
          const [yearStr, monthStr] = (block.yearMonth || '').split('-');
          const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
          const monthLabel = monthNames[parseInt(monthStr, 10) - 1] + ' ' + yearStr;

          const hdr = document.createElement('div');
          hdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;';

          const monthEl = document.createElement('span');
          monthEl.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-11);font-weight:600;';
          monthEl.textContent = monthLabel;
          hdr.appendChild(monthEl);

          // Order status if processed
          const orderStatus = block.processedSubscriptionOrder?.status;
          if (orderStatus) {
            const statusColor = (orderStatus === 'SHIPPED' || orderStatus === 'DELIVERED') ? 'var(--sb-success-soft)'
              : orderStatus === 'PENDING' ? 'var(--sb-warning-strong)' : 'var(--sb-text-muted)';
            const statusEl = document.createElement('span');
            statusEl.style.cssText = `font-size:var(--sb-fs-10);font-weight:600;color:${statusColor};`;
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
            nameEl.style.cssText = 'color:var(--sb-text-secondary);font-size:var(--sb-fs-12);flex:1;margin-right:6px;display:flex;align-items:center;gap:6px;flex-wrap:wrap;';
            const nameText = document.createElement('span');
            nameText.textContent = prodName;
            nameEl.appendChild(nameText);
            const plaque = makeProductStatusPlaque(prod.tradingItem?.status);
            if (plaque) nameEl.appendChild(plaque);
            row.appendChild(nameEl);

            // Upcharge indicator
            const upcharge = prod.upchargePrice || pi.upchargePrice;
            if (upcharge && upcharge > 0) {
              const upEl = document.createElement('span');
              upEl.style.cssText = 'color:var(--sb-warning-strong);font-size:var(--sb-fs-10);font-weight:600;white-space:nowrap;margin-right:4px;';
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
              delBtn.style.cssText = 'background:none;border:none;cursor:pointer;color:var(--sb-text-dim);font-size:var(--sb-fs-13);padding:0 3px;flex-shrink:0;';
              delBtn.onmouseenter = () => delBtn.style.color = 'var(--sb-danger)';
              delBtn.onmouseleave = () => delBtn.style.color = 'var(--sb-text-dim)';
              delBtn.onclick = () => {
                delBtn.disabled = true;
                delBtn.textContent = '⏳';
                delBtn.style.color = 'var(--sb-text-muted)';

                const doDelete = () => {
                  customerApiDelete(prodFlatIdx, (data, err) => {
                    if (err) {
                      delBtn.textContent = '✕';
                      delBtn.style.color = 'var(--sb-danger)';
                      delBtn.disabled = false;
                      return;
                    }
                    mutPostComment(user.id, `Queue item removed: ${prodName} (${monthLabel})`, () => {});
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
                      delBtn.style.color = 'var(--sb-danger)';
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
          copyAllBtn.style.cssText = 'background:none;border:1px solid var(--sb-border);border-radius:5px;color:var(--sb-accent-strong);font-size:var(--sb-fs-11);padding:4px 10px;cursor:pointer;';
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

  /** Shipment month/year selects. A year either side of now is the whole useful
   *  range — a custom shipment is always near the present. */
  function _csMonthPicker(form) {
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

    return { monthSel, yearSel };
  }

  /** The products list. Starts empty by design — a custom shipment has no source
   *  order to seed from, so everything arrives through Search Product. */
  function _csProductRows(form) {
    // Products
    form.appendChild(makeLabel('PRODUCTS'));
    const itemsWrap = document.createElement('div');
    itemsWrap.style.cssText = 'background:var(--sb-bg-deep);border-radius:6px;padding:8px 10px;min-height:30px;';
    const itemChecks = [];

    const noItems = document.createElement('div');
    noItems.style.cssText = 'color:var(--sb-text-dim);font-size:var(--sb-fs-11);font-style:italic;';
    noItems.textContent = 'No products added. Use Search Product below.';
    itemsWrap.appendChild(noItems);
    form.appendChild(itemsWrap);
    return { itemsWrap, itemChecks, noItems };
  }

  function showCustomShipmentForm(user) {
    const existingForm = document.getElementById('sb-cs-form');
    if (existingForm) { existingForm.remove(); document.getElementById('sb-product-search-panel')?.remove(); return; }

    const form = createPanel({
      id: 'sb-cs-form',
      title: '📦 Custom Shipment',
      width: 400,
      maxHeight: 720,
      centered: true,
      draggable: true,
      onClose: () => document.getElementById('sb-product-search-panel')?.remove(),
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
    const { monthSel, yearSel } = _csMonthPicker(form);

    // Products
    const { itemsWrap, itemChecks, noItems } = _csProductRows(form);

    // Product search button
    const searchRow = document.createElement('div');
    searchRow.style.cssText = 'margin-top:8px;';
    const searchProductBtn = document.createElement('button');
    searchProductBtn.textContent = '🔍 Search Product';
    searchProductBtn.style.cssText = 'padding:5px 10px;border-radius:5px;border:1px solid var(--sb-accent-strong);background:transparent;color:var(--sb-accent-strong);font-weight:600;font-size:var(--sb-fs-11);cursor:pointer;';
    searchProductBtn.onmouseenter = () => searchProductBtn.style.background = 'rgba(129,140,248,0.15)';
    searchProductBtn.onmouseleave = () => searchProductBtn.style.background = 'transparent';
    searchProductBtn.onclick = () => {
      if (noItems.parentNode) noItems.remove();
      showProductSearchPanel(itemChecks, itemsWrap, () => csUpchargeUI.refresh(), user.id);
    };
    searchRow.appendChild(searchProductBtn);
    form.appendChild(searchRow);

    // Additional comment
    form.appendChild(makeLabel('ADDITIONAL COMMENT (OPTIONAL)'));
    const commentInput = document.createElement('textarea');
    commentInput.placeholder = 'Add a note...';
    commentInput.rows = 2;
    commentInput.style.cssText = 'width:100%;background:var(--sb-bg-deep);color:var(--sb-text);border:1px solid var(--sb-border);border-radius:5px;padding:6px 8px;font-size:var(--sb-fs-12);box-sizing:border-box;resize:vertical;font-family:Arial,sans-serif;';
    form.appendChild(commentInput);

    // Credit info (loaded async)
    const creditInfo = document.createElement('div');
    creditInfo.style.cssText = 'margin-top:12px;font-size:var(--sb-fs-11);color:var(--sb-text-muted);';
    creditInfo.textContent = 'Loading credit info...';
    form.appendChild(creditInfo);

    let _subscriptionId = null;
    let _availableCredits = 0;

    // Fetch subscription for credits
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
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
            creditInfo.style.color = _availableCredits > 0 ? 'var(--sb-success-soft)' : 'var(--sb-warning-strong)';
          } else {
            creditInfo.textContent = 'No subscription found';
            creditInfo.style.color = 'var(--sb-danger)';
          }
        } catch(e) {
          creditInfo.textContent = 'Failed to load credits';
          creditInfo.style.color = 'var(--sb-danger)';
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
      border:none;background:var(--sb-solid-primary);color:var(--sb-on-solid);font-weight:700;
      font-size:var(--sb-fs-13);cursor:pointer;box-sizing:border-box;
    `;
    confirmBtn.onmouseenter = () => confirmBtn.style.background = 'var(--sb-solid-primary-hi)';
    confirmBtn.onmouseleave = () => confirmBtn.style.background = 'var(--sb-solid-primary)';

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
        statusEl.style.color = 'var(--sb-danger)';
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
      detailsEl.style.cssText = 'background:var(--sb-bg-deep);border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:var(--sb-fs-12);line-height:1.8;color:var(--sb-text-secondary);';

      let detailsHtml = `
        <div><span style="color:var(--sb-text-dim);">Month</span>&nbsp;&nbsp;&nbsp;&nbsp; ${monthNames[csMonth - 1]} ${csYear}</div>
        <div><span style="color:var(--sb-text-dim);">Reason</span>&nbsp;&nbsp;&nbsp; ${reasonLabel}</div>
        <div><span style="color:var(--sb-text-dim);">Products</span>&nbsp; ${selectedItems.length}</div>
      `;
      if (creditsToAdd > 0) {
        detailsHtml += `<div style="margin-top:6px;padding-top:6px;border-top:1px solid var(--sb-control-bg);color:var(--sb-warning-strong);font-weight:600;">${creditsToAdd} credit(s) will be added</div>`;
      }
      detailsEl.innerHTML = detailsHtml;
      confBox.appendChild(detailsEl);

      // Product list
      const prodList = document.createElement('div');
      prodList.style.cssText = 'margin-bottom:12px;font-size:var(--sb-fs-11);color:var(--sb-text-muted);';
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
        confirmColor: 'var(--sb-solid-primary)',
        onCancel: () => confOverlay.remove(),
        onConfirm: (cBtn, cancelBtn) => {
          cBtn.disabled = true;
          cancelBtn.disabled = true;
          cBtn.textContent = 'Processing...';

          const doShipment = () => {
            confStatus.textContent = 'Creating custom shipment...';

            const variables = {
              input: {
                kustomer: window.location.href,
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
                  confStatus.style.color = 'var(--sb-danger)';
                  confStatus.textContent = '\u2718 ' + errMsg;
                  cBtn.disabled = false;
                  cancelBtn.disabled = false;
                  cBtn.textContent = 'Confirm';
                  return;
                }

                confStatus.style.color = 'var(--sb-text-muted)';
                confStatus.textContent = 'Finding automated comment...';

                const csOrderNumber = data?.customShipmentSave?.data?.orderNumber;
                const monthLabel = monthNames[csMonth - 1] + ' ' + csYear;
                const commentParts = ['Custom Shipment - ' + reasonLabel + ' - ' + monthLabel];
                if (creditsToAdd > 0) commentParts.push(creditsToAdd + ' Credit(s) Added for CS to ship');
                const extra = commentInput.value.trim();
                if (extra) commentParts.push(extra);
                const commentText = commentParts.join('\n');

                const postComment = () => {
                  confStatus.textContent = 'Posting comment...';
                  gqlMutate('createUserComment', CREATE_COMMENT_MUTATION,
                    { input: { userId: user.id, comment: commentText, kustomerUrl: window.location.href } },
                    () => {
                      confStatus.style.color = 'var(--sb-success-soft)';
                      confStatus.textContent = '\u2714 Custom shipment created!';
                      cBtn.textContent = '\u2714 Done';
                      confirmBtn.disabled = true;
                      confirmBtn.textContent = '\u2714 Created';
                      confirmBtn.style.background = 'rgba(110,231,183,0.15)';
                      confirmBtn.style.color = 'var(--sb-success-soft)';
                      refreshOpenPanels('sb-order-panel');
                      setTimeout(() => {
                        confOverlay.remove();
                        form.remove();
                        document.getElementById('sb-product-search-panel')?.remove();
                      }, 1800);
                    }
                  );
                };

                const pollForAutoComment = (attemptsLeft) => {
                  fetchComments(user.id, (comments) => {
                    const autoComment = comments
                      .filter(c => c.comment && c.comment.toLowerCase().includes('created a custom shipment'))
                      .sort((a, b) => (b.created || 0) - (a.created || 0))[0] || null;
                    if (autoComment) {
                      confStatus.textContent = 'Removing automated comment...';
                      gqlMutate('deleteUserComment', DELETE_COMMENT_MUTATION, { id: autoComment.id },
                        () => postComment()
                      );
                    } else if (attemptsLeft > 1) {
                      confStatus.textContent = 'Finding automated comment...';
                      setTimeout(() => pollForAutoComment(attemptsLeft - 1), 1500);
                    } else {
                      postComment();
                    }
                  });
                };

                pollForAutoComment(4);
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
                  confStatus.style.color = 'var(--sb-danger)';
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

  /** Email row + Update. Posts the change, then re-resolves the customer. */
  function _editEmailSection(panel, user) {
  // Email section
  panel.appendChild(makeLabel('EMAIL'));
  const emailRow = document.createElement('div');
  emailRow.style.cssText = 'display:flex;gap:6px;margin-bottom:12px;';

  const emailInput = document.createElement('input');
  emailInput.type = 'email';
  emailInput.value = user.email || '';
  emailInput.style.cssText = `
    flex:1;padding:7px 10px;border-radius:6px;border:1px solid var(--sb-hover);
    background:var(--sb-bg-deep);color:var(--sb-text);font-size:var(--sb-fs-12);outline:none;box-sizing:border-box;
  `;

  const updateBtn = document.createElement('button');
  updateBtn.textContent = 'Update';
  updateBtn.style.cssText = 'padding:7px 14px;border-radius:6px;border:none;background:var(--sb-solid-primary);color:var(--sb-on-solid);font-weight:700;font-size:var(--sb-fs-12);cursor:pointer;';
  updateBtn.onmouseenter = () => updateBtn.style.background = 'var(--sb-solid-primary-hi)';
  updateBtn.onmouseleave = () => updateBtn.style.background = 'var(--sb-solid-primary)';

  const emailStatus = makeStatusEl();

  updateBtn.onclick = () => {
    const newEmail = emailInput.value.trim();
    if (!newEmail || !newEmail.includes('@')) {
      emailStatus.style.color = 'var(--sb-danger)';
      emailStatus.textContent = 'Enter a valid email.';
      return;
    }
    if (newEmail === user.email) {
      emailStatus.style.color = 'var(--sb-warning-strong)';
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
          emailStatus.style.color = 'var(--sb-danger)';
          emailStatus.textContent = '\u2718 ' + errMsg;
          updateBtn.disabled = false;
          updateBtn.textContent = 'Update';
          return;
        }
        emailStatus.style.color = 'var(--sb-success-soft)';
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

  }

  /** Gender select — drives which catalogue the customer is shown. */
  function _editGenderSection(panel, user) {
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
  genderBtn.style.cssText = 'padding:7px 14px;border-radius:6px;border:none;background:var(--sb-solid-primary);color:var(--sb-on-solid);font-weight:700;font-size:var(--sb-fs-12);cursor:pointer;';
  genderBtn.onmouseenter = () => genderBtn.style.background = 'var(--sb-solid-primary-hi)';
  genderBtn.onmouseleave = () => genderBtn.style.background = 'var(--sb-solid-primary)';

  const genderStatus = makeStatusEl();

  genderBtn.onclick = () => {
    const newGender = genderSel.value;
    if (newGender === currentGender) {
      genderStatus.style.color = 'var(--sb-warning-strong)';
      genderStatus.textContent = 'Gender is unchanged.';
      return;
    }
    genderBtn.disabled = true;
    genderBtn.textContent = '⏳';
    genderStatus.textContent = 'Authenticating...';

    ensureCustomerAuth((ok) => {
      if (!ok) {
        genderStatus.style.color = 'var(--sb-danger)';
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
            genderStatus.style.color = 'var(--sb-danger)';
            genderStatus.textContent = '\u2718 ' + errMsg;
            genderBtn.disabled = false;
            genderBtn.textContent = 'Update';
            return;
          }
          genderStatus.style.color = 'var(--sb-success-soft)';
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

  }

  /** Shipping address: the current value, then an autocomplete-backed editor
   *  that stays hidden until the customer API has authenticated. */
  function _editAddressSection(panel, user) {
  // ── Address section ─────────────────────────────────────────────────
  panel.appendChild(makeLabel('SHIPPING ADDRESS'));

  // Current address display
  const shipping = user.userAddress?.shipping;
  if (shipping) {
    const addrDisplay = document.createElement('div');
    addrDisplay.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-muted);margin-bottom:8px;line-height:1.5;';
    addrDisplay.textContent = [shipping.street1, shipping.street2, shipping.city, shipping.region, shipping.postcode].filter(Boolean).join(', ');
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
  updateAddrBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid var(--sb-solid-primary);background:transparent;
    color:var(--sb-accent);font-weight:600;font-size:var(--sb-fs-12);cursor:pointer;box-sizing:border-box;`;
  updateAddrBtn.onmouseenter = () => updateAddrBtn.style.background = 'rgba(79,70,229,0.1)';
  updateAddrBtn.onmouseleave = () => updateAddrBtn.style.background = 'transparent';
  updateAddrBtn.onclick = () => {
    updateAddrBtn.disabled = true;
    updateAddrBtn.textContent = '⏳ Authenticating...';
    addrAuthStatus.textContent = '';

    ensureCustomerAuth((ok) => {
      if (!ok) {
        addrAuthStatus.style.color = 'var(--sb-danger)';
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
    width:100%;padding:7px 10px;border-radius:6px;border:1px solid var(--sb-hover);
    background:var(--sb-bg-deep);color:var(--sb-text);font-size:var(--sb-fs-12);outline:none;box-sizing:border-box;
  `;
  addrEditWrap.appendChild(addrInput);

  const addrResults = document.createElement('div');
  addrResults.style.cssText = 'max-height:200px;overflow-y:auto;';
  addrEditWrap.appendChild(addrResults);

  const selectedAddr = document.createElement('div');
  selectedAddr.style.cssText = 'display:none;background:var(--sb-bg-deep);border-radius:6px;padding:8px 10px;margin-top:8px;font-size:var(--sb-fs-11);color:var(--sb-text-secondary);line-height:1.6;';
  addrEditWrap.appendChild(selectedAddr);

  const addrStatus = makeStatusEl();
  addrEditWrap.appendChild(addrStatus);

  const addrSaveBtn = document.createElement('button');
  addrSaveBtn.textContent = 'Save Address';
  addrSaveBtn.style.cssText = 'display:none;width:100%;padding:9px;border-radius:6px;border:none;background:var(--sb-solid-primary);color:var(--sb-on-solid);font-weight:700;font-size:var(--sb-fs-12);cursor:pointer;box-sizing:border-box;margin-top:8px;';
  addrSaveBtn.onmouseenter = () => addrSaveBtn.style.background = 'var(--sb-solid-primary-hi)';
  addrSaveBtn.onmouseleave = () => addrSaveBtn.style.background = 'var(--sb-solid-primary)';
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
            row.style.cssText = 'padding:6px 10px;cursor:pointer;font-size:var(--sb-fs-11);color:var(--sb-text-secondary);border-bottom:1px solid var(--sb-border-soft);';
            row.onmouseenter = () => row.style.background = 'rgba(99,102,241,0.15)';
            row.onmouseleave = () => row.style.background = '';

            const main = document.createElement('div');
            main.textContent = suggestion.mainText;
            main.style.fontWeight = '600';
            row.appendChild(main);

            if (suggestion.secondaryText) {
              const sec = document.createElement('div');
              sec.style.cssText = 'color:var(--sb-text-dim);font-size:var(--sb-fs-10);';
              sec.textContent = suggestion.secondaryText;
              row.appendChild(sec);
            }

            row.onclick = () => {
              if (suggestion.type === 'CONTAINER') {
                addrInput.value = suggestion.mainText;
                addrInput.dispatchEvent(new Event('input', { bubbles: true }));
                return;
              }

              addrResults.innerHTML = '<div style="color:var(--sb-text-muted);font-size:var(--sb-fs-11);padding:6px 10px;">Loading details...</div>';
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
                    addrStatus.style.color = 'var(--sb-danger)';
                    addrStatus.textContent = '\u2718 Failed to load address details';
                    return;
                  }
                  const addr = detData.addressAutocompleteDetails.data;
                  _selectedAddrData = addr;
                  addrInput.value = '';
                  selectedAddr.style.display = '';
                  selectedAddr.innerHTML = `
                    <div style="font-weight:600;color:var(--sb-text);">${addr.street1 || ''}${addr.street2 ? ', ' + addr.street2 : ''}</div>
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
          addrStatus.style.color = 'var(--sb-danger)';
          addrStatus.textContent = '\u2718 ' + valErrMsg;
          addrSaveBtn.disabled = false;
          addrSaveBtn.textContent = 'Save Address';
          return;
        }

        const validatedAddr = valResult?.data?.address;
        if (!validatedAddr) {
          addrStatus.style.color = 'var(--sb-danger)';
          addrStatus.textContent = '\u2718 Validation returned no address data';
          addrSaveBtn.disabled = false;
          addrSaveBtn.textContent = 'Save Address';
          return;
        }

        const addrId = validatedAddr.id || cachedCustomerCtx?._shippingAddrId;
        if (!addrId) {
          addrStatus.style.color = 'var(--sb-danger)';
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
              addrStatus.style.color = 'var(--sb-danger)';
              addrStatus.textContent = '\u2718 ' + updErrMsg;
              addrSaveBtn.disabled = false;
              addrSaveBtn.textContent = 'Save Address';
              return;
            }
            addrStatus.style.color = 'var(--sb-success-soft)';
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
  }

  /** Edit Customer panel. Three independent sections, each owning its own inputs,
   *  status line and mutation — they never talked to each other, which is what made
   *  the 439-line original separable without changing any behaviour. */
  function showEditCustomerPanel(btn) {
    if (document.getElementById('sb-edit-customer-panel')) {
      removePanel('sb-edit-customer-panel'); return;
    }

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;

      const panel = createPanel({ id: 'sb-edit-customer-panel', title: '✎ Edit Customer', width: 340 });
      panel.appendChild(renderUserRow(user));

      _editEmailSection(panel, user);
      _editGenderSection(panel, user);
      _editAddressSection(panel, user);
    });
  }

  // ── Downgrade Panel ───────────────────────────────────────────────────────
  // Plan: CRM's own "Change plan" dialog asks subscriptionAvailablePlansInquire for the
  // plans reachable from the current one, keyed by planTransitionType (UPGRADE /
  // DOWNGRADE / FREQUENCY). We ask for DOWNGRADE, so the list is whatever CRM considers
  // lower for this customer — no plan ladder hardcoded here. The change itself is the
  // dialog's subscriptionChangePlan, skipCharge:true like the UI sends.

  const AVAILABLE_PLANS_QUERY = `
    query subscriptionAvailablePlansInquire($input: SubscriptionAvailablePlansInput) {
      subscriptionAvailablePlansInquire(input: $input) { data { code planName } error { message } }
    }`;
  const CHANGE_PLAN_MUTATION = `
    mutation subscriptionChangePlan($input: SubscriptionChangePlanInput) {
      subscriptionChangePlan(input: $input) { data { id planName } error { message } }
    }`;

  function showDowngradePanel(btn) {
    if (document.getElementById('sb-downgrade-panel')) { removePanel('sb-downgrade-panel'); return; }
    loadCustomer(btn, ({ user }) => {
      const panel = createPanel({ id: 'sb-downgrade-panel', title: '⬇ Downgrade', width: 360 });
      panel.appendChild(renderUserRow(user));
      const status = makeStatusEl();
      status.textContent = 'Loading subscription…';
      panel.appendChild(status);

      fetchUserDetails(user.id, (details, err) => {
        const sub = (details?.subscriptionList || []).find(s => s.status === 'Active');
        if (!sub) {
          status.style.color = 'var(--sb-danger)';
          status.textContent = '✘ ' + (err || 'No active subscription');
          return;
        }
        status.textContent = '';
        _downgradeAddOnSection(panel, user, sub);
        _downgradePlanSection(panel, user, sub);
      });
    });
  }

  // Add-ons: CRM has one mutation per add-on, all named after the addOnSettings key
  // (samplesSubscription → samplesSubscriptionUpdate / SamplesSubscriptionUpdateInput)
  // and all taking { subscriptionId, enabled }. "On" means `selected`, same as the info bar.
  const ADDON_LABELS = {
    caseSubscription:         'Case',
    samplesSubscription:      'Samples',
    candleSubscription:       'Candle',
    carScentSubscription:     'Car Scent',
    homeDiffuserSubscription: 'Diffuser',
  };

  async function setAddOn(subscriptionId, key, on) {
    const op = key + 'Update';
    const type = key[0].toUpperCase() + key.slice(1) + 'UpdateInput';
    const data = await crmQueryP(op,
      `mutation ${op}($input: ${type}!) { ${op}(input: $input) { data { id addOnSettings { ${key} { selected } } } error { message } } }`,
      { input: { subscriptionId, enabled: on } });
    const r = data[op];
    if (r?.error?.message) throw new Error(r.error.message);
    if (!!r?.data?.addOnSettings?.[key]?.selected !== on) throw new Error(`still ${on ? 'off' : 'on'} after update`);
  }

  /** One switch per add-on the subscription can have (`enabled` = offered, `selected`
   *  = on). Each flip fires its own mutation straight away — it's a reversible
   *  preference, so no arm-to-confirm — and leaves a CRM comment. */
  function _downgradeAddOnSection(panel, user, sub) {
    panel.appendChild(makeLabel('ADD-ONS'));
    const keys = Object.keys(ADDON_LABELS).filter(k => sub.addOnSettings?.[k]?.enabled || sub.addOnSettings?.[k]?.selected);
    if (!keys.length) {
      panel.appendChild(el('div', { style: 'font-size:var(--sb-fs-12);color:var(--sb-text-dim);margin-bottom:8px;', text: 'No add-ons available.' }));
      return;
    }
    const row = el('div', { cls: 'sb-row', style: 'gap:6px;flex-wrap:wrap;margin-bottom:4px;' });
    panel.appendChild(row);
    const status = makeStatusEl();
    panel.appendChild(status);

    const paint = (b, on) => {
      b.setAttribute('aria-checked', String(on));
      b.textContent = (on ? '● ' : '○ ') + ADDON_LABELS[b.dataset.addon];
      b.style.cssText = on
        ? 'padding:4px 10px;border-radius:20px;cursor:pointer;font-weight:600;font-size:var(--sb-fs-11);color:var(--sb-success-soft);background:var(--sb-tag-success-bg);border:1px solid var(--sb-tag-success-border);'
        : 'padding:4px 10px;border-radius:20px;cursor:pointer;font-weight:600;font-size:var(--sb-fs-11);color:var(--sb-text-dim);background:transparent;border:1px solid var(--sb-control-bg);';
    };

    keys.forEach(k => {
      const b = el('button', { data: { addon: k }, attrs: { role: 'switch' }, title: 'Turn ' + ADDON_LABELS[k] + ' on/off' });
      paint(b, !!sub.addOnSettings[k].selected);
      b.addEventListener('click', async () => {
        const on = !sub.addOnSettings[k].selected;
        b.disabled = true; b.style.opacity = '0.6';
        status.style.color = ''; status.textContent = `Turning ${ADDON_LABELS[k]} ${on ? 'on' : 'off'}…`;
        try {
          await setAddOn(sub.id, k, on);
          sub.addOnSettings[k].selected = on;
          paint(b, on);
          status.style.color = 'var(--sb-success-soft)';
          status.textContent = `✔ ${ADDON_LABELS[k]} ${on ? 'on' : 'off'}`;
          mutPostComment(user.id, `Add-on ${ADDON_LABELS[k]} turned ${on ? 'on' : 'off'}`, () => {});
          refreshSubscriptionBar();
        } catch (e) {
          status.style.color = 'var(--sb-danger)';
          status.textContent = `✘ ${ADDON_LABELS[k]}: ${e.message}`;
        }
        b.disabled = false; b.style.opacity = '';
      });
      row.appendChild(b);
    });
  }

  function _downgradePlanSection(panel, user, sub) {
    panel.appendChild(makeLabel('PLAN'));
    const cur = el('div', { style: 'font-size:var(--sb-fs-12);color:var(--sb-text-muted);margin-bottom:6px;',
      text: 'Current: ' + (PLAN_LABELS[sub.planName] || sub.planName || '—') });
    panel.appendChild(cur);
    const row = el('div', { cls: 'sb-row', style: 'gap:6px;' });
    const sel = el('select', { cls: 'sb-select', style: 'flex:1;min-width:0;' });
    sel.disabled = true;
    sel.appendChild(el('option', { text: 'Loading plans…' }));
    const go = el('button', { cls: 'sb-btn sb-btn--danger' });
    go.disabled = true;
    row.appendChild(sel); row.appendChild(go);
    panel.appendChild(row);
    const status = makeStatusEl();
    panel.appendChild(status);
    const fail = (msg) => { status.style.color = 'var(--sb-danger)'; status.textContent = '✘ ' + msg; };

    const cashbirdId = sub.cashbirdDetails?.data?.cashbirdId;
    if (!cashbirdId) { sel.innerHTML = ''; return fail('No billing id on this subscription'); }

    crmQueryP('subscriptionAvailablePlansInquire', AVAILABLE_PLANS_QUERY,
      { input: { userId: user.id, cashbirdSubscriptionId: cashbirdId, planTransitionType: 'DOWNGRADE' } })
      .then((data) => {
        const r = data.subscriptionAvailablePlansInquire;
        if (r?.error?.message) throw new Error(r.error.message);
        const plans = r?.data || [];
        sel.innerHTML = '';
        if (!plans.length) { sel.appendChild(el('option', { text: 'Already on the lowest plan' })); return; }
        plans.forEach(p => sel.appendChild(el('option', { text: p.planName.replace(/ subscription$/i, ''), attrs: { value: p.code } })));
        sel.disabled = false; go.disabled = false;
      })
      .catch(e => { sel.innerHTML = ''; fail(e.message); });

    armedButton(go, {
      idle: 'Downgrade',
      armed: '⚠ Confirm',
      onFire: async () => {
        const code = sel.value;
        const label = sel.selectedOptions[0]?.textContent || code;
        go.disabled = true; sel.disabled = true;
        status.style.color = ''; status.textContent = 'Changing plan…';
        try {
          const data = await crmQueryP('subscriptionChangePlan', CHANGE_PLAN_MUTATION,
            { input: { subscriptionId: sub.id, planCode: code, skipCharge: true, paymentMethodId: null } });
          const r = data.subscriptionChangePlan;
          if (r?.error?.message || !r?.data) throw new Error(r?.error?.message || 'No result');
          status.style.color = 'var(--sb-success-soft)';
          status.textContent = '✔ Plan changed to ' + label;
          cur.textContent = 'Current: ' + (PLAN_LABELS[r.data.planName] || r.data.planName);
          const from = PLAN_LABELS[sub.planName] || sub.planName;
          mutPostComment(user.id, `Plan downgraded: ${from} → ${label}`, () => {});
          refreshSubscriptionBar();
        } catch (e) {
          fail(e.message);
          go.disabled = false; sel.disabled = false;
        }
      },
    });
  }

  // ── Cancel Subscription Panel ─────────────────────────────────────────────

  /** Name + email header line. */
  function _cancelInfoSection(C) {
    // Customer info
    const name = [C.user.firstName, C.user.lastName].filter(Boolean).join(' ') || '\u2014';
    const infoEl = document.createElement('div');
    infoEl.style.cssText = 'font-size:var(--sb-fs-12);color:var(--sb-text-muted);margin-bottom:10px;line-height:1.6;';
    infoEl.innerHTML = `<span style="color:var(--sb-text);font-weight:600;">${name}</span> &middot; ${C.user.email}`;
    C.panel.appendChild(infoEl);

  }

  /** Last subscription order, with a Refund button when the order still
   *  qualifies. Refundability is the same NON_REFUNDABLE_STATUSES test used
   *  everywhere else. */
  function _cancelOrderSection(C) {
    // Last subscription order summary
    const lastOrder = C.orders?.find(o => o.type === 'SUBSCRIPTION');
    if (lastOrder) {
      const orderEl = document.createElement('div');
      orderEl.style.cssText = 'background:var(--sb-bg-row);border-radius:6px;padding:8px 10px;margin-bottom:12px;border-left:2px solid var(--sb-solid-primary);';
      const monthLbl = orderMonthLabel(lastOrder);
      const ws = lastOrder.warehouseOrder?.data?.status || lastOrder.status || '';
      const wsUC = ws.toUpperCase();
      const statusColor = wsUC === 'DONE' ? 'var(--sb-success-soft)' : 'var(--sb-danger)';
      const lines = orderProductLines(lastOrder);

      const orderTop = document.createElement('div');
      orderTop.style.cssText = 'display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px;';

      const orderInfo = document.createElement('div');
      orderInfo.innerHTML = `<div style="display:flex;gap:6px;align-items:center;margin-bottom:3px;">
        <span style="color:var(--sb-text-muted);font-size:var(--sb-fs-11);">${monthLbl}</span>
        <span style="color:${statusColor};font-size:var(--sb-fs-11);font-weight:600;">${ws}</span>
      </div>` + lines.map(l => `<div style="color:var(--sb-text-secondary);font-size:var(--sb-fs-11);">${l}</div>`).join('');

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
          a.style.cssText = 'color:var(--sb-accent-strong);font-size:var(--sb-fs-11);text-decoration:none;';
          tRow.appendChild(a);
          const copyBtn = document.createElement('button');
          copyBtn.textContent = '\uD83D\uDD17';
          copyBtn.title = 'Copy tracking link';
          copyBtn.style.cssText = 'background:none;border:none;cursor:pointer;color:var(--sb-accent-strong);font-size:var(--sb-fs-12);padding:0 2px;';
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
          sp.style.cssText = 'color:var(--sb-accent-strong);font-size:var(--sb-fs-11);';
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
          d.style.cssText = 'color:var(--sb-text-dim);font-size:var(--sb-fs-11);margin-top:1px;';
          const dt = new Date(item.date).toLocaleDateString('en-US', { month:'short', day:'numeric' });
          d.textContent = dt + ' · ' + item.description;
          evEl.appendChild(d);
        });
        orderInfo.appendChild(evEl);
      }
      orderTop.appendChild(orderInfo);

      // Refund button
      // Always offered on the last order — no fulfilment-status gate (2026-10-13).
      const refundable = true;
      const orderRefBtn = document.createElement('button');
      orderRefBtn.textContent = '💰 Refund';
      orderRefBtn.title = 'Open refund charges';
      orderRefBtn.style.cssText = `
        margin-left:8px;padding:3px 8px;border-radius:5px;border:none;font-size:var(--sb-fs-11);font-weight:600;
        white-space:nowrap;flex-shrink:0;
        cursor:${refundable ? 'pointer' : 'not-allowed'};
        background:${refundable ? 'var(--sb-solid-violet)' : 'var(--sb-solid-violet-dim)'};
        color:${refundable ? 'var(--sb-on-solid)' : 'var(--sb-text-dim)'};
      `;
      if (refundable) {
        orderRefBtn.onmouseenter = () => orderRefBtn.style.background = 'var(--sb-solid-violet-hi)';
        orderRefBtn.onmouseleave = () => orderRefBtn.style.background = 'var(--sb-solid-violet)';
        orderRefBtn.onclick = () => showRefundChargesPanel(C.user);
      }
      orderTop.appendChild(orderRefBtn);
      orderEl.appendChild(orderTop);
      C.panel.appendChild(orderEl);
    }

  }

  /** The cancel-subscription control itself, plus the status line every other
   *  section writes into. Disabled — not hidden — when the sub is already
   *  cancelled, so the panel still says why. */
  function _cancelCancelSection(C) {
    const statusEl = C.statusEl = makeStatusEl();
    C.panel.appendChild(statusEl);

    if (C.isAlreadyCancelled) {
      const subStatusEl = document.createElement('div');
      subStatusEl.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-warning-strong);margin-bottom:10px;';
      subStatusEl.textContent = C.isAwaitingCancel
        ? '⚠ Subscription is already scheduled for cancellation'
        : '⚠ Subscription status: ' + C.subStatus;
      C.panel.appendChild(subStatusEl);
    }

    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'Cancel Subscription';
    confirmBtn.disabled = C.isAlreadyCancelled;
    confirmBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:none;font-weight:700;font-size:var(--sb-fs-13);box-sizing:border-box;
      background:${C.isAlreadyCancelled ? 'var(--sb-solid-disabled)' : 'var(--sb-danger-border)'};
      color:${C.isAlreadyCancelled ? 'var(--sb-text-dim)' : 'var(--sb-on-solid)'};
      cursor:${C.isAlreadyCancelled ? 'not-allowed' : 'pointer'};`;
    if (!C.isAlreadyCancelled) {
      confirmBtn.onmouseenter = () => confirmBtn.style.background = 'var(--sb-solid-danger-hi)';
      confirmBtn.onmouseleave = () => confirmBtn.style.background = 'var(--sb-danger-border)';
    }

    confirmBtn.onclick = async () => {
      confirmBtn.disabled = true;
      statusEl.style.color = 'var(--sb-text-muted)';
      statusEl.textContent = '';
      try {
        await executeCancelSubscription(C.user, (msg) => {
          confirmBtn.textContent = msg;
          statusEl.textContent = msg;
        });
        statusEl.style.color = 'var(--sb-success-soft)';
        statusEl.textContent = '\u2714 Subscription cancelled.';
        confirmBtn.textContent = '\u2714 Done';
      } catch (e) {
        statusEl.style.color = 'var(--sb-danger)';
        statusEl.textContent = '\u2718 ' + (e.message || e);
        confirmBtn.disabled = false;
        confirmBtn.textContent = 'Confirm Cancellation';
      }
    };

    C.panel.appendChild(confirmBtn);

  }

  /** "Unsubscribe from Emails" — red outline to match Delete Payment Methods, which it
   *  sits beside. Confirm-on-second-click. Returns { btn, status }. */
  function _cancelUnsubButton(C) {
    const idleCss = `flex:1;padding:9px;border-radius:6px;border:1px solid var(--sb-danger-border);background:transparent;
      color:var(--sb-danger);font-weight:600;font-size:var(--sb-fs-12);cursor:pointer;box-sizing:border-box;`;
    const btn = el('button', { style: idleCss });
    btn.onmouseenter = () => { if (!btn.disabled) btn.style.background = 'rgba(220,38,38,0.1)'; };
    btn.onmouseleave = () => { btn.style.background = 'transparent'; };
    const status = makeStatusEl();
    armedButton(btn, {
      idle: '✉ Unsubscribe from Emails',
      armed: '⚠ Click again',
      onFire: async () => {
        btn.disabled = true; btn.textContent = '⏳ Unsubscribing…';
        try {
          await crmUnsubscribeEmail(C.user.email);
          btn.textContent = '✔ Unsubscribed';
          btn.style.color = 'var(--sb-success-soft)';
          btn.style.borderColor = 'var(--sb-success-soft)';
          btn.style.cursor = 'default';
          status.textContent = '';
          mutPostComment(C.user.id, 'Unsubscribed from emails', () => {});
        } catch (e) {
          btn.disabled = false; btn.textContent = '✉ Unsubscribe from Emails';
          status.style.color = 'var(--sb-danger)';
          status.textContent = '✘ ' + (e.message || e);
        }
      },
    });
    return { btn, status };
  }

  /** Delete stored payment methods. */
  function _cancelPmSection(C) {
    // ── Delete Payment Methods ────────────────────────────────────────
    const pmSep = document.createElement('div');
    pmSep.style.cssText = 'height:1px;background:var(--sb-control-bg);margin:12px 0;';
    C.panel.appendChild(pmSep);

    const delPmBtn = document.createElement('button');
    delPmBtn.textContent = '🗑 Delete Payment Methods';
    delPmBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid var(--sb-danger-border);background:transparent;
      color:var(--sb-danger);font-weight:600;font-size:var(--sb-fs-12);cursor:pointer;box-sizing:border-box;`;
    delPmBtn.onmouseenter = () => delPmBtn.style.background = 'rgba(220,38,38,0.1)';
    delPmBtn.onmouseleave = () => delPmBtn.style.background = 'transparent';
    delPmBtn.onclick = () => {
      const { overlay, box } = createModal({ id: 'sb-del-pm-overlay', title: '🗑 Delete All Payment Methods', width: 320 });

      const warnEl = document.createElement('div');
      warnEl.style.cssText = 'background:rgba(220,38,38,0.1);border:1px solid rgba(220,38,38,0.3);border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:var(--sb-fs-12);color:var(--sb-danger);line-height:1.6;';
      warnEl.innerHTML = `This will permanently remove <strong>all stored payment methods</strong> for:<br><br>
        <span style="color:var(--sb-text);font-weight:600;">${name}</span><br>
        <span style="color:var(--sb-text-muted);">${C.user.email}</span><br><br>
        This action cannot be undone.`;
      box.appendChild(warnEl);

      const pmStatus = makeStatusEl();
      box.appendChild(pmStatus);

      box.appendChild(makeDialogButtons({
        confirmLabel: 'Delete All',
        confirmColor: 'var(--sb-danger-border)',
        onCancel: () => overlay.remove(),
        onConfirm: (cBtn) => {
          cBtn.disabled = true;
          cBtn.textContent = 'Deleting...';
          gqlMutate('deleteAllPaymentMethods', DELETE_PAYMENT_METHODS_MUTATION,
            { input: { userId: C.user.id } },
            (data, err) => {
              const respErr = data?.paymentMethodDeleteAll?.error;
              const errMsg = err || respErr?.message || respErr?.serverErrorMessage;
              if (errMsg) {
                pmStatus.style.color = 'var(--sb-danger)';
                pmStatus.textContent = '\u2718 ' + errMsg;
                cBtn.disabled = false;
                cBtn.textContent = 'Delete All';
                return;
              }
              pmStatus.style.color = 'var(--sb-success-soft)';
              pmStatus.textContent = '\u2714 All payment methods deleted.';
              cBtn.textContent = '\u2714 Done';
              delPmBtn.disabled = true;
              delPmBtn.textContent = '\u2714 Payment Methods Deleted';
              delPmBtn.style.color = 'var(--sb-success-soft)';
              delPmBtn.style.borderColor = 'var(--sb-success-soft)';
              delPmBtn.style.cursor = 'default';
              mutPostComment(C.user.id, 'Billing method removed', () => {});
              autoClose(() => overlay.remove());
            }
          );
        },
      }));
    };
    // Delete Payment Methods + Unsubscribe from Emails share one row.
    delPmBtn.style.width = 'auto'; delPmBtn.style.flex = '1';
    const unsub = _cancelUnsubButton(C);
    C.panel.appendChild(el('div', { style: 'display:flex;gap:6px;' }, [delPmBtn, unsub.btn]));
    C.panel.appendChild(unsub.status);

  }

  /** Delete the account outright. */
  function _cancelAccSection(C) {
    // ── Delete Account ────────────────────────────────────────────────
    const delAccBtn = document.createElement('button');
    delAccBtn.textContent = '⛔ Delete Account';
    delAccBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid var(--sb-danger-strong);background:transparent;
      color:var(--sb-danger);font-weight:600;font-size:var(--sb-fs-12);cursor:pointer;box-sizing:border-box;margin-top:6px;`;
    delAccBtn.onmouseenter = () => delAccBtn.style.background = 'rgba(153,27,27,0.15)';
    delAccBtn.onmouseleave = () => delAccBtn.style.background = 'transparent';
    delAccBtn.onclick = () => {
      const { overlay, box } = createModal({ id: 'sb-del-acc-overlay', title: '⛔ Delete Account', width: 320 });

      const warnEl = document.createElement('div');
      warnEl.style.cssText = 'background:rgba(153,27,27,0.15);border:1px solid rgba(153,27,27,0.4);border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:var(--sb-fs-12);color:var(--sb-danger);line-height:1.6;';
      warnEl.innerHTML = `This will:<br>
        <span style="color:var(--sb-text);">1.</span> Remove all payment methods<br>
        <span style="color:var(--sb-text);">2.</span> Disable the account email<br><br>
        <span style="color:var(--sb-text);font-weight:600;">${name}</span><br>
        <span style="color:var(--sb-text-muted);">${C.user.email}</span><br><br>
        This action cannot be undone.`;
      box.appendChild(warnEl);

      const accStatus = makeStatusEl();
      box.appendChild(accStatus);

      box.appendChild(makeDialogButtons({
        confirmLabel: 'Delete Account',
        confirmColor: 'var(--sb-danger-strong)',
        onCancel: () => overlay.remove(),
        onConfirm: (cBtn) => {
          cBtn.disabled = true;
          cBtn.textContent = 'Deleting...';
          accStatus.textContent = 'Removing payment methods...';

          // Step 1: Delete payment methods (skip if fails)
          gqlMutate('deleteAllPaymentMethods', DELETE_PAYMENT_METHODS_MUTATION,
            { input: { userId: C.user.id } },
            (pmData, pmErr) => {
              // Step 2: Disable email
              accStatus.textContent = 'Disabling account email...';
              const [localPart, domain] = C.user.email.split('@');
              const rand = Math.floor(Math.random() * 900000000) + 100000000;
              const disabledEmail = localPart + '_disabled' + rand + '@' + domain;

              gqlMutate('userChangeEmail', CHANGE_EMAIL_MUTATION,
                { input: { userId: C.user.id, email: disabledEmail } },
                (emailData, emailErr) => {
                  const respErr = emailData?.userChangeEmail?.error;
                  const errMsg = emailErr || respErr?.message;
                  if (errMsg) {
                    accStatus.style.color = 'var(--sb-danger)';
                    accStatus.textContent = '\u2718 ' + errMsg;
                    cBtn.disabled = false;
                    cBtn.textContent = 'Delete Account';
                    return;
                  }
                  accStatus.style.color = 'var(--sb-success-soft)';
                  accStatus.textContent = '\u2714 Account deleted.';
                  cBtn.textContent = '\u2714 Done';
                  delAccBtn.disabled = true;
                  delAccBtn.textContent = '\u2714 Account Deleted';
                  delAccBtn.style.color = 'var(--sb-success-soft)';
                  delAccBtn.style.borderColor = 'var(--sb-success-soft)';
                  delAccBtn.style.cursor = 'default';
                  mutPostComment(C.user.id, 'Account deleted', () => {});
                  autoClose(() => overlay.remove());
                }
              );
            }
          );
        },
      }));
    };
    C.panel.appendChild(delAccBtn);

  }

  /** Fraud blocklist add/remove. Reads its current state from cachedCustomerCtx. */
  function _cancelFraudSection(C) {
    // ── Fraud Blocklist Toggle ────────────────────────────────────────
    const fraudBtn = document.createElement('button');
    fraudBtn.textContent = '🚫 Blocklist (Fraud)';
    fraudBtn.style.cssText = `width:100%;padding:9px;border-radius:6px;border:1px solid var(--sb-solid-violet);background:transparent;
      color:var(--sb-tag-plan-fg);font-weight:600;font-size:var(--sb-fs-12);cursor:pointer;box-sizing:border-box;margin-top:6px;`;
    fraudBtn.onmouseenter = () => fraudBtn.style.background = 'rgba(124,58,237,0.1)';
    fraudBtn.onmouseleave = () => fraudBtn.style.background = 'transparent';

    // Use cached fraud status from initial user details query
    const cachedFraudStatus = cachedCustomerCtx?._fraudStatus || 'DEFAULT';
    fraudBtn._currentStatus = cachedFraudStatus;
    if (cachedFraudStatus === 'DECLINE') {
      fraudBtn.textContent = '✔ Blocklisted (Fraud)';
      fraudBtn.style.color = 'var(--sb-danger-strong)';
      fraudBtn.style.borderColor = 'var(--sb-danger-strong)';
    }

    fraudBtn.onclick = () => {
      const newStatus = (fraudBtn._currentStatus === 'DECLINE') ? 'DEFAULT' : 'DECLINE';
      const isAdding = newStatus === 'DECLINE';

      const { overlay, box } = createModal({
        id: 'sb-fraud-confirm-dialog',
        title: isAdding ? '🚫 Confirm Blocklist' : '✅ Remove from Blocklist',
        width: 360,
      });

      const intro = document.createElement('div');
      intro.style.cssText = 'font-size:var(--sb-fs-12);color:var(--sb-text-muted);margin-bottom:10px;';
      intro.textContent = isAdding
        ? `Add ${C.user.email} to the fraud blocklist?`
        : `Remove ${C.user.email} from the fraud blocklist?`;
      box.appendChild(intro);

      box.appendChild(makeLabel('Comment (optional)'));
      const ta = document.createElement('textarea');
      ta.rows = 3;
      ta.placeholder = isAdding ? 'Reason for blocklisting…' : 'Reason for unblocklisting…';
      ta.style.cssText = 'width:100%;background:var(--sb-bg-deep);color:var(--sb-text);border:1px solid var(--sb-border);border-radius:5px;padding:6px 8px;font-size:var(--sb-fs-12);box-sizing:border-box;resize:vertical;font-family:Arial,sans-serif;margin-bottom:10px;';
      box.appendChild(ta);

      const dlgStatus = makeStatusEl();
      box.appendChild(dlgStatus);

      const btns = makeDialogButtons({
        confirmLabel: isAdding ? 'Blocklist' : 'Remove',
        confirmColor: isAdding ? 'var(--sb-danger-border)' : 'var(--sb-solid-success)',
        onCancel: () => overlay.remove(),
        onConfirm: (confirmBtn, cancelBtn) => {
          confirmBtn.disabled = true; cancelBtn.disabled = true;
          confirmBtn.textContent = isAdding ? 'Blocklisting…' : 'Removing…';
          dlgStatus.style.color = 'var(--sb-text-muted)';
          dlgStatus.textContent = '';

          gqlMutate('userSetFraudStatus', SET_FRAUD_STATUS_MUTATION,
            { input: { userId: C.user.id, status: newStatus } },
            (data, err) => {
              const respErr = data?.userSetFraudStatus?.error;
              const errMsg = err || respErr?.message;
              if (errMsg) {
                dlgStatus.style.color = 'var(--sb-danger)';
                dlgStatus.textContent = '\u2718 ' + errMsg;
                confirmBtn.disabled = false; cancelBtn.disabled = false;
                confirmBtn.textContent = isAdding ? 'Blocklist' : 'Remove';
                return;
              }

              const base = isAdding ? 'Fraud blocklist applied' : 'Fraud blocklist removed';
              const note = ta.value.trim();
              const comment = note ? `${base} — ${note}` : base;

              dlgStatus.textContent = 'Posting comment…';
              mutPostComment(C.user.id, comment, () => {
                fraudBtn._currentStatus = newStatus;
                if (newStatus === 'DECLINE') {
                  fraudBtn.textContent = '✔ Blocklisted (Fraud)';
                  fraudBtn.style.color = 'var(--sb-danger-strong)';
                  fraudBtn.style.borderColor = 'var(--sb-danger-strong)';
                } else {
                  fraudBtn.textContent = '🚫 Blocklist (Fraud)';
                  fraudBtn.style.color = 'var(--sb-tag-plan-fg)';
                  fraudBtn.style.borderColor = 'var(--sb-solid-violet)';
                }
                fraudBtn.disabled = false;
                dlgStatus.style.color = 'var(--sb-success-soft)';
                dlgStatus.textContent = '✔ Done!';
                autoClose(() => overlay.remove());
              });
            }
          );
        },
      });
      box.appendChild(btns);
    };
    C.panel.appendChild(fraudBtn);
  }

  /** Cancel Subscription panel. Six sections that share one context object and
   *  one re-render: every destructive action here ends by calling C.render(), so
   *  the panel always reflects the account state it just changed. Split out of a
   *  402-line builder whose sections communicated only through those locals. */
  function showCancelPanel(btn) {
    if (document.getElementById('sb-cancel-panel')) {
      removePanel('sb-cancel-panel'); return;
    }

    loadCustomer(btn, (result) => {
      if (!result) return;
      const { user } = result;
      const panel = createPanel({ id: 'sb-cancel-panel', title: '❌ Cancel Subscription', width: 320 });

      function render() {
        // Everything but the user row is rebuilt; the row is index 0 after the header.
        [...panel.children].slice(1).forEach(c => c.remove());
        loadCustomerOrders(user.id, (orders) => {
          if (!document.getElementById('sb-cancel-panel')) return;

          const subStatus = user.subscription?.status || '';
          const isAwaitingCancel = !!cachedCustomerCtx?._isAwaitCancellation;
          const C = {
            panel, user, orders, subStatus, isAwaitingCancel, render,
            hasRecentSubOrder: (orders || []).some(o => o.type === 'SUBSCRIPTION' || !o.type),
            isAlreadyCancelled: subStatus.toLowerCase() === 'cancelled' || isAwaitingCancel,
            // Set by the cancel section, written to by all the others.
            statusEl: null,
          };

          _cancelInfoSection(C);
          _cancelOrderSection(C);
          _cancelCancelSection(C);
          _cancelPmSection(C);
          _cancelAccSection(C);
          _cancelFraudSection(C);
        });
      }
      render();
      panel._refresh = render;
    });
  }

  // ── Refund Charges Panel ──────────────────────────────────────────────────

  function showRefundChargesPanel(user) {
    if (document.getElementById('sb-refund-charges-panel')) {
      document.getElementById('sb-refund-charges-panel').remove(); return;
    }

    const panel = createPanel({ id: 'sb-refund-charges-panel', title: '💳 Charges — Refund', width: 360, right: 350 });

    const loadingEl = document.createElement('div');
    loadingEl.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-12);';
    loadingEl.textContent = 'Loading charges...';
    panel.appendChild(loadingEl);

    fetchCharges(user.id, (result) => {
      loadingEl.remove();

      if (!result || !result.all?.length) {
        const empty = document.createElement('div');
        empty.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-12);';
        empty.textContent = result ? 'No charges found for this month.' : 'Failed to load charges.';
        panel.appendChild(empty);
        return;
      }

      // Refunds are open on the latest billing group regardless of the linked order's
      // fulfilment status (Ivan, 2026-10-13) — the order-status gate is gone.

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
          sep.style.cssText = 'height:1px;background:var(--sb-control-bg);margin:10px 0;';
          panel.appendChild(sep);
        }

        const block = document.createElement('div');
        block.style.cssText = 'background:var(--sb-bg-deep);border-radius:8px;padding:10px 12px;';

        // Charge header
        const chgHdr = document.createElement('div');
        chgHdr.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;';
        const chgLeft = document.createElement('div');
        chgLeft.style.cssText = 'font-size:var(--sb-fs-12);color:var(--sb-text-muted);';
        chgLeft.textContent = chargeDateLabel(charge.paymentDate);
        const chgTotal = document.createElement('div');
        chgTotal.style.cssText = 'display:flex;align-items:center;gap:2px;font-weight:700;font-size:var(--sb-fs-13);color:var(--sb-text);';
        chgTotal.textContent = fmtMoney(totalCents, currency);
        chgTotal.appendChild(makeCopyBtn(fmtMoney(totalCents, currency), '📋'));
        chgHdr.appendChild(chgLeft);
        chgHdr.appendChild(chgTotal);
        block.appendChild(chgHdr);

        // Payment method
        const pmMethod = charge?.cashbirdDetails?.paymentMethod?.methodName || '';
        if (pmMethod) {
          const pmEl = document.createElement('div');
          pmEl.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-dim);margin-bottom:6px;';
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
            liEl.style.cssText = 'display:flex;justify-content:space-between;font-size:var(--sb-fs-11);color:var(--sb-text-muted);padding:2px 0;';
            const liName = document.createElement('span');
            liName.style.cssText = 'flex:1;margin-right:8px;';
            liName.textContent = li.description || li.productCode || li.type || '—';
            const liRight = document.createElement('span');
            liRight.style.cssText = 'display:flex;align-items:center;gap:6px;white-space:nowrap;';
            const liAmt = document.createElement('span');
            liAmt.style.color = 'var(--sb-text-secondary)';
            liAmt.textContent = fmtMoney(li.total || li.price || 0, currency);
            liRight.appendChild(liAmt);
            if (li.refundedAmount !== null && li.refundedAmount !== undefined) {
              const liRefTag = document.createElement('span');
              const fullyRefunded = li.refundedAmount >= (li.total || li.price || 0);
              liRefTag.style.cssText = 'font-size:var(--sb-fs-10);font-weight:600;color:' + (fullyRefunded ? 'var(--sb-success-soft)' : 'var(--sb-warning-strong)') + ';';
              liRefTag.textContent = fullyRefunded ? '\u21A9 refunded' : '\u21A9 partial';
              liRight.appendChild(liRefTag);
            }
            liEl.appendChild(liName);
            liEl.appendChild(liRight);
            breakdown.appendChild(liEl);
          });
          if (invoice.tax) {
            const taxEl = document.createElement('div');
            taxEl.style.cssText = 'display:flex;justify-content:space-between;font-size:var(--sb-fs-11);color:var(--sb-text-dim);padding:2px 0;border-top:1px solid var(--sb-border-soft);margin-top:3px;';
            taxEl.innerHTML = `<span>Tax</span><span>${fmtMoney(invoice.tax, currency)}</span>`;
            breakdown.appendChild(taxEl);
          }
          block.appendChild(breakdown);
        }

        // Credit summary
        if (credits.length) {
          const creditEl = document.createElement('div');
          creditEl.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-muted);margin-bottom:8px;';
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
          warnEl.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-warning-strong);margin-bottom:8px;';
          warnEl.textContent = alreadyRefunded
            ? '⚠ Partially refunded: ' + fmtMoney(Math.round(charge.refundAmount * 100), currency) + ' already returned'
            : '⚠ Refund records exist for this charge';
          block.appendChild(warnEl);
        }

        // Refund button
        const hasCreditsFlag = credits.length > 0 || shippingCredits.length > 0;
        const lineItemsForRefund = invoice?.lineItems || [];
        const linkedOrder = charge?.orders?.[0];
        const canRefund = hasCreditsFlag || lineItemsForRefund.length > 0;
        const refBtn = document.createElement('button');
        refBtn.textContent = 'Refund ' + fmtMoney(totalCents, currency);
        refBtn.disabled = !canRefund;
        refBtn.style.cssText = `
          width:100%;padding:7px;border-radius:6px;border:none;margin-top:4px;
          font-weight:700;font-size:var(--sb-fs-12);cursor:${canRefund ? 'pointer' : 'not-allowed'};
          background:${canRefund ? 'var(--sb-solid-violet)' : 'var(--sb-solid-violet-dim)'};
          color:${canRefund ? 'var(--sb-on-solid)' : 'var(--sb-text-dim)'};
          box-sizing:border-box;
        `;
        if (canRefund) {
          refBtn.onmouseenter = () => refBtn.style.background = 'var(--sb-solid-violet-hi)';
          refBtn.onmouseleave = () => refBtn.style.background = 'var(--sb-solid-violet)';
          const linkedStatus = (linkedOrder?.status || '').toUpperCase();
          const orderId = linkedStatus === 'UPGRADED' ? null : (linkedOrder?.id || null);
          refBtn.onclick = () => showRefundConfirmDialog(user, charge, credits, shippingCredits, currency, totalCents, refBtn, block, orderId);
        }
        block.appendChild(refBtn);
        panel.appendChild(block);
      });

      // Refund total row — updates dynamically after each refund
      const refundTotalRow = document.createElement('div');
      refundTotalRow.id = 'sb-refund-total-row';
      refundTotalRow.style.cssText = 'display:none;margin-top:10px;padding:8px 12px;background:var(--sb-bg-deep);border-radius:8px;border:1px solid rgba(110,231,183,0.2);';
      const refundTotalInner = document.createElement('div');
      refundTotalInner.style.cssText = 'display:flex;align-items:center;justify-content:space-between;';
      const refundTotalLabel = document.createElement('span');
      refundTotalLabel.style.cssText = 'font-size:var(--sb-fs-12);color:var(--sb-success-soft);font-weight:700;';
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

  // ── Cancel Order Dialog ───────────────────────────────────────────────────

  const CANCEL_ORDER_REASONS = [
    { value: 'charge clarification', label: 'Charge clarification' },
    { value: 'order change needed',  label: 'Order change needed' },
    { value: 'clarification needed', label: 'Clarification needed' },
    { value: '__OTHER__',            label: 'Other (custom comment)' },
  ];

  // Exactly what CRM's own Confirm Refund dialog offers (labels and enum values from
  // its bundle, 2026-10-13). CRM's full reason map has 9 more (Multiple Charges,
  // Charged While Cancelled / On Hold, Use Customer Credits, Subscription Plan Changed,
  // Cancelled gift subscription, Digital Genius issue, Premium Upcharges / Add-on,
  // Out-of-stock item) that its refund dialog doesn't show for a used credit — every
  // refund here force-refunds used credits, so they're left out.
  const REFUND_REASONS = [
    { value: 'APPROVED_REFUND',              label: 'Approved By Manager' },
    { value: 'THREAT',                       label: 'BBB / Social media / etc threat' },
    { value: 'CHARGEBACK',                   label: 'Chargeback' },
    { value: 'CUSTOMER_CHANGED_THEIR_MIND',  label: 'Customer Changed Their Mind' },
    { value: 'ISSUE_SHIPMENT',               label: 'Damaged/Empty/Missing package (no replacement)' },
    { value: 'FAMILY_EMERGENCY',             label: 'Family Emergency' },
    { value: 'FRAUD',                        label: 'Fraud' },
    { value: 'INCOMPLETE_SHIPMENT',          label: 'Incomplete Shipment' },
    { value: 'MISSED_SHIPMENT',              label: 'Missed Shipment' },
    { value: 'MULTIPLE_SUBSCRIBED_ACCOUNTS', label: 'Multiple Subscribed Accounts' },
    { value: 'PRICE_INCREASE',               label: 'Price Increase' },
    { value: 'PROMO_ISSUE',                  label: 'Promo Not Applied' },
    { value: 'UNDELIVERABLE_SHIPMENT',       label: 'Undeliverable / Returned Shipment' },
  ];

  function showCancelOrderDialog(order, user, triggerBtn) {
    const { overlay, box } = createModal({ id: 'sb-cancel-order-dialog', title: '❌ Cancel Order', width: 320 });

    // Order info
    const info = document.createElement('div');
    info.style.cssText = 'font-size:var(--sb-fs-12);color:var(--sb-text-muted);margin-bottom:14px;';
    info.textContent = orderMonthLabel(order) + ' order';
    box.appendChild(info);

    // Reason
    box.appendChild(makeLabel('Reason'));
    const select = makeSelect(CANCEL_ORDER_REASONS, 'charge clarification');
    select.style.marginBottom = '10px';
    box.appendChild(select);

    // Custom comment textarea (shown when "Other" is selected)
    const otherTa = document.createElement('textarea');
    otherTa.placeholder = 'Custom reason / comment…';
    otherTa.style.cssText = 'width:100%;min-height:60px;padding:6px 8px;background:var(--sb-bg-deep);border:1px solid var(--sb-border);border-radius:5px;color:var(--sb-text);font-size:var(--sb-fs-12);font-family:inherit;box-sizing:border-box;margin-bottom:14px;display:none;';
    box.appendChild(otherTa);
    select.addEventListener('change', () => {
      otherTa.style.display = select.value === '__OTHER__' ? 'block' : 'none';
    });

    // Status
    const statusEl = makeStatusEl();
    box.appendChild(statusEl);

    // Buttons
    const btns = makeDialogButtons({
      confirmLabel: 'Cancel Order',
      confirmColor: 'var(--sb-danger-border)',
      onCancel: () => overlay.remove(),
      onConfirm: (confirmBtn, cancelBtn) => {
        const isOther = select.value === '__OTHER__';
        const otherNote = otherTa.value.trim();
        if (isOther && !otherNote) {
          statusEl.style.color = 'var(--sb-danger)';
          statusEl.textContent = '✘ Please enter a custom comment';
          otherTa.focus();
          return;
        }

        confirmBtn.disabled = true; cancelBtn.disabled = true;
        select.disabled = true; otherTa.disabled = true;
        confirmBtn.textContent = 'Cancelling...';
        statusEl.textContent = 'Cancelling order...';

        mutCancelOrder(order.id, (data, err) => {
          if (err) {
            statusEl.style.color = 'var(--sb-danger)';
            statusEl.textContent = '✘ ' + err;
            confirmBtn.disabled = false; cancelBtn.disabled = false;
            select.disabled = false; otherTa.disabled = false;
            confirmBtn.textContent = 'Cancel Order';
            return;
          }

          statusEl.style.color = 'var(--sb-text-muted)';
          statusEl.textContent = 'Posting comment...';

          const d = new Date(order.year + '-' + String(order.month).padStart(2, '0') + '-01');
          const mLabel = d.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }).replace(' ', "'");
          const reasonText = isOther ? otherNote : select.value;
          const comment = 'Cancelled ' + mLabel + ' order, ' + reasonText;

          mutPostComment(user.id, comment, () => {
            statusEl.style.color = 'var(--sb-success-soft)';
            statusEl.textContent = '✔ Done!';
            confirmBtn.textContent = '✔ Done';
            triggerBtn.disabled = true;
            triggerBtn.textContent = '✔ Cancelled';
            triggerBtn.style.color = 'var(--sb-success-soft)';
            triggerBtn.style.borderColor = 'var(--sb-success-soft)';
            refreshOpenPanels('sb-cancel-panel', 'sb-order-panel');
            autoClose(() => overlay.remove());
          });
        });
      },
    });
    box.appendChild(btns);
  }

  // ── Refund Confirm Dialog ─────────────────────────────────────────────────

  function showRefundConfirmDialog(user, charge, credits, shippingCredits, currency, totalCents, refBtn, parentBlock, orderId) {
    const { overlay, box } = createModal({ id: 'sb-refund-confirm-overlay', title: 'Confirm Refund', width: 300 });
    const orderFulfilled = chargeRefundBlock(charge).blocked;
    let cancelledOrder = false;

    const invoice = charge?.cashbirdDetails?.invoice;
    const pmMethod = charge?.cashbirdDetails?.paymentMethod?.methodName || 'Unknown';
    const dateStr = chargeDateLabel(charge.paymentDate);

    // Details
    const details = document.createElement('div');
    details.style.cssText = 'background:var(--sb-bg-deep);border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:var(--sb-fs-12);line-height:1.8;color:var(--sb-text-secondary);';
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
      <div><span style="color:var(--sb-text-dim);">Date</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ${dateStr}</div>
      <div><span style="color:var(--sb-text-dim);">Method</span>&nbsp;&nbsp;&nbsp; ${pmMethod}</div>
      <div><span style="color:var(--sb-text-dim);">Credits</span>&nbsp;&nbsp;&nbsp; ${creditLine}${shippingLine}</div>
      <div style="margin-top:6px;padding-top:6px;border-top:1px solid var(--sb-control-bg);font-weight:700;font-size:var(--sb-fs-13);color:var(--sb-text);">
        Total refund: ${fmtMoney(totalCents, currency)}
      </div>`;
    box.appendChild(details);

    // Existing refund warning
    const existingRefunds = charge?.cashbirdDetails?.invoice?.refunds || [];
    if (existingRefunds.length) {
      const refWarn = document.createElement('div');
      refWarn.style.cssText = 'background:rgba(245,158,11,0.1);border:1px solid rgba(245,158,11,0.3);border-radius:6px;padding:8px 10px;margin-bottom:10px;font-size:var(--sb-fs-11);color:var(--sb-warning-strong);line-height:1.6;';
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
      liWarn.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-muted);margin-bottom:10px;line-height:1.8;';
      refundedItems.forEach(li => {
        const row = document.createElement('div');
        const fullyRefunded = li.refundedAmount >= li.total;
        row.style.color = fullyRefunded ? 'var(--sb-success-soft)' : 'var(--sb-warning-strong)';
        row.textContent = (fullyRefunded ? '\u2714' : '\u21A9') + ' ' + (li.description || li.type) + ': ' + fmtMoney(li.refundedAmount, currency) + ' refunded';
        liWarn.appendChild(row);
      });
      box.appendChild(liWarn);
    }

    // Upgraded order — must confirm which order to cancel before refund is allowed
    const orders = charge?.orders || [];
    const isUpgraded = (orders[0]?.status || '').toUpperCase() === 'UPGRADED';
    let cancelCandidateId = null;
    let cancelCandidateChk = null;
    if (isUpgraded) {
      const candidate = orders
        .filter(o => (o.status || '').toUpperCase() !== 'UPGRADED')
        .sort((a, b) => b.id - a.id)[0] || null;
      const upgradeWarn = document.createElement('div');
      upgradeWarn.style.cssText = 'background:rgba(245,158,11,0.08);border:1px solid rgba(245,158,11,0.3);border-radius:6px;padding:8px 10px;margin-bottom:10px;font-size:var(--sb-fs-11);color:var(--sb-warning-strong);';
      if (candidate) {
        cancelCandidateId = candidate.id;
        const chkRow = document.createElement('label');
        chkRow.style.cssText = 'display:flex;align-items:center;gap:8px;cursor:pointer;';
        cancelCandidateChk = document.createElement('input');
        cancelCandidateChk.type = 'checkbox';
        cancelCandidateChk.style.cssText = 'width:14px;height:14px;flex-shrink:0;cursor:pointer;accent-color:var(--sb-warning-strong);';
        const chkLabel = document.createElement('span');
        const candidateStatus = (candidate.warehouseOrder?.data?.status || candidate.status || '').toUpperCase() || 'UNKNOWN';
        chkLabel.innerHTML = `Original order was <strong>UPGRADED</strong>. Confirm cancel of order <strong>#${candidate.id}</strong> (${candidateStatus}) to proceed:`;
        chkRow.appendChild(cancelCandidateChk);
        chkRow.appendChild(chkLabel);
        upgradeWarn.appendChild(chkRow);
      } else {
        upgradeWarn.textContent = '⚠ Original order was UPGRADED — no active order found to cancel. Verify manually before refunding.';
      }
      box.appendChild(upgradeWarn);
    }

    const reasonLabel = makeLabel('Refund reason');
    reasonLabel.style.marginTop = '0';
    box.appendChild(reasonLabel);
    const reasonSel = makeSelect(REFUND_REASONS, 'CUSTOMER_CHANGED_THEIR_MIND');
    reasonSel.style.marginBottom = '12px';
    box.appendChild(reasonSel);

    const warn = document.createElement('div');
    warn.style.cssText = 'font-size:var(--sb-fs-12);color:var(--sb-danger);margin-bottom:14px;line-height:1.5;';
    warn.textContent = 'This will refund the payment to the customer. This cannot be undone.';
    box.appendChild(warn);

    const statusEl = makeStatusEl();
    box.appendChild(statusEl);

    const btns = makeDialogButtons({
      confirmLabel: 'Refund',
      confirmColor: 'var(--sb-solid-violet)',
      onCancel: () => overlay.remove(),
      onConfirm: (confirmBtn, cancelBtn) => {
        if (!orderFulfilled && cancelCandidateChk && !cancelCandidateChk.checked) return;
        confirmBtn.disabled = true;
        cancelBtn.disabled = true;
        reasonSel.disabled = true;
        confirmBtn.textContent = 'Refunding...';
        statusEl.textContent = '';

        const reason = reasonSel.value;
        const creditIds = credits.map(c => c.id);
        const shippingCreditIds = shippingCredits.map(c => c.id);

        const doRefund = () => {
          const onResult = (data, err) => {
            if (err) {
              statusEl.style.color = 'var(--sb-danger)';
              statusEl.textContent = '\u2718 ' + err;
              confirmBtn.disabled = false;
              cancelBtn.disabled = false;
              reasonSel.disabled = false;
              confirmBtn.textContent = 'Refund';
              return;
            }

            statusEl.style.color = 'var(--sb-text-muted)';
            statusEl.textContent = 'Posting comment...';

            // Track total refunded
            _totalRefundedCents += totalCents;
            _totalRefundedCurrency = currency;

            const refDate = new Date(charge.paymentDate);
            const refMonthLabel = refDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            const commentText = (cancelledOrder ? 'Cancelled & Refunded ' : 'Refunded ') + refMonthLabel + ' order';

            const afterComment = () => {
              statusEl.style.color = 'var(--sb-success-soft)';
              statusEl.textContent = '\u2714 Refunded!';
              confirmBtn.textContent = '\u2714 Done';
              refBtn.disabled = true;
              refBtn.textContent = '\u2714 Refunded';
              refBtn.style.background = 'rgba(110,231,183,0.15)';
              refBtn.style.color = 'var(--sb-success-soft)';
              refBtn.style.cursor = 'not-allowed';
              refreshOpenPanels('sb-cancel-panel', 'sb-order-panel');
              autoClose(() => overlay.remove());
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
            mutRefundCredits(user.id, creditIds, shippingCreditIds, reason, (data, err) => {
              if (err) return onResult(null, err);
              statusEl.textContent = 'Refunding upcharge...';
              mutRefundInvoiceItems(user.id, unrefundedItems, reason, onResult);
            });
          } else if (hasCreditsFlag) {
            mutRefundCredits(user.id, creditIds, shippingCreditIds, reason, onResult);
          } else if (hasItems) {
            mutRefundInvoiceItems(user.id, unrefundedItems, reason, onResult);
          } else {
            onResult({}, null);
          }
        };

        // Refunds no longer wait on fulfilment, so only cancel an order that can still be
        // cancelled — a shipped/delivered one is refund-only (its cancel would fail and
        // block the refund).
        const effectiveOrderId = orderFulfilled ? null : (orderId || (cancelCandidateChk?.checked ? cancelCandidateId : null));
        cancelledOrder = !!effectiveOrderId;
        if (effectiveOrderId) {
          statusEl.style.color = 'var(--sb-text-muted)';
          statusEl.textContent = 'Cancelling order...';
          mutCancelOrder(effectiveOrderId, reason, (data, cancelErr) => {
            if (cancelErr) {
              statusEl.style.color = 'var(--sb-danger)';
              statusEl.textContent = '✘ Cancel failed: ' + cancelErr;
              confirmBtn.disabled = false; cancelBtn.disabled = false;
              reasonSel.disabled = false;
              confirmBtn.textContent = 'Refund';
              return;
            }
            doRefund();
          });
        } else {
          doRefund();
        }
      },
    });
    if (cancelCandidateChk) {
      btns._confirmBtn.disabled = true;
      btns._confirmBtn.style.opacity = '0.4';
      btns._confirmBtn.style.cursor = 'not-allowed';
      cancelCandidateChk.onchange = () => {
        btns._confirmBtn.disabled = !cancelCandidateChk.checked;
        btns._confirmBtn.style.opacity = cancelCandidateChk.checked ? '1' : '0.4';
        btns._confirmBtn.style.cursor = cancelCandidateChk.checked ? 'pointer' : 'not-allowed';
      };
    }
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
    infoEl.style.cssText = 'display:none;background:rgba(245,158,11,0.1);border:1px solid rgba(245,158,11,0.3);border-radius:6px;padding:8px 10px;margin-bottom:8px;font-size:var(--sb-fs-12);color:var(--sb-warning-strong);line-height:1.6;';
    wrap.appendChild(infoEl);

    const btnRow = document.createElement('div');
    btnRow.style.cssText = 'display:none;gap:6px;';
    wrap.appendChild(btnRow);

    const billBtn = document.createElement('button');
    billBtn.textContent = '💳 Bill Upcharge';
    billBtn.style.cssText = 'flex:1;padding:7px;border-radius:6px;border:none;background:var(--sb-warning-strong);color:var(--sb-on-warning);font-weight:700;font-size:var(--sb-fs-12);cursor:pointer;';
    billBtn.onmouseenter = () => billBtn.style.background = 'var(--sb-warning-hi)';
    billBtn.onmouseleave = () => billBtn.style.background = 'var(--sb-warning-strong)';
    btnRow.appendChild(billBtn);

    const skipBtn = document.createElement('button');
    skipBtn.textContent = 'Billed Already';
    skipBtn.style.cssText = 'padding:7px 12px;border-radius:6px;border:1px solid var(--sb-hover);background:transparent;color:var(--sb-text-muted);font-size:var(--sb-fs-12);cursor:pointer;';
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
      detailsEl.style.cssText = 'background:var(--sb-bg-deep);border-radius:6px;padding:10px 12px;margin-bottom:12px;font-size:var(--sb-fs-12);color:var(--sb-text-secondary);line-height:1.8;';

      // List upcharge products
      let detailsHtml = '';
      itemChecks.filter(({ chk }) => chk.checked).forEach(({ item }) => {
        const pi = item.product?.productInfo || item.productInfo;
        const up = pi?.upchargePrice || 0;
        if (up > 0) {
          detailsHtml += `<div>${pi.name} by ${pi.brand} — <span style="color:var(--sb-warning-strong);font-weight:600;">+$${up}</span></div>`;
        }
      });
      detailsHtml += `<div style="margin-top:6px;padding-top:6px;border-top:1px solid var(--sb-control-bg);font-weight:700;font-size:var(--sb-fs-13);color:var(--sb-text);">Total: $${total}</div>`;
      detailsEl.innerHTML = detailsHtml;
      box.appendChild(detailsEl);

      const chargeStatus = makeStatusEl();
      box.appendChild(chargeStatus);

      box.appendChild(makeDialogButtons({
        confirmLabel: 'Charge $' + total,
        confirmColor: 'var(--sb-warning-strong)',
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
                chargeStatus.style.color = 'var(--sb-danger)';
                chargeStatus.textContent = '\u2718 ' + errMsg;
                cBtn.disabled = false;
                cancelBtn.disabled = false;
                cBtn.textContent = 'Charge $' + total;
                return;
              }

              const state = data?.charge?.data?.state || 'UNKNOWN';
              const paid = data?.charge?.data?.paid;
              if (state === 'PAID') {
                chargeStatus.style.color = 'var(--sb-success-soft)';
                chargeStatus.textContent = '\u2714 Charged successfully' + (paid ? ' — $' + (paid / 100).toFixed(2) + ' (incl. tax)' : '');
                cBtn.textContent = '\u2714 Done';
                _upchargeCleared = true;
                statusEl.style.color = 'var(--sb-success-soft)';
                statusEl.textContent = '\u2714 Upcharge billed — $' + total;
                const upchargeLines = itemChecks.filter(({ chk }) => chk.checked).map(({ item }) => { const pi = item.product?.productInfo || item.productInfo; const up = pi?.upchargePrice || 0; return up > 0 ? `${pi.name} by ${pi.brand} (+$${up})` : null; }).filter(Boolean); if (upchargeLines.length) { mutPostComment(userId, 'Premium upcharge billed: ' + upchargeLines.join(', '), () => {}); }
                refresh();
                autoClose(() => overlay.remove());
              } else {
                chargeStatus.style.color = 'var(--sb-danger)';
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
      statusEl.style.color = 'var(--sb-text-muted)';
      statusEl.textContent = 'Upcharge skipped — billed separately.';
      refresh();
    };

    return { el: wrap, refresh };
  }

  function showProductSearchPanel(itemChecks, itemsWrap, onItemAdded, userId) {
    const existingPanel = document.getElementById('sb-product-search-panel');
    if (existingPanel) { existingPanel.remove(); return; }

    // Position to the left of the replacement form overlay
    const panel = document.createElement('div');
    panel.id = 'sb-product-search-panel';
    panel.style.cssText = `
      all:initial;position:fixed;top:50%;left:50%;transform:translate(-110%, -50%);
      width:520px;max-height:80vh;overflow-y:auto;
      background:var(--sb-bg-panel);color:var(--sb-text);font-family:Arial,sans-serif;font-size:var(--sb-fs-13);
      border-radius:10px;box-shadow:var(--sb-shadow);padding:14px;
      z-index:10000001;border:1px solid var(--sb-border);box-sizing:border-box;
    `;

    // Header
    const hdr = document.createElement('div');
    hdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;';
    hdr.innerHTML = '<span style="font-weight:700;font-size:var(--sb-fs-14);">🔍 Product Search</span>';
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close \u2715';
    closeBtn.style.cssText = 'background:var(--sb-control-bg);border:1px solid var(--sb-hover);color:var(--sb-text-muted);font-size:var(--sb-fs-11);font-weight:600;cursor:pointer;padding:3px 10px;border-radius:20px;';
    closeBtn.onmouseenter = () => closeBtn.style.background = 'var(--sb-hover)';
    closeBtn.onmouseleave = () => closeBtn.style.background = 'var(--sb-control-bg)';
    closeBtn.onclick = () => panel.remove();
    hdr.appendChild(closeBtn);
    panel.appendChild(hdr);

    // Search input row
    const inputRow = document.createElement('div');
    inputRow.style.cssText = 'display:flex;gap:6px;margin-bottom:10px;';
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search by name...';
    searchInput.style.cssText = 'flex:1;padding:7px 10px;border-radius:6px;border:1px solid var(--sb-hover);background:var(--sb-bg-input);color:var(--sb-text);font-size:var(--sb-fs-13);outline:none;box-sizing:border-box;';
    const searchBtn = document.createElement('button');
    searchBtn.textContent = 'Search';
    searchBtn.style.cssText = 'padding:7px 14px;border-radius:6px;border:none;background:var(--sb-solid-primary);color:var(--sb-on-solid);font-weight:700;font-size:var(--sb-fs-13);cursor:pointer;';
    inputRow.appendChild(searchInput);
    inputRow.appendChild(searchBtn);
    panel.appendChild(inputRow);

    const resultsEl = document.createElement('div');
    panel.appendChild(resultsEl);

    function doSearch() {
      const q = searchInput.value.trim();
      if (!q || q.length < 2) { resultsEl.innerHTML = '<div style="color:var(--sb-text-muted);font-size:var(--sb-fs-12);">Type at least 2 characters.</div>'; return; }
      resultsEl.innerHTML = '<div style="color:var(--sb-text-muted);font-size:var(--sb-fs-12);">Searching…</div>';

      crmRequest({
        method: 'POST', url: GRAPHQL_URL,
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({
          operationName: 'productSuggestion',
          query: PRODUCT_SEARCH_QUERY,
          // userId scopes results to the customer's market/currency — without it the
          // backend can return a legacy trading item that has no valid id, which
          // makes replacementSave fail its trading-item lookup (e.g. fragrance cases).
          variables: { input: { name: q, sections: ['Subscription', 'Extras', 'AddonSubscription', 'Ecommerce'], statuses: ['LIVE', 'OUT_OF_STOCK', 'NOT_AVAILABLE_FOR_NEW_ORDERS'], userId } },
        }),
        onload(res) {
          if (res.status === 401) { handle401(); resultsEl.innerHTML = '<div style="color:var(--sb-danger);">Sign in required.</div>'; return; }
          if (res.status === 403) { handle403(); resultsEl.innerHTML = '<div style="color:var(--sb-danger);">CRM captcha required.</div>'; return; }
          try {
            const json = JSON.parse(res.responseText);
            console.log('[BirdsEye] Product search parsed:', json?.data?.productSuggestion?.data?.length, 'results, error:', json?.data?.productSuggestion?.error);
            const products = json?.data?.productSuggestion?.data || [];
            const err = json?.data?.productSuggestion?.error?.message;
            if (err) { resultsEl.innerHTML = `<div style="color:var(--sb-danger);">${err}</div>`; return; }
            if (!products.length) { resultsEl.innerHTML = '<div style="color:var(--sb-text-muted);font-size:var(--sb-fs-12);">No products found.</div>'; return; }
            renderProductResults(products);
          } catch(e) {
            resultsEl.innerHTML = '<div style="color:var(--sb-danger);">Parse error.</div>';
          }
        },
        onerror() { resultsEl.innerHTML = '<div style="color:var(--sb-danger);">Network error.</div>'; },
      });
    }

    function renderProductResults(products) {
      resultsEl.innerHTML = '';
      const count = document.createElement('div');
      count.style.cssText = 'color:var(--sb-text-muted);margin-bottom:8px;font-size:var(--sb-fs-11);';
      count.textContent = products.length + ' result(s)';
      resultsEl.appendChild(count);

      // Split into sections
      const subProducts = products.filter(p => p.section === 'Subscription' || p.section === 'Extras');
      const sampleProducts = products.filter(p => p.section === 'AddonSubscription');
      const ecommProducts = products.filter(p => p.section === 'Ecommerce');

      const columnsWrap = document.createElement('div');

      function renderColumn(title, items, container) {
        const col = document.createElement('div');
        col.style.cssText = 'flex:1 1 150px;min-width:0;';
        const hdr = document.createElement('div');
        hdr.style.cssText = 'font-size:var(--sb-fs-10);font-weight:600;color:var(--sb-text-dim);letter-spacing:0.05em;margin-bottom:6px;';
        hdr.textContent = title;
        col.appendChild(hdr);

        if (!items.length) {
          const empty = document.createElement('div');
          empty.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-dim);';
          empty.textContent = 'No results';
          col.appendChild(empty);
        }

        items.forEach(product => {
          const info = product.productInfo;
          if (!info) return;

          const card = document.createElement('div');
          card.style.cssText = `
            padding:6px 8px;margin-bottom:4px;border-radius:5px;background:var(--sb-bg-input);
            border:1px solid transparent;transition:0.15s;display:flex;justify-content:space-between;align-items:center;gap:6px;
          `;
          card.onmouseenter = () => card.style.borderColor = 'rgba(99,102,241,0.4)';
          card.onmouseleave = () => card.style.borderColor = 'transparent';

          const leftCol = document.createElement('div');
          leftCol.style.cssText = 'flex:1;min-width:0;';

          const nameEl = document.createElement('div');
          nameEl.style.cssText = 'font-size:var(--sb-fs-11);font-weight:600;color:var(--sb-text);';
          nameEl.textContent = info.name + ' by ' + info.brand;
          leftCol.appendChild(nameEl);

          const tagsRow = document.createElement('div');
          tagsRow.style.cssText = 'display:flex;gap:3px;margin-top:2px;flex-wrap:wrap;align-items:center;';

          // Add-on tag for Extras
          if (product.section === 'Extras') {
            const addonBadge = document.createElement('span');
            addonBadge.style.cssText = 'font-size:var(--sb-fs-9);padding:1px 4px;border-radius:3px;background:var(--sb-tag-plan-bg);color:var(--sb-tag-plan-fg);border:1px solid var(--sb-tag-plan-border);';
            addonBadge.textContent = 'add-on';
            tagsRow.appendChild(addonBadge);
          }

          // Size badge for ecomm — full-size bottles vary per SKU, unlike the
          // fixed 0.27oz / 1.5ml of the other sections.
          if (product.section === 'Ecommerce' && product.volume) {
            const sizeBadge = document.createElement('span');
            sizeBadge.style.cssText = 'font-size:var(--sb-fs-9);padding:1px 4px;border-radius:3px;background:rgba(96,165,250,0.15);color:var(--sb-info);border:1px solid rgba(96,165,250,0.3);';
            sizeBadge.textContent = product.volumeUnit === 'oz' ? product.volume.toFixed(2) + 'oz' : product.volume + (product.volumeUnit || '');
            tagsRow.appendChild(sizeBadge);
          }

          // Upcharge badge
          if (info.upchargePrice && info.upchargePrice > 0) {
            const upBadge = document.createElement('span');
            upBadge.style.cssText = 'font-size:var(--sb-fs-9);padding:1px 4px;border-radius:3px;background:rgba(245,158,11,0.15);color:var(--sb-warning-strong);border:1px solid rgba(245,158,11,0.3);';
            upBadge.textContent = '+$' + info.upchargePrice;
            tagsRow.appendChild(upBadge);
          }

          // Stock badge
          if (product.status && product.status !== 'LIVE') {
            const stockBadge = document.createElement('span');
            const stockLabel = product.status.replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
            stockBadge.style.cssText = 'font-size:var(--sb-fs-9);padding:1px 4px;border-radius:3px;background:rgba(239,68,68,0.15);color:var(--sb-danger-strong);border:1px solid rgba(239,68,68,0.3);';
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
          addBtn.style.cssText = 'padding:3px 8px;border-radius:4px;border:1px solid var(--sb-solid-primary-lo);background:transparent;color:var(--sb-accent);font-weight:700;font-size:var(--sb-fs-12);cursor:pointer;flex-shrink:0;';
          addBtn.onmouseenter = () => addBtn.style.background = 'rgba(99,102,241,0.15)';
          addBtn.onmouseleave = () => addBtn.style.background = 'transparent';
          addBtn.onclick = () => {
            const fakeItem = {
              // Searched products aren't existing source-order line items, so their
              // orderItem id must be null (matches CRM's own replacement payload);
              // product.id carries the market-scoped trading item id from the search.
              id: null,
              product: { id: product.id, status: product.status, productInfo: info },
            };

            const chk = document.createElement('input');
            chk.type = 'checkbox'; chk.checked = true;
            chk.style.cssText = 'cursor:pointer;accent-color:var(--sb-accent-strong);flex-shrink:0;';
            itemChecks.push({ item: fakeItem, chk, starterChk: null });
            chk.addEventListener('change', () => { if (onItemAdded) onItemAdded(); });

            const row = document.createElement('div');
            row.style.cssText = 'margin-bottom:6px;border-left:2px solid var(--sb-solid-primary-lo);padding-left:6px;';
            const lbl = document.createElement('label');
            lbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:var(--sb-fs-12);color:var(--sb-accent);flex-wrap:wrap;';
            lbl.appendChild(chk);
            lbl.appendChild(document.createTextNode(info.name + ' by ' + info.brand));

            if (vol) {
              const volTag = document.createElement('span');
              const volText = unit === 'oz' ? vol.toFixed(2) + 'oz' : vol + unit;
              volTag.style.cssText = 'font-size:var(--sb-fs-10);color:var(--sb-text-muted);';
              volTag.textContent = '(' + volText + ')';
              lbl.appendChild(volTag);
            }

            const addedTag = document.createElement('span');
            addedTag.style.cssText = 'font-size:var(--sb-fs-10);padding:1px 5px;border-radius:3px;background:rgba(99,102,241,0.2);color:var(--sb-accent);border:1px solid rgba(99,102,241,0.3);';
            addedTag.textContent = 'Added';
            lbl.appendChild(addedTag);

            row.appendChild(lbl);
            itemsWrap.appendChild(row);

            addBtn.textContent = '✔';
            addBtn.style.color = 'var(--sb-success-soft)';
            addBtn.style.borderColor = 'var(--sb-success-soft)';
            addBtn.disabled = true;
            if (onItemAdded) onItemAdded();
          };

          card.appendChild(addBtn);
          col.appendChild(card);
        });

        container.appendChild(col);
      }

      // Subscription column always shows (even empty); the others only when they
      // have hits, so a plain subscription search keeps its full-width layout.
      const columns = [{ title: 'SUBSCRIPTION (0.27oz)', items: subProducts }];
      if (sampleProducts.length) columns.push({ title: 'SAMPLES (1.5ml)', items: sampleProducts });
      if (ecommProducts.length) columns.push({ title: 'ECOMM (full size)', items: ecommProducts });
      columnsWrap.style.cssText = columns.length > 1 ? 'display:flex;flex-wrap:wrap;gap:8px;' : '';
      columns.forEach(({ title, items }) => renderColumn(title, items, columnsWrap));
      resultsEl.appendChild(columnsWrap);
    }

    searchBtn.onclick = doSearch;
    searchInput.addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });

    document.body.appendChild(panel);
    setTimeout(() => searchInput.focus(), 50);
  }

  /** The "items to replace" checkbox list. Returns the wrapper and the check
   *  array, which the confirm handler and the upcharge manager both read —
   *  itemChecks entries are the { chk, item } pairs, not raw elements. */
  function _replacementItemRows(form, order, onToggle) {
    // Items
    form.appendChild(makeLabel('ITEMS TO REPLACE'));
    const itemsWrap = document.createElement('div');
    itemsWrap.style.cssText = 'background:var(--sb-bg-deep);border-radius:6px;padding:8px 10px;';
    const itemChecks = [];

    (order.orderItems || []).forEach(item => {
      const info = item?.product?.productInfo;
      const name = info ? `${info.name} by ${info.brand}` : `Item #${item.id}`;
      const productType = info?.productType || '';
      const isDrift = productType === 'CarFreshenerRefill';
      const stockStatus = item?.product?.status;
      const stockBadge = {
        'OUT_OF_STOCK':              { text: 'Out of Stock',    color: 'var(--sb-danger-strong)' },
        'NOT_AVAILABLE_FOR_NEW_ORDERS': { text: 'Not Available', color: 'var(--sb-plaque-orange)' },
        'INTERNAL_USE':              { text: 'Internal Use',    color: 'var(--sb-plaque-yellow)' },
      }[stockStatus];

      const row = document.createElement('div');
      row.style.cssText = 'margin-bottom:6px;';

      const lbl = document.createElement('label');
      lbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:var(--sb-fs-12);color:var(--sb-text-secondary);flex-wrap:wrap;';
      const chk = document.createElement('input');
      chk.type = 'checkbox'; chk.checked = true;
      chk.style.cssText = 'cursor:pointer;accent-color:var(--sb-accent-strong);flex-shrink:0;';
      lbl.appendChild(chk);
      lbl.appendChild(document.createTextNode(name));
      if (stockBadge) {
        const badge = document.createElement('span');
        badge.style.cssText = `font-size:var(--sb-fs-10);font-weight:600;color:${stockBadge.color};white-space:nowrap;`;
        badge.textContent = '● ' + stockBadge.text;
        lbl.appendChild(badge);
      }
      row.appendChild(lbl);

      let starterChk = null;
      if (isDrift) {
        const starterRow = document.createElement('div');
        starterRow.style.cssText = 'margin-left:22px;margin-top:3px;';
        const sLbl = document.createElement('label');
        sLbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:var(--sb-fs-11);color:var(--sb-text-muted);';
        starterChk = document.createElement('input');
        starterChk.type = 'checkbox'; starterChk.checked = true;
        starterChk.style.cssText = 'cursor:pointer;accent-color:var(--sb-accent-strong);';
        sLbl.appendChild(starterChk);
        sLbl.appendChild(document.createTextNode('Starter Set'));
        starterRow.appendChild(sLbl);
        row.appendChild(starterRow);
      }

      itemsWrap.appendChild(row);
      itemChecks.push({ item, chk, starterChk });
      chk.addEventListener('change', () => onToggle && onToggle());
    });
    form.appendChild(itemsWrap);
    return { itemsWrap, itemChecks };
  }

  function showReplacementForm(order, user, totalOrders) {
    const existingForm = document.getElementById('sb-replacement-form');
    if (existingForm) { existingForm.remove(); document.getElementById('sb-product-search-panel')?.remove(); return; }

    const isEcomm = order.type && order.type !== 'SUBSCRIPTION' && order.type !== 'MANUAL';
    const monthLabel = order.month
      ? new Date(order.year, order.month - 1).toLocaleString('en-US', { month:'short', year:'numeric' })
      : 'Order';

    const form = createPanel({
      id: 'sb-replacement-form',
      title: '\uD83D\uDD04 Replacement — ' + monthLabel,
      width: 400,
      maxHeight: 720,
      centered: true,
      draggable: true,
      onClose: () => document.getElementById('sb-product-search-panel')?.remove(),
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
    const { itemsWrap, itemChecks } = _replacementItemRows(form, order, () => upchargeUI.refresh());

    // Product search button
    const searchRow = document.createElement('div');
    searchRow.style.cssText = 'margin-top:8px;';
    const searchProductBtn = document.createElement('button');
    searchProductBtn.textContent = '🔍 Search Product';
    searchProductBtn.style.cssText = 'padding:5px 10px;border-radius:5px;border:1px solid var(--sb-accent-strong);background:transparent;color:var(--sb-accent-strong);font-weight:600;font-size:var(--sb-fs-11);cursor:pointer;';
    searchProductBtn.onmouseenter = () => searchProductBtn.style.background = 'rgba(129,140,248,0.15)';
    searchProductBtn.onmouseleave = () => searchProductBtn.style.background = 'transparent';
    searchProductBtn.onclick = () => showProductSearchPanel(itemChecks, itemsWrap, () => upchargeUI.refresh(), user.id);
    searchRow.appendChild(searchProductBtn);
    form.appendChild(searchRow);

    // Comment
    form.appendChild(makeLabel('ADDITIONAL COMMENT (OPTIONAL)'));
    const commentInput = document.createElement('textarea');
    commentInput.placeholder = 'Add a note...';
    commentInput.rows = 2;
    commentInput.style.cssText = 'width:100%;background:var(--sb-bg-deep);color:var(--sb-text);border:1px solid var(--sb-border);border-radius:5px;padding:6px 8px;font-size:var(--sb-fs-12);box-sizing:border-box;resize:vertical;font-family:Arial,sans-serif;';
    form.appendChild(commentInput);

    // Options
    form.appendChild(makeLabel('OPTIONS'));
    const optsWrap = document.createElement('div');
    optsWrap.style.cssText = 'background:var(--sb-bg-deep);border-radius:6px;padding:8px 10px;display:flex;flex-direction:column;gap:6px;';

    function makeCheckRow(labelText, defaultChecked = false) {
      const lbl = document.createElement('label');
      lbl.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:var(--sb-fs-12);color:var(--sb-text-secondary);';
      const chk = document.createElement('input');
      chk.type = 'checkbox'; chk.checked = defaultChecked;
      chk.style.cssText = 'cursor:pointer;accent-color:var(--sb-accent-strong);';
      const warn = document.createElement('span');
      warn.textContent = '⚠';
      warn.style.cssText = 'color:var(--sb-warning-strong);font-size:var(--sb-fs-14);display:none;margin-left:2px;';
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
      border:none;background:var(--sb-solid-primary-lo);color:var(--sb-on-solid);font-weight:700;
      font-size:var(--sb-fs-13);cursor:pointer;box-sizing:border-box;
    `;
    confirmBtn.onmouseenter = () => confirmBtn.style.background = 'var(--sb-solid-primary)';
    confirmBtn.onmouseleave = () => confirmBtn.style.background = 'var(--sb-solid-primary-lo)';

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
        statusEl.style.color = 'var(--sb-danger)';
        statusEl.textContent = 'Select at least one item.';
        return;
      }

      confirmBtn.disabled = true;
      confirmBtn.textContent = 'Submitting...';
      statusEl.textContent = '';

      const variables = {
        input: {
          kustomer: window.location.href,
          reason: reasonSel.value,
          bom: bomSel.value,
          customerId: user.id,
          addressConfirmed: addressChk.checked,
          picturesProvided: picturesChk.checked,
          orderNumber: order.orderNumber,
          oneTimeCourtesy: false,
          sourceOrder: {
            id: order.id,
            orderItems: selectedItems.map(i => ({
              id: i.id, starterSet: i.starterSet, product: { id: i.product.id },
            })),
          },
        },
      };

      crmRequest({
        method: 'POST', url: GRAPHQL_URL,
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({ operationName: 'replacementSave', query: REPLACEMENT_MUTATION, variables }),
        onload(res) {
          if (res.status === 401) {
            handle401();
            statusEl.style.color = 'var(--sb-danger)';
            statusEl.textContent = '\u2718 Token expired — click 🔑 Token to update.';
            confirmBtn.disabled = false;
            confirmBtn.textContent = 'Confirm Replacement';
            return;
          }
          if (res.status === 403) {
            handle403();
            statusEl.style.color = 'var(--sb-danger)';
            statusEl.textContent = '\u2718 CRM captcha required — open CRM in browser.';
            confirmBtn.disabled = false;
            confirmBtn.textContent = 'Confirm Replacement';
            return;
          }
          try {
            const json = JSON.parse(res.responseText);
            // The backend can fail two ways: a business error inside
            // replacementSave.error, or a top-level GraphQL `errors` array with
            // replacementSave === null (e.g. "Could not find trading item \u2026" when
            // a non-replaceable line item is included). Treat BOTH as failure, and
            // require a saved id \u2014 otherwise the mutation silently no-ops and we
            // falsely report success.
            const err = json?.data?.replacementSave?.error?.message || json?.errors?.[0]?.message;
            const savedId = json?.data?.replacementSave?.data?.id;
            if (err || !savedId) {
              statusEl.style.color = 'var(--sb-danger)';
              statusEl.textContent = '\u2718 ' + (err || 'Replacement was not created (no confirmation from CRM).');
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

            statusEl.style.color = 'var(--sb-text-muted)';
            statusEl.textContent = 'Finding automated comment...';

            const postComment = () => {
              statusEl.textContent = 'Posting comment...';
              gqlMutate('createUserComment', CREATE_COMMENT_MUTATION,
                { input: { userId: user.id, comment: commentText, kustomerUrl: window.location.href } },
                (data, err) => {
                  if (err) {
                    statusEl.style.color = 'var(--sb-warning-strong)';
                    statusEl.textContent = '\u2714 Replacement created, but comment failed.';
                  } else {
                    statusEl.style.color = 'var(--sb-success-soft)';
                    statusEl.textContent = '\u2714 Done!';
                  }
                  confirmBtn.textContent = '\u2714 Done';
                  refreshOpenPanels('sb-order-panel');
                  autoClose(() => { form.remove(); document.getElementById('sb-product-search-panel')?.remove(); });
                }
              );
            };

            statusEl.textContent = 'Finding automated comment...';
            fetchComments(user.id, (comments) => {
              const autoComment = order.orderNumber
                ? comments.find(c => c.comment && c.comment.includes(order.orderNumber))
                : null;
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
            statusEl.style.color = 'var(--sb-danger)';
            statusEl.textContent = '\u2718 Unexpected error.';
            confirmBtn.disabled = false;
            confirmBtn.textContent = 'Confirm Replacement';
          }
        },
        onerror() {
          statusEl.style.color = 'var(--sb-danger)';
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

  // Kustomer's own save updates its header by feeding the PATCH response into its
  // React store; our out-of-band PATCH can't reach that store, and Kustomer no
  // longer pushes the change back over its websocket, so the header stays stale
  // until a reload. Reflect the rename in the DOM directly: rewrite leaf elements
  // whose entire text is the old name (specific enough to be safe — composite
  // strings like "Conversation with <name>" are left alone). A later navigation
  // re-fetches the same value, so this never drifts from the server.
  function nudgeKustomerName(oldName, newName) {
    const from = (oldName || '').trim();
    const to   = (newName || '').trim();
    if (!from || !to || from === to) return;  // nothing to do (incl. blank → first name)
    let touched = 0;
    document.querySelectorAll('body *').forEach(el => {
      if (el.childElementCount === 0 && el.textContent.trim() === from) {
        el.textContent = to;
        touched++;
      }
    });
    return touched;
  }

  /** Core name-fill for a Kustomer customer: resolve the CRM name (or case the
   *  existing Kustomer name), then PATCH it. Returns { changed, label, color }
   *  describing the outcome; throws only on a hard Kustomer GET/PATCH error.
   *  `emailHint` lets a background caller (e.g. CCPA) skip the URL/cache email
   *  lookup. When `nudge` is true (the customer is the one on screen), also nudges
   *  Kustomer's UI to reflect the rename without a reload. Extracted so both the
   *  Fill Name toolbar button and the CCPA flow share one implementation. */
  // Two different guards, previously conflated into one — which is what made
  // "Scentbird User" block the fill it was supposed to trigger.
  //
  // PROTECTED: names on the KUSTOMER record we must never overwrite, whatever CRM says.
  // "DO NOT MERGE OR REPLY" is a deliberate agent flag on specific tickets; renaming it
  // defeats its whole purpose.
  const PROTECTED_KUSTOMER_NAMES = new Set(['do not merge or reply']);
  function isProtectedKustomerName(n) {
    return PROTECTED_KUSTOMER_NAMES.has(String(n || '').trim().toLowerCase());
  }

  // PLACEHOLDER: names never worth WRITING INTO Kustomer. CRM stores "Scentbird User"
  // on accounts with no real name on file, so filling it in replaces whatever Kustomer
  // had with a value that identifies nobody. It belongs on the value we're about to
  // write, not on the value we found — a Kustomer record already reading "Scentbird
  // User" is the single best candidate for a fill, and guarding it here skipped exactly
  // the tickets that needed the button.
  const PLACEHOLDER_NAMES = new Set(['scentbird user']);
  function isPlaceholderName(n) {
    return PLACEHOLDER_NAMES.has(String(n || '').trim().toLowerCase());
  }

  async function fillKustomerName(kustomerId, emailHint, nudge) {
    // Round-trip the full attribute block: PATCH is full-replace on the editable subset.
    const attrs = await kustomerGetCustomer(kustomerId);

    // Bail before any CRM lookup if the current name is protected — never overwrite it.
    if (isProtectedKustomerName((attrs.name || '').trim())) {
      return { changed: false, label: '⏭ Protected name', color: 'var(--sb-text-muted)' };
    }

    const user = await new Promise(async (resolve) => {
      if (!emailHint && cachedCustomerCtx?.user) return resolve(cachedCustomerCtx.user);
      if (emailHint && cachedCustomerCtx?.user && cachedCustomerCtx.email === emailHint) return resolve(cachedCustomerCtx.user);
      const email = emailHint || await resolveEmail();
      if (!email) return resolve(null);
      searchCRM(email, (users, err) => {
        if (err || !users?.length) return resolve(null);
        const sbUsers = users.filter(u => !u.origin || u.origin === 'SCENTBIRD');
        const exactMatch = sbUsers.find(u => u.email?.toLowerCase() === email.toLowerCase());
        if (exactMatch) cachedCustomerCtx = { email, user: exactMatch, _kustomerId: kustomerId };
        resolve(exactMatch || null);
      });
    });

    const rawCrmName = user ? toProperCase([user.firstName, user.lastName].filter(Boolean).join(' ')) : null;
    // A placeholder is not a name: treat it as if CRM held nothing, so the casing
    // fallback below can still improve whatever Kustomer has.
    const crmName = isPlaceholderName(rawCrmName) ? null : rawCrmName;
    const currentName = (attrs.name || '').trim();

    let newName, label, color;
    if (crmName) {
      newName = crmName;                   label = '✔ Done';   color = 'var(--sb-success-soft)';
    } else if (currentName && !isPlaceholderName(currentName)) {
      newName = toProperCase(currentName); label = '✔ Cased';  color = 'var(--sb-warning)';
    } else if (rawCrmName) {
      // CRM resolved, but only to the placeholder — and Kustomer has nothing better.
      return { changed: false, label: '⏭ No real name on file', color: 'var(--sb-text-muted)' };
    } else if (!getAccessToken() || _reauthRequired) {
      return { changed: false, label: '⚠ Sign in with Okta', color: 'var(--sb-danger)' };
    } else {
      return { changed: false, label: '✘ Not Found', color: 'var(--sb-danger)' };
    }

    if (newName === attrs.name) return { changed: false, label: '✔ Already set', color: 'var(--sb-text-muted)' };

    attrs.name = newName;
    await kustomerPatchCustomer(kustomerId, attrs);
    if (nudge) { try { nudgeKustomerName(currentName, newName); } catch (_) {} }
    return { changed: true, label, color };
  }

  function registerFillName() {
    if (fillNameRegistered) return;
    fillNameRegistered = true;

    addToolbarButton('sb-fill-name-btn', 'Fill Name', async (btn) => {
      const orig = btn.textContent;
      btn.textContent = '⏳'; btn.disabled = true;
      const finish = (label, color) => {
        btn.textContent = label;
        btn.style.color = color || '';
        setTimeout(() => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; }, 2000);
      };

      const kustomerId = getCustomerIdFromURL();
      if (!kustomerId) return finish('✘ No customer', 'var(--sb-danger)');

      try {
        const { label, color } = await fillKustomerName(kustomerId, null, true);
        finish(label, color);
      } catch (e) {
        finish('✘ ' + e.message, 'var(--sb-danger)');
      }
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // CCPA — one-click privacy-request handler
  // ══════════════════════════════════════════════════════════════════════════

  /** Resolve the exact Scentbird CRM user for an email (Promise wrapper over searchCRM). */
  function resolveScentbirdUser(email) {
    return new Promise((resolve) => {
      searchCRM(email, (users, err) => {
        if (err || !users?.length) return resolve(null);
        const sbUsers = users.filter(u => !u.origin || u.origin === 'SCENTBIRD');
        resolve(sbUsers.find(u => u.email?.toLowerCase() === email.toLowerCase()) || null);
      });
    });
  }

  /** Kustomer deep link for a conversation. */
  function kustomerTicketUrl(customerId, conversationId) {
    return `https://scentbird.kustomerapp.com/app/customers/${customerId}/event/${conversationId}`;
  }

  /** Post a note on the ORIGINATING ticket pointing at the ticket CCPA just created.
   *  The mirror of the ticketLink note: that one carries the source ticket into the new
   *  conversation, this one carries the new conversation back to the source, so the
   *  privacy request and the ticket it came from are linked in both directions.
   *  Best-effort by design — the CCPA work is already done by the time it runs. */
  async function postNewTicketBacklink(originCustomerId, originConversationId, newCustomerId, newConversationId) {
    if (!originCustomerId || !originConversationId) return null;
    return postConversationNote(
      originCustomerId,
      originConversationId,
      kustomerTicketUrl(newCustomerId, newConversationId),
      'CCPA ticket created:',
    );
  }

  /** The CCPA sequence itself, shared by the toolbar button and the automatic
   *  info@scentbird.com trigger. Creates a Kustomer customer + conversation, fills its
   *  name from CRM, cancels a live subscription, cross-links the originating ticket,
   *  prefills the cancellation reply from the macro shortcut and sets the conversation
   *  categories. `onStatus` receives a short label per step so callers can attribute a
   *  mid-sequence failure. Throws on hard failure — callers must not navigate then.
   *  Returns { cid, convId, url, cancelled }. */
  async function runCcpaFlow({ email, ticketLink, originCustomerId, originConversationId, onStatus }) {
    const status = (m) => { try { onStatus?.(m); } catch (_) {} };

    status('Creating…');
    const cid = await kustomerCreateCustomer(email);

    status('Conversation…');
    const convId = await kustomerCreateConversation(cid);

    status('Name…');
    try { await fillKustomerName(cid, email, false); }
    catch (e) { console.warn('[BirdsEye] CCPA fill name failed', e); }

    status('Cancel sub…');
    const user = await resolveScentbirdUser(email);
    // Cancel whenever there's a live (cancellable) subscription — Active, Unpaid
    // or OnHold. Cancelled/absent subs and missing CRM accounts are skipped.
    const subs = user?.subscriptionList || [];
    const hasCancellableSub = CANCELLABLE_SUB_STATUSES.includes(user?.subscription?.status)
      || subs.some(s => CANCELLABLE_SUB_STATUSES.includes(s.status));
    if (hasCancellableSub) {
      await executeCancelSubscription(user, (msg) => status(msg));
    }

    // Unsubscribe from emails — part of deletion. The mutation is keyed by email, so it
    // runs for the CRM account's address when there is one, else the requester's. A
    // failure is logged, not thrown: the customer and ticket already exist by now, and
    // aborting here would strand the run half-done; the agent sees it in the console
    // and can use Cancel Sub → Unsubscribe from Emails.
    status('Unsubscribe…');
    let unsubscribed = false;
    try { await crmUnsubscribeEmail(user?.email || email); unsubscribed = true; }
    catch (e) { console.warn('[BirdsEye] Account deletion: email unsubscribe failed', e); }

    if (ticketLink) {
      status('Note…');
      await postConversationNote(cid, convId, ticketLink);
    }

    // Back-link onto the ticket this request came from. Best-effort: the privacy
    // request is already handled, and a failed cross-reference must not read as a
    // failed CCPA run or strand the agent away from the new ticket.
    if (originCustomerId && originConversationId) {
      status('Back-link…');
      try { await postNewTicketBacklink(originCustomerId, originConversationId, cid, convId); }
      catch (e) { console.warn('[BirdsEye] CCPA back-link note failed', e); }
    }

    // Two different replies. With a CRM account we confirm a cancellation; without
    // one there is nothing to confirm — the request can't be actioned until the
    // customer tells us which account they mean, which is what CCPA-loc asks. That
    // case is common now, so it gets its own macro rather than a wrong one.
    const located = !!user;
    const shortcutId = located ? getCcpaShortcutId() : getCcpaLocShortcutId();
    if (shortcutId) {
      status('Reply…');
      await prefillReply(cid, convId, shortcutId, user);
      status('Macro…');
      // Both paths get the conversation macro. These tickets are closed as part of the
      // flow, and Kustomer wants a categorised conversation when that happens — the
      // categories are bookkeeping for the close, not a claim about what was done, so
      // skipping them on the locator path just leaves closed tickets uncategorised.
      await applyCcpaConversationMacro(convId, shortcutId);
    } else {
      console.warn(`[BirdsEye] CCPA: no ${located ? 'cancellation' : 'CCPA-loc'} macro shortcut configured — skipping reply/categories`);
    }

    // Close the request this came from. The privacy tooling files it, we action it on
    // a ticket of our own, and it has nothing left to say — but it is the ORIGIN
    // ticket, so this runs last and never fails the flow: the customer-facing work is
    // already done by now and the agent is about to be navigated away.
    if (originConversationId) {
      status('Closing origin…');
      try { await closeCcpaOriginTicket(originConversationId); }
      catch (e) { console.warn('[BirdsEye] CCPA: could not close the originating ticket', e); }
    }

    return { cid, convId, url: kustomerTicketUrl(cid, convId), cancelled: hasCancellableSub, unsubscribed, located };
  }

  /** Specials ▾ — the rarely-used one-shot actions, out of the main toolbar row. */
  function registerSpecialsMenu() {
    // No module-level guard: addToolbarMenu dedupes by id, and a guard that isn't reset
    // on toolbar teardown is exactly what once made CCPA vanish after a rebuild.
    const menu = addToolbarMenu('sb-specials-menu', '★ Specials');
    if (!menu) return;
    addOvertimeItem(menu);
    addMenuItem(menu, 'sb-ccpa-btn', '🗂 Account Deletion Ticket', () => showAccountDeletionDialog());
  }

  /** Account Deletion Ticket (formerly the CCPA button + the AI panel's CCPA card).
   *  On an info@scentbird.com privacy hand-off the requester is read off the ticket and
   *  the dialog is a plain confirm; anywhere else the email is typed in. Either way
   *  the open ticket is the origin: it gets the back-link note and is closed at the end.
   *  Run is arm-then-fire — the sequence creates a Kustomer customer and cancels a live
   *  subscription. On success it lands the agent on the new ticket; on any hard error it
   *  stays put. Runs are recorded per origin ticket (GM storage, survives the navigation)
   *  so a second run warns — a re-run creates a duplicate customer. */
  async function showAccountDeletionDialog() {
    const originCustomerId = getCustomerIdFromURL();
    const originConversationId = getConversationIdFromURL();
    const { overlay, box } = createModal({ id: 'sb-ccpa-input', title: '🗂 Account Deletion Ticket', width: 360 });
    const info = el('div', { style: 'font-size:var(--sb-fs-12);color:var(--sb-text-muted);margin-bottom:10px;line-height:1.5;', text: 'Checking ticket…' });
    box.appendChild(info);

    let handoff = null;
    try { handoff = await detectCcpaHandoff(originConversationId); } catch (_) {}
    if (!document.body.contains(overlay)) return;
    const parsed = handoff?.parsed || null;

    const warn = (html) => box.appendChild(el('div', { cls: 'sb-note sb-note--warn', html, style: 'margin-bottom:8px;' }));
    if (originConversationId && ccpaDoneList().includes(originConversationId)) {
      warn('<b>⚠ Already run from this ticket.</b> Running again creates a second Kustomer customer and conversation.');
    }
    if (!getCcpaShortcutId()) warn('<b>⚠ No macro shortcut configured</b> — reply and categories will be skipped (Settings).');
    const signedOut = !getAccessToken() || _reauthRequired;
    if (signedOut) warn('<b>⚠ Not signed in to CRM</b> — sign in with Okta first.');

    // Read off the request → just confirm. Otherwise ask for the requester.
    let emailIn = null, linkIn = null;
    if (parsed) {
      info.innerHTML = '';
      info.append('Create an account deletion ticket for ',
        el('b', { style: 'color:var(--sb-text);', text: parsed.email }),
        '? This request ticket gets linked and closed.');
    } else {
      info.textContent = handoff
        ? `Filed by ${CCPA_SOURCE_EMAIL}, but no “Email:” line was found — enter the requester.`
        : 'Enter the requester. This ticket gets linked and closed when the run finishes.';
      const mkField = (labelText, ph) => {
        box.appendChild(makeLabel(labelText));
        const i = el('input', { cls: 'sb-input', attrs: { type: 'text', placeholder: ph, spellcheck: 'false' }, style: 'width:100%;box-sizing:border-box;margin-bottom:12px;' });
        box.appendChild(i);
        return i;
      };
      emailIn = mkField('Customer email', 'customer@example.com');
      linkIn  = mkField('Ticket link (posted as a note — optional)', 'https://…');
    }

    const status = makeStatusEl();
    const say = (m, c) => { status.style.color = c || ''; status.textContent = m; };
    const cancelBtn = el('button', { cls: 'sb-btn sb-btn--ghost', text: 'Cancel', on: { click: () => overlay.remove() } });
    const runBtn = el('button', { cls: 'sb-btn sb-btn--danger' });
    runBtn.disabled = signedOut;
    box.appendChild(el('div', { cls: 'sb-row', style: 'justify-content:flex-end;gap:6px;margin-top:4px;' }, [cancelBtn, runBtn]));
    box.appendChild(status);
    const inputs = [emailIn, linkIn].filter(Boolean);

    const run = async () => {
      const email = parsed ? parsed.email : emailIn.value.trim().toLowerCase();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return say('✘ Bad email', 'var(--sb-danger)');
      const ticketLink = parsed ? kustomerTicketUrl(originCustomerId, originConversationId) : linkIn.value.trim();
      runBtn.disabled = true; cancelBtn.disabled = true; inputs.forEach(i => { i.disabled = true; });
      // Marked before the sequence starts: it navigates away on success, and a crash
      // part-way must not read as "never ran" — the customer may already exist.
      if (originConversationId) markCcpaDone(originConversationId);
      try {
        const { url, unsubscribed } = await runCcpaFlow({ email, ticketLink, originCustomerId, originConversationId, onStatus: (m) => say('⏳ ' + m) });
        if (!unsubscribed) {
          // Everything else is done; don't navigate past the one step that wasn't.
          say('⚠ Ticket created, but the email unsubscribe failed — do it from Cancel Sub.', 'var(--sb-warning-strong)');
          cancelBtn.disabled = false; cancelBtn.textContent = 'Close';
          runBtn.replaceWith(el('button', { cls: 'sb-btn sb-btn--primary', text: 'Open ticket', on: { click: () => { location.href = url; } } }));
          return;
        }
        say('✔ Ticket created — opening it…', 'var(--sb-success-soft)');
        location.href = url;
      } catch (e) {
        console.warn('[BirdsEye] Account deletion failed', e);
        say('✘ ' + (e.message || 'failed'), 'var(--sb-danger)');
        runBtn.disabled = false; cancelBtn.disabled = false; inputs.forEach(i => { i.disabled = false; });
        runBtn.textContent = 'Retry';
      }
    };

    if (parsed) {
      // The prompt itself is the confirmation — one click runs.
      runBtn.textContent = 'Create ticket';
      runBtn.addEventListener('click', run);
      setTimeout(() => runBtn.focus(), 30);
    } else {
      // Typed in by hand: keep arm-then-fire so a stray Enter can't run it.
      armedButton(runBtn, { idle: 'Run', armed: '⚠ Click again to run', onFire: run });
      inputs.forEach(i => i.addEventListener('keydown', e => { if (e.key === 'Enter') runBtn.click(); }));
      setTimeout(() => emailIn.focus(), 30);
    }
  }

  // ── Overtime mode ─────────────────────────────────────────────────────────
  // One click stamps the open ticket as overtime work: an internal note
  // "OT - DD_MM_YYYY" and a tag OT_ASSIGNED_DD_MM (local date). The tag is dated, so
  // a new one is needed every day — Kustomer's tag picker handles that by PATCHing
  // the conversation's tag list with the new tag's *name* in place of an id, which
  // creates it on the fly. An existing tag must go by id, so look it up first.

  function _otDateParts(d = new Date()) {
    const p = (n) => String(n).padStart(2, '0');
    return { dd: p(d.getDate()), mm: p(d.getMonth() + 1), yyyy: String(d.getFullYear()) };
  }

  /** Id of the tag named exactly `name` (case-insensitive), or null if none exists. */
  async function kustomerFindTagId(name) {
    const reqId = _aiUuid();
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/search/platform/tags?pageSize=100&page=1&client-request-id=${reqId}`,
      {
        method: 'POST',
        credentials: 'include',
        headers: { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId },
        body: JSON.stringify({ and: [{ tag_name: { contains: name } }], queryContext: 'tag', includeDeleted: false }),
      },
    );
    if (!res.ok) throw new Error('Kustomer tag search ' + res.status);
    const want = name.toLowerCase();
    const hit = ((await res.json())?.data || []).find(t => (t.attributes?.name || '').toLowerCase() === want);
    return hit?.id || null;
  }

  /** Add a tag (by name) to a conversation. Returns false if it was already there. */
  async function kustomerAddConversationTag(conversationId, name) {
    const res = await fetch(`${KUSTOMER_API_BASE}/v1/conversations/${conversationId}`, { credentials: 'include', headers: _aiHeaders() });
    if (!res.ok) throw new Error('Kustomer conversation GET ' + res.status);
    const current = (await res.json())?.data?.attributes?.tags || [];
    const id = await kustomerFindTagId(name);
    if (id && current.includes(id)) return false;
    await patchConversation(conversationId, { tags: [...current, id || name] });
    return true;
  }

  /** Post a plain-text internal note (one DraftJS block per line). Same create → fill →
   *  send sequence as postConversationNote. `body` is Kustomer's markdown, so markdown
   *  punctuation is escaped the way the composer does it ('OT - x' goes out as 'OT \- x'). */
  async function postPlainNote(customerId, conversationId, text) {
    const body = text.replace(/[\\`*_[\]#+\-!>~|]/g, '\\$&');
    const blocks = text.split('\n').map(line => ({ key: _draftBlockKey(), text: line, type: 'unstyled', depth: 0, inlineStyleRanges: [], entityRanges: [], data: {} }));
    const draftId = await kustomerCreateDraft(customerId, 'note', conversationId);
    await kustomerUpdateDraft(draftId, {
      body,
      channel: 'note',
      payload: { draftJs: { blocks, entityMap: {} } },
    });
    await kustomerUpdateDraft(draftId, {
      body,
      channel: 'note',
      userMentions: [],
      conversation: conversationId,
      sendAt: new Date().toISOString(),
    });
    return draftId;
  }

  function addOvertimeItem(menu) {
    const created = addMenuItem(menu, 'sb-ot-btn', '🕒 OT', async (btn) => {
      const orig = btn.textContent;
      const finish = (label, color) => {
        btn.textContent = label;
        btn.style.color = color || '';
        setTimeout(() => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; }, 2500);
      };

      const customerId = getCustomerIdFromURL();
      const conversationId = getConversationIdFromURL();
      if (!customerId || !conversationId) return finish('✘ No ticket', 'var(--sb-danger)');

      const { dd, mm, yyyy } = _otDateParts();
      btn.disabled = true;
      btn.textContent = '⏳';
      try {
        // Tag first: if today's tag is already on the ticket, OT was already stamped —
        // skip the note rather than post a duplicate.
        const added = await kustomerAddConversationTag(conversationId, `OT_ASSIGNED_${dd}_${mm}`);
        if (!added) return finish('✔ Already OT', 'var(--sb-warning)');
        await postPlainNote(customerId, conversationId, `OT - ${dd}_${mm}_${yyyy}`);
        finish('✔ OT', 'var(--sb-success)');
      } catch (e) {
        console.warn('[BirdsEye] OT failed', e);
        finish('✘ ' + (e.message || 'failed'), 'var(--sb-danger)');
      }
    });
    if (created) created.title = 'Overtime: note "OT - DD_MM_YYYY" + tag OT_ASSIGNED_DD_MM';
  }

  // ── Snooze note ───────────────────────────────────────────────────────────
  // Two entry points, one form: "What are you waiting on?" → internal note
  // "Waiting on <thing>".
  //
  // 1. Kustomer's snooze button (svg[data-icon=clock-snooze]) opens a popover of
  //    options. Once one is picked and the popover closes, the form goes up and the
  //    note posts. Escape / click-away (nothing picked) shows nothing.
  // 2. The composer's ⋯ menu "Send & snooze 1 day" ([data-kt=popoverMenuItem-snooze])
  //    is held before it acts (it fires on press — every pointer phase is swallowed on
  //    window, capture, ahead of React). "Post & send" then does the whole job through
  //    the API, in this order: send the reply (sendConversationReply, the same path as
  //    the AI panel's Send + Pending), snooze 1 day, post the note — so the note lands
  //    after the message. Kustomer's own auto-advance doesn't run; the ticket stays open.
  //
  // Ticket ids are captured when the action starts, so moving on can't post the note
  // onto the next ticket.
  const SNOOZE_POPOVER_SEL = '[class*="snoozePopoverKomponent__popover"]';
  const SEND_SNOOZE_SEL = '[data-kt="popoverMenuItem-snooze"]';
  const SEND_SNOOZE_MS = 24 * 60 * 60 * 1000;
  const WAITING_ON = ['images', 'address confirmation', 'fragrance confirmation'];
  let _snoozePick = null;   // { customerId, conversationId } once a popover option is clicked

  document.addEventListener('click', (e) => {
    const pop = e.target.closest?.(SNOOZE_POPOVER_SEL);
    if (!pop || e.target.closest('[class*="snoozePopoverKomponent__header"]')) return;
    const customerId = getCustomerIdFromURL(), conversationId = getConversationIdFromURL();
    if (!customerId || !conversationId || _snoozePick) return;
    _snoozePick = { customerId, conversationId };
    const deadline = Date.now() + 120000;   // Pick Date & Time can take a while
    const wait = setInterval(() => {
      if (document.querySelector(SNOOZE_POPOVER_SEL) && Date.now() < deadline) return;
      clearInterval(wait);
      const pick = _snoozePick; _snoozePick = null;
      if (Date.now() < deadline) showWaitingOnPrompt(pick);
    }, 300);
  }, true);

  ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(type => window.addEventListener(type, (e) => {
    if (!e.isTrusted || !e.target.closest?.(SEND_SNOOZE_SEL)) return;
    e.preventDefault(); e.stopImmediatePropagation();
    if (type !== 'pointerdown' || document.getElementById('sb-snooze-note')) return;
    const customerId = getCustomerIdFromURL(), conversationId = getConversationIdFromURL();
    if (customerId && conversationId) showWaitingOnPrompt({ customerId, conversationId }, { sendFirst: true });
  }, true));

  /** "Waiting on …" form. `sendFirst`: send the composer reply and snooze 1 day before
   *  the note (Send & snooze); otherwise just post the note (after a native snooze). */
  function showWaitingOnPrompt({ customerId, conversationId }, { sendFirst = false } = {}) {
    const id = 'sb-snooze-note';
    const panel = createPanel({ id, title: sendFirst ? '💤 Send & snooze 1 day' : '💤 Snooze note', width: 340, centered: true, draggable: true });

    panel.appendChild(makeLabel('WHAT ARE YOU WAITING ON?'));
    const sel = el('select', { cls: 'sb-select', style: 'width:100%;box-sizing:border-box;' });
    [...WAITING_ON, 'other'].forEach(v => sel.appendChild(el('option', {
      text: v === 'other' ? 'Other (specify)' : v[0].toUpperCase() + v.slice(1), attrs: { value: v } })));
    const other = el('input', { cls: 'sb-input', attrs: { placeholder: 'Waiting on…' },
      style: 'display:none;width:100%;box-sizing:border-box;margin-top:6px;' });
    sel.addEventListener('change', () => { other.style.display = sel.value === 'other' ? '' : 'none'; if (sel.value === 'other') other.focus(); });
    const status = makeStatusEl();
    const cancel = el('button', { cls: 'sb-btn sb-btn--ghost', text: sendFirst ? 'Cancel' : 'Skip', on: { click: () => removePanel(id) } });
    const go = el('button', { cls: 'sb-btn sb-btn--primary', text: sendFirst ? 'Post & send' : 'Post note' });

    const noteText = () => {
      const what = sel.value === 'other' ? other.value.trim() : sel.value;
      return what ? 'Waiting on ' + what : '';
    };
    const say = (msg, color) => { status.style.color = color || ''; status.textContent = msg; };

    let sent = false, snoozed = false;
    const run = async () => {
      const text = noteText();
      if (!text) { say('✘ Say what you are waiting on', 'var(--sb-danger)'); other.focus(); return; }
      go.disabled = true; cancel.disabled = true; sel.disabled = true; other.disabled = true;
      go.textContent = '⏳';
      try {
        if (sendFirst && !sent) {
          const r = await sendConversationReply({ customerId, conversationId, onStatus: (m) => say(m) });
          sent = true;
          clearComposerAfterSend({ conversationId, draftId: r.draftId, sentText: r.composerText }).catch(() => {});
        }
        if (sendFirst && !snoozed) {
          say('Snoozing 1 day…');
          await patchConversation(conversationId, {
            status: 'snoozed',
            snooze: { time: new Date(Date.now() + SEND_SNOOZE_MS).toISOString(), status: 'scheduled' },
          });
          snoozed = true;
        }
        say('Posting note…');
        await postPlainNote(customerId, conversationId, text);
        removePanel(id);
      } catch (err) {
        console.warn('[BirdsEye] waiting-on flow failed', err);
        // A retry resumes from the failed step — never re-sends a reply that went out.
        const where = !sendFirst ? '' : !sent ? 'Not sent: ' : !snoozed ? 'Sent, snooze failed: ' : 'Sent & snoozed, note failed: ';
        say('✘ ' + where + (err.message || 'failed'), 'var(--sb-danger)');
        go.disabled = false; sel.disabled = false; other.disabled = false;
        cancel.disabled = false;
        if (sent) cancel.textContent = 'Close';
        go.textContent = sent ? 'Retry' : (sendFirst ? 'Post & send' : 'Post note');
      }
    };
    go.addEventListener('click', run);
    panel.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey || e.target === other)) { e.preventDefault(); if (!go.disabled) run(); }
      if (e.key === 'Escape' && !cancel.disabled) removePanel(id);
    });

    panel.appendChild(sel);
    panel.appendChild(other);
    panel.appendChild(el('div', { cls: 'sb-row', style: 'justify-content:flex-end;gap:6px;margin-top:10px;' }, [cancel, go]));
    panel.appendChild(status);
    setTimeout(() => sel.focus(), 0);
  }

  // ── CCPA hand-offs from info@scentbird.com ────────────────────────────────
  // The privacy tooling files these as a ticket from info@scentbird.com whose first
  // message carries the requester's identifiers:
  //     Requested by: User in Web App
  //     Email: ssjahna@gmail.com
  //     User ID: 12714325
  // Those tickets are always the same job, so the whole sequence can be automated —
  // but it is NOT automatic. It used to fire unattended off the MutationObserver the
  // moment an agent opened one of these tickets, which meant a customer was created
  // and a subscription cancelled with no one having pressed anything. The detection
  // now lives in the AI panel (see _aiRenderCcpaAction): the panel recognises the
  // ticket type and offers the automation as the flow, the agent approves it.
  //
  // Conversations that have been run are still recorded in GM storage: the flow
  // navigates away to the ticket it just created, so an in-memory mark would be lost
  // exactly when it matters, and a re-run creates a duplicate customer. The record is
  // advisory now rather than a silent skip — the panel warns and lets the agent decide.

  const CCPA_SOURCE_EMAIL = 'info@scentbird.com';
  const CCPA_DONE_KEY = 'birdseye_ccpa_auto_done';
  const CCPA_DONE_MAX = 200;

  function ccpaDoneList() {
    try { const v = JSON.parse(GM_getValue(CCPA_DONE_KEY, '[]')); return Array.isArray(v) ? v : []; }
    catch (_) { return []; }
  }
  function markCcpaDone(conversationId) {
    const list = ccpaDoneList().filter(id => id !== conversationId);
    list.push(conversationId);
    try { GM_setValue(CCPA_DONE_KEY, JSON.stringify(list.slice(-CCPA_DONE_MAX))); } catch (_) {}
  }

  /** Pull the requester's email (and user id, when present) out of a hand-off message.
   *  Returns null when the message isn't one — the caller then leaves the ticket alone. */
  function parseCcpaRequestMessage(raw) {
    const text = String(raw || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(?:p|div|tr|li)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&');
    const email = (text.match(/Email:\s*([^\s<>()[\],;:]+@[^\s<>()[\],;:]+)/i) || [])[1];
    if (!email) return null;
    const userId = (text.match(/User\s*ID:\s*(\d+)/i) || [])[1] || null;
    return { email: email.trim().toLowerCase(), userId };
  }

  /** Is the open ticket a privacy-tooling CCPA hand-off? Detection only — nothing
   *  runs off the back of this, the AI panel renders it as an approvable flow.
   *  Returns null when the ticket isn't from info@scentbird.com at all; otherwise
   *  `{ parsed }`, where `parsed` is null for a hand-off whose messages carry no
   *  "Email:" line (the panel then sends the agent to the manual button). */
  async function detectCcpaHandoff(conversationId) {
    if (!conversationId) return null;

    let email = null;
    try { email = await resolveEmail(); } catch (_) { return null; }
    if (email !== CCPA_SOURCE_EMAIL) return null;

    let parsed = null;
    try {
      const messages = await kustomerGetConversationMessagesRaw(conversationId, 10);
      for (const m of messages) {
        parsed = parseCcpaRequestMessage(m.body || m.preview);
        if (parsed) break;
      }
    } catch (e) {
      console.warn('[BirdsEye] CCPA: could not read messages', e);
    }
    return { parsed };
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
        return [s.street1, s.street2, s.city, s.region, s.postcode].filter(Boolean).join(', ') || null;
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
        // Raw chargeList entries ({ date, kind, charge }) — there is no `success` flag
        // on them, which is why this used to filter everything out and never fill.
        const successEntries = charges.filter(e => e.charge?.status === 'SUCCESS');
        if (!successEntries.length) return null;
        return successEntries.map(e => buildBillExplanation(e)).filter(Boolean).join('\n\n') || null;
      },
    },
  ];

  function getComposer() {
    return document.querySelector('.public-DraftEditor-content[contenteditable="true"]');
  }

  // Draft.js splits the composer's text across <span data-offset-key> nodes, and it
  // splits exactly where inline styling or a link begins — so a macro that wraps a token
  // in a link serializes as "[TRACK</span><span…>ING]". Substituting over raw innerHTML
  // therefore misses the token while innerText still shows it as present, which is how a
  // literal [TRACKING] ended up pasted into DNR replies while the plain-sentence macros
  // (cant-ref-shipped) filled fine. Match against the text NODES instead, edit them on a
  // detached clone, and leave behind a contiguous sentinel: it survives serialization in
  // one piece, so the (possibly HTML) value can be spliced in afterwards.
  const _FILL_SENTINEL = (i) => `⟦BEFILL${i}⟧`;

  /** Fill every [TOKEN] in `rootEl`, tolerating tokens split across element boundaries.
   *  Returns { html, count } — count is how many distinct tokens matched — or null when
   *  nothing matched at all. Does not mutate `rootEl`. */
  function _fillTokensInHtml(rootEl, replacements) {
    const root = rootEl.cloneNode(true);
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodes = [];
    let full = '';
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      // `len` is captured up front: the edits below change node text, and every offset
      // here is in ORIGINAL coordinates.
      nodes.push({ node: n, start: full.length, len: n.nodeValue.length });
      full += n.nodeValue;
    }
    if (!nodes.length) return null;

    const hits = [];
    replacements.forEach((r, i) => {
      const token = '[' + r.key + ']';
      for (let at = full.indexOf(token); at !== -1; at = full.indexOf(token, at + token.length)) {
        hits.push({ at, end: at + token.length, i });
      }
    });
    if (!hits.length) return null;

    // Queue edits per node, then apply each node's edits right-to-left so the
    // still-unapplied original coordinates stay valid.
    const edits = new Map();
    for (const h of hits) {
      let isFirstNode = true;
      for (const nd of nodes) {
        const nStart = nd.start, nEnd = nd.start + nd.len;
        if (nEnd <= h.at || nStart >= h.end) continue;
        const list = edits.get(nd.node) || [];
        list.push({
          from:   Math.max(h.at, nStart) - nStart,
          to:     Math.min(h.end, nEnd) - nStart,
          // The sentinel goes in the node holding the token's START; the tail nodes
          // just lose their slice of it.
          insert: isFirstNode ? _FILL_SENTINEL(h.i) : '',
        });
        edits.set(nd.node, list);
        isFirstNode = false;
      }
    }
    for (const [node, list] of edits) {
      list.sort((a, b) => b.from - a.from);
      let v = node.nodeValue;
      for (const e of list) v = v.slice(0, e.from) + e.insert + v.slice(e.to);
      node.nodeValue = v;
    }

    let html = root.innerHTML;
    replacements.forEach((r, i) => {
      const str = (r.value && typeof r.value === 'object') ? r.value.text : String(r.value);
      const htmlVal = str.split('\n').map((line, k) =>
        k === 0 ? line : '<div>' + (line || '<br>') + '</div>'
      ).join('');
      html = html.split(_FILL_SENTINEL(i)).join(htmlVal);
    });
    return { html, count: new Set(hits.map(h => h.i)).size };
  }

  function fillComposer(replacements) {
    const composer = getComposer();
    if (!composer) return 0;
    const res = _fillTokensInHtml(composer, replacements);
    if (!res) return 0;
    GM_setClipboard(res.html, 'html');
    return res.count;
  }

  // Fill "pick an order" state: set while the ✏️ Fill button waits for the agent
  // to click an order in the Last Orders panel. Null = normal panel (unchanged).
  let _fillPickState = null;

  function _fillPickEsc(e) { if (e.key === 'Escape') cancelFillPick(); }

  function enterFillPick(needed, btn) {
    _fillPickState = { needed, btn };
    document.addEventListener('keydown', _fillPickEsc, true);
    // Open the panel if closed; otherwise re-render it into pick mode in place.
    if (document.getElementById('sb-order-panel')) refreshOpenPanels('sb-order-panel');
    else showLastOrderPanel(btn);
  }

  function cancelFillPick() {
    if (!_fillPickState) return;
    _fillPickState = null;
    document.removeEventListener('keydown', _fillPickEsc, true);
    refreshOpenPanels('sb-order-panel');
  }

  function resolveFillPick(order) {
    if (!_fillPickState) return;
    const { needed, btn } = _fillPickState;
    _fillPickState = null;
    document.removeEventListener('keydown', _fillPickEsc, true);
    refreshOpenPanels('sb-order-panel');
    doFill(btn, needed, order);
  }

  /** In pick mode, wrap an order block so the whole card is a click target with a
   *  hover border; otherwise return it unchanged. */
  function pickWrap(orderEl, order) {
    if (!_fillPickState) return orderEl;
    const wrap = document.createElement('div');
    wrap.style.cssText = 'position:relative;border-radius:8px;cursor:pointer;';
    wrap.appendChild(orderEl);
    const ov = document.createElement('div');
    ov.style.cssText = 'position:absolute;inset:0;border-radius:8px;border:2px solid transparent;';
    ov.addEventListener('mouseenter', () => ov.style.borderColor = 'var(--sb-accent)');
    ov.addEventListener('mouseleave', () => ov.style.borderColor = 'transparent');
    ov.addEventListener('click', () => resolveFillPick(order));
    wrap.appendChild(ov);
    return wrap;
  }

  function fillVariables(btn) {
    const composer = getComposer();
    if (!composer) {
      const orig = btn.textContent;
      btn.textContent = '✘ No editor'; btn.style.color = 'var(--sb-danger)';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; }, 2000);
      return;
    }

    const text = composer.innerText;
    const needed = FILL_VARIABLES.filter(v => text.includes('[' + v.key + ']'));
    if (!needed.length) {
      const orig = btn.textContent;
      btn.textContent = '✘ No tokens'; btn.style.color = 'var(--sb-danger)';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; }, 2000);
      return;
    }

    // Order-dependent tokens need a specific order. Instead of a cramped dropdown,
    // let the agent pick it directly in the Last Orders panel.
    const needsOrder = needed.some(v => ['MONTH','DATE','FRAGRANCE','TRACKING'].includes(v.key));
    if (!needsOrder) { doFill(btn, needed, null); return; }
    enterFillPick(needed, btn);
  }

  // Tokens whose values come off the subscription record rather than the order.
  const SUBSCRIPTION_FILL_KEYS = ['BD', 'BD-1', 'DATE_SUBSCRIBED', 'DATE_CANCELLED'];

  /** Fresh subscription for the date tokens. Always refetched, never read off
   *  cachedCustomerCtx: the cached user predates whatever CRM op the flow just ran
   *  (a cancellation sets subscriptionEndDate, which [DATE_CANCELLED] needs), and the
   *  lean user.subscription from CRM_QUERY carries no billing or end dates at all.
   *  Prefers the active sub, else the newest entry. Resolves null on any failure. */
  function fetchFillSubscription(userId) {
    return new Promise((resolve) => {
      if (!userId) return resolve(null);
      crmRequest({
        method: 'POST', url: GRAPHQL_URL,
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({ operationName: 'userDetailsById', query: USER_DETAILS_QUERY, variables: { id: userId } }),
        onload(res) {
          try {
            const list = JSON.parse(res.responseText)?.data?.userById?.data?.subscriptionList || [];
            resolve(list.find(s => s.status === 'Active' || s.subscribed) || list[0] || null);
          } catch (e) { resolve(null); }
        },
        onerror() { resolve(null); },
      });
    });
  }

  function doFill(btn, needed, pickedOrder) {
    const orig = btn.textContent;
    btn.textContent = '⏳'; btn.disabled = true;
    const restore = (label, color) => {
      btn.textContent = label; btn.style.color = color || '';
      setTimeout(() => { btn.textContent = orig; btn.style.color = ''; btn.disabled = false; }, 2000);
    };

    const needsOrder        = needed.some(v => ['MONTH','DATE','FRAGRANCE','TRACKING'].includes(v.key));
    const needsSubscription = needed.some(v => SUBSCRIPTION_FILL_KEYS.includes(v.key));
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
            if (count === 0 && skipped > 0) return restore('✘ No data', 'var(--sb-danger)');
            const label = (skipped > 0 ? '⚠ Paste (missing)' : '📋 Paste now');
            restore(label, 'var(--sb-success-soft)');
          };

          if (needsCharges) {
            loadCustomerCharges(user.id, (chargeResult) => {
              // fetchCharges hands back { success, all, failed } — never `entries`.
              // `all` is the recent successful cluster, as the Charges panel reads it.
              afterCharges(chargeResult?.all || (chargeResult?.success ? [chargeResult.success] : []));
            });
          } else {
            afterCharges(null);
          }
        };

        if (needsSubscription) {
          fetchFillSubscription(user.id).then(afterSubscription);
        } else {
          afterSubscription(null);
        }
      };

      if (needsOrder) {
        // pickedOrder comes straight from the Last Orders panel; fall back to the
        // most recent main order if somehow absent.
        if (pickedOrder) {
          afterOrders(pickedOrder);
        } else {
          loadCustomerOrders(user.id, (orders) => {
            const mainOrders = (orders || []).filter(o => o.type !== 'REPLACEMENT' && o.type !== 'REPLACEMENT_ORDER');
            afterOrders(mainOrders[0] || null);
          });
        }
      } else {
        afterOrders(null);
      }
    });
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
    addToolbarButton('sb-downgrade-btn', '⬇ Downgrade', (btn) => {
      showDowngradePanel(btn);
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
        const label = !getAccessToken() ? '⚠ Sign in with Okta' : '✘ No customer';
        btn.textContent = label;
        btn.style.color = 'var(--sb-danger)';
        setTimeout(() => { btn.textContent = '🔗 Open Account'; btn.style.color = ''; }, 2000);
        return;
      }
      window.open(`https://crm.scentbird.com/user/${user.id}/profile/subscription`, '_blank');
    });
    addToolbarButton('sb-crm-token-btn', '🔐 Sign in with Okta', () => {
      if (_captchaRequired) openCaptchaPopup();
      else startOktaLogin();
    });
    refreshOktaButtonState();
    addToolbarButton('sb-settings-btn', '⚙', () => {
      if (document.getElementById('sb-settings-panel')) { removePanel('sb-settings-panel'); return; }
      showSettingsPanel();
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SETTINGS PANEL
  // ══════════════════════════════════════════════════════════════════════════

  /** Dark/light segmented toggle. */
  function _settingsThemeSection(panel, s) {
    // ── Theme toggle (segmented) ────────────────────────────────────────────
    const themeLabel = document.createElement('div');
    themeLabel.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-11);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;';
    themeLabel.textContent = 'Theme';
    panel.appendChild(themeLabel);

    const themeRow = document.createElement('div');
    themeRow.style.cssText = 'display:flex;gap:0;border:1px solid var(--sb-border);border-radius:6px;overflow:hidden;margin-bottom:14px;';
    const makeThemeBtn = (key, label) => {
      const b = document.createElement('button');
      b.textContent = label;
      b.dataset.theme = key;
      b.style.cssText = 'flex:1;padding:7px;border:none;background:transparent;color:var(--sb-text);font-size:var(--sb-fs-12);font-weight:600;cursor:pointer;';
      b.onclick = () => { setSettings({ theme: key }); paintTheme(); };
      return b;
    };
    const darkBtn  = makeThemeBtn('dark',  '🌙 Dark');
    const lightBtn = makeThemeBtn('light', '☀ Light');
    themeRow.appendChild(darkBtn);
    themeRow.appendChild(lightBtn);
    panel.appendChild(themeRow);

    const paintTheme = () => {
      const cur = getSettings().theme;
      [darkBtn, lightBtn].forEach(b => {
        const active = b.dataset.theme === cur;
        b.style.background = active ? 'var(--sb-control-bg)' : 'transparent';
        b.style.color      = active ? 'var(--sb-text)'       : 'var(--sb-text-muted)';
      });
    };
    paintTheme();

  }

  /** Font and modal scale sliders, plus the range-row
   *  helper they share. */
  function _settingsScaleSection(panel, s) {
    // ── Range slider helper ─────────────────────────────────────────────────
    const makeRangeRow = (key, label, min, max, step, current, onInput) => {
      const wrap = document.createElement('div');
      wrap.style.cssText = 'margin-bottom:14px;';
      const lbl = document.createElement('div');
      lbl.style.cssText = 'display:flex;justify-content:space-between;color:var(--sb-text-muted);font-size:var(--sb-fs-11);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;';
      const lblName = document.createElement('span');
      lblName.textContent = label;
      const lblVal = document.createElement('span');
      lblVal.textContent = current.toFixed(2) + '×';
      lblVal.style.color = 'var(--sb-text)';
      lbl.appendChild(lblName); lbl.appendChild(lblVal);
      wrap.appendChild(lbl);

      const range = document.createElement('input');
      range.type  = 'range';
      range.min   = String(min);
      range.max   = String(max);
      range.step  = String(step);
      range.value = String(current);
      range.style.cssText = 'width:100%;accent-color:var(--sb-solid-primary-lo);';
      range.oninput = () => {
        const v = parseFloat(range.value);
        lblVal.textContent = v.toFixed(2) + '×';
        onInput(v);
      };
      wrap.appendChild(range);
      return wrap;
    };

    // ── Font scale ──────────────────────────────────────────────────────────
    panel.appendChild(makeRangeRow('fontScale', 'Font size', 0.85, 1.4, 0.05, s.fontScale, (v) => {
      setSettings({ fontScale: v });
    }));

    // ── Modal scale ─────────────────────────────────────────────────────────
    panel.appendChild(makeRangeRow('modalScale', 'Window size', 0.9, 1.3, 0.05, s.modalScale, (v) => {
      setSettings({ modalScale: v });
    }));

  }

  /** Close-delay and hide-tax toggles. */
  function _settingsBehaviourSection(panel, s) {
    // ── Close delay toggle ──────────────────────────────────────────────────
    const delayRow = document.createElement('div');
    delayRow.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;';
    const delayLabel = document.createElement('div');
    delayLabel.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-11);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;';
    delayLabel.textContent = 'Delay window close on success';
    const delayToggle = document.createElement('button');
    const paintDelay = () => {
      const on = getSettings().closeDelay;
      delayToggle.textContent = on ? 'On' : 'Off';
      delayToggle.style.cssText = `padding:4px 14px;border-radius:20px;border:none;font-size:var(--sb-fs-12);font-weight:700;cursor:pointer;
        background:${on ? 'var(--sb-solid-primary-lo)' : 'var(--sb-control-bg)'};color:${on ? 'var(--sb-on-solid)' : 'var(--sb-text-muted)'};`;
    };
    paintDelay();
    delayToggle.onclick = () => { setSettings({ closeDelay: !getSettings().closeDelay }); paintDelay(); };
    delayRow.appendChild(delayLabel);
    delayRow.appendChild(delayToggle);
    panel.appendChild(delayRow);

    // ── Hide tax toggle ─────────────────────────────────────────────────────
    const taxRow = document.createElement('div');
    taxRow.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;';
    const taxLabel = document.createElement('div');
    taxLabel.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-11);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;';
    taxLabel.textContent = 'Hide tax in bill explanations';
    const taxToggle = document.createElement('button');
    const paintTax = () => {
      const on = getSettings().hideTax;
      taxToggle.textContent = on ? 'On' : 'Off';
      taxToggle.style.cssText = `padding:4px 14px;border-radius:20px;border:none;font-size:var(--sb-fs-12);font-weight:700;cursor:pointer;
        background:${on ? 'var(--sb-solid-primary-lo)' : 'var(--sb-control-bg)'};color:${on ? 'var(--sb-on-solid)' : 'var(--sb-text-muted)'};`;
    };
    paintTax();
    taxToggle.onclick = () => { setSettings({ hideTax: !getSettings().hideTax }); paintTax(); };
    taxRow.appendChild(taxLabel);
    taxRow.appendChild(taxToggle);
    panel.appendChild(taxRow);

  }

  /** The two CCPA macro shortcut links. */
  function _settingsCcpaSection(panel, s) {
    // ── CCPA macro shortcuts ────────────────────────────────────────────────
    const ccpaLabel = document.createElement('div');
    ccpaLabel.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-11);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;';
    ccpaLabel.textContent = 'CCPA macro shortcut links';
    panel.appendChild(ccpaLabel);

    /** Shortcut-link field with a live "Shortcut id: …" readout. `key` is the
     *  settings key it writes. `required` decides whether a blank value is an error
     *  or just an unconfigured optional path. */
    const addShortcutField = (key, caption, required) => {
      const cap = document.createElement('div');
      cap.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-dim);margin-bottom:3px;';
      cap.textContent = caption;
      panel.appendChild(cap);

      const inp = document.createElement('input');
      inp.type = 'text';
      inp.value = s[key] || '';
      inp.placeholder = 'https://…/settings/shortcuts/<id>/edit';
      inp.spellcheck = false;
      inp.style.cssText = 'width:100%;box-sizing:border-box;padding:7px 10px;margin-bottom:4px;border-radius:6px;border:1px solid var(--sb-border);background:var(--sb-bg-input);color:var(--sb-text);font-size:var(--sb-fs-12);';
      const hint = document.createElement('div');
      hint.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-dim);margin-bottom:10px;';
      const paint = () => {
        const id = parseShortcutId(inp.value);
        if (id) { hint.textContent = 'Shortcut id: ' + id; hint.style.color = 'var(--sb-text-dim)'; return; }
        const blank = !inp.value.trim();
        hint.textContent = blank && !required ? 'Not configured — this reply will be skipped.' : '⚠ No shortcut id found in link';
        hint.style.color = blank && !required ? 'var(--sb-text-dim)' : 'var(--sb-danger)';
      };
      paint();
      inp.oninput = () => { setSettings({ [key]: inp.value.trim() }); paint(); };
      panel.appendChild(inp);
      panel.appendChild(hint);
    };

    addShortcutField('ccpaShortcut',    'Cancellation reply (account found)', true);
    addShortcutField('ccpaLocShortcut', 'CCPA-loc — no CRM account found',    false);
    const ccpaSpacer = document.createElement('div');
    ccpaSpacer.style.cssText = 'height:4px;';
    panel.appendChild(ccpaSpacer);

  }

  /** Entry point to the bill-explanation templates panel. */
  function _settingsBillSection(panel, s) {
    // ── Bill Explanations ───────────────────────────────────────────────────
    const billBtn = document.createElement('button');
    billBtn.textContent = '✎ Edit Bill Explanations';
    billBtn.style.cssText = `
      width:100%;padding:8px;border-radius:6px;border:1px solid var(--sb-border);
      background:transparent;color:var(--sb-text-muted);font-size:var(--sb-fs-12);
      font-weight:600;cursor:pointer;text-align:left;margin-bottom:14px;
    `;
    billBtn.onmouseenter = () => billBtn.style.background = 'var(--sb-control-bg)';
    billBtn.onmouseleave = () => billBtn.style.background = 'transparent';
    billBtn.onclick = () => showBillTemplatesPanel();
    panel.appendChild(billBtn);

  }

  /** Reset to defaults / Done. */
  function _settingsFooterSection(panel, s) {
    // ── Reset / Done ────────────────────────────────────────────────────────
    const btnRow = document.createElement('div');
    btnRow.style.cssText = 'display:flex;gap:8px;margin-top:6px;';

    const resetBtn = document.createElement('button');
    resetBtn.textContent = 'Reset';
    resetBtn.style.cssText = 'padding:7px 12px;border-radius:6px;border:1px solid var(--sb-border);background:transparent;color:var(--sb-text-muted);font-size:var(--sb-fs-12);font-weight:600;cursor:pointer;';
    resetBtn.onclick = () => {
      setSettings(SETTINGS_DEFAULTS);
      removePanel('sb-settings-panel');
      showSettingsPanel();
    };
    btnRow.appendChild(resetBtn);

    const spacer = document.createElement('div');
    spacer.style.cssText = 'flex:1;';
    btnRow.appendChild(spacer);

    const doneBtn = document.createElement('button');
    doneBtn.textContent = 'Done';
    doneBtn.style.cssText = 'padding:7px 16px;border-radius:6px;border:none;background:var(--sb-solid-primary-lo);color:var(--sb-on-solid);font-size:var(--sb-fs-12);font-weight:700;cursor:pointer;';
    doneBtn.onclick = () => removePanel('sb-settings-panel');
    btnRow.appendChild(doneBtn);

    panel.appendChild(btnRow);
  }

  /** Settings panel. Eight independent sections, each reading and writing its own
   *  slice of getSettings()/setSettings — none of them shares state with another,
   *  which is what made the 288-line original a pure concatenation. */
  function showSettingsPanel() {
    const panel = createPanel({
      id: 'sb-settings-panel',
      title: '⚙ Settings',
      width: 320,
      maxHeight: 480,
      centered: true,
      noScale: true,
    });

    const s = getSettings();

    _settingsThemeSection(panel, s);
    _settingsScaleSection(panel, s);
    _settingsBehaviourSection(panel, s);
    _settingsCcpaSection(panel, s);
    _settingsBillSection(panel, s);
    _settingsFooterSection(panel, s);
  }

  function showBillTemplatesPanel() {
    if (document.getElementById('sb-bill-templates-panel')) { removePanel('sb-bill-templates-panel'); return; }
    const panel = createPanel({
      id: 'sb-bill-templates-panel',
      title: '✎ Bill Explanations',
      width: 500,
      maxHeight: 620,
      centered: true,
      draggable: true,
    });

    const legend = document.createElement('div');
    legend.style.cssText = 'font-size:var(--sb-fs-11);color:var(--sb-text-dim);margin-bottom:14px;line-height:1.6;';
    legend.innerHTML = '<span style="color:var(--sb-text-muted);font-weight:700;">Placeholders:</span> ' +
      ['{total}', '{price}', '{month}', '{taxClause}', '{discStr}', '{taxStr}', '{name}']
        .map(p => `<code style="background:var(--sb-bg-deep);padding:1px 5px;border-radius:3px;font-size:var(--sb-fs-10);">${p}</code>`)
        .join(' ');
    panel.appendChild(legend);

    const textareas = {};
    const templates = getBillTemplates();

    Object.keys(BILL_TEMPLATE_LABELS).forEach(key => {
      const wrap = document.createElement('div');
      wrap.style.cssText = 'margin-bottom:6px;border:1px solid var(--sb-border-soft);border-radius:6px;overflow:hidden;';

      const header = document.createElement('div');
      header.style.cssText = `
        display:flex;justify-content:space-between;align-items:center;
        padding:8px 10px;background:var(--sb-bg-deep);cursor:pointer;
        font-size:var(--sb-fs-12);font-weight:600;color:var(--sb-text-secondary);
        user-select:none;
      `;
      const chevron = document.createElement('span');
      chevron.textContent = '▶';
      chevron.style.cssText = 'font-size:9px;color:var(--sb-text-dim);transition:transform 0.15s;';
      header.appendChild(document.createTextNode(BILL_TEMPLATE_LABELS[key]));
      header.appendChild(chevron);

      const body = document.createElement('div');
      body.style.cssText = 'display:none;padding:8px;background:var(--sb-bg-row);';

      const ta = document.createElement('textarea');
      ta.value = templates[key] || BILL_TEMPLATES_DEFAULTS[key] || '';
      ta.style.cssText = `
        width:100%;box-sizing:border-box;min-height:72px;padding:7px;resize:vertical;
        background:var(--sb-bg-input);color:var(--sb-text);border:1px solid var(--sb-border);
        border-radius:4px;font-size:var(--sb-fs-12);line-height:1.5;font-family:inherit;
      `;
      textareas[key] = ta;
      body.appendChild(ta);

      header.onclick = () => {
        const open = body.style.display !== 'none';
        body.style.display = open ? 'none' : 'block';
        chevron.style.transform = open ? '' : 'rotate(90deg)';
      };

      wrap.appendChild(header);
      wrap.appendChild(body);
      panel.appendChild(wrap);
    });

    const actionRow = document.createElement('div');
    actionRow.style.cssText = 'display:flex;gap:8px;margin-top:14px;';

    const resetBtn = document.createElement('button');
    resetBtn.textContent = 'Reset to Defaults';
    resetBtn.style.cssText = 'padding:7px 12px;border-radius:6px;border:1px solid var(--sb-border);background:transparent;color:var(--sb-text-muted);font-size:var(--sb-fs-12);font-weight:600;cursor:pointer;';
    resetBtn.onclick = () => {
      GM_setValue(BILL_TEMPLATES_KEY, JSON.stringify(BILL_TEMPLATES_DEFAULTS));
      Object.keys(textareas).forEach(k => { textareas[k].value = BILL_TEMPLATES_DEFAULTS[k] || ''; });
      resetBtn.textContent = '✔ Reset';
      setTimeout(() => { resetBtn.textContent = 'Reset to Defaults'; }, 1500);
    };

    const spacer = document.createElement('div');
    spacer.style.flex = '1';

    const saveBtn = document.createElement('button');
    saveBtn.textContent = 'Save';
    saveBtn.style.cssText = 'padding:7px 20px;border-radius:6px;border:none;background:var(--sb-solid-primary-lo);color:var(--sb-on-solid);font-size:var(--sb-fs-12);font-weight:700;cursor:pointer;';
    saveBtn.onclick = () => {
      const patch = {};
      Object.keys(textareas).forEach(k => { patch[k] = textareas[k].value; });
      GM_setValue(BILL_TEMPLATES_KEY, JSON.stringify(patch));
      saveBtn.textContent = '✔ Saved';
      saveBtn.style.background = 'var(--sb-solid-success)';
      setTimeout(() => { saveBtn.textContent = 'Save'; saveBtn.style.background = 'var(--sb-solid-primary-lo)'; }, 1500);
    };

    actionRow.appendChild(resetBtn);
    actionRow.appendChild(spacer);
    actionRow.appendChild(saveBtn);
    panel.appendChild(actionRow);
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

  function renderUserInfoBar(sub, addressMatch, orderAddrStr, coupons = []) {
    const existing = document.getElementById('sb-user-info-bar');
    if (existing) existing.remove();

    const bar = document.createElement('div');
    bar.id = 'sb-user-info-bar';
    bar.style.cssText = 'position:relative;display:flex;flex-direction:column;align-items:flex-start;gap:3px;padding:4px 30px 4px 10px;background:var(--sb-bg-deep);border-top:1px solid var(--sb-border-soft);font-family:Arial,sans-serif;font-size:var(--sb-fs-11);';

    // Manual refresh — re-fetches subscription + shipping address and re-renders.
    const refreshBtn = document.createElement('button');
    refreshBtn.textContent = '⟳';
    refreshBtn.title = 'Refresh info bar';
    refreshBtn.style.cssText = 'position:absolute;top:4px;right:8px;background:var(--sb-control-bg);border:1px solid var(--sb-border-soft);color:var(--sb-text-muted);border-radius:5px;cursor:pointer;font-size:var(--sb-fs-13);line-height:1;padding:2px 6px;';
    refreshBtn.onmouseenter = () => refreshBtn.style.color = 'var(--sb-text)';
    refreshBtn.onmouseleave = () => refreshBtn.style.color = 'var(--sb-text-muted)';
    refreshBtn.onclick = () => { refreshBtn.textContent = '⏳'; refreshSubscriptionBar(); };
    bar.appendChild(refreshBtn);

    // Identity line: name · email · shipping address (above the tag groups)
    const idLine = document.createElement('div');
    idLine.style.cssText = 'color:var(--sb-text-muted);font-size:var(--sb-fs-13);';
    const _u  = cachedCustomerCtx?.user;
    const _sh = cachedCustomerCtx?._shippingAddress;
    const _idParts = [];
    if (_u) {
      const nameEmail = `${_u.firstName || ''} ${_u.lastName || ''} · ${_u.email || ''}`.replace(/\s+/g, ' ').trim();
      if (nameEmail) _idParts.push(nameEmail);
    }
    const _addrStr = formatShippingAddress(_sh);
    if (_addrStr) _idParts.push(_addrStr);
    if (_idParts.length) {
      idLine.textContent = _idParts.join(' · ');
      bar.appendChild(idLine);
    }

    // Tags row container (the original horizontal flow)
    const tagsRow = document.createElement('div');
    tagsRow.style.cssText = 'display:flex;align-items:center;flex-wrap:wrap;gap:6px;';
    bar.appendChild(tagsRow);

    function makeGroup() {
      const g = document.createElement('div');
      g.style.cssText = 'display:inline-flex;align-items:center;gap:4px;border:1px solid var(--sb-control-bg);border-radius:6px;padding:2px 4px;';
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
      isActive ? 'var(--sb-success-soft)' : 'var(--sb-danger)',
      isActive ? 'var(--sb-tag-success-bg)' : 'var(--sb-tag-danger-bg)',
      isActive ? 'var(--sb-tag-success-border)' : 'var(--sb-tag-danger-border)'
    ));

    if (sub.cashbirdDetails?.data?.isAwaitCancellation) {
      group1.appendChild(makeTag('Scheduled Cancellation', 'var(--sb-danger-strong)', 'var(--sb-tag-danger-bg)', 'var(--sb-tag-danger-border)'));
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
      group1.appendChild(makeTag(bdLabel, 'var(--sb-text-muted)', 'var(--sb-tag-muted-bg)', 'var(--sb-tag-muted-border)'));
    }

    // Gender tag from cache
    const cachedGender = cachedCustomerCtx?._gender;
    if (cachedGender) {
      const isMale = cachedGender === 'MALE';
      const gTag = makeTag(
        isMale ? '♂ Colognes' : '♀ Perfumes',
        isMale ? 'var(--sb-info)' : 'var(--sb-tag-pink-fg)',
        isMale ? 'var(--sb-tag-info-bg)' : 'var(--sb-tag-pink-bg)',
        isMale ? 'var(--sb-tag-info-border)' : 'var(--sb-tag-pink-border)'
      );
      gTag.dataset.tag = 'gender';
      group1.appendChild(gTag);
    }

    tagsRow.appendChild(group1);

    // Group 2: Plan + add-ons
    const group2 = makeGroup();

    if (sub.planName) {
      const planLabel = PLAN_LABELS[sub.planName] || sub.planName.replace(/_/g, ' ').toLowerCase();
      group2.appendChild(makeTag(planLabel, 'var(--sb-tag-plan-fg)', 'var(--sb-tag-plan-bg)', 'var(--sb-tag-plan-border)'));
    }

    const remainingCredits = (sub.credits || []).filter(c => c.status === 'NEW').length;
    if (remainingCredits > 0) {
      const label = `🎟 ${remainingCredits} credit${remainingCredits === 1 ? '' : 's'}`;
      group2.appendChild(makeTag(label, 'var(--sb-success-soft)', 'var(--sb-tag-success-bg)', 'var(--sb-tag-success-border)'));
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
        group2.appendChild(makeTag(label, 'var(--sb-warning)', 'var(--sb-tag-warning-bg)', 'var(--sb-tag-warning-border)'));
      }
    });

    if (group2.children.length) tagsRow.appendChild(group2);

    // Coupons (every subscription on the account): active ones as green pills, past
    // ones folded into one dim "+N past" pill. Detail lives in the tooltips.
    const couponTip = (c) => [c.code, c.type, c.amount != null ? 'amount ' + c.amount : null, c.status,
      c.applyToNextBillingDates ? 'next ' + c.applyToNextBillingDates + ' billing(s)' : null].filter(Boolean).join(' · ');
    const isActiveCoupon = (c) => c.status === 'ACTIVE' || c.status === 'NEW';
    const activeCoupons = coupons.filter(isActiveCoupon);
    const pastCoupons = coupons.filter(c => !isActiveCoupon(c));
    if (coupons.length) {
      const couponGroup = makeGroup();
      activeCoupons.forEach(c => {
        const t = makeTag('🏷 ' + (c.code || c.type || 'coupon'), 'var(--sb-success-soft)', 'var(--sb-tag-success-bg)', 'var(--sb-tag-success-border)');
        t.title = couponTip(c);
        couponGroup.appendChild(t);
      });
      if (pastCoupons.length) {
        const t = makeTag(activeCoupons.length ? `+${pastCoupons.length} past` : `🏷 ${pastCoupons.length} past coupon${pastCoupons.length === 1 ? '' : 's'}`,
          'var(--sb-text-dim)', 'var(--sb-tag-muted-bg)', 'var(--sb-tag-muted-border)');
        t.title = pastCoupons.map(couponTip).join('\n');
        couponGroup.appendChild(t);
      }
      tagsRow.appendChild(couponGroup);
    }

    // Group 3: Location tags (Canada, UK, Hawaii/Alaska/Puerto Rico)
    const shipping = cachedCustomerCtx?.user?.userAddress?.shipping;
    if (shipping) {
      const country = (shipping.country || '').trim().toUpperCase();
      const region  = (shipping.region  || '').trim().toUpperCase();
      const locationGroup = makeGroup();
      if (country === 'CA' || country === 'CAN' || country === 'CANADA') {
        locationGroup.appendChild(makeTag('🍁 Canada', 'var(--sb-warning-strong)', 'var(--sb-tag-warning-bg)', 'var(--sb-tag-warning-border)'));
      } else if (country === 'GB' || country === 'UK' || country === 'UNITED KINGDOM') {
        locationGroup.appendChild(makeTag('🇬🇧 UK', 'var(--sb-info)', 'var(--sb-tag-info-bg)', 'var(--sb-tag-info-border)'));
      } else if (['HI', 'HAWAII'].includes(region)) {
        locationGroup.appendChild(makeTag('🌺 Hawaii', 'var(--sb-tag-pink-fg)', 'var(--sb-tag-pink-bg)', 'var(--sb-tag-pink-border)'));
      } else if (['AK', 'ALASKA'].includes(region)) {
        locationGroup.appendChild(makeTag('🏔 Alaska', 'var(--sb-text-muted)', 'var(--sb-tag-muted-bg)', 'var(--sb-tag-muted-border)'));
      } else if (['PR', 'PUERTO RICO'].includes(region)) {
        locationGroup.appendChild(makeTag('🌴 Puerto Rico', 'var(--sb-tag-emerald-fg)', 'var(--sb-tag-emerald-bg)', 'var(--sb-tag-emerald-border)'));
      }
      if (locationGroup.children.length) tagsRow.appendChild(locationGroup);
    }

    // Group 4: Warnings (addr, fraud, chargeback — fraud/chargeback appended async)
    const group3 = makeGroup();
    group3.id = 'sb-info-group-warnings';
    group3.style.display = 'none'; // hidden until a warning is added

    if (addressMatch === false) {
      const addrTag = makeTag('⚠ Addr Updated', 'var(--sb-warning-strong)', 'var(--sb-tag-warning-bg)', 'var(--sb-tag-warning-border)');
      if (orderAddrStr) addrTag.title = 'Last order shipped to: ' + orderAddrStr;
      group3.appendChild(addrTag);
      group3.style.display = '';
    }

    tagsRow.appendChild(group3);

    const slot = document.getElementById('sb-info-bar-slot');
    if (slot) {
      slot.innerHTML = '';
      slot.appendChild(bar);
    } else {
      const toolbar = document.getElementById('sb-toolbar');
      if (toolbar && toolbar.parentNode) toolbar.insertAdjacentElement('afterend', bar);
    }
  }

  /** Tokenize address into a sorted, punctuation-stripped word multiset for comparison. */
  function _addrTokens(s) {
    return (s || '')
      .toLowerCase()
      .replace(/[^a-z0-9\s]/g, ' ')
      .split(/\s+/)
      .filter(Boolean)
      .sort()
      .join(' ');
  }

  /** Compare user's current address against last order's shipping address. */
  function _compareAddresses(user, order) {
    const current = user?.userAddress?.shipping;
    const orderAddr = order?.initialShippingAddress;
    if (!current?.street1 || !orderAddr?.street1) return null; // can't compare
    const currentFull = [current.street1, current.street2].filter(Boolean).join(' ');
    const orderFull   = [orderAddr.street1, orderAddr.street2].filter(Boolean).join(' ');
    return _addrTokens(currentFull) === _addrTokens(orderFull);
  }

  /** Fetch user details (subscription + fraud + gender + gweb) + last order and render the info bar. */
  function _fetchAndRenderSubBar(userId) {
    const user = cachedCustomerCtx?.user;
    const startId = getCustomerIdFromURL();

    // Single query: subscriptions + fraud + gweb + gender + shipping addr
    crmRequest({
      method: 'POST', url: GRAPHQL_URL,
      headers: { 'Content-Type': 'application/json' },
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
            if (userData.userAddress?.shipping)     cachedCustomerCtx._shippingAddress = userData.userAddress.shipping;
          }

          // Subscription
          const subs = userData.subscriptionList || [];
          const active = subs.find(s => s.status === 'Active') || subs[0];
          if (!active) return;

          if (cachedCustomerCtx) {
            const awaitCancel = !!active.cashbirdDetails?.data?.isAwaitCancellation;
            cachedCustomerCtx._isAwaitCancellation = awaitCancel;
            // When the cancellation actually takes effect, cached here because this is
            // the only path holding the full subscription — the lean user.subscription
            // from CRM_QUERY has no billing dates. An await-cancellation sub still reads
            // Active and ends at the next billing date; an already-Cancelled one carries
            // subscriptionEndDate. pickCancellationFlow sanity-checks it.
            cachedCustomerCtx._subCancelDate = awaitCancel
              ? (active.cashbirdDetails?.data?.nextBillingDate || active.nextBillingDate || null)
              : (active.subscriptionEndDate || null);
          }

          // Fetch last order for address comparison + chargeback check
          fetchLastOrders(userId, (orders) => {
            if (getCustomerIdFromURL() !== startId) return; // stale
            const lastSub = (orders || []).find(o => o.type === 'SUBSCRIPTION' || !o.type);
            const addrMatch = lastSub ? _compareAddresses(user, lastSub) : null;
            const orderAddr = lastSub?.initialShippingAddress;
            const orderAddrStr = orderAddr ? [orderAddr.street1, orderAddr.street2, orderAddr.city, orderAddr.region, orderAddr.postcode].filter(Boolean).join(', ') : '';
            renderUserInfoBar(active, addrMatch, orderAddrStr, subs.flatMap(s => s.coupons || []));

            // Async: fraud tag
            if (userData.fraudInfo?.status === 'DECLINE') {
              const warnGroup = document.getElementById('sb-info-group-warnings');
              if (warnGroup) {
                warnGroup.appendChild(makeTag('⚠ Fraud', 'var(--sb-danger-strong)', 'var(--sb-tag-danger-bg)', 'var(--sb-tag-danger-border)'));
                warnGroup.style.display = '';
              }
            }

            // Async: chargeback tag
            crmRequest({
              method: 'POST', url: GRAPHQL_URL,
              headers: { 'Content-Type': 'application/json' },
              data: JSON.stringify({ operationName: 'chargesByUserId', query: CHARGES_QUERY, variables: { id: userId } }),
              onload(cRes) {
                if (getCustomerIdFromURL() !== startId) return; // stale
                try {
                  const cData = JSON.parse(cRes.responseText);
                  const chargeList = cData?.data?.userById?.data?.chargeList || [];
                  const hasChargeback = chargeList.some(c => c.charge?.status === 'CHARGEBACK');
                  if (hasChargeback) {
                    const warnGroup = document.getElementById('sb-info-group-warnings');
                    if (warnGroup && !warnGroup.querySelector('[data-tag="chargeback"]')) {
                      const t = makeTag('⚠ Chargeback', 'var(--sb-danger-strong)', 'var(--sb-tag-danger-bg)', 'var(--sb-tag-danger-border)');
                      t.dataset.tag = 'chargeback';
                      warnGroup.appendChild(t);
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

      const noEmailKey = `__no_email:${startId}`;
      if (userInfoLastEmail === noEmailKey) return; // already determined no email for this customer

      const email = await resolveEmail();
      if (getCustomerIdFromURL() !== startId) return; // stale
      if (!email) { userInfoLastEmail = noEmailKey; return; }

      // Full resolve
      if (email === userInfoLastEmail && document.getElementById('sb-user-info-bar')) return;
      userInfoLastEmail = email;
      document.getElementById('sb-user-info-bar')?.remove();

      searchCRM(email, (users, err) => {
        if (getCustomerIdFromURL() !== startId) return; // stale
        if (err || !users) return;
        const sbUsers = users.filter(u => !u.origin || u.origin === 'SCENTBIRD');
        const exactMatch = sbUsers.find(u => u.email?.toLowerCase() === email.toLowerCase());
        if (!exactMatch) return;
        cachedCustomerCtx = { email, user: exactMatch, _kustomerId: getCustomerIdFromURL() };
        _fetchAndRenderSubBar(exactMatch.id);
      });
    }, 800);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // AI ASSIST — piggybacks on Kustomer's internal Copilot (AIR) endpoint.
  // No third-party API key: uses the agent's existing Kustomer auth via cookies.
  // Each analysis creates a fresh thread (server upserts on first POST).
  // ══════════════════════════════════════════════════════════════════════════

  const KUSTOMER_API_BASE = 'https://scentbird.api.kustomerapp.com';

  // Kustomer attaches a couple of headers to every API call that we can't derive
  // ourselves: x-csrf-token (NOT the _csrf cookie — a separate token from a boot
  // endpoint) and x-kustomer-version (the live client build, which Kustomer bumps
  // often and rejects stale values of). Sniff outbound fetch/XHR to capture both
  // from Kustomer's own traffic, then reuse them on our requests.
  let _aiCapturedCsrf = null;
  let _aiCapturedVersion = null;
  (function captureKustomerHeaders() {
    const readHeader = (headers, name) => {
      if (!headers) return null;
      if (typeof Headers !== 'undefined' && headers instanceof Headers) return headers.get(name);
      if (Array.isArray(headers)) {
        const p = headers.find(x => (x?.[0] || '').toLowerCase() === name);
        return p ? p[1] : null;
      }
      for (const k of Object.keys(headers)) if (k.toLowerCase() === name) return headers[k];
      return null;
    };
    const captureFrom = (headers) => {
      if (!headers) return;
      const csrf = readHeader(headers, 'x-csrf-token');
      if (csrf) _aiCapturedCsrf = csrf;
      const ver = readHeader(headers, 'x-kustomer-version');
      if (ver) _aiCapturedVersion = ver;
    };
    const origFetch = window.fetch;
    window.fetch = function(input, init) {
      try { captureFrom(init?.headers); captureFrom(input?.headers); } catch(e) {}
      return origFetch.apply(this, arguments);
    };
    const origSetRH = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
      const n = (name || '').toLowerCase();
      if (n === 'x-csrf-token' && value) _aiCapturedCsrf = value;
      else if (n === 'x-kustomer-version' && value) _aiCapturedVersion = value;
      return origSetRH.apply(this, arguments);
    };
  })();

  /** Current Kustomer client build = the x-kustomer-version the API expects.
   *  Kustomer serves its JS bundles from a versioned path, e.g.
   *    https://scentbird.kustomerapp.com/release-v0.1.8314/runtime.<hash>.js
   *  so we read it straight out of the loaded <script> URLs — synchronous and
   *  always current, with no hardcoded value to go stale (which is what broke
   *  Fill Name). Falls back to a version sniffed off Kustomer's own traffic;
   *  throws if neither is available rather than guessing a stale one. */
  function _aiVersion() {
    const scan = (urls) => urls.map(u => (String(u).match(/\/(release-v[\d.]+)\//) || [])[1]).find(Boolean);
    const v = scan([...document.scripts].map(s => s.src))
           || scan(performance.getEntriesByType('resource').map(e => e.name))
           || _aiCapturedVersion;
    if (!v) throw new Error('Kustomer version not found — reload the ticket and retry');
    return v;
  }

  function _aiCsrf() {
    // Kustomer stores its live CSRF token in localStorage.csrfToken (the _csrf cookie
    // is HttpOnly and unreadable). Fall back to the interceptor capture if storage fails.
    try {
      const stored = localStorage.getItem('csrfToken');
      if (stored) {
        // localStorage sometimes wraps strings in JSON quotes — strip them.
        try {
          const parsed = JSON.parse(stored);
          if (typeof parsed === 'string' && parsed) return parsed;
        } catch(e) {}
        return stored;
      }
    } catch(e) {}
    return _aiCapturedCsrf || '';
  }
  function _aiUuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
      const r = Math.random() * 16 | 0;
      return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
    });
  }
  function _aiHeaders() {
    return {
      'accept':                       'application/json',
      'content-type':                 'application/json',
      'x-csrf-token':                 _aiCsrf(),
      'x-kustomer-client':            'web',
      'x-kustomer-client-request-id': _aiUuid(),
      'x-kustomer-version':           _aiVersion(),
    };
  }

  // Kustomer's Edit-Customer modal PATCHes a fixed subset of editable attributes,
  // full-replace style. Sending fields outside this list (activeUsers, rev, etc.)
  // risks 4xx or unintended mutation; sending fewer fields wipes them.
  const KUSTOMER_PATCH_FIELDS = [
    'phones', 'sharedPhones', 'whatsapps', 'sharedWhatsapps',
    'emails', 'sharedEmails', 'socials', 'tiktoks', 'urls',
    'locations', 'externalIds',
    'name', 'company', 'externalId', 'defaultLang', 'timeZone', 'dedicatedUser', 'gender',
  ];

  async function kustomerGetCustomer(id) {
    const res = await fetch(`${KUSTOMER_API_BASE}/v1/customers/${id}`, {
      credentials: 'include', headers: _aiHeaders(),
    });
    if (!res.ok) throw new Error('Kustomer GET ' + res.status);
    const attrs = (await res.json())?.data?.attributes;
    if (!attrs) throw new Error('Kustomer GET: no attributes');
    return attrs;
  }

  /** Locate Kustomer customers that already hold the given email. Mirrors the
   *  search Kustomer's UI fires after a duplicate-email 409 (source=fetchDuplicateCustomer). */
  async function kustomerSearchByEmail(email) {
    const reqId = _aiUuid();
    const url = `${KUSTOMER_API_BASE}/v1/customers/search?include=customers%2CsatisfactionResponse&page=1&pageSize=5&source=fetchDuplicateCustomer&trackTotalHits=10000&client-request-id=${reqId}`;
    const res = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId },
      body: JSON.stringify({
        or: [{ customer_email: { equals: (email || '').toLowerCase() } }],
        starrocksMode: 'disabled',
      }),
    });
    if (!res.ok) throw new Error('Kustomer search ' + res.status);
    return (await res.json())?.data || [];
  }

  /** Fold `source` customer into `target`. After merge, source ceases to exist
   *  and target retains a combined record. Used to resolve duplicate-email conflicts. */
  async function kustomerMergeCustomers({ target, source }) {
    const reqId = _aiUuid();
    const url = `${KUSTOMER_API_BASE}/v1/customers/merges?client-request-id=${reqId}`;
    const res = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId },
      body: JSON.stringify({ target, source }),
    });
    if (!res.ok) {
      let detail = 'Kustomer merge ' + res.status;
      try {
        const j = await res.json();
        if (j?.errors?.[0]?.detail) detail += ': ' + j.errors[0].detail;
        console.warn('[BirdsEye] Kustomer merge error', res.status, j);
      } catch(_) {}
      throw new Error(detail);
    }
    return (await res.json())?.data;
  }

  /** Create a fresh Kustomer customer record holding `email` and nothing else
   *  (name filled in later by Fill Name). Cookie-authenticated like the rest of the
   *  kustomer* helpers. On a duplicate-email 409 we reuse the existing record:
   *  look it up via kustomerSearchByEmail and return its id so the CCPA flow
   *  continues on it. Returns the customer id string. */
  async function kustomerCreateCustomer(email) {
    const reqId = _aiUuid();
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/customers?client-request-id=${reqId}`,
      {
        method: 'POST',
        credentials: 'include',
        headers: { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId },
        body: JSON.stringify({
          phones: [], sharedPhones: [], whatsapps: [], sharedWhatsapps: [],
          emails: [{ type: 'home', email }], sharedEmails: [],
          socials: [], tiktoks: [], urls: [],
          locations: [{ type: 'home', address: '', address2: '', cityName: '', countryName: '', regionName: '', zipCode: '' }],
          externalIds: [],
          name: null, company: null, externalId: null,
          defaultLang: null, timeZone: null, dedicatedUser: null, gender: null,
        }),
      },
    );
    if (res.status === 409) {
      const list = await kustomerSearchByEmail(email);
      const existing = (list || [])[0];
      if (existing?.id) return existing.id;
      throw new Error('Duplicate email but existing customer not found');
    }
    if (!res.ok) {
      let detail = 'Kustomer create ' + res.status;
      try {
        const j = await res.json();
        if (j?.errors?.[0]?.detail) detail += ': ' + j.errors[0].detail;
        console.warn('[BirdsEye] Kustomer create error', res.status, j);
      } catch(_) {}
      throw new Error(detail);
    }
    const id = (await res.json())?.data?.id;
    if (!id) throw new Error('Kustomer create: no id in response');
    return id;
  }

  /** Open an outbound, open-status conversation on a Kustomer customer.
   *  Returns the conversation id (used to navigate the agent onto it). */
  async function kustomerCreateConversation(customerId) {
    const reqId = _aiUuid();
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/customers/${customerId}/conversations?client-request-id=${reqId}`,
      {
        method: 'POST',
        credentials: 'include',
        headers: { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId },
        body: JSON.stringify({ direction: 'out', status: 'open' }),
      },
    );
    if (!res.ok) {
      let detail = 'Kustomer conversation ' + res.status;
      try {
        const j = await res.json();
        if (j?.errors?.[0]?.detail) detail += ': ' + j.errors[0].detail;
        console.warn('[BirdsEye] Kustomer conversation error', res.status, j);
      } catch(_) {}
      throw new Error(detail);
    }
    const id = (await res.json())?.data?.id;
    if (!id) throw new Error('Kustomer conversation: no id in response');
    return id;
  }

  // ── Drafts / notes / macros ────────────────────────────────────────────────
  // Kustomer composes every outbound message (reply or note) as a draft: POST to
  // /v1/customers/{id}/drafts creates an empty draft on a conversation, then a PUT
  // to /v1/drafts/{draftId} fills it (body, or a shortcut/macro). These helpers are
  // deliberately granular so the AI layer can reuse them (e.g. prefillReply).

  /** Current signed-in Kustomer agent { id, displayName }. Cached. Used for
   *  self-assigning conversations and the {{currentUser.displayName}} macro var. */
  let _kustomerAgent = null;
  async function kustomerCurrentUser() {
    if (_kustomerAgent) return _kustomerAgent;
    const res = await fetch(`${KUSTOMER_API_BASE}/v1/users/current`, {
      credentials: 'include', headers: _aiHeaders(),
    });
    if (!res.ok) throw new Error('Kustomer users/current ' + res.status);
    const data = (await res.json())?.data;
    const a = data?.attributes || {};
    if (!data?.id) throw new Error('Kustomer users/current: no id');
    _kustomerAgent = { id: data.id, displayName: a.displayName || a.name || a.firstName || '' };
    return _kustomerAgent;
  }
  async function kustomerCurrentUserId() { return (await kustomerCurrentUser()).id; }

  /** Create an empty draft (channel 'note' or 'email') on a conversation. Returns draft id. */
  async function kustomerCreateDraft(customerId, channel, conversationId) {
    const reqId = _aiUuid();
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/customers/${customerId}/drafts?client-request-id=${reqId}`,
      {
        method: 'POST',
        credentials: 'include',
        headers: { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId },
        body: JSON.stringify({ channel, conversation: conversationId }),
      },
    );
    const raw = await res.text();
    if (!res.ok) throw new Error('Kustomer draft create ' + res.status + ' ' + raw.slice(0, 200));
    let json = null; try { json = JSON.parse(raw); } catch (_) {}
    const id = json?.data?.id;
    if (!id) throw new Error('Kustomer draft create: no id');
    return id;
  }

  /** Update a draft. `sourceShortcutId` attributes the change to a Kustomer shortcut
   *  (mirrors the UI's x-kustomer-ui-source-* headers when applying a macro). */
  async function kustomerUpdateDraft(draftId, payload, sourceShortcutId) {
    const reqId = _aiUuid();
    const headers = { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId };
    if (sourceShortcutId) {
      headers['x-kustomer-ui-source-id'] = sourceShortcutId;
      headers['x-kustomer-ui-source-type'] = 'shortcut';
    }
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/drafts/${draftId}?client-request-id=${reqId}`,
      { method: 'PUT', credentials: 'include', headers, body: JSON.stringify(payload) },
    );
    const raw = await res.text();
    if (!res.ok) throw new Error('Kustomer draft update ' + res.status + ' ' + raw.slice(0, 200));
    let json = null; try { json = JSON.parse(raw); } catch (_) {}
    return json?.data;
  }

  /** Random 5-char DraftJS block key. */
  function _draftBlockKey() {
    let k = '';
    for (let i = 0; i < 5; i++) k += (Math.random() * 36 | 0).toString(36);
    return k;
  }

  /** Note draft body holding a single clickable link. `body` is the plain URL —
   *  the send PUT drops `payload`, so the posted note renders from `body` alone and
   *  a markdown mirror would post literally. DraftJS entity carries the link in the
   *  composer only. */
  function buildLinkNotePayload(linkUrl, prefix) {
    const lead = prefix ? prefix + ' ' : '';
    const text = lead + linkUrl + ' ';
    return {
      body: lead + linkUrl,
      payload: {
        draftJs: {
          blocks: [{
            text,
            key: _draftBlockKey(),
            type: 'unstyled',
            depth: 0,
            data: {},
            inlineStyleRanges: [],
            entityRanges: [{ offset: lead.length, length: linkUrl.length, key: 0 }],
          }],
          entityMap: {
            '0': { type: 'LINK', mutability: 'MUTABLE', data: { href: linkUrl, rel: 'noopener', target: '_blank', url: linkUrl } },
          },
        },
      },
      channel: 'note',
    };
  }

  /** Post an internal note containing a single link onto a conversation.
   *
   *  Kustomer has no "send message" endpoint — sending is a *second* PUT to the same
   *  draft carrying `sendAt`, which flips it to status:'scheduled'. Writing the body
   *  alone (what we used to do) leaves an unsent draft sitting in the composer.
   *  Order matters: once scheduled the draft is frozen ('Cannot edit scheduled draft',
   *  409), so the content PUT has to land first. The send PUT mirrors the web client:
   *  body + conversation + userMentions, no `payload`. */
  async function postConversationNote(customerId, conversationId, linkUrl, prefix) {
    const draftId = await kustomerCreateDraft(customerId, 'note', conversationId);
    const payload = buildLinkNotePayload(linkUrl, prefix);
    await kustomerUpdateDraft(draftId, payload);
    await kustomerUpdateDraft(draftId, {
      body: payload.body,
      channel: 'note',
      userMentions: [],
      conversation: conversationId,
      sendAt: new Date().toISOString(),
    });
    return draftId;
  }

  // ── Shortcut (macro) templates ─────────────────────────────────────────────
  // Kustomer interpolates shortcuts *client-side*, so PUT {shortcuts:[id]} leaves the
  // draft body empty. We instead fetch the shortcut's own template (draft.text +
  // payload.text draftJs), interpolate the vars we can, and PUT the filled body
  // ourselves. Shortcuts rarely change → cache the whole list for the session.
  let _shortcutCache = null;   // Map(id → attributes)
  async function kustomerFetchShortcuts() {
    if (_shortcutCache) return _shortcutCache;
    const map = new Map();
    let page = 1;
    // Hard cap on pages so a schema change can't spin forever.
    for (let guard = 0; guard < 50; guard++) {
      const res = await fetch(
        `${KUSTOMER_API_BASE}/v1/shortcuts?page=${page}&pageSize=100&client-request-id=${_aiUuid()}`,
        { credentials: 'include', headers: _aiHeaders() },
      );
      if (!res.ok) throw new Error('Kustomer shortcuts ' + res.status);
      const json = await res.json();
      (json?.data || []).forEach(sc => { if (sc?.id) map.set(sc.id, sc.attributes || {}); });
      if (!json?.links?.next) break;
      page++;
    }
    _shortcutCache = map;
    return map;
  }
  /** Shortcut attributes by id ({ name, draft:{text}, payload:{text:{blocks,entityMap}} }). */
  async function kustomerGetShortcut(id) {
    const map = await kustomerFetchShortcuts();
    const attrs = map.get(id);
    if (!attrs) throw new Error('Shortcut not found: ' + id);
    return attrs;
  }

  /** The conversation's composer draft — `{ id, attributes }` — or null if none saved
   *  yet. Kustomer autosaves the live composer into this object, so its `body` and
   *  `payload.draftJs` are the rich (links, italics) version of what the agent sees. */
  async function kustomerGetConversationDraft(conversationId) {
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/conversations/${conversationId}/drafts?client-request-id=${_aiUuid()}`,
      { credentials: 'include', headers: _aiHeaders() },
    );
    if (!res.ok) return null;
    const d = ((await res.json())?.data || [])[0];
    return d?.id ? { id: d.id, attributes: d.attributes || {} } : null;
  }

  /** Existing composer draft id for a conversation, or null if none saved yet. */
  async function kustomerGetConversationDraftId(conversationId) {
    return (await kustomerGetConversationDraft(conversationId))?.id || null;
  }

  // ── Sending an email reply ────────────────────────────────────────────────
  // Same shape as postConversationNote: there is no send endpoint, a second PUT to
  // the draft carrying `sendAt` IS the send. Email needs four fields a note doesn't —
  // from/to/app/template — and none of them live on the draft object; Kustomer's web
  // client computes them at send time, so we supply them the same way it does.
  //
  // `from` is deliberately a constant rather than derived from the inbound message's
  // recipient: tickets that arrived at the old support@scentbird.mail.kustomerapp.com
  // mailbox would otherwise reply out of it, and the live client no longer does that.
  const KUSTOMER_REPLY_FROM = 'support@scentbird.com';
  const KUSTOMER_REPLY_APP = 'postmark';
  // Org email template (the branded HTML wrapper). Stable per org; if Scentbird ever
  // changes it, sends fail loudly and this is the one line to update.
  const KUSTOMER_REPLY_TEMPLATE = '67ea39f7a280fa09cc90f178';

  /** Who a reply goes to: the address the customer actually wrote in from (newest
   *  inbound email's `meta.from`), falling back to the Kustomer customer record. */
  async function resolveReplyRecipient(conversationId) {
    try {
      const messages = await kustomerGetConversationMessagesRaw(conversationId, 20);
      for (let i = messages.length - 1; i >= 0; i--) {
        const a = messages[i];
        if (a.direction !== 'in') continue;
        const from = a.meta?.from || (typeof a.from === 'string' ? a.from : a.from?.email);
        if (from && /@/.test(from)) return String(from).trim();
      }
    } catch (e) {
      console.warn('[BirdsEye] reply recipient: message scan failed', e);
    }
    return await resolveEmail();
  }

  /** Plain text → draftJs blocks. Only used when the server draft is too stale to
   *  trust; it flattens links and italics, which is why it isn't the primary path. */
  function draftJsFromText(text) {
    return {
      blocks: String(text || '').split('\n').map(line => ({
        key: _draftBlockKey(), text: line, type: 'unstyled', depth: 0,
        data: {}, inlineStyleRanges: [], entityRanges: [],
      })),
      entityMap: {},
    };
  }

  const _normReplyText = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  /** A draft's text as the composer renders it. Reads the draftJs blocks, NOT `body`:
   *  `body` is a markdown serialization ("future \\- whenever", "*italic*"), so it
   *  never compares equal to the composer's innerText. Blocks carry the clean text. */
  function draftPlainText(attributes) {
    const blocks = attributes?.payload?.draftJs?.blocks;
    if (Array.isArray(blocks)) return blocks.map(b => b?.text || '').join('\n');
    return attributes?.body || '';
  }

  const _sleep = (ms) => new Promise(r => setTimeout(r, ms));

  /** Poll the conversation's draft until Kustomer has autosaved the composer's current
   *  text into it. Returns the draft once it matches, or null on timeout — the caller
   *  then falls back to serializing the composer itself.
   *
   *  The gap between checks ramps 120ms → 400ms rather than sitting flat at 400ms.
   *  Autosave normally lands within a few hundred ms of the composer settling, so the
   *  flat interval spent most of its time asleep just past the moment the draft became
   *  ready; the ramp catches that case without turning a genuinely slow save into a
   *  request flood. `abort` (optional) stops the loop early once nobody is waiting. */
  async function waitForDraftToMatch(conversationId, composerText, timeoutMs = 4000, abort) {
    const want = _normReplyText(composerText);
    const deadline = Date.now() + timeoutMs;
    let gap = 120;
    while (Date.now() < deadline && !abort?.()) {
      const d = await kustomerGetConversationDraft(conversationId).catch(() => null);
      if (d && _normReplyText(draftPlainText(d.attributes)) === want) return d;
      await _sleep(gap);
      gap = Math.min(400, Math.round(gap * 1.5));
    }
    return null;
  }
  let _replyPrime = null;

  /** The prefetch for this exact send, or null if there isn't a usable one. Composer
   *  text is the key: an edit between arming and confirming invalidates the draft we
   *  were waiting on, and we would rather re-poll than send the previous wording. */
  function _takeReplyPrime(conversationId, composerText) {
    const p = _replyPrime;
    _replyPrime = null;
    if (!p) return null;
    const usable = p.conversationId === conversationId
      && _normReplyText(p.composerText) === _normReplyText(composerText);
    if (!usable) { p.cancelled = true; return null; }
    return p;
  }

  /** Wait out the prefetch's own poll rather than starting a second one against the
   *  same endpoint. Resolves the instant it lands, or null once `timeoutMs` passes —
   *  at which point the poller is called off and the caller falls back. */
  async function _primedDraft(prime, timeoutMs) {
    const deadline = Date.now() + timeoutMs;
    while (!prime.draft && !prime.done && Date.now() < deadline) await _sleep(50);
    prime.cancelled = true;   // no-op if it already landed
    return prime.draft;
  }

  /** Send the reply currently sitting in the composer, as an email to the customer.
   *
   *  Body source is Kustomer's own autosaved draft whenever it agrees with what's on
   *  screen — that copy carries the link entities and inline styles a plain-text
   *  serialization would drop. If autosave hasn't caught up (our token fill lands via
   *  a synthetic paste, so it can lag), we write the composer's text ourselves rather
   *  than send something the agent isn't looking at.
   *
   *  Two PUTs, in this order: content, then the same body again with `sendAt`. A
   *  scheduled draft is frozen (409 'Cannot edit scheduled draft'), so content first.
   *  Throws on any hard failure — nothing is sent then. */
  async function sendConversationReply({ customerId, conversationId, onStatus }) {
    const status = (m) => { try { onStatus?.(m); } catch (_) {} };
    if (!customerId || !conversationId) throw new Error('No open conversation');

    const composer = getComposer();
    const composerText = (composer?.innerText || '').trim();

    // Started when the button armed; usually already resolved. Falling back costs a
    // poll, never correctness — see primeReplySend.
    const prime = _takeReplyPrime(conversationId, composerText);
    // The recipient is a separate GET that depends on nothing below it, so it runs
    // alongside the draft wait instead of after it.
    const recipientP = prime ? prime.recipient
                             : resolveReplyRecipient(conversationId).catch(() => null);

    let draft = null;
    if (composerText) {
      // Only say "Reading draft…" if we are actually about to wait on one — with a
      // warm prefetch this resolves in a tick and the agent sees "Sending…" directly.
      if (!prime?.draft) status('Reading draft…');
      draft = prime ? await _primedDraft(prime, 4000)
                    : await waitForDraftToMatch(conversationId, composerText);
    }
    let body, payload;

    if (draft) {
      // Kustomer's own copy — keeps links and italics.
      body = draft.attributes.body || composerText;
      payload = draft.attributes.payload || { draftJs: draftJsFromText(body) };
    } else if (composerText) {
      // Autosave never caught up. Send what's on screen: `body` goes out unescaped,
      // so any markdown characters in it render as formatting — an acceptable trade
      // against sending a body the agent never saw.
      console.warn('[BirdsEye] reply: draft autosave lagged — sending composer text (formatting flattened).');
      body = composerText;
      payload = { draftJs: draftJsFromText(composerText) };
      draft = await kustomerGetConversationDraft(conversationId)
           || { id: await kustomerCreateDraft(customerId, 'email', conversationId) };
      status('Writing draft…');
      await kustomerUpdateDraft(draft.id, { body, channel: 'email', payload });
    } else {
      // Composer closed or empty — the saved draft is all there is.
      draft = await kustomerGetConversationDraft(conversationId);
      if (!draft?.attributes?.body) throw new Error('Nothing to send — the reply is empty');
      body = draft.attributes.body;
      payload = draft.attributes.payload || { draftJs: draftJsFromText(body) };
    }

    // Last line of defence against a half-filled macro reaching a customer. The token
    // fill is best-effort (it degrades to a clipboard fallback the agent completes by
    // hand), so a live "[MONTH]" in the body is a real possibility — and it is the one
    // failure that is embarrassing rather than merely broken. Checked unescaped: the
    // draft's `body` is markdown, where the brackets come through as \[MONTH\].
    const unfilled = FILL_VARIABLES
      .map(v => '[' + v.key + ']')
      .filter(t => String(body).replace(/\\/g, '').includes(t) || composerText.includes(t));
    if (unfilled.length) throw new Error('Unfilled token(s) in the reply: ' + unfilled.join(', '));

    const to = await recipientP;
    if (!to) throw new Error('No recipient address for this conversation');

    status('Sending…');
    await kustomerUpdateDraft(draft.id, {
      body,
      channel: 'email',
      from: { email: KUSTOMER_REPLY_FROM },
      to: [{ email: to }],
      sendAt: new Date().toISOString(),
      payload,
      template: KUSTOMER_REPLY_TEMPLATE,
      app: KUSTOMER_REPLY_APP,
    });
    // composerText is the agent's copy as they saw it on screen; `body` is the
    // markdown serialization that went to Kustomer. The log wants the former to
    // compare against the macro, and the latter as the record of what was sent.
    return { to, body, composerText, draftId: draft.id };
  }

  /** After an API send the composer still shows the reply: the send went around
   *  Kustomer's editor, so its in-memory draft never learned the text went out.
   *  Clear it through Kustomer's own discard (trash → "Yes, discard") so its store
   *  resets too — wiping the editor ourselves would leave Kustomer autosaving into
   *  a draft that no longer exists.
   *
   *  Waits until the server has dropped the sent draft first: the send is a PUT with
   *  `sendAt`, and discarding before Kustomer consumes it could delete the email with
   *  it. Then only acts while the same conversation is open and the composer still
   *  holds exactly what went out, so nothing the agent typed since is ever lost.
   *  Returns true when the composer ends up clear. */
  async function clearComposerAfterSend({ conversationId, draftId, sentText, timeoutMs = 15000 }) {
    // Not kustomerGetConversationDraft: it reads a failed GET as "no draft", and a
    // false "gone" here is the one mistake that could cancel the send.
    const draftStillThere = async () => {
      const res = await fetch(
        `${KUSTOMER_API_BASE}/v1/conversations/${conversationId}/drafts?client-request-id=${_aiUuid()}`,
        { credentials: 'include', headers: _aiHeaders() },
      );
      if (!res.ok) throw new Error('drafts ' + res.status);
      return ((await res.json())?.data || []).some(d => d.id === draftId);
    };
    const deadline = Date.now() + timeoutMs;
    for (;;) {
      try { if (!(await draftStillThere())) break; } catch (_) {}
      if (Date.now() > deadline) return false;
      await _sleep(500);
    }

    if (getConversationIdFromURL() !== conversationId) return false;
    const composer = getComposer();
    const current = _normReplyText(composer?.innerText || '');
    if (!current) return true;
    if (current !== _normReplyText(sentText || '')) return false;

    // Nearest trash icon walking up from the editor is the reply box's own toolbar.
    let trash = null;
    for (let n = composer.parentElement, i = 0; n && !trash && i < 14; n = n.parentElement, i++) {
      trash = n.querySelector('svg[data-icon="trash"]')?.closest('button') || null;
    }
    if (!trash) return false;
    trash.click();

    let yes = null;
    const until = Date.now() + 3000;
    while (!yes && Date.now() < until) {
      const dlg = [...document.querySelectorAll('.ReactModalPortal [role="dialog"]')]
        .find(x => /discard draft/i.test(x.innerText || ''));
      yes = dlg ? [...dlg.querySelectorAll('button')].find(b => /^yes,?\s*discard$/i.test((b.innerText || '').trim())) : null;
      if (!yes) await _sleep(100);
    }
    if (yes) { yes.click(); await _sleep(300); }
    return !getComposer();
  }

  /** Fill the macro vars we can resolve. customer.firstName + currentUser.displayName
   *  only; BirdsEye [TOKENS] and integration vars ({{recharge…}}, {{shopify…}}) are
   *  left as-is (the former for the Fill step, the latter dropped intentionally). */
  function interpolateMacro(text, { firstName, agentName }) {
    return String(text || '')
      .replace(/\{\{\s*customer\.firstName\s*\}\}/g, firstName || '')
      .replace(/\{\{\s*currentUser\.displayName\s*\}\}/g, agentName || '');
  }

  /** Build a draft PUT payload from a shortcut's template, interpolated. Reuses the
   *  shortcut's own draftJs blocks; clears inline styles/entities on any block we
   *  rewrite so substitution can't misalign their offsets. */
  function buildFilledDraftPayload(shortcutAttrs, vars) {
    const body = interpolateMacro(shortcutAttrs?.draft?.text || '', vars);
    const srcDraftJs = shortcutAttrs?.payload?.text;
    let draftJs;
    if (srcDraftJs?.blocks) {
      draftJs = {
        blocks: srcDraftJs.blocks.map(b => {
          const text = interpolateMacro(b.text || '', vars);
          if (text === b.text) return { ...b };
          // Block changed by substitution → drop ranges to avoid offset drift.
          return { ...b, text, inlineStyleRanges: [], entityRanges: [] };
        }),
        entityMap: srcDraftJs.entityMap || {},
      };
    } else {
      // No draftJs template — synthesize plain unstyled blocks from the body text.
      draftJs = {
        blocks: body.split('\n').map(line => ({
          key: _draftBlockKey(), text: line, type: 'unstyled', depth: 0,
          data: {}, inlineStyleRanges: [], entityRanges: [],
        })),
        entityMap: {},
      };
    }
    return { channel: 'email', body, payload: { draftJs } };
  }

  /** Fill the conversation's reply draft from a Kustomer shortcut (macro). Interpolates
   *  name vars, writes the result into the conversation's *existing* composer draft
   *  (creating one only if none exists), and mirrors it to the clipboard as a fallback
   *  for when the open Draft.js composer doesn't live-refresh. BirdsEye [TOKENS] are
   *  left for the Fill step. Returns { draftId, body }. */
  async function prefillReply(customerId, conversationId, shortcutId, user) {
    const attrs = await kustomerGetShortcut(shortcutId);
    let agentName = '';
    try { agentName = (await kustomerCurrentUser()).displayName; } catch (_) {}
    const vars = {
      firstName: toProperCase(user?.firstName || '') || '',
      agentName,
    };
    const payload = buildFilledDraftPayload(attrs, vars);
    const draftId = await kustomerGetConversationDraftId(conversationId)
                 || await kustomerCreateDraft(customerId, 'email', conversationId);
    await kustomerUpdateDraft(draftId, payload);
    try { GM_setClipboard(payload.body); } catch (_) {}
    return { draftId, body: payload.body };
  }

  // Conversation-level attributes the CCPA macro sets (its name + category fields).
  // Configurable bit (the shortcut id) lives in settings; these belong to the one
  // Cancel-My-Subscription macro CCPA drives.
  const CCPA_CONVERSATION_NAME = 'Regarding your Scentbird Subscription cancellation...';
  const CCPA_CONVERSATION_CUSTOM = {
    categoryStr: 'Cancel My Subscription',
    cancelMySubscriptionWrapUpStr: 'Customer Requested Cancellation::Cancelled Subscription via Admin',
  };

  /** Apply the CCPA macro's conversation actions: rename, self-assign, set categories.
   *  Self-assign is best-effort (skipped if the agent id can't be resolved). */
  /** PATCH a conversation. `sourceShortcutId` attributes the change to a Kustomer
   *  shortcut, mirroring the UI's x-kustomer-ui-source-* headers when a macro drives it. */
  async function patchConversation(conversationId, body, sourceShortcutId) {
    const reqId = _aiUuid();
    const headers = { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId };
    if (sourceShortcutId) {
      headers['x-kustomer-ui-source-id'] = sourceShortcutId;
      headers['x-kustomer-ui-source-type'] = 'shortcut';
    }
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/conversations/${conversationId}?client-request-id=${reqId}`,
      { method: 'PATCH', credentials: 'include', headers, body: JSON.stringify(body) },
    );
    if (!res.ok) throw new Error('Kustomer conversation PATCH ' + res.status);
    return (await res.json())?.data;
  }

  async function applyCcpaConversationMacro(conversationId, shortcutId) {
    const body = { name: CCPA_CONVERSATION_NAME, custom: { ...CCPA_CONVERSATION_CUSTOM } };
    try { body.assignedUsers = [await kustomerCurrentUserId()]; }
    catch (e) { console.warn('[BirdsEye] CCPA self-assign skipped', e); }
    return patchConversation(conversationId, body, shortcutId);
  }

  // Wrap-up the ORIGINATING info@scentbird.com ticket carries when it's closed. Same
  // two fields the new ticket gets — the captured PATCHes hold identical values — so
  // it aliases rather than duplicates them. Split the day they genuinely differ.
  const CCPA_ORIGIN_CLOSE_CUSTOM = { ...CCPA_CONVERSATION_CUSTOM };

  /** Close the info@scentbird.com hand-off ticket that triggered a CCPA run.
   *
   *  Two PATCHes, categories before the close. Kustomer's UI sets each wrap-up field
   *  with its own request and only then marks the conversation done, and that order is
   *  worth keeping: a conversation that reaches `done` uncategorised is the state
   *  nobody goes back to fix. */
  async function closeCcpaOriginTicket(conversationId) {
    await patchConversation(conversationId, { custom: { ...CCPA_ORIGIN_CLOSE_CUSTOM } });
    return patchConversation(conversationId, { status: 'done' });
  }

  /** When an Apply hits a duplicate-email 409, find the conflicting Kustomer
   *  customer and rewire the Apply button into a two-stage Merge & retry control.
   *  Returns true when handled (caller should NOT also show the duplicate error). */
  async function offerMergeOnDuplicateError(btn, candidateUser) {
    const targetEmail = (candidateUser?.email || '').trim().toLowerCase();
    const currentId = getCustomerIdFromURL();
    if (!targetEmail || !currentId) return false;

    btn.disabled = true;
    btn.textContent = '🔍 Finding conflict…';
    btn.style.background = 'var(--sb-solid-warn)';

    let conflict;
    try {
      const list = await kustomerSearchByEmail(targetEmail);
      conflict = (list || []).find(c => c.id !== currentId);
    } catch (e) {
      btn.disabled = false;
      btn.textContent = '✘ Conflict lookup failed';
      btn.style.background = 'var(--sb-danger-border)';
      return true;
    }
    if (!conflict) {
      btn.disabled = false;
      btn.textContent = '✘ Duplicate but conflict not found';
      btn.style.background = 'var(--sb-danger-border)';
      return true;
    }

    let armed = false, armTimer = null;
    const idleText = `🔗 Merge into ${conflict.id.slice(-6)} & retry`;
    btn.disabled = false;
    btn.textContent = idleText;
    btn.title = `Conflicting Kustomer customer: ${conflict.id} (${conflict?.attributes?.displayName || ''})`;
    btn.onclick = async () => {
      if (!armed) {
        armed = true;
        btn.textContent = '⚠ Click again to merge';
        btn.style.background = 'var(--sb-solid-warn-hi)';
        armTimer = setTimeout(() => { armed = false; btn.textContent = idleText; btn.style.background = 'var(--sb-solid-warn)'; }, 4000);
        return;
      }
      clearTimeout(armTimer);
      btn.disabled = true;
      btn.textContent = '⏳ Merging…';
      try {
        // target = existing customer that already holds the email (preserves its history);
        // source = current conversation's customer (gets folded in)
        await kustomerMergeCustomers({ target: conflict.id, source: currentId });
        btn.textContent = '✓ Merged';
        btn.style.background = 'var(--sb-solid-success)';
        // After merge, Kustomer's UI re-routes the conversation to target on its own.
      } catch (e) {
        btn.disabled = false;
        btn.textContent = '✘ Merge failed: ' + e.message;
        btn.style.background = 'var(--sb-danger-border)';
      }
    };
    return true;
  }

  async function kustomerPatchCustomer(id, attrs) {
    const reqId = _aiUuid();
    const body = {};
    for (const k of KUSTOMER_PATCH_FIELDS) body[k] = attrs[k] ?? null;
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/customers/${id}?client-request-id=${reqId}`,
      {
        method: 'PATCH',
        credentials: 'include',
        headers: { ..._aiHeaders(), 'x-kustomer-client-request-id': reqId },
        body: JSON.stringify(body),
      },
    );
    if (!res.ok) {
      // Parse Kustomer's structured error envelope when present, so the UI surfaces
      // an actionable message instead of just an HTTP code.
      let detail = 'Kustomer PATCH ' + res.status;
      try {
        const errBody = await res.json();
        const e = errBody?.errors?.[0];
        if (e?.code === 'duplicate' && e?.source?.parameter === 'email') {
          detail = 'Duplicate email — already on another Kustomer customer. Merge manually.';
        } else if (e?.detail) {
          detail = `Kustomer ${res.status}: ${e.detail}`;
        }
        console.warn('[BirdsEye] Kustomer PATCH error', res.status, errBody);
      } catch(_) {}
      throw new Error(detail);
    }
    return (await res.json())?.data?.attributes;
  }

  /** Parse `/app/customers/{customerId}/event/{conversationId}` from URL. */
  function getConversationIdFromURL() {
    const m = window.location.href.match(/\/event\/([0-9a-f]{24})/);
    return m ? m[1] : null;
  }

  // ── Generic CRM query as a promise ─────────────────────────────────────────
  // The callback fetchers each shape their result for one panel (fetchLastOrders
  // slices to 4, fetchCharges keeps one billing group). The agent wants the whole
  // record, so this goes to the same queries directly and hands back `data`.
  function crmQueryP(operationName, query, variables) {
    return new Promise((resolve, reject) => {
      if (!getAccessToken()) return reject(new Error('Not signed in to CRM (Okta).'));
      crmRequest({
        method: 'POST', url: GRAPHQL_URL,
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({ operationName, query, variables }),
        onload(res) {
          if (res.status === 401) { handle401(); return reject(new Error('CRM token expired')); }
          if (res.status === 403) { handle403(); return reject(new Error('CRM captcha required')); }
          try {
            const json = JSON.parse(res.responseText);
            if (json.errors?.length) return reject(new Error(json.errors[0]?.message || 'GraphQL error'));
            resolve(json?.data || {});
          } catch (e) { reject(new Error('CRM parse error')); }
        },
        onerror() { reject(new Error('CRM network error')); },
      });
    });
  }

  // ── Live carrier tracking: Asendia ────────────────────────────────────────
  // Asendia parcels are looked up in Asendia's own tracking system — the same API its
  // public tracking page (a1.asendiausa.com/tracking) calls. That page ships a public
  // API key and brand key in its script; they are read from the script at runtime
  // (cached, re-read on a 401/403) rather than hardcoded, so a key rotation on their
  // side doesn't break this. Undocumented endpoint: every failure falls back to the CRM
  // tracking timeline, labelled as such. Street addresses in scan locations are dropped.
  const ASENDIA_PAGE_SCRIPT = 'https://a1.asendiausa.com/tracking/js/main.js';
  const ASENDIA_DEFAULT_TRACKING_KEY = 'AE654169-0B14-45F9-8498-A8E464E13D26';
  const ASENDIA_CONFIG_TTL_MS = 6 * 60 * 60 * 1000;
  const ASENDIA_RESULT_TTL_MS = 10 * 60 * 1000;
  let _asendiaConfig = null;
  const _asendiaResults = new Map();   // TRACKINGNO → { at, data }

  function _gmGet(url, headers = {}) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'GET', url, headers, timeout: 15000,
        onload: (res) => resolve({ status: res.status, text: res.responseText || '' }),
        onerror: () => reject(new Error('network error')),
        ontimeout: () => reject(new Error('timed out')),
      });
    });
  }

  /** "courier" from the order's tracking link (scentbird.com/tracking?courier=asendia-us&…). */
  function _orderCourier(o) {
    const url = o?.tracking?.trackingUrl || o?.shipment?.trackingUrl || '';
    const m = /[?&]courier=([^&#]+)/i.exec(url);
    return m ? decodeURIComponent(m[1]).toLowerCase() : null;
  }

  async function _asendiaConfigGet(force = false) {
    if (!force && _asendiaConfig && Date.now() - _asendiaConfig.at < ASENDIA_CONFIG_TTL_MS) return _asendiaConfig;
    const { status, text } = await _gmGet(ASENDIA_PAGE_SCRIPT);
    if (status !== 200) throw new Error(`Asendia page script ${status}`);
    const field = (name) => (new RegExp(`"${name}"\\s*:\\s*"([^"]+)"`).exec(text) || [])[1] || null;
    const cfg = {
      baseUrl: field('baseUrl'),
      auth: field('authorizationHeaderValue'),
      apiKey: field('a1ApiKeyHeaderValue'),
      trackingKey: (/"([0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12})";let\s/i.exec(text) || [])[1] || ASENDIA_DEFAULT_TRACKING_KEY,
      at: Date.now(),
    };
    if (!cfg.baseUrl || !cfg.auth || !cfg.apiKey) throw new Error('Asendia page script changed — keys not found');
    _asendiaConfig = cfg;
    return cfg;
  }

  const _finalMileCarrier = (link) => {
    const host = (() => { try { return new URL(link).hostname.replace(/^www\./, ''); } catch (_) { return ''; } })();
    const known = [[/royalmail/, 'Royal Mail'], [/uniuni/, 'UniUni'], [/canadapost|postescanada/, 'Canada Post'], [/usps/, 'USPS'], [/auspost/, 'Australia Post'], [/dhl/, 'DHL'], [/evri|hermes/, 'Evri'], [/purolator/, 'Purolator']];
    return (known.find(([re]) => re.test(host)) || [null, host || null])[1];
  };

  /** Live Asendia tracking for one tracking number. Resolves { found:false } when
   *  Asendia has no record, throws on transport/API failure. */
  async function asendiaTracking(trackingNo, { force = false } = {}) {
    const tn = String(trackingNo || '').trim().toUpperCase();
    if (!tn) return { found: false };
    const hit = _asendiaResults.get(tn);
    if (!force && hit && Date.now() - hit.at < ASENDIA_RESULT_TTL_MS) return hit.data;

    const call = async (cfg) => _gmGet(
      `${cfg.baseUrl}/api/A1/TrackingBranded/Tracking?trackingKey=${encodeURIComponent(cfg.trackingKey)}&trackingNumber=${encodeURIComponent(tn)}`,
      { Authorization: cfg.auth, 'X-AsendiaOne-ApiKey': cfg.apiKey, Accept: 'application/json' },
    );
    let res = await call(await _asendiaConfigGet());
    if (res.status === 401 || res.status === 403) res = await call(await _asendiaConfigGet(true));   // keys rotated
    if (res.status !== 200) throw new Error(`Asendia tracking ${res.status}`);
    let j;
    try { j = JSON.parse(res.text); } catch (_) { throw new Error('Asendia tracking: unreadable response'); }

    const s = j?.trackingBrandedSummary;
    const code = j?.responseStatus?.responseStatusCode;
    let data;
    if (!s || code === 204 || code === 404) {
      data = { found: false };
    } else {
      const events = (Array.isArray(j.trackingBrandedDetail) ? j.trackingBrandedDetail : [])
        .map(e => {
          const L = e.eventLocationDetails || {};
          return { code: e.eventCode || null, description: e.eventDescription || '', at: e.eventOn || null,
            location: [L.city, L.province, L.countryIso2].filter(Boolean).join(', ') || null, source: e.eventSource || null };
        })
        .filter(e => e.at)
        .sort((a, b) => new Date(a.at) - new Date(b.at));   // oldest first, like CRM
      const vendorNo = s.trackingNumberVendor && String(s.trackingNumberVendor).toUpperCase() !== tn ? s.trackingNumberVendor : null;
      data = {
        found: true,
        service: s.service || null, weight: s.weight || null, destination: s.destinationCountry || null,
        progress: s.trackingProgress ? { completed: s.trackingProgress.completed, total: s.trackingProgress.total } : null,
        finalMile: s.finalMileTrackingLink || vendorNo
          ? { carrier: _finalMileCarrier(s.finalMileTrackingLink), trackingNumber: vendorNo, link: s.finalMileTrackingLink || null }
          : null,
        lastMileMessage: s.lastMileMessage || null,
        events,
        fetchedAt: new Date().toISOString(),
      };
    }
    _asendiaResults.set(tn, { at: Date.now(), data });
    return data;
  }

  /** Raw message attributes for a conversation, oldest first. The formatted variant
   *  above flattens to "[CUSTOMER] preview" strings; the CCPA parser needs the actual
   *  fields, since `preview` is truncated and the identifiers can sit past the cut. */
  async function kustomerGetConversationMessagesRaw(conversationId, max = 10) {
    const res = await fetch(
      `${KUSTOMER_API_BASE}/v1/conversations/${conversationId}/messages?page=1&pageSize=${max}&client-request-id=${_aiUuid()}`,
      { credentials: 'include', headers: _aiHeaders() },
    );
    if (!res.ok) throw new Error('Kustomer messages ' + res.status);
    return ((await res.json())?.data || [])
      .map(m => m.attributes || {})
      .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  }

  // ══════════════════════════════════════════════════════════════════════════
  // MUTATION OBSERVER — wire everything on header appear
  // ══════════════════════════════════════════════════════════════════════════

  let _lastObservedHref = location.href;
  const observer = new MutationObserver(() => {
    if (location.href !== _lastObservedHref) {
      _lastObservedHref = location.href;
      window.dispatchEvent(new Event('sb-route-change'));
    }
    if (toolbarEl && !document.body.contains(toolbarEl)) {
      toolbarEl = null;
      fillNameRegistered = false;
      clearCustomerCtx();
      ['sb-order-panel','sb-charges-panel','sb-cancel-panel','sb-downgrade-panel','sb-search-panel','sb-queue-panel','sb-edit-customer-panel'].forEach(id => document.getElementById(id)?.remove());
    }
    if (!ensureToolbar()) return;
    registerFillName();
    registerCrmButtons();
    registerSpecialsMenu();
    refreshOktaButtonState();
    refreshUserInfoBar();
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
    if (!popup) console.warn('[BirdsEye] customer auth: popup blocked');

    setTimeout(() => {
      try { popup?.close(); } catch(e) {}

      GM_xmlhttpRequest({
        method: 'POST',
        url: 'https://api.scentbird.com/graphql?opname=authProbe',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Origin': 'https://www.scentbird.com',
          'Referer': 'https://www.scentbird.com/',
        },
        data: JSON.stringify({
          // Introspection is disabled on this API, and currentUser answers guests too —
          // personalInfo is only non-null once the gwebLink session has landed.
          operationName: 'authProbe',
          variables: {},
          query: `query authProbe { currentUser { data { personalInfo { firstName } } } }`
        }),
        onload(res) {
          const csrfMatch = res.responseHeaders.match(/x-csrf-token:\s*(.+)/i);
          if (csrfMatch) _customerCsrfToken = csrfMatch[1].trim();

          try {
            const body = JSON.parse(res.responseText);
            if (body?.data?.currentUser?.data?.personalInfo) {
              _customerApiReady = true;
              callback(true);
            } else {
              console.warn('[BirdsEye] customer auth probe: no data', res.status, (body?.errors || []).map(e => e.message).join(' | ').slice(0, 300));
              callback(false);
            }
          } catch(e) {
            console.warn('[BirdsEye] customer auth probe: non-JSON', res.status, String(res.responseText || '').slice(0, 200));
            callback(false);
          }
        },
        onerror(res) {
          console.warn('[BirdsEye] customer auth probe: network error', res?.status, res?.error);
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
