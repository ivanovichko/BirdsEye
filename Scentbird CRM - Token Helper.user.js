// ==UserScript==
// @name         Scentbird CRM - Token Helper
// @namespace    scentbird-kustomer
// @version      1.0
// @description  Adds a "Copy Token" button on crm.scentbird.com for easy token capture
// @author       You
// @match        https://crm.scentbird.com/*
// @updateURL    https://raw.githubusercontent.com/ivanovichko/BirdsEye/main/Scentbird%20CRM%20-%20Token%20Helper.user.js
// @downloadURL  https://raw.githubusercontent.com/ivanovichko/BirdsEye/main/Scentbird%20CRM%20-%20Token%20Helper.user.js
// @grant        none
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  let capturedToken = '';

  // ── Intercept requests to grab the token ─────────────────────────────────────

  const origFetch = window.fetch;
  window.fetch = function (input, init) {
    const auth = (init?.headers || {})['Authorization'] || (init?.headers || {})['authorization'];
    if (auth?.startsWith('Bearer ')) setToken(auth.slice(7));
    return origFetch.apply(this, arguments);
  };

  const origSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    if (name.toLowerCase() === 'authorization' && value?.startsWith('Bearer ')) setToken(value.slice(7));
    return origSetRequestHeader.apply(this, arguments);
  };

  function setToken(token) {
    if (token === capturedToken) return;
    capturedToken = token;
    updateButton();
  }

  // ── UI ────────────────────────────────────────────────────────────────────────

  function injectButton() {
    if (document.getElementById('sb-copy-token-btn')) return;

    const btn = document.createElement('button');
    btn.id = 'sb-copy-token-btn';
    btn.textContent = '⏳ Waiting for token…';
    btn.style.cssText = `
      position: fixed;
      bottom: 20px;
      left: 20px;
      z-index: 999999;
      padding: 8px 16px;
      border-radius: 8px;
      border: none;
      background: #374151;
      color: #9ca3af;
      font-size: 13px;
      font-weight: 600;
      font-family: Arial, sans-serif;
      cursor: default;
      box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    `;

    btn.addEventListener('click', () => {
      if (!capturedToken) return;
      navigator.clipboard.writeText(capturedToken).then(() => {
        const orig = btn.textContent;
        btn.textContent = '✔ Copied!';
        btn.style.background = '#059669';
        btn.style.color = '#fff';
        setTimeout(() => {
          btn.textContent = orig;
          btn.style.background = '#4f46e5';
          btn.style.color = '#fff';
        }, 2000);
      });
    });

    document.body.appendChild(btn);
  }

  function updateButton() {
    const btn = document.getElementById('sb-copy-token-btn');
    if (!btn) return;
    btn.textContent = '🔑 Copy Token';
    btn.style.background = '#4f46e5';
    btn.style.color = '#fff';
    btn.style.cursor = 'pointer';
  }

  // Wait for body then inject
  if (document.body) {
    injectButton();
  } else {
    document.addEventListener('DOMContentLoaded', injectButton);
  }

})();