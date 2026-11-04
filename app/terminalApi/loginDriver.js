// app/terminalApi/loginDriver.js

/**
 * Drives the Microsoft web sign-in from the terminal (`teams-cli login`) when
 * the app runs headless and nobody can see the window.
 *
 * buildLoginStepScript() returns a script that runs inside the sign-in page,
 * looks at which step is showing and acts on it at most once per page:
 * email -> password -> (pick account) -> MFA approval -> "Stay signed in?".
 * The main process calls it repeatedly until the sign-in page goes away.
 *
 * MFA is never bypassed: for number matching the code is reported so the
 * user can enter it in Microsoft Authenticator. One-time-code (SMS/TOTP)
 * pages are reported as unsupported.
 */

/**
 * @param {{ user: string|null, password: string|null }} input
 * @returns {string} script for frame.executeJavaScript(); resolves to
 *   { state, code?, error?, title? }
 */
function buildLoginStepScript({ user, password }) {
  return `(() => new Promise((resolve) => {
    const USER = ${JSON.stringify(user || null)};
    const PASSWORD = ${JSON.stringify(password || null)};
    // The email page also carries the password field, parked off-screen as
    // aria-hidden ".moveOffScreen" in a fixed box, so it has a size and no
    // offsetParent: check those markers rather than layout alone.
    const visible = (el) => {
      if (!el || !el.getClientRects().length) return false;
      if (el.closest('[aria-hidden="true"], .moveOffScreen')) return false;
      if (getComputedStyle(el).visibility === 'hidden') return false;
      const r = el.getBoundingClientRect();
      return r.width > 1 && r.height > 1;
    };
    const usable = (el) => visible(el) && !el.disabled && !el.readOnly;
    const first = (sel, test = visible) => Array.from(document.querySelectorAll(sel)).find(test);
    const setValue = (el, v) => {
      try { el.focus(); } catch (e) { void e; }
      const desc = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
      if (desc && desc.set) desc.set.call(el, v); else el.value = v;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    };
    const SUBMIT = '#idSIButton9, button[type=submit], input[type=submit]';
    // Click after a short delay so the page's data binding sees the new value.
    const submitSoon = () => setTimeout(() => { const b = first(SUBMIT, usable); if (b) b.click(); }, 400);
    // One action per page: a new page load gets a fresh window object.
    const done = (window.__tflLoginDone = window.__tflLoginDone || {});
    const errEl = first('#usernameError, #passwordError, #idTD_Error, #errorText, [role=alert]');
    const error = errEl ? (errEl.textContent || '').trim() : '';

    const codeEl = first('#idRichContext_DisplaySign, [data-testid*="DisplaySign" i]');
    const code = codeEl ? (codeEl.textContent || '').trim() : '';
    if (/^\\d{1,3}$/.test(code)) return resolve({ state: 'mfa', code });
    if (first('#idDiv_SAOTCAS_Title, #idDiv_SAOTCAS_Description')) return resolve({ state: 'mfa-wait' });
    if (first('#idTxtBx_SAOTCC_OTC, input[name=otc]', usable)) return resolve({ state: 'otc-unsupported' });

    if (first('#KmsiCheckboxField, #kmsiTitle, input[name=DontShowAgain]')) {
      if (!done.kmsi) { done.kmsi = true; const yes = first('#idSIButton9, #acceptButton', usable); if (yes) yes.click(); }
      return resolve({ state: 'kmsi-accepted' });
    }

    const pwd = first('input[type=password]', usable);
    if (pwd) {
      // Never retype a password the page has rejected: repeated failures can lock the account.
      if (error) return resolve({ state: 'error', error });
      if (!PASSWORD) return resolve({ state: 'password', error });
      if (!done.password) { done.password = true; setValue(pwd, PASSWORD); submitSoon(); }
      return resolve({ state: 'password-submitted' });
    }

    const email = first('input[type=email], input[name=loginfmt]', usable);
    if (email) {
      if (error && done.email) return resolve({ state: 'error', error });
      if (!USER) return resolve({ state: 'email', error });
      if (!done.email) { done.email = true; setValue(email, USER); submitSoon(); }
      return resolve({ state: 'email-submitted' });
    }

    const tiles = Array.from(document.querySelectorAll('#tilesHolder [role=button], #tilesHolder .table, [data-test-id]')).filter(visible);
    if (tiles.length) {
      const mine = USER && tiles.find((t) => (t.textContent || '').toLowerCase().includes(USER.toLowerCase()));
      const other = first('#otherTile, #otherTileText');
      if (!done.tile && (mine || other)) { done.tile = true; (mine || other).click(); }
      return resolve({ state: mine ? 'account-picked' : 'account-other' });
    }

    if (error) return resolve({ state: 'error', error });
    resolve({ state: 'unknown', title: document.title });
  }))()`;
}

module.exports = { buildLoginStepScript };
