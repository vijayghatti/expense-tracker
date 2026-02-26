/**
 * Security Module - Simple passcode lock for the app
 * NOTE: This is an app-level lock (to keep casual users out),
 * not a full cryptographic encryption of your data.
 */
const Security = (() => {
  const STORAGE_KEY = 'et_lock';

  function getConfig() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }

  function saveConfig(cfg) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
  }

  function randomSalt(len = 16) {
    const bytes = new Uint8Array(len);
    crypto.getRandomValues(bytes);
    return Array.from(bytes).map(b => b.toString(16).padStart(2, '0')).join('');
  }

  async function hashPassword(password, salt) {
    const enc = new TextEncoder();
    const data = enc.encode(password + ':' + salt);
    const buf = await crypto.subtle.digest('SHA-256', data);
    const bytes = new Uint8Array(buf);
    return Array.from(bytes).map(b => b.toString(16).padStart(2, '0')).join('');
  }

  function createOverlay(innerHtml) {
    let overlay = document.getElementById('security-overlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'security-overlay';
      overlay.style.position = 'fixed';
      overlay.style.inset = '0';
      overlay.style.display = 'flex';
      overlay.style.alignItems = 'center';
      overlay.style.justifyContent = 'center';
      overlay.style.background = 'radial-gradient(circle at top, rgba(76, 175, 80, 0.20), rgba(0,0,0,0.90))';
      overlay.style.zIndex = '9999';
      overlay.style.backdropFilter = 'blur(8px)';
      document.body.appendChild(overlay);
    }
    overlay.innerHTML = `
      <div style="
        width: 100%;
        max-width: 420px;
        margin: 0 16px;
        background: var(--bg-elevated, #111827);
        border-radius: 16px;
        padding: 24px 24px 20px;
        box-shadow: 0 18px 45px rgba(0,0,0,0.6);
        border: 1px solid rgba(148,163,184,0.35);
        color: var(--text-primary, #e5e7eb);
        font-family: system-ui, -apple-system, BlinkMacSystemFont, 'SF Pro Text', sans-serif;
      ">
        ${innerHtml}
      </div>
    `;
    return overlay;
  }

  function renderSetup(onUnlocked) {
    const overlay = createOverlay(`
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="font-size:1.4rem;">🔒</div>
        <div>
          <div style="font-weight:600;font-size:1.05rem;">Set up passcode</div>
          <div style="font-size:0.85rem;color:var(--text-secondary,#9ca3af);">
            Protect access to your expense data on this device.
          </div>
        </div>
      </div>
      <form id="lock-setup-form" style="display:flex;flex-direction:column;gap:12px;margin-top:12px;">
        <div style="display:flex;flex-direction:column;gap:4px;">
          <label style="font-size:0.85rem;color:var(--text-secondary,#9ca3af);">Passcode</label>
          <input type="password" name="password" required minlength="4"
            style="border-radius:8px;padding:8px 10px;border:1px solid rgba(148,163,184,0.6);background:#020617;color:inherit;">
        </div>
        <div style="display:flex;flex-direction:column;gap:4px;">
          <label style="font-size:0.85rem;color:var(--text-secondary,#9ca3af);">Confirm passcode</label>
          <input type="password" name="confirm" required minlength="4"
            style="border-radius:8px;padding:8px 10px;border:1px solid rgba(148,163,184,0.6);background:#020617;color:inherit;">
        </div>
        <div id="lock-error" style="min-height:16px;font-size:0.8rem;color:#f97373;"></div>
        <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:8px;">
          <button type="submit"
            style="border:none;border-radius:999px;padding:8px 16px;font-size:0.9rem;font-weight:500;
                   background:linear-gradient(135deg,#22c55e,#16a34a);color:white;cursor:pointer;">
            Enable Lock
          </button>
        </div>
      </form>
    `);

    const form = overlay.querySelector('#lock-setup-form');
    const errEl = overlay.querySelector('#lock-error');
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      errEl.textContent = '';
      const pwd = form.password.value.trim();
      const confirm = form.confirm.value.trim();
      if (pwd.length < 4) {
        errEl.textContent = 'Passcode must be at least 4 characters.';
        return;
      }
      if (pwd !== confirm) {
        errEl.textContent = 'Passcodes do not match.';
        return;
      }
      const salt = randomSalt();
      const hash = await hashPassword(pwd, salt);
      saveConfig({ salt, hash, enabled: true });
      overlay.remove();
      if (typeof onUnlocked === 'function') onUnlocked();
    });
  }

  function renderUnlock(config, onUnlocked) {
    const overlay = createOverlay(`
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="font-size:1.4rem;">🔐</div>
        <div>
          <div style="font-weight:600;font-size:1.05rem;">Enter passcode</div>
          <div style="font-size:0.85rem;color:var(--text-secondary,#9ca3af);">
            This protects your expense data on this device.
          </div>
        </div>
      </div>
      <form id="lock-unlock-form" style="display:flex;flex-direction:column;gap:12px;margin-top:12px;">
        <div style="display:flex;flex-direction:column;gap:4px;">
          <label style="font-size:0.85rem;color:var(--text-secondary,#9ca3af);">Passcode</label>
          <input type="password" name="password" required minlength="4" autofocus
            style="border-radius:8px;padding:8px 10px;border:1px solid rgba(148,163,184,0.6);background:#020617;color:inherit;">
        </div>
        <div id="lock-error" style="min-height:16px;font-size:0.8rem;color:#f97373;"></div>
        <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:8px;">
          <button type="submit"
            style="border:none;border-radius:999px;padding:8px 16px;font-size:0.9rem;font-weight:500;
                   background:linear-gradient(135deg,#22c55e,#16a34a);color:white;cursor:pointer;">
            Unlock
          </button>
        </div>
      </form>
    `);

    const form = overlay.querySelector('#lock-unlock-form');
    const errEl = overlay.querySelector('#lock-error');
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      errEl.textContent = '';
      const pwd = form.password.value.trim();
      const hash = await hashPassword(pwd, config.salt);
      if (hash !== config.hash) {
        errEl.textContent = 'Incorrect passcode. Try again.';
        form.password.value = '';
        form.password.focus();
        return;
      }
      overlay.remove();
      if (typeof onUnlocked === 'function') onUnlocked();
    });
  }

  function init(onUnlocked) {
    const cfg = getConfig();
    if (!cfg || !cfg.enabled) {
      renderSetup(onUnlocked);
    } else {
      renderUnlock(cfg, onUnlocked);
    }
  }

  return { init };
})();

