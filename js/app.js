/**
 * App Module - Initialization, tab routing, settings, modal management
 */
const App = (() => {
  let activeTab = 'dashboard';

  function init() {
    Store.init();
    Dashboard.init();
    applyTheme();
    renderShell();
    switchTab('dashboard');
  }

  function applyTheme() {
    const theme = Store.getSettings().theme || 'dark';
    document.documentElement.setAttribute('data-theme', theme);
  }

  function toggleTheme() {
    const current = Store.getSettings().theme || 'dark';
    const next = current === 'dark' ? 'light' : 'dark';
    Store.updateSettings({ theme: next });
    applyTheme();
    updateThemeIcon();
  }

  function updateThemeIcon() {
    const btn = document.getElementById('theme-btn');
    if (btn) btn.textContent = Store.getSettings().theme === 'dark' ? '☀️' : '🌙';
  }

  function renderShell() {
    // Header and tabs are static in HTML, just update theme icon
    updateThemeIcon();
  }

  function switchTab(tab) {
    activeTab = tab;
    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.tab === tab);
    });
    // Update content
    document.querySelectorAll('.tab-content').forEach(el => {
      el.classList.toggle('active', el.id === 'tab-' + tab);
    });
    // Render tab content
    switch (tab) {
      case 'dashboard': Dashboard.render(); break;
      case 'expenses': Expenses.render(); break;
      case 'recurring': Recurring.render(); break;
      case 'budgets': Budgets.render(); break;
      case 'settings': renderSettings(); break;
    }
  }

  // --- Modal ---
  function openModal(contentHtml) {
    let overlay = document.getElementById('modal-overlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'modal-overlay';
      overlay.className = 'modal-overlay';
      overlay.innerHTML = '<div class="modal" id="modal-content"></div>';
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) closeModal();
      });
      document.body.appendChild(overlay);
    }
    document.getElementById('modal-content').innerHTML = contentHtml;
    requestAnimationFrame(() => overlay.classList.add('active'));
    // Focus first input
    setTimeout(() => {
      const input = overlay.querySelector('input:not([type=hidden]):not([disabled]), select');
      if (input) input.focus();
    }, 100);
  }

  function closeModal() {
    const overlay = document.getElementById('modal-overlay');
    if (overlay) {
      overlay.classList.remove('active');
      setTimeout(() => overlay.remove(), 300);
    }
  }

  // --- Confirm Dialog ---
  function confirm(message, onConfirm) {
    const overlay = document.createElement('div');
    overlay.className = 'confirm-overlay';
    overlay.innerHTML = `
      <div class="confirm-box">
        <div class="icon">⚠️</div>
        <div class="message">${message}</div>
        <div class="actions">
          <button class="btn btn-secondary" id="confirm-cancel">Cancel</button>
          <button class="btn btn-danger" id="confirm-ok">Delete</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    overlay.querySelector('#confirm-cancel').onclick = () => overlay.remove();
    overlay.querySelector('#confirm-ok').onclick = () => { overlay.remove(); onConfirm(); };
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
  }

  // --- Settings Tab ---
  function renderSettings() {
    const container = document.getElementById('tab-settings');
    const settings = Store.getSettings();
    const categories = Store.getCategories();
    const usage = Store.getStorageUsage();
    const usageKB = (usage / 1024).toFixed(1);
    const usagePct = ((usage / (5 * 1024 * 1024)) * 100).toFixed(1);

    container.innerHTML = `
      <div class="settings-section">
        <div class="settings-section-title">General</div>
        <div class="settings-row">
          <div>
            <div class="settings-label">Currency Symbol</div>
            <div class="settings-desc">Displayed before amounts</div>
          </div>
          <input type="text" class="form-input" style="width:80px;text-align:center" value="${Utils.escapeHtml(settings.currency)}"
            onchange="App.updateCurrency(this.value)">
        </div>
        <div class="settings-row">
          <div>
            <div class="settings-label">Theme</div>
            <div class="settings-desc">Switch between dark and light mode</div>
          </div>
          <button class="btn btn-secondary" onclick="App.toggleTheme()">
            ${settings.theme === 'dark' ? '☀️ Switch to Light' : '🌙 Switch to Dark'}
          </button>
        </div>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">Categories</div>
        <div class="category-list" id="category-list">
          ${categories.map(c => `
            <div class="category-item">
              <div class="category-color" style="background:${c.color}"></div>
              <span style="font-size:1.2rem">${c.icon}</span>
              <span class="category-name">${Utils.escapeHtml(c.name)}</span>
              <button class="btn-icon" onclick="App.editCategory('${Utils.escapeHtml(c.name)}')" title="Edit">✏️</button>
              <button class="btn-icon danger" onclick="App.deleteCategory('${Utils.escapeHtml(c.name)}')" title="Delete">🗑️</button>
            </div>
          `).join('')}
        </div>
        <button class="btn btn-secondary" style="margin-top:12px" onclick="App.addCategory()">+ Add Category</button>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">Data Management</div>
        <div class="settings-row">
          <div>
            <div class="settings-label">Storage Used</div>
            <div class="settings-desc">${usageKB} KB of ~5 MB (${usagePct}%)</div>
          </div>
          <div class="progress-bar" style="width:120px;height:6px">
            <div class="progress-fill ${Number(usagePct) > 80 ? 'red' : 'green'}" style="width:${usagePct}%"></div>
          </div>
        </div>
        <div class="settings-row">
          <div>
            <div class="settings-label">Export Backup</div>
            <div class="settings-desc">Download all data as JSON</div>
          </div>
          <button class="btn btn-secondary" onclick="App.exportBackup()">📥 Export JSON</button>
        </div>
        <div class="settings-row">
          <div>
            <div class="settings-label">Import Backup</div>
            <div class="settings-desc">Restore from a JSON backup</div>
          </div>
          <div class="file-input-wrapper">
            <button class="btn btn-secondary">📤 Import JSON</button>
            <input type="file" accept=".json" onchange="App.importBackup(event)">
          </div>
        </div>
        <div class="settings-row">
          <div>
            <div class="settings-label">Clear All Data</div>
            <div class="settings-desc">Remove all expenses, budgets, and settings</div>
          </div>
          <button class="btn btn-danger btn-sm" onclick="App.clearAllData()">🗑️ Clear All</button>
        </div>
      </div>
    `;
  }

  function updateCurrency(val) {
    Store.updateSettings({ currency: val.trim() || '₹' });
    Utils.showToast('Currency updated');
  }

  function addCategory() {
    const html = `
      <div class="modal-header">
        <h3 class="modal-title">Add Category</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <form onsubmit="App.saveCategory(event)">
        <div class="form-group">
          <label class="form-label">Name</label>
          <input type="text" class="form-input" name="name" required placeholder="e.g. Travel">
        </div>
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Color</label>
            <input type="color" class="form-input" name="color" value="#4CAF50" style="height:42px;padding:4px">
          </div>
          <div class="form-group">
            <label class="form-label">Icon (emoji)</label>
            <input type="text" class="form-input" name="icon" placeholder="✈️" maxlength="4" value="📦">
          </div>
        </div>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button type="submit" class="btn btn-primary">Add</button>
        </div>
      </form>
    `;
    openModal(html);
  }

  function editCategory(name) {
    const cat = Store.getCategories().find(c => c.name === name);
    if (!cat) return;
    const html = `
      <div class="modal-header">
        <h3 class="modal-title">Edit Category</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <form onsubmit="App.saveCategory(event, '${Utils.escapeHtml(name)}')">
        <div class="form-group">
          <label class="form-label">Name</label>
          <input type="text" class="form-input" name="name" required value="${Utils.escapeHtml(cat.name)}">
        </div>
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Color</label>
            <input type="color" class="form-input" name="color" value="${cat.color}" style="height:42px;padding:4px">
          </div>
          <div class="form-group">
            <label class="form-label">Icon (emoji)</label>
            <input type="text" class="form-input" name="icon" maxlength="4" value="${cat.icon}">
          </div>
        </div>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button type="submit" class="btn btn-primary">Update</button>
        </div>
      </form>
    `;
    openModal(html);
  }

  function saveCategory(e, oldName = null) {
    e.preventDefault();
    const form = e.target;
    const cat = {
      name: form.name.value.trim(),
      color: form.color.value,
      icon: form.icon.value || '📦'
    };
    if (!cat.name) return;

    if (oldName) {
      Store.updateCategory(oldName, cat);
      Utils.showToast('Category updated');
    } else {
      if (Store.getCategories().find(c => c.name === cat.name)) {
        Utils.showToast('Category already exists', 'error');
        return;
      }
      Store.addCategory(cat);
      Utils.showToast('Category added');
    }
    closeModal();
    renderSettings();
  }

  function deleteCategory(name) {
    const expenses = Store.getExpenses().filter(e => e.category === name);
    const msg = expenses.length
      ? `"${name}" has ${expenses.length} expenses. They will be moved to "Other". Continue?`
      : `Delete category "${name}"?`;
    confirm(msg, () => {
      if (expenses.length) {
        expenses.forEach(e => Store.updateExpense(e.id, { category: 'Other' }));
      }
      Store.deleteCategory(name);
      Utils.showToast('Category deleted');
      renderSettings();
    });
  }

  function exportBackup() {
    const json = Store.exportBackup();
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `expense-tracker-backup-${Utils.today()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    Utils.showToast('Backup exported');
  }

  function importBackup(event) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = '';
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        Store.importBackup(e.target.result);
        Utils.showToast('Backup restored successfully');
        switchTab(activeTab); // Re-render current tab
      } catch (err) {
        Utils.showToast('Invalid backup file', 'error');
      }
    };
    reader.readAsText(file);
  }

  function clearAllData() {
    confirm('This will permanently delete ALL your data. Are you sure?', () => {
      Store.clearAll();
      Utils.showToast('All data cleared');
      switchTab('dashboard');
    });
  }

  // Keyboard shortcuts — use capture phase so we get the event first
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeModal();
      document.querySelector('.confirm-overlay')?.remove();
      return;
    }

    // Ctrl+I (Windows/Linux) or Cmd+I (macOS) => open Excel import file picker
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i') {
      e.preventDefault();
      e.stopPropagation();

      // Switch to expenses tab if not already there
      if (activeTab !== 'expenses') switchTab('expenses');

      // Create a temporary file input and click it immediately
      // This is more reliable than clicking the hidden DOM input
      const tmp = document.createElement('input');
      tmp.type = 'file';
      tmp.accept = '.csv,.tsv,.xls,.xlsx';
      tmp.style.display = 'none';
      tmp.addEventListener('change', (ev) => {
        Excel.handleImport(ev);
        tmp.remove();
      });
      document.body.appendChild(tmp);
      tmp.click();
    }
  }, true);

  return {
    init, switchTab, toggleTheme, openModal, closeModal, confirm,
    updateCurrency, addCategory, editCategory, saveCategory, deleteCategory,
    exportBackup, importBackup, clearAllData
  };
})();

// Boot
document.addEventListener('DOMContentLoaded', App.init);
