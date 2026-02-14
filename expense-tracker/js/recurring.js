/**
 * Recurring Expenses Module
 */
const Recurring = (() => {

  function render() {
    const container = document.getElementById('tab-recurring');
    const items = Store.getRecurring();

    container.innerHTML = `
      <div class="toolbar">
        <div class="toolbar-left">
          <button class="btn btn-primary" onclick="Recurring.openModal()">+ Add Recurring</button>
          <button class="btn btn-secondary" onclick="Recurring.generateAll()">🔄 Generate Due Expenses</button>
        </div>
      </div>
      <div class="recurring-list" id="recurring-list"></div>
    `;

    const listEl = document.getElementById('recurring-list');
    if (!items.length) {
      listEl.innerHTML = `<div class="empty-state">
        <div class="icon">🔄</div>
        <div class="title">No recurring expenses</div>
        <div class="desc">Add subscriptions, bills, and regular payments to track them automatically</div>
      </div>`;
      return;
    }

    listEl.innerHTML = items.map(item => {
      const nextDue = getNextDueDate(item);
      return `
        <div class="recurring-item" style="opacity:${item.enabled ? 1 : 0.5}">
          <div style="display:flex;align-items:center;gap:12px">
            <span style="font-size:1.5rem">${Store.getCategoryIcon(item.category)}</span>
            <div class="recurring-info">
              <div class="recurring-name">${Utils.escapeHtml(item.name)}</div>
              <div class="recurring-meta">
                <span>${Utils.escapeHtml(item.category)}</span>
                <span>Every ${item.frequency}</span>
                ${nextDue ? `<span>Next: ${Utils.formatDate(nextDue)}</span>` : ''}
                ${item.lastGenerated ? `<span>Last: ${Utils.formatDate(item.lastGenerated)}</span>` : ''}
              </div>
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:16px">
            <span class="recurring-amount">${Utils.formatCurrency(item.amount)}</span>
            <div class="recurring-actions">
              <label class="toggle" title="${item.enabled ? 'Disable' : 'Enable'}">
                <input type="checkbox" ${item.enabled ? 'checked' : ''}
                  onchange="Recurring.toggleEnabled('${item.id}', this.checked)">
                <span class="toggle-slider"></span>
              </label>
              <button class="btn-icon" onclick="Recurring.openModal('${item.id}')" title="Edit">✏️</button>
              <button class="btn-icon danger" onclick="Recurring.confirmDelete('${item.id}')" title="Delete">🗑️</button>
            </div>
          </div>
        </div>
      `;
    }).join('');
  }

  function getNextDueDate(item) {
    if (!item.enabled) return null;
    const startDate = new Date(item.startDate + 'T00:00:00');
    const lastGen = item.lastGenerated ? new Date(item.lastGenerated + 'T00:00:00') : null;
    const baseDate = lastGen || startDate;
    const next = new Date(baseDate);

    switch (item.frequency) {
      case 'weekly': next.setDate(next.getDate() + 7); break;
      case 'monthly': next.setMonth(next.getMonth() + 1); break;
      case 'yearly': next.setFullYear(next.getFullYear() + 1); break;
    }
    return Utils.toISODate(next);
  }

  function generateAll() {
    const items = Store.getRecurring().filter(r => r.enabled);
    const today = Utils.today();
    let count = 0;

    items.forEach(item => {
      let nextDue = getNextDueDate(item);
      // Generate all past-due entries
      while (nextDue && nextDue <= today) {
        Store.addExpense({
          date: nextDue,
          amount: item.amount,
          category: item.category,
          description: item.name + ' (recurring)',
          paymentMethod: item.paymentMethod || 'Other',
          isRecurring: true
        });
        Store.updateRecurring(item.id, { lastGenerated: nextDue });
        item.lastGenerated = nextDue;
        nextDue = getNextDueDate(item);
        count++;
      }
    });

    if (count > 0) {
      Utils.showToast(`Generated ${count} recurring expense${count > 1 ? 's' : ''}`);
    } else {
      Utils.showToast('No recurring expenses are due yet', 'info');
    }
    render();
    Dashboard.render();
  }

  function openModal(editId = null) {
    const item = editId ? Store.getRecurring().find(r => r.id === editId) : null;
    const isEdit = !!item;
    const categories = Store.getCategories();

    const html = `
      <div class="modal-header">
        <h3 class="modal-title">${isEdit ? 'Edit' : 'Add'} Recurring Expense</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <form id="recurring-form" onsubmit="Recurring.save(event, '${editId || ''}')">
        <div class="form-group">
          <label class="form-label">Name</label>
          <input type="text" class="form-input" name="name" required placeholder="e.g. Netflix"
            value="${item ? Utils.escapeHtml(item.name) : ''}">
        </div>
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Amount</label>
            <input type="number" step="0.01" min="0" class="form-input" name="amount" required
              placeholder="0.00" value="${item ? item.amount : ''}">
          </div>
          <div class="form-group">
            <label class="form-label">Frequency</label>
            <select class="form-select" name="frequency">
              ${['weekly', 'monthly', 'yearly'].map(f =>
                `<option value="${f}" ${item && item.frequency === f ? 'selected' : ''}>${f.charAt(0).toUpperCase() + f.slice(1)}</option>`
              ).join('')}
            </select>
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Category</label>
            <select class="form-select" name="category">
              ${categories.map(c => `<option value="${c.name}" ${item && item.category === c.name ? 'selected' : ''}>${c.icon} ${c.name}</option>`).join('')}
            </select>
          </div>
          <div class="form-group">
            <label class="form-label">Start Date</label>
            <input type="date" class="form-input" name="startDate" required
              value="${item ? item.startDate : Utils.today()}">
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">Payment Method</label>
          <select class="form-select" name="paymentMethod">
            ${['Cash', 'Credit Card', 'Debit Card', 'UPI', 'Net Banking', 'Other'].map(m =>
              `<option ${item && item.paymentMethod === m ? 'selected' : ''}>${m}</option>`
            ).join('')}
          </select>
        </div>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button type="submit" class="btn btn-primary">${isEdit ? 'Update' : 'Add'}</button>
        </div>
      </form>
    `;
    App.openModal(html);
  }

  function save(e, editId) {
    e.preventDefault();
    const form = e.target;
    const data = {
      name: form.name.value.trim(),
      amount: parseFloat(form.amount.value),
      frequency: form.frequency.value,
      category: form.category.value,
      startDate: form.startDate.value,
      paymentMethod: form.paymentMethod.value
    };

    if (editId) {
      Store.updateRecurring(editId, data);
      Utils.showToast('Recurring expense updated');
    } else {
      Store.addRecurring(data);
      Utils.showToast('Recurring expense added');
    }
    App.closeModal();
    render();
  }

  function toggleEnabled(id, enabled) {
    Store.updateRecurring(id, { enabled });
    render();
  }

  function confirmDelete(id) {
    App.confirm('Delete this recurring expense?', () => {
      Store.deleteRecurring(id);
      Utils.showToast('Recurring expense deleted');
      render();
    });
  }

  return { render, openModal, save, generateAll, toggleEnabled, confirmDelete };
})();
