/**
 * Budgets Module - Per-category budget limits with visual progress
 */
const Budgets = (() => {

  function render() {
    const container = document.getElementById('tab-budgets');
    const budgets = Store.getBudgets();
    const categories = Store.getCategories();
    const { start, end } = Utils.getCurrentMonthRange();
    const monthExpenses = Store.getExpenses().filter(e => e.date >= start && e.date <= end);

    // Calculate spending per category
    const spending = {};
    monthExpenses.forEach(e => {
      spending[e.category] = (spending[e.category] || 0) + Number(e.amount);
    });

    const now = new Date();
    const monthName = Utils.getMonthName(now.getMonth()) + ' ' + now.getFullYear();

    container.innerHTML = `
      <div class="toolbar">
        <div class="toolbar-left">
          <h3 style="font-weight:600;margin:0">Budget for ${monthName}</h3>
        </div>
        <div class="toolbar-right">
          <button class="btn btn-primary" onclick="Budgets.openModal()">+ Set Budget</button>
        </div>
      </div>

      ${renderOverview(budgets, spending)}

      <div class="budget-grid" id="budget-grid"></div>
    `;

    const gridEl = document.getElementById('budget-grid');
    const budgetEntries = Object.entries(budgets);

    if (!budgetEntries.length) {
      gridEl.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
        <div class="icon">🎯</div>
        <div class="title">No budgets set</div>
        <div class="desc">Set monthly spending limits per category to stay on track</div>
      </div>`;
      return;
    }

    gridEl.innerHTML = budgetEntries.map(([cat, limit]) => {
      const spent = spending[cat] || 0;
      const pct = limit > 0 ? Math.min((spent / limit) * 100, 100) : 0;
      const remaining = limit - spent;
      let colorClass = 'green';
      if (pct >= 100) colorClass = 'red';
      else if (pct >= 80) colorClass = 'yellow';

      return `
        <div class="budget-card">
          <div class="budget-card-header">
            <div class="budget-category">
              <span>${Store.getCategoryIcon(cat)}</span>
              <span>${Utils.escapeHtml(cat)}</span>
            </div>
            <div style="display:flex;gap:6px">
              <button class="btn-icon" onclick="Budgets.openModal('${Utils.escapeHtml(cat)}')" title="Edit">✏️</button>
              <button class="btn-icon danger" onclick="Budgets.confirmDelete('${Utils.escapeHtml(cat)}')" title="Remove">🗑️</button>
            </div>
          </div>
          <div class="progress-bar">
            <div class="progress-fill ${colorClass}" style="width:${pct}%"></div>
          </div>
          <div class="budget-info">
            <span><span class="spent">${Utils.formatCurrency(spent)}</span> of ${Utils.formatCurrency(limit)}</span>
            <span style="color:${remaining >= 0 ? 'var(--primary)' : 'var(--danger)'}">
              ${remaining >= 0 ? Utils.formatCurrency(remaining) + ' left' : Utils.formatCurrency(Math.abs(remaining)) + ' over!'}
            </span>
          </div>
          ${pct >= 100 ? '<div style="margin-top:8px;font-size:0.8rem;color:var(--danger);font-weight:500">⚠️ Budget exceeded!</div>' :
            pct >= 80 ? '<div style="margin-top:8px;font-size:0.8rem;color:var(--warning);font-weight:500">⚠️ Approaching limit</div>' : ''}
        </div>
      `;
    }).join('');
  }

  function renderOverview(budgets, spending) {
    const totalBudget = Object.values(budgets).reduce((s, v) => s + v, 0);
    const totalSpent = Object.entries(budgets).reduce((s, [cat]) => s + (spending[cat] || 0), 0);
    if (totalBudget === 0) return '';

    const pct = Math.min((totalSpent / totalBudget) * 100, 100);
    let colorClass = 'green';
    if (pct >= 100) colorClass = 'red';
    else if (pct >= 80) colorClass = 'yellow';

    return `
      <div class="card" style="margin-bottom:20px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <span style="font-weight:600">Overall Budget</span>
          <span style="font-size:0.9rem;color:var(--text-secondary)">${Math.round(pct)}% used</span>
        </div>
        <div class="progress-bar" style="height:12px">
          <div class="progress-fill ${colorClass}" style="width:${pct}%"></div>
        </div>
        <div class="budget-info" style="margin-top:8px">
          <span><span class="spent">${Utils.formatCurrency(totalSpent)}</span> spent</span>
          <span>${Utils.formatCurrency(totalBudget)} total budget</span>
        </div>
      </div>
    `;
  }

  function openModal(editCategory = null) {
    const budgets = Store.getBudgets();
    const categories = Store.getCategories();
    const isEdit = !!editCategory;
    const currentAmount = isEdit ? budgets[editCategory] || 0 : '';

    // For new budget, filter out categories that already have budgets
    const availableCategories = isEdit ? categories :
      categories.filter(c => budgets[c.name] === undefined);

    if (!isEdit && !availableCategories.length) {
      Utils.showToast('All categories already have budgets', 'info');
      return;
    }

    const html = `
      <div class="modal-header">
        <h3 class="modal-title">${isEdit ? 'Edit' : 'Set'} Budget</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <form onsubmit="Budgets.save(event, '${isEdit ? Utils.escapeHtml(editCategory) : ''}')">
        <div class="form-group">
          <label class="form-label">Category</label>
          ${isEdit ?
            `<input type="text" class="form-input" value="${Utils.escapeHtml(editCategory)}" disabled>
             <input type="hidden" name="category" value="${Utils.escapeHtml(editCategory)}">` :
            `<select class="form-select" name="category" required>
              ${availableCategories.map(c => `<option value="${c.name}">${c.icon} ${c.name}</option>`).join('')}
            </select>`
          }
        </div>
        <div class="form-group">
          <label class="form-label">Monthly Budget Amount</label>
          <input type="number" step="0.01" min="0" class="form-input" name="amount" required
            placeholder="0.00" value="${currentAmount}">
        </div>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button type="submit" class="btn btn-primary">${isEdit ? 'Update' : 'Set'} Budget</button>
        </div>
      </form>
    `;
    App.openModal(html);
  }

  function save(e, editCategory) {
    e.preventDefault();
    const form = e.target;
    const category = editCategory || form.category.value;
    const amount = parseFloat(form.amount.value);
    Store.setBudget(category, amount);
    Utils.showToast('Budget saved');
    App.closeModal();
    render();
  }

  function confirmDelete(category) {
    App.confirm(`Remove budget for "${category}"?`, () => {
      Store.deleteBudget(category);
      Utils.showToast('Budget removed');
      render();
    });
  }

  return { render, openModal, save, confirmDelete };
})();
