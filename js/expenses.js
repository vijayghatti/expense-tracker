/**
 * Expenses Module - CRUD, filtering, sorting, table rendering
 */
const Expenses = (() => {
  let sortField = 'date';
  let sortDir = 'desc';
  let filterCategory = '';
  let filterSearch = '';
  let filterDateFrom = '';
  let filterDateTo = '';
  let currentPage = 1;
  const pageSize = 20;

  function render() {
    const container = document.getElementById('tab-expenses');
    const { start, end } = Utils.getCurrentMonthRange();
    container.innerHTML = `
      <div class="toolbar">
      <div class="toolbar-left">
      <button class="btn btn-primary" onclick="Expenses.openModal()">+ Add Expense</button>
      <div class="file-input-wrapper">
      <button class="btn btn-secondary" title="Ctrl+R (Windows) / ⌘+R (Mac) to import quickly">📥 Import</button>
      <input id="excel-import-input" type="file" accept=".csv,.tsv,.xls,.xlsx" onchange="Excel.handleImport(event)">
      </div>
      <button class="btn btn-secondary" onclick="Excel.exportExpenses()">📤 Export</button>
      </div>
        <div class="toolbar-right">
          <input type="text" class="filter-input" placeholder="🔍 Search..." id="expense-search"
            oninput="Expenses.setSearch(this.value)" value="${Utils.escapeHtml(filterSearch)}">
        </div>
      </div>
      <div class="filters-bar">
        <select class="filter-input" id="filter-category" onchange="Expenses.setCategory(this.value)">
          <option value="">All Categories</option>
          ${Store.getCategories().map(c =>
            `<option value="${c.name}" ${filterCategory === c.name ? 'selected' : ''}>${c.icon} ${c.name}</option>`
          ).join('')}
        </select>
        <input type="date" class="filter-input" id="filter-from" value="${filterDateFrom}"
          onchange="Expenses.setDateFrom(this.value)" title="From date">
        <input type="date" class="filter-input" id="filter-to" value="${filterDateTo}"
          onchange="Expenses.setDateTo(this.value)" title="To date">
        ${(filterCategory || filterSearch || filterDateFrom || filterDateTo) ?
          '<button class="btn btn-sm btn-secondary" onclick="Expenses.clearFilters()">✕ Clear</button>' : ''}
      </div>
      <div class="table-wrapper">
        <table>
          <thead>
            <tr>
              ${renderSortHeader('date', 'Date')}
              ${renderSortHeader('description', 'Description')}
              ${renderSortHeader('category', 'Category')}
              ${renderSortHeader('amount', 'Amount')}
              <th>Payment</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody id="expense-tbody"></tbody>
        </table>
      </div>
      <div id="expense-pagination" style="margin-top:12px;display:flex;justify-content:space-between;align-items:center;"></div>
    `;
    renderRows();
  }

  function renderSortHeader(field, label) {
    const isActive = sortField === field;
    const icon = isActive ? (sortDir === 'asc' ? '↑' : '↓') : '↕';
    return `<th class="${isActive ? 'sorted' : ''}" onclick="Expenses.sort('${field}')">
      ${label} <span class="sort-icon">${icon}</span>
    </th>`;
  }

  function getFilteredExpenses() {
    let expenses = Store.getExpenses();
    if (filterCategory) expenses = expenses.filter(e => e.category === filterCategory);
    if (filterSearch) {
      const q = filterSearch.toLowerCase();
      expenses = expenses.filter(e =>
        (e.description || '').toLowerCase().includes(q) ||
        (e.category || '').toLowerCase().includes(q) ||
        (e.paymentMethod || '').toLowerCase().includes(q)
      );
    }
    if (filterDateFrom) expenses = expenses.filter(e => e.date >= filterDateFrom);
    if (filterDateTo) expenses = expenses.filter(e => e.date <= filterDateTo);

    // Sort
    expenses.sort((a, b) => {
      let va = a[sortField], vb = b[sortField];
      if (sortField === 'amount') { va = Number(va); vb = Number(vb); }
      if (va < vb) return sortDir === 'asc' ? -1 : 1;
      if (va > vb) return sortDir === 'asc' ? 1 : -1;
      return 0;
    });
    return expenses;
  }

  function renderRows() {
    const filtered = getFilteredExpenses();
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    if (currentPage > totalPages) currentPage = totalPages;
    const start = (currentPage - 1) * pageSize;
    const page = filtered.slice(start, start + pageSize);

    const tbody = document.getElementById('expense-tbody');
    if (!page.length) {
      tbody.innerHTML = `<tr><td colspan="6">
        <div class="empty-state">
          <div class="icon">📝</div>
          <div class="title">No expenses found</div>
          <div class="desc">Add your first expense or adjust your filters</div>
        </div>
      </td></tr>`;
    } else {
      tbody.innerHTML = page.map(e => `
        <tr>
          <td>${Utils.formatDate(e.date)}</td>
          <td>${Utils.escapeHtml(e.description || '—')}</td>
          <td><span class="category-badge" style="background:${Store.getCategoryColor(e.category)}20;color:${Store.getCategoryColor(e.category)}">
            ${Store.getCategoryIcon(e.category)} ${Utils.escapeHtml(e.category)}</span></td>
          <td class="amount-cell">${Utils.formatCurrency(e.amount)}</td>
          <td style="color:var(--text-secondary)">${Utils.escapeHtml(e.paymentMethod || '—')}</td>
          <td class="action-btns">
            <button class="btn-icon" onclick="Expenses.openModal('${e.id}')" title="Edit">✏️</button>
            <button class="btn-icon danger" onclick="Expenses.confirmDelete('${e.id}')" title="Delete">🗑️</button>
          </td>
        </tr>
      `).join('');
    }

    // Pagination
    const pagEl = document.getElementById('expense-pagination');
    pagEl.innerHTML = `
      <span style="color:var(--text-secondary);font-size:0.85rem">${filtered.length} expense${filtered.length !== 1 ? 's' : ''}</span>
      ${totalPages > 1 ? `<div style="display:flex;gap:6px;align-items:center">
        <button class="btn-icon" onclick="Expenses.prevPage()" ${currentPage <= 1 ? 'disabled' : ''}>◀</button>
        <span style="font-size:0.85rem">${currentPage} / ${totalPages}</span>
        <button class="btn-icon" onclick="Expenses.nextPage()" ${currentPage >= totalPages ? 'disabled' : ''}>▶</button>
      </div>` : ''}
    `;
  }

  function sort(field) {
    if (sortField === field) sortDir = sortDir === 'asc' ? 'desc' : 'asc';
    else { sortField = field; sortDir = field === 'date' ? 'desc' : 'asc'; }
    render();
  }

  function setSearch(val) { filterSearch = val; currentPage = 1; renderRows(); }
  function setCategory(val) { filterCategory = val; currentPage = 1; renderRows(); }
  function setDateFrom(val) { filterDateFrom = val; currentPage = 1; renderRows(); }
  function setDateTo(val) { filterDateTo = val; currentPage = 1; renderRows(); }
  function clearFilters() { filterCategory = ''; filterSearch = ''; filterDateFrom = ''; filterDateTo = ''; currentPage = 1; render(); }
  function prevPage() { if (currentPage > 1) { currentPage--; renderRows(); } }
  function nextPage() { currentPage++; renderRows(); }

  function openModal(editId = null) {
    const expense = editId ? Store.getExpenses().find(e => e.id === editId) : null;
    const isEdit = !!expense;
    const categories = Store.getCategories();

    const html = `
      <div class="modal-header">
        <h3 class="modal-title">${isEdit ? 'Edit' : 'Add'} Expense</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <form id="expense-form" onsubmit="Expenses.saveExpense(event, '${editId || ''}')">
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Date</label>
            <input type="date" class="form-input" name="date" required value="${expense ? expense.date : Utils.today()}">
          </div>
          <div class="form-group">
            <label class="form-label">Amount</label>
            <input type="number" step="0.01" min="0" class="form-input" name="amount" required
              placeholder="0.00" value="${expense ? expense.amount : ''}">
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">Category</label>
          <select class="form-select" name="category" required>
            ${categories.map(c => `<option value="${c.name}" ${expense && expense.category === c.name ? 'selected' : ''}>${c.icon} ${c.name}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Description</label>
          <input type="text" class="form-input" name="description" placeholder="e.g. Grocery shopping"
            value="${expense ? Utils.escapeHtml(expense.description || '') : ''}">
        </div>
        <div class="form-group">
          <label class="form-label">Payment Method</label>
          <select class="form-select" name="paymentMethod">
            ${['Cash', 'Credit Card', 'Debit Card', 'UPI', 'Net Banking', 'Other'].map(m =>
              `<option ${expense && expense.paymentMethod === m ? 'selected' : ''}>${m}</option>`
            ).join('')}
          </select>
        </div>
        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
          <button type="submit" class="btn btn-primary">${isEdit ? 'Update' : 'Add'} Expense</button>
        </div>
      </form>
    `;
    App.openModal(html);
  }

  function saveExpense(e, editId) {
    e.preventDefault();
    const form = e.target;
    const data = {
      date: form.date.value,
      amount: parseFloat(form.amount.value),
      category: form.category.value,
      description: form.description.value.trim(),
      paymentMethod: form.paymentMethod.value,
      isRecurring: false
    };

    if (editId) {
      Store.updateExpense(editId, data);
      Utils.showToast('Expense updated');
    } else {
      Store.addExpense(data);
      Utils.showToast('Expense added');
    }
    App.closeModal();
    render();
    Dashboard.render(); // refresh dashboard too
  }

  function confirmDelete(id) {
    App.confirm('Delete this expense?', () => {
      Store.deleteExpense(id);
      Utils.showToast('Expense deleted');
      render();
      Dashboard.render();
    });
  }

  return {
    render, sort, setSearch, setCategory, setDateFrom, setDateTo,
    clearFilters, openModal, saveExpense, confirmDelete,
    prevPage, nextPage, getFilteredExpenses
  };
})();
