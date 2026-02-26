/**
 * Excel Import/Export Module
 */
const Excel = (() => {

  function exportExpenses() {
    const expenses = Expenses.getFilteredExpenses();
    if (!expenses.length) {
      Utils.showToast('No expenses to export', 'info');
      return;
    }

    const rows = expenses.map(e => ({
      Date: e.date,
      Description: e.description || '',
      Category: e.category,
      Amount: Number(e.amount),
      'Payment Method': e.paymentMethod || '',
      Recurring: e.isRecurring ? 'Yes' : 'No'
    }));

    // Show export options
    const html = `
      <div class="modal-header">
        <h3 class="modal-title">Export Expenses</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <p style="margin-bottom:16px;color:var(--text-secondary)">${rows.length} expenses will be exported</p>
      <div style="display:flex;flex-direction:column;gap:10px">
        <button class="btn btn-primary" onclick="Excel.doExport('excel')" style="width:100%">
          📊 Export as Excel (.xlsx)
        </button>
        <button class="btn btn-secondary" onclick="Excel.doExport('csv')" style="width:100%">
          📄 Export as CSV (.csv)
        </button>
      </div>
    `;
    App.openModal(html);
  }

  function formatMonth(yearMonth) {
    if (!yearMonth || yearMonth === 'Unknown') return 'Unknown';
    const [year, month] = yearMonth.split('-');
    const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const m = parseInt(month, 10);
    return `${names[m - 1] || month} ${year}`;
  }

  function doExport(format) {
    const expenses = Expenses.getFilteredExpenses();
    const rows = expenses.map(e => ({
      Date: e.date,
      Description: e.description || '',
      Category: e.category,
      Amount: Number(e.amount),
      'Payment Method': e.paymentMethod || '',
      Recurring: e.isRecurring ? 'Yes' : 'No'
    }));

    const timestamp = new Date().toISOString().slice(0, 10);
    if (format === 'excel') {
      // Compute category totals
      const catTotals = {};
      expenses.forEach(e => {
        const cat = e.category || 'Other';
        catTotals[cat] = (catTotals[cat] || 0) + Number(e.amount);
      });
      const catData = Object.entries(catTotals)
        .map(([name, amount]) => ({ name, amount: Math.round(amount * 100) / 100 }))
        .sort((a, b) => b.amount - a.amount);

      // Compute monthly totals
      const monthTotals = {};
      expenses.forEach(e => {
        const month = e.date ? e.date.slice(0, 7) : 'Unknown';
        monthTotals[month] = (monthTotals[month] || 0) + Number(e.amount);
      });
      const monthData = Object.entries(monthTotals)
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([month, amount]) => ({
          month: formatMonth(month),
          amount: Math.round(amount * 100) / 100
        }));

      MiniXLSX.exportToExcelWithCharts(rows, catData, monthData, `expenses_${timestamp}.xlsx`);
    } else {
      MiniXLSX.exportToCSV(rows, `expenses_${timestamp}.csv`);
    }
    App.closeModal();
    Utils.showToast('Export complete');
  }

  function handleImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = ''; // Reset input

    MiniXLSX.importFile(file).then(rows => {
      if (!rows.length) {
        Utils.showToast('No data found in file', 'error');
        return;
      }
      showMappingUI(rows, file.name);
    }).catch(err => {
      Utils.showToast('Error reading file: ' + err.message, 'error');
    });
  }

  function showMappingUI(rows, filename) {
    const headers = Object.keys(rows[0]);
    const fields = [
      { key: 'date', label: 'Date', required: true },
      { key: 'amount', label: 'Amount', required: true },
      { key: 'category', label: 'Category', required: false },
      { key: 'description', label: 'Description', required: false },
      { key: 'paymentMethod', label: 'Payment Method', required: false }
    ];

    // Auto-detect mapping
    const autoMap = {};
    fields.forEach(f => {
      const match = headers.find(h =>
        h.toLowerCase().includes(f.key.toLowerCase()) ||
        h.toLowerCase().replace(/[_\s]/g, '') === f.key.toLowerCase()
      );
      if (match) autoMap[f.key] = match;
    });

    // Extra detection
    if (!autoMap.date) {
      const dateCol = headers.find(h => /date|time|when/i.test(h));
      if (dateCol) autoMap.date = dateCol;
    }
    if (!autoMap.amount) {
      const amtCol = headers.find(h => /amount|total|price|cost|value/i.test(h));
      if (amtCol) autoMap.amount = amtCol;
    }
    if (!autoMap.description) {
      const descCol = headers.find(h => /desc|note|detail|memo|name|item/i.test(h));
      if (descCol) autoMap.description = descCol;
    }

    // Store data for later
    window._importData = rows;

    const html = `
      <div class="modal-header">
        <h3 class="modal-title">Import from ${Utils.escapeHtml(filename)}</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <p style="margin-bottom:12px;color:var(--text-secondary)">
        Found ${rows.length} rows. Map the columns to expense fields:
      </p>
      <table class="mapping-table">
        <thead><tr><th>Expense Field</th><th>File Column</th></tr></thead>
        <tbody>
          ${fields.map(f => `
            <tr>
              <td>${f.label} ${f.required ? '<span style="color:var(--danger)">*</span>' : ''}</td>
              <td>
                <select class="mapping-select" id="map-${f.key}">
                  <option value="">— Skip —</option>
                  ${headers.map(h => `<option value="${Utils.escapeHtml(h)}" ${autoMap[f.key] === h ? 'selected' : ''}>${Utils.escapeHtml(h)}</option>`).join('')}
                </select>
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <p style="font-size:0.8rem;color:var(--text-muted);margin-top:8px">
        Preview: first row → Date: "${rows[0][autoMap.date] || '?'}", Amount: "${rows[0][autoMap.amount] || '?'}"
      </p>
      <div class="modal-actions">
        <button class="btn btn-secondary" onclick="App.closeModal()">Cancel</button>
        <button class="btn btn-primary" onclick="Excel.doImport()">Import ${rows.length} Rows</button>
      </div>
    `;
    App.openModal(html);
  }

  function doImport() {
    const rows = window._importData;
    if (!rows) return;

    const dateCol = document.getElementById('map-date').value;
    const amountCol = document.getElementById('map-amount').value;
    const categoryCol = document.getElementById('map-category').value;
    const descCol = document.getElementById('map-description').value;
    const paymentCol = document.getElementById('map-paymentMethod').value;

    if (!dateCol || !amountCol) {
      Utils.showToast('Date and Amount columns are required', 'error');
      return;
    }

    // Replace any existing expenses so only the latest imported file's
    // data is kept in the app.
    Store.setExpenses([]);

    const categories = Store.getCategories().map(c => c.name);
    let imported = 0;
    let skipped = 0;

    rows.forEach(row => {
      const dateVal = String(row[dateCol] || '').trim();
      const amountVal = parseFloat(String(row[amountCol] || '').replace(/[^0-9.\-]/g, ''));

      if (!dateVal || isNaN(amountVal)) { skipped++; return; }

      // Parse date - try multiple formats
      let date = parseDate(dateVal);
      if (!date) { skipped++; return; }

      let category = categoryCol ? String(row[categoryCol] || '').trim() : 'Other';
      if (!categories.includes(category)) category = 'Other';

      Store.addExpense({
        date: date,
        amount: Math.abs(amountVal),
        category: category,
        description: descCol ? String(row[descCol] || '').trim() : '',
        paymentMethod: paymentCol ? String(row[paymentCol] || '').trim() : 'Other',
        isRecurring: false
      });
      imported++;
    });

    delete window._importData;
    App.closeModal();
    Utils.showToast(`Imported ${imported} expenses${skipped ? `, ${skipped} skipped` : ''}`);
    Expenses.render();
    Dashboard.render();
  }

  function parseDate(str) {
    // Try YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    // Try DD/MM/YYYY or DD-MM-YYYY
    let m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    // Try MM/DD/YYYY
    m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m && parseInt(m[1]) <= 12) return `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
    // Try Date.parse
    const d = new Date(str);
    if (!isNaN(d.getTime())) return Utils.toISODate(d);
    return null;
  }

  return { exportExpenses, doExport, handleImport, doImport };
})();
