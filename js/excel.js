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

      // Pass ALL categories (not just ones with expenses) so the Excel
      // dropdown shows the full list regardless of what's currently recorded.
      const allCategoryNames = Store.getCategories().map(c => c.name);
      MiniXLSX.exportToExcelWithCharts(rows, catData, monthData, allCategoryNames, `expenses_${timestamp}.xlsx`);
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

  const VALID_PAYMENT_METHODS = ['Cash', 'Credit Card', 'Debit Card', 'UPI', 'Net Banking', 'Other'];

  /**
   * Validates all rows against required field rules.
   * Returns an array of error strings; empty array means all rows are valid.
   */
  function validateImportRows(rows, dateCol, amountCol, categoryCol, paymentCol) {
    const validCategories = Store.getCategories().map(c => c.name);
    const errors = [];
    const MAX_ERRORS = 25; // cap to keep the modal readable

    rows.forEach((row, idx) => {
      if (errors.length >= MAX_ERRORS) return;
      const rowNum = idx + 2; // row 1 = header

      // --- Date (required) ---
      const dateVal = String(row[dateCol] || '').trim();
      if (!dateVal) {
        errors.push(`Row ${rowNum}: Date is required`);
      } else if (!parseDate(dateVal)) {
        errors.push(`Row ${rowNum}: Date "${dateVal}" is not a recognised date — use YYYY-MM-DD or DD/MM/YYYY`);
      }

      // --- Amount (required, must be a positive number) ---
      const rawAmt = String(row[amountCol] || '').trim();
      if (!rawAmt) {
        errors.push(`Row ${rowNum}: Amount is required`);
      } else {
        const num = parseFloat(rawAmt.replace(/[^0-9.\-]/g, ''));
        if (isNaN(num)) {
          errors.push(`Row ${rowNum}: Amount "${rawAmt}" is not a valid number`);
        } else if (num <= 0) {
          errors.push(`Row ${rowNum}: Amount must be greater than 0 (got ${num})`);
        }
      }

      // --- Category (optional, but must match a known category if provided) ---
      if (categoryCol) {
        const catVal = String(row[categoryCol] || '').trim();
        if (catVal && !validCategories.includes(catVal)) {
          errors.push(`Row ${rowNum}: Category "${catVal}" is not recognised — valid values: ${validCategories.join(', ')}`);
        }
      }

      // --- Payment Method (optional, but must match a known method if provided) ---
      if (paymentCol) {
        const payVal = String(row[paymentCol] || '').trim();
        if (payVal && !VALID_PAYMENT_METHODS.includes(payVal)) {
          errors.push(`Row ${rowNum}: Payment method "${payVal}" is not recognised — valid values: ${VALID_PAYMENT_METHODS.join(', ')}`);
        }
      }
    });

    return errors;
  }

  /** Renders a validation-error modal and aborts the import. */
  function showValidationErrors(errors, totalRows) {
    const hasMore = errors.length >= 25;
    const html = `
      <div class="modal-header">
        <h3 class="modal-title">⚠️ Import Validation Failed</h3>
        <button class="modal-close" onclick="App.closeModal()">&times;</button>
      </div>
      <p style="color:var(--danger);margin-bottom:12px;font-size:0.9rem">
        Import was cancelled — found <strong>${errors.length}${hasMore ? '+' : ''} error(s)</strong> across ${totalRows} rows.
        Please fix your file and try again.
      </p>
      <div style="max-height:280px;overflow-y:auto;background:var(--bg-secondary);border-radius:8px;padding:10px 12px;">
        ${errors.map(e => `
          <div style="padding:5px 0;border-bottom:1px solid var(--border);font-size:0.82rem;color:var(--danger);display:flex;gap:6px;align-items:flex-start">
            <span style="flex-shrink:0">⛔</span><span>${Utils.escapeHtml(e)}</span>
          </div>`).join('')}
        ${hasMore ? `<div style="padding:6px 0;color:var(--text-muted);font-size:0.8rem;font-style:italic">… and more errors not shown. Fix the above and re-import.</div>` : ''}
      </div>
      <div class="modal-actions">
        <button class="btn btn-secondary" onclick="App.closeModal()">Close</button>
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

    // --- Validate all rows before touching any data ---
    const validationErrors = validateImportRows(rows, dateCol, amountCol, categoryCol, paymentCol);
    if (validationErrors.length > 0) {
      showValidationErrors(validationErrors, rows.length);
      return; // abort — do NOT import
    }

    const categories = Store.getCategories().map(c => c.name);
    let imported = 0;
    const newExpenses = [];

    rows.forEach(row => {
      const dateVal = String(row[dateCol] || '').trim();
      const amountVal = parseFloat(String(row[amountCol] || '').replace(/[^0-9.\-]/g, ''));
      const date = parseDate(dateVal);

      // These checks are redundant after validation, but kept as a safety net
      if (!date || isNaN(amountVal) || amountVal <= 0) return;

      const category = categoryCol ? (String(row[categoryCol] || '').trim() || 'Other') : 'Other';
      const paymentMethod = paymentCol ? (String(row[paymentCol] || '').trim() || 'Other') : 'Other';

      newExpenses.push({
        id: Utils.generateId(),
        createdAt: new Date().toISOString(),
        date,
        amount: Math.abs(amountVal),
        category,
        description: descCol ? String(row[descCol] || '').trim() : '',
        paymentMethod,
        isRecurring: false
      });
      imported++;
    });

    // Replace any existing expenses so only the latest imported file's data is kept.
    Store.setExpenses(newExpenses);

    delete window._importData;
    App.closeModal();
    Utils.showToast(`Imported ${imported} expense${imported !== 1 ? 's' : ''} successfully`);
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
