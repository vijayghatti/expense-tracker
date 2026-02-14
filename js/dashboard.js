/**
 * Dashboard Module - Summary cards, charts, recent transactions
 */
const Dashboard = (() => {
  let viewYear, viewMonth;

  function init() {
    const now = new Date();
    viewYear = now.getFullYear();
    viewMonth = now.getMonth();
  }

  function render() {
    if (viewYear === undefined) init();
    const container = document.getElementById('tab-dashboard');
    const { start, end } = Utils.getMonthRange(viewYear, viewMonth);
    const allExpenses = Store.getExpenses();
    const monthExpenses = allExpenses.filter(e => e.date >= start && e.date <= end);
    const totalSpent = monthExpenses.reduce((s, e) => s + Number(e.amount), 0);
    const daysInMonth = Utils.daysInMonth(viewYear, viewMonth);
    const daysPassed = (viewYear === new Date().getFullYear() && viewMonth === new Date().getMonth())
      ? new Date().getDate() : daysInMonth;
    const avgDaily = daysPassed > 0 ? totalSpent / daysPassed : 0;
    const txCount = monthExpenses.length;

    // Category breakdown
    const catTotals = {};
    monthExpenses.forEach(e => {
      catTotals[e.category] = (catTotals[e.category] || 0) + Number(e.amount);
    });
    const topCategory = Object.entries(catTotals).sort((a, b) => b[1] - a[1])[0];

    // Budget status
    const budgets = Store.getBudgets();
    const totalBudget = Object.values(budgets).reduce((s, v) => s + v, 0);

    container.innerHTML = `
      <div class="month-picker" style="margin-bottom:20px;">
        <button class="btn-icon" onclick="Dashboard.prevMonth()">◀</button>
        <span class="current-month">${Utils.getMonthName(viewMonth)} ${viewYear}</span>
        <button class="btn-icon" onclick="Dashboard.nextMonth()">▶</button>
      </div>

      <div class="summary-grid">
        <div class="summary-card">
          <div class="icon">💰</div>
          <div class="label">Total Spent</div>
          <div class="value">${Utils.formatCurrency(totalSpent)}</div>
          ${totalBudget > 0 ? `<div class="sub">of ${Utils.formatCurrency(totalBudget)} budget</div>` : ''}
        </div>
        <div class="summary-card">
          <div class="icon">📊</div>
          <div class="label">Daily Average</div>
          <div class="value">${Utils.formatCurrency(avgDaily)}</div>
          <div class="sub">${daysPassed} days tracked</div>
        </div>
        <div class="summary-card">
          <div class="icon">📋</div>
          <div class="label">Transactions</div>
          <div class="value">${txCount}</div>
          <div class="sub">this month</div>
        </div>
        <div class="summary-card">
          <div class="icon">🏷️</div>
          <div class="label">Top Category</div>
          <div class="value" style="font-size:1.1rem">${topCategory ? topCategory[0] : '—'}</div>
          <div class="sub">${topCategory ? Utils.formatCurrency(topCategory[1]) : 'No expenses yet'}</div>
        </div>
      </div>

      <div class="charts-grid">
        <div class="card">
          <div class="card-header"><span class="card-title">Spending by Category</span></div>
          <div class="chart-container" id="chart-pie"></div>
          <div class="chart-legend" id="chart-pie-legend"></div>
        </div>
        <div class="card">
          <div class="card-header"><span class="card-title">Daily Spending Trend</span></div>
          <div class="chart-container" id="chart-line"></div>
        </div>
      </div>

      <div class="card" style="margin-top:16px">
        <div class="card-header">
          <span class="card-title">Recent Transactions</span>
          <button class="btn btn-sm btn-secondary" onclick="App.switchTab('expenses')">View All</button>
        </div>
        <div id="recent-transactions"></div>
      </div>
    `;

    renderCharts(monthExpenses, catTotals);
    renderRecent(allExpenses);
  }

  function renderCharts(monthExpenses, catTotals) {
    const categories = Store.getCategories();

    // Pie chart
    const pieData = Object.entries(catTotals)
      .sort((a, b) => b[1] - a[1])
      .map(([name, value]) => ({
        label: name,
        value,
        color: Store.getCategoryColor(name)
      }));

    const pieContainer = document.getElementById('chart-pie');
    const parentWidth = pieContainer.parentElement.clientWidth - 40;
    const pieSize = Math.min(280, parentWidth);
    MiniChart.drawPie(pieContainer, pieData, { width: pieSize, height: pieSize, doughnut: true });

    // Legend
    const legendEl = document.getElementById('chart-pie-legend');
    legendEl.innerHTML = pieData.map(d => `
      <div class="legend-item">
        <span class="legend-dot" style="background:${d.color}"></span>
        ${d.label}: ${Utils.formatCurrency(d.value)}
      </div>
    `).join('');

    // Line chart - daily spending
    const { start, end } = Utils.getMonthRange(viewYear, viewMonth);
    const daysInMonth = Utils.daysInMonth(viewYear, viewMonth);
    const dailySpending = new Array(daysInMonth).fill(0);
    monthExpenses.forEach(e => {
      const day = parseInt(e.date.split('-')[2]) - 1;
      if (day >= 0 && day < daysInMonth) dailySpending[day] += Number(e.amount);
    });

    const labels = dailySpending.map((_, i) => String(i + 1));
    const lineContainer = document.getElementById('chart-line');
    const lineWidth = Math.min(480, parentWidth);
    MiniChart.drawLine(lineContainer, { labels, values: dailySpending }, {
      width: lineWidth, height: 240, color: '#4CAF50'
    });
  }

  function renderRecent(allExpenses) {
    const recent = allExpenses.slice(0, 5);
    const el = document.getElementById('recent-transactions');
    if (!recent.length) {
      el.innerHTML = '<div class="empty-state"><div class="desc">No transactions yet</div></div>';
      return;
    }
    el.innerHTML = `<div style="display:flex;flex-direction:column;gap:8px">
      ${recent.map(e => `
        <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-bottom:1px solid var(--border-color)">
          <div style="display:flex;align-items:center;gap:12px">
            <span style="font-size:1.3rem">${Store.getCategoryIcon(e.category)}</span>
            <div>
              <div style="font-weight:500">${Utils.escapeHtml(e.description || e.category)}</div>
              <div style="font-size:0.8rem;color:var(--text-secondary)">${Utils.formatDate(e.date)} · ${e.paymentMethod || ''}</div>
            </div>
          </div>
          <span class="amount-cell">${Utils.formatCurrency(e.amount)}</span>
        </div>
      `).join('')}
    </div>`;
  }

  function prevMonth() {
    viewMonth--;
    if (viewMonth < 0) { viewMonth = 11; viewYear--; }
    render();
  }
  function nextMonth() {
    viewMonth++;
    if (viewMonth > 11) { viewMonth = 0; viewYear++; }
    render();
  }

  return { init, render, prevMonth, nextMonth };
})();
