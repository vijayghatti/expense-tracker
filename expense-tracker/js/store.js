/**
 * Store - LocalStorage abstraction layer
 */
const Store = (() => {
  const KEYS = {
    expenses: 'et_expenses',
    recurring: 'et_recurring',
    budgets: 'et_budgets',
    categories: 'et_categories',
    settings: 'et_settings'
  };

  const DEFAULT_CATEGORIES = [
    { name: 'Food & Dining', color: '#4CAF50', icon: '🍔' },
    { name: 'Transport', color: '#2196F3', icon: '🚗' },
    { name: 'Housing', color: '#9C27B0', icon: '🏠' },
    { name: 'Utilities', color: '#FF9800', icon: '💡' },
    { name: 'Entertainment', color: '#E91E63', icon: '🎬' },
    { name: 'Healthcare', color: '#00BCD4', icon: '🏥' },
    { name: 'Shopping', color: '#FF5722', icon: '🛍️' },
    { name: 'Education', color: '#3F51B5', icon: '📚' },
    { name: 'Subscriptions', color: '#607D8B', icon: '📱' },
    { name: 'Other', color: '#795548', icon: '📦' }
  ];

  const DEFAULT_SETTINGS = {
    currency: '₹',
    theme: 'dark'
  };

  function get(key) {
    try {
      const data = localStorage.getItem(KEYS[key]);
      return data ? JSON.parse(data) : null;
    } catch {
      return null;
    }
  }

  function set(key, value) {
    try {
      localStorage.setItem(KEYS[key], JSON.stringify(value));
    } catch (e) {
      console.error('Storage error:', e);
      Utils.showToast('Storage is full! Consider exporting and clearing old data.', 'error');
    }
  }

  function init() {
    if (!get('categories')) set('categories', DEFAULT_CATEGORIES);
    if (!get('expenses')) set('expenses', []);
    if (!get('recurring')) set('recurring', []);
    if (!get('budgets')) set('budgets', {});
    if (!get('settings')) set('settings', DEFAULT_SETTINGS);
  }

  // Expenses
  function getExpenses() { return get('expenses') || []; }
  function setExpenses(expenses) { set('expenses', expenses); }
  function addExpense(expense) {
    const expenses = getExpenses();
    expense.id = Utils.generateId();
    expense.createdAt = new Date().toISOString();
    expenses.unshift(expense);
    setExpenses(expenses);
    return expense;
  }
  function updateExpense(id, updates) {
    const expenses = getExpenses();
    const idx = expenses.findIndex(e => e.id === id);
    if (idx !== -1) {
      expenses[idx] = { ...expenses[idx], ...updates };
      setExpenses(expenses);
      return expenses[idx];
    }
    return null;
  }
  function deleteExpense(id) {
    setExpenses(getExpenses().filter(e => e.id !== id));
  }

  // Recurring
  function getRecurring() { return get('recurring') || []; }
  function setRecurring(items) { set('recurring', items); }
  function addRecurring(item) {
    const items = getRecurring();
    item.id = Utils.generateId();
    item.enabled = true;
    item.lastGenerated = null;
    items.push(item);
    setRecurring(items);
    return item;
  }
  function updateRecurring(id, updates) {
    const items = getRecurring();
    const idx = items.findIndex(r => r.id === id);
    if (idx !== -1) {
      items[idx] = { ...items[idx], ...updates };
      setRecurring(items);
      return items[idx];
    }
    return null;
  }
  function deleteRecurring(id) {
    setRecurring(getRecurring().filter(r => r.id !== id));
  }

  // Budgets
  function getBudgets() { return get('budgets') || {}; }
  function setBudgets(budgets) { set('budgets', budgets); }
  function setBudget(category, amount) {
    const budgets = getBudgets();
    budgets[category] = amount;
    setBudgets(budgets);
  }
  function deleteBudget(category) {
    const budgets = getBudgets();
    delete budgets[category];
    setBudgets(budgets);
  }

  // Categories
  function getCategories() { return get('categories') || DEFAULT_CATEGORIES; }
  function setCategories(cats) { set('categories', cats); }
  function addCategory(cat) {
    const cats = getCategories();
    cats.push(cat);
    setCategories(cats);
  }
  function updateCategory(oldName, cat) {
    const cats = getCategories();
    const idx = cats.findIndex(c => c.name === oldName);
    if (idx !== -1) {
      // Also update references in expenses, recurring, budgets
      if (oldName !== cat.name) {
        const expenses = getExpenses();
        expenses.forEach(e => { if (e.category === oldName) e.category = cat.name; });
        setExpenses(expenses);
        const recurring = getRecurring();
        recurring.forEach(r => { if (r.category === oldName) r.category = cat.name; });
        setRecurring(recurring);
        const budgets = getBudgets();
        if (budgets[oldName] !== undefined) {
          budgets[cat.name] = budgets[oldName];
          delete budgets[oldName];
          setBudgets(budgets);
        }
      }
      cats[idx] = cat;
      setCategories(cats);
    }
  }
  function deleteCategory(name) {
    setCategories(getCategories().filter(c => c.name !== name));
  }
  function getCategoryColor(name) {
    const cat = getCategories().find(c => c.name === name);
    return cat ? cat.color : '#795548';
  }
  function getCategoryIcon(name) {
    const cat = getCategories().find(c => c.name === name);
    return cat ? cat.icon : '📦';
  }

  // Settings
  function getSettings() { return get('settings') || DEFAULT_SETTINGS; }
  function updateSettings(updates) {
    set('settings', { ...getSettings(), ...updates });
  }

  // Backup / Restore
  function exportBackup() {
    return JSON.stringify({
      expenses: getExpenses(),
      recurring: getRecurring(),
      budgets: getBudgets(),
      categories: getCategories(),
      settings: getSettings(),
      exportedAt: new Date().toISOString()
    }, null, 2);
  }

  function importBackup(jsonStr) {
    const data = JSON.parse(jsonStr);
    if (data.expenses) setExpenses(data.expenses);
    if (data.recurring) setRecurring(data.recurring);
    if (data.budgets) setBudgets(data.budgets);
    if (data.categories) setCategories(data.categories);
    if (data.settings) set('settings', { ...DEFAULT_SETTINGS, ...data.settings });
  }

  function clearAll() {
    Object.values(KEYS).forEach(k => localStorage.removeItem(k));
    init();
  }

  function getStorageUsage() {
    let total = 0;
    Object.values(KEYS).forEach(k => {
      const item = localStorage.getItem(k);
      if (item) total += item.length * 2; // UTF-16
    });
    return total;
  }

  return {
    init, getExpenses, setExpenses, addExpense, updateExpense, deleteExpense,
    getRecurring, setRecurring, addRecurring, updateRecurring, deleteRecurring,
    getBudgets, setBudgets, setBudget, deleteBudget,
    getCategories, setCategories, addCategory, updateCategory, deleteCategory,
    getCategoryColor, getCategoryIcon,
    getSettings, updateSettings,
    exportBackup, importBackup, clearAll, getStorageUsage
  };
})();
