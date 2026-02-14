/**
 * MiniChart - Lightweight Canvas charting library
 * Supports: Pie, Doughnut, Line, Bar charts
 * Zero dependencies
 */
const MiniChart = (() => {
  const defaultColors = [
    '#4CAF50', '#2196F3', '#FF9800', '#E91E63', '#9C27B0',
    '#00BCD4', '#FF5722', '#607D8B', '#8BC34A', '#FFC107',
    '#3F51B5', '#795548', '#009688', '#CDDC39', '#FF4081'
  ];

  function createCanvas(container, width, height) {
    const canvas = document.createElement('canvas');
    const dpr = window.devicePixelRatio || 1;
    canvas.width = width * dpr;
    canvas.height = height * dpr;
    canvas.style.width = width + 'px';
    canvas.style.height = height + 'px';
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    if (typeof container === 'string') container = document.querySelector(container);
    container.innerHTML = '';
    container.appendChild(canvas);
    return { canvas, ctx, width, height };
  }

  function drawPie(container, data, options = {}) {
    const w = options.width || 320;
    const h = options.height || 320;
    const { ctx, width, height } = createCanvas(container, w, h);
    const isDoughnut = options.doughnut || false;
    const total = data.reduce((s, d) => s + d.value, 0);
    if (total === 0) {
      ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#999';
      ctx.font = '14px system-ui';
      ctx.textAlign = 'center';
      ctx.fillText('No data', width / 2, height / 2);
      return;
    }

    const cx = width / 2;
    const cy = height / 2;
    const radius = Math.min(cx, cy) - 10;
    let startAngle = -Math.PI / 2;

    data.forEach((d, i) => {
      const sliceAngle = (d.value / total) * Math.PI * 2;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, radius, startAngle, startAngle + sliceAngle);
      ctx.closePath();
      ctx.fillStyle = d.color || defaultColors[i % defaultColors.length];
      ctx.fill();

      // Label
      if (d.value / total > 0.05) {
        const midAngle = startAngle + sliceAngle / 2;
        const labelR = radius * (isDoughnut ? 0.75 : 0.65);
        const lx = cx + Math.cos(midAngle) * labelR;
        const ly = cy + Math.sin(midAngle) * labelR;
        ctx.fillStyle = '#fff';
        ctx.font = 'bold 11px system-ui';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        const pct = Math.round((d.value / total) * 100);
        ctx.fillText(pct + '%', lx, ly);
      }
      startAngle += sliceAngle;
    });

    if (isDoughnut) {
      ctx.beginPath();
      ctx.arc(cx, cy, radius * 0.5, 0, Math.PI * 2);
      ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--card-bg').trim() || '#fff';
      ctx.fill();
    }
  }

  function drawLine(container, data, options = {}) {
    const w = options.width || 500;
    const h = options.height || 280;
    const { ctx, width, height } = createCanvas(container, w, h);
    const padding = { top: 20, right: 20, bottom: 40, left: 60 };
    const chartW = width - padding.left - padding.right;
    const chartH = height - padding.top - padding.bottom;

    if (!data.labels || !data.labels.length) {
      ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#999';
      ctx.font = '14px system-ui';
      ctx.textAlign = 'center';
      ctx.fillText('No data', width / 2, height / 2);
      return;
    }

    const values = data.values;
    const maxVal = Math.max(...values, 1);
    const minVal = 0;
    const range = maxVal - minVal || 1;

    const textColor = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#999';
    const gridColor = getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim() || '#eee';

    // Grid lines
    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 0.5;
    const gridLines = 5;
    for (let i = 0; i <= gridLines; i++) {
      const y = padding.top + (chartH / gridLines) * i;
      ctx.beginPath();
      ctx.moveTo(padding.left, y);
      ctx.lineTo(padding.left + chartW, y);
      ctx.stroke();

      // Y labels
      const val = maxVal - (range / gridLines) * i;
      ctx.fillStyle = textColor;
      ctx.font = '11px system-ui';
      ctx.textAlign = 'right';
      ctx.textBaseline = 'middle';
      ctx.fillText(formatNum(val), padding.left - 8, y);
    }

    // X labels
    const step = Math.max(1, Math.floor(data.labels.length / 8));
    data.labels.forEach((label, i) => {
      if (i % step === 0 || i === data.labels.length - 1) {
        const x = padding.left + (chartW / (data.labels.length - 1 || 1)) * i;
        ctx.fillStyle = textColor;
        ctx.font = '11px system-ui';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        ctx.fillText(label, x, height - padding.bottom + 8);
      }
    });

    // Line + fill
    const lineColor = options.color || '#4CAF50';
    ctx.beginPath();
    const points = values.map((v, i) => ({
      x: padding.left + (chartW / (values.length - 1 || 1)) * i,
      y: padding.top + chartH - ((v - minVal) / range) * chartH
    }));

    // Fill area
    ctx.beginPath();
    ctx.moveTo(points[0].x, padding.top + chartH);
    points.forEach(p => ctx.lineTo(p.x, p.y));
    ctx.lineTo(points[points.length - 1].x, padding.top + chartH);
    ctx.closePath();
    ctx.fillStyle = lineColor + '20';
    ctx.fill();

    // Line
    ctx.beginPath();
    points.forEach((p, i) => i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y));
    ctx.strokeStyle = lineColor;
    ctx.lineWidth = 2.5;
    ctx.lineJoin = 'round';
    ctx.stroke();

    // Dots
    points.forEach(p => {
      ctx.beginPath();
      ctx.arc(p.x, p.y, 3.5, 0, Math.PI * 2);
      ctx.fillStyle = lineColor;
      ctx.fill();
      ctx.strokeStyle = getComputedStyle(document.documentElement).getPropertyValue('--card-bg').trim() || '#fff';
      ctx.lineWidth = 2;
      ctx.stroke();
    });
  }

  function drawBar(container, data, options = {}) {
    const w = options.width || 500;
    const h = options.height || 280;
    const { ctx, width, height } = createCanvas(container, w, h);
    const padding = { top: 20, right: 20, bottom: 50, left: 60 };
    const chartW = width - padding.left - padding.right;
    const chartH = height - padding.top - padding.bottom;

    if (!data.labels || !data.labels.length) {
      ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#999';
      ctx.font = '14px system-ui';
      ctx.textAlign = 'center';
      ctx.fillText('No data', width / 2, height / 2);
      return;
    }

    const maxVal = Math.max(...data.values, 1);
    const textColor = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#999';
    const gridColor = getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim() || '#eee';
    const barWidth = Math.min(40, (chartW / data.labels.length) * 0.6);
    const gap = chartW / data.labels.length;

    // Grid
    for (let i = 0; i <= 5; i++) {
      const y = padding.top + (chartH / 5) * i;
      ctx.strokeStyle = gridColor;
      ctx.lineWidth = 0.5;
      ctx.beginPath();
      ctx.moveTo(padding.left, y);
      ctx.lineTo(padding.left + chartW, y);
      ctx.stroke();
      ctx.fillStyle = textColor;
      ctx.font = '11px system-ui';
      ctx.textAlign = 'right';
      ctx.fillText(formatNum(maxVal - (maxVal / 5) * i), padding.left - 8, y + 4);
    }

    data.values.forEach((v, i) => {
      const barH = (v / maxVal) * chartH;
      const x = padding.left + gap * i + (gap - barWidth) / 2;
      const y = padding.top + chartH - barH;
      const color = data.colors ? data.colors[i] : defaultColors[i % defaultColors.length];

      // Bar with rounded top
      ctx.beginPath();
      const r = Math.min(4, barWidth / 2);
      ctx.moveTo(x, y + r);
      ctx.quadraticCurveTo(x, y, x + r, y);
      ctx.lineTo(x + barWidth - r, y);
      ctx.quadraticCurveTo(x + barWidth, y, x + barWidth, y + r);
      ctx.lineTo(x + barWidth, padding.top + chartH);
      ctx.lineTo(x, padding.top + chartH);
      ctx.closePath();
      ctx.fillStyle = color;
      ctx.fill();

      // X label
      ctx.save();
      ctx.translate(x + barWidth / 2, height - padding.bottom + 8);
      ctx.rotate(-Math.PI / 6);
      ctx.fillStyle = textColor;
      ctx.font = '11px system-ui';
      ctx.textAlign = 'right';
      const lbl = data.labels[i].length > 10 ? data.labels[i].slice(0, 9) + '…' : data.labels[i];
      ctx.fillText(lbl, 0, 0);
      ctx.restore();
    });
  }

  function formatNum(n) {
    if (n >= 1000000) return (n / 1000000).toFixed(1) + 'M';
    if (n >= 1000) return (n / 1000).toFixed(1) + 'K';
    return Math.round(n).toString();
  }

  return { drawPie, drawLine, drawBar, defaultColors };
})();
