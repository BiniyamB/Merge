/* ============================================
   DIGITAL TRANSACTION VALUE SNAPSHOT - APP.JS
   ============================================ */

(function () {
  'use strict';

  // ---- STATE ----
  let reportState = null;
  let servicesState = [];
  let formOpen = true;
  const STORAGE_KEY = 'dtsnapshot.config.v1';

  // ---- INIT ----
  document.addEventListener('DOMContentLoaded', function () {
    document.getElementById('form-body').classList.add('open');
    initNavDate();
    initSettingsControls();
    const saved = loadState();
    if (saved) {
      restoreFormData(saved);
    }
    renderReport();
    lucide.createIcons();
  });

  function initSettingsControls() {
    var modeEl = document.getElementById('setting-takeaway-mode');
    var overrideEl = document.getElementById('setting-takeaway-override');
    function syncTakeaway() {
      if (overrideEl) overrideEl.disabled = !(modeEl && modeEl.value === 'custom');
    }
    if (modeEl) modeEl.addEventListener('change', function () {
      syncTakeaway();
      renderReport();
    });
    syncTakeaway();

    ['setting-metric-bars', 'setting-auto-highlight'].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener('change', debounce(function () { renderReport(); }, 200));
    });
    if (overrideEl) overrideEl.addEventListener('input', debounce(function () { renderReport(); }, 300));
  }

  function initNavDate() {
    const el = document.getElementById('nav-date');
    if (el) el.textContent = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  // ---- FORM TOGGLE ----
  window.toggleFormSection = function () {
    formOpen = !formOpen;
    const body = document.getElementById('form-body');
    const chevron = document.getElementById('form-chevron');
    body.classList.toggle('open', formOpen);
    if (chevron) chevron.style.transform = formOpen ? 'rotate(180deg)' : '';
  };

  // ---- READ FORM STATE ----
  function readReportFromForm() {
    return {
      title: gv('report-title'),
      organization: gv('report-org'),
      brand: gv('report-brand'),
      tagline1: gv('report-tagline1'),
      tagline2: gv('report-tagline2'),
      subtitle: gv('report-subtitle'),
      date: gv('report-date')
    };
  }

  function readSettings() {
    var el = function (id) { return document.getElementById(id); };
    return {
      showMetricBars: el('setting-metric-bars') ? el('setting-metric-bars').checked : true,
      autoHighlight: el('setting-auto-highlight') ? el('setting-auto-highlight').checked : true,
      takeawayMode: el('setting-takeaway-mode') ? el('setting-takeaway-mode').value : 'auto',
      takeawayOverride: el('setting-takeaway-override') ? el('setting-takeaway-override').value.trim() : ''
    };
  }

  function readServicesFromForm() {
    const cards = document.querySelectorAll('.service-card');
    const services = [];
    cards.forEach(function (card, idx) {
      const volInput = card.querySelector('[data-field="transactionVolume"]');
      const valInput = card.querySelector('[data-field="totalValue"]');
      const targetInput = card.querySelector('[data-field="target"]');
      const hlCheckbox = card.querySelector('[data-field="highlighted"]');

      const volume = parseNum(volInput ? volInput.value : '0');
      const totalVal = parseNum(valInput ? valInput.value : '0');
      const target = parseNum(targetInput ? targetInput.value : '0');
      const typeEl = card.querySelector('[data-field="type"]');
      const type = typeEl ? typeEl.value : 'financial';

      services.push({
        id: card.dataset.id || 'svc-' + idx,
        name: (card.querySelector('[data-field="name"]') || {}).value || 'Service ' + (idx + 1),
        icon: getServiceIcon(card.querySelector('[data-field="name"]') ? card.querySelector('[data-field="name"]').value : ''),
        type: type,
        transactionVolume: volume,
        totalValue: totalVal,
        target: target,
        keyMessage: (card.querySelector('[data-field="keyMessage"]') || {}).value || '',
        highlighted: hlCheckbox ? hlCheckbox.checked : false,
        highlightStyle: ''
      });
    });
    return services;
  }

  function getServiceIcon(name) {
    var n = name.toUpperCase();
    if (n.indexOf('CASH') >= 0 || n.indexOf('WITHDRAW') >= 0) return 'banknote';
    if (n.indexOf('POS') >= 0 || n.indexOf('PURCHASE') >= 0) return 'credit-card';
    if (n.indexOf('P2P') >= 0 || n.indexOf('IPS') >= 0) return 'users';
    if (n.indexOf('QR') >= 0) return 'qr-code';
    if (n.indexOf('BALANCE') >= 0 || n.indexOf('INQUIRY') >= 0 || n.indexOf('MINI') >= 0) return 'landmark';
    if (n.indexOf('RTP') >= 0) return 'arrow-left-right';
    if (n.indexOf('NPG') >= 0 || n.indexOf('ONLINE') >= 0) return 'globe';
    if (n.indexOf('SUCCESS RATE') >= 0) return 'percent';
    return 'circle-dot';
  }

  function getServiceIconClass(name) {
    var n = (name || '').toUpperCase();
    if (n.indexOf('CASH') >= 0 || n.indexOf('WITHDRAW') >= 0) return 'cash';
    if (n.indexOf('POS') >= 0 || n.indexOf('PURCHASE') >= 0) return 'pos';
    if (n.indexOf('P2P') >= 0 || n.indexOf('IPS') >= 0) return 'p2p';
    if (n.indexOf('QR') >= 0) return 'qr';
    if (n.indexOf('BALANCE') >= 0 || n.indexOf('INQUIRY') >= 0 || n.indexOf('MINI') >= 0) return 'landmark';
    if (n.indexOf('RTP') >= 0 || n.indexOf('NPG') >= 0 || n.indexOf('ONLINE') >= 0) return 'landmark';
    if (n.indexOf('SUCCESS RATE') >= 0) return 'pos';
    return 'default';
  }

  // ---- UTILITIES ----
  function gv(id) {
    var el = document.getElementById(id);
    return el ? el.value.trim() : '';
  }

  function parseNum(str) {
    if (str === '' || str === null || str === undefined) return 0;
    var cleaned = String(str).replace(/,/g, '').trim();
    var n = parseFloat(cleaned);
    return isNaN(n) || !isFinite(n) ? 0 : n;
  }

  function fmtNum(n) {
    if (n === null || n === undefined || isNaN(n) || !isFinite(n)) return '-';
    if (Number.isInteger(n)) {
      return n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    }
    return n.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 2 });
  }

  function fmtDecimal(n) {
    if (n === null || n === undefined || isNaN(n) || !isFinite(n)) return '-';
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // ---- CALCULATIONS ----
  function isSuccessRate(s) {
    return (s.type === 'success-rate') || /SUCCESS RATE/i.test((s.name || ''));
  }

  function calcService(s) {
    var avg = 0;
    if (s.type === 'financial' && s.transactionVolume > 0 && s.totalValue > 0) {
      avg = s.totalValue / s.transactionVolume;
    }
    var rate = isSuccessRate(s);
    var ach = null;
    if (!rate && s.target > 0) {
      ach = (s.transactionVolume / s.target) * 100;
    } else if (rate && s.target > 0) {
      ach = (s.transactionVolume / s.target) * 100;
    }
    return {
      avgTransactionValue: avg,
      isFinancial: s.type === 'financial' && s.totalValue > 0,
      isSuccessRate: rate,
      achievementPercent: ach
    };
  }

  function calcAll(services) {
    var enriched = services.map(function (s) {
      var c = calcService(s);
      return Object.assign({}, s, {
        averageTransactionValue: c.avgTransactionValue,
        isFinancial: c.isFinancial,
        isSuccessRate: c.isSuccessRate,
        achievementPercent: c.achievementPercent
      });
    });

    var maxVolume = Math.max.apply(null, enriched.filter(function (s) { return !s.isSuccessRate; }).map(function (s) { return s.transactionVolume || 0; }));
    var maxAvg = Math.max.apply(null, enriched.filter(function (s) { return s.isFinancial; }).map(function (s) { return s.averageTransactionValue; }));

    enriched.forEach(function (s) {
      s.volumePercent = maxVolume > 0 ? (s.transactionVolume / maxVolume) * 100 : 0;
      s.avgPercent = maxAvg > 0 ? (s.averageTransactionValue / maxAvg) * 100 : 0;
    });

    // Sort financial by volume desc for ranking
    var financial = enriched.filter(function (s) { return s.isFinancial; });
    financial.sort(function (a, b) { return b.transactionVolume - a.transactionVolume; });

    var volumeLeader = financial[0] || null;

    // Sort by avg desc
    var byAvg = enriched.filter(function (s) { return s.isFinancial; }).slice().sort(function (a, b) {
      return b.averageTransactionValue - a.averageTransactionValue;
    });
    var highestAvg = byAvg[0] || null;
    var lowestAvg = byAvg[byAvg.length - 1] || null;

    // Highest total value
    var byTotal = financial.slice().sort(function (a, b) { return b.totalValue - a.totalValue; });
    var highestTotal = byTotal[0] || null;

    // QR service
    var qr = enriched.find(function (s) { return s.id === 'qr'; }) || null;
    var qrAdvantages = [];
    if (qr && qr.isFinancial) {
      // Match reference ordering: IPS P2P, POS PURCHASE, CASH WITHDRAWAL
      var order = ['ips-p2p', 'pos-purchase', 'cash-withdrawal'];
      var others = [].concat(enriched)
        .filter(function (s) { return s.isFinancial && s.id !== 'qr'; })
        .sort(function (a, b) {
          var ia = order.indexOf(a.id);
          var ib = order.indexOf(b.id);
          ia = ia === -1 ? 99 : ia;
          ib = ib === -1 ? 99 : ib;
          return ia - ib;
        });
      others.forEach(function (o) {
        if (o.averageTransactionValue > 0) {
          var ratio = qr.averageTransactionValue / o.averageTransactionValue;
          qrAdvantages.push({
            service: o.name,
            ratio: ratio,
            label: ratio.toFixed(1) + 'x',
            description: 'Higher than ' + formatServiceShort(o.name)
          });
        }
      });
    }

    // Dynamic key message for QR
    if (qr && qr.isFinancial && !qr.keyMessage) {
      if (highestAvg && highestAvg.id === 'qr') {
        qr.keyMessage = 'HIGHEST average transaction value';
      }
    }

    // Dynamic key takeaway
    var takeaway = generateTakeaway(enriched, qr, volumeLeader, highestAvg);

    // Total row (excludes success-rate rows - they are percentages)
    var total = enriched.reduce(function (acc, s) {
      if (s.isSuccessRate) return acc;
      acc.performance += s.transactionVolume || 0;
      acc.target += s.target || 0;
      acc.totalValue += s.totalValue || 0;
      return acc;
    }, { performance: 0, target: 0, totalValue: 0 });

    var achList = enriched
      .map(function (s) { return s.achievementPercent; })
      .filter(function (v) { return v !== null && v !== undefined && !isNaN(v); });
    total.achievementPercent = achList.length > 0 ? (achList.reduce(function (a, b) { return a + b; }, 0) / achList.length) : null;

    return {
      services: enriched,
      volumeLeader: volumeLeader,
      highestAvg: highestAvg,
      lowestAvg: lowestAvg,
      highestTotal: highestTotal,
      qr: qr,
      qrAdvantages: qrAdvantages,
      takeaway: takeaway,
      total: total
    };
  }

  function formatServiceShort(name) {
    var n = name.toUpperCase();
    if (n.indexOf('CASH') >= 0) return 'Cash Withdrawal';
    if (n.indexOf('POS') >= 0) return 'POS';
    if (n.indexOf('IPS') >= 0 || n.indexOf('P2P') >= 0) return 'P2P';
    if (n.indexOf('QR') >= 0) return 'QR';
    if (n.indexOf('RTP') >= 0) return 'RTP';
    if (n.indexOf('NPG') >= 0) return 'NPG';
    return name;
  }

  function generateTakeaway(services, qr, volumeLeader, highestAvg) {
    if (!qr || !qr.isFinancial) return '';

    var qrVolRank = services.filter(function (s) { return s.isFinancial; })
      .sort(function (a, b) { return b.transactionVolume - a.transactionVolume; })
      .findIndex(function (s) { return s.id === 'qr'; }) + 1;

    var totalFinancial = services.filter(function (s) { return s.isFinancial; }).length;
    var qrAvgRank = services.filter(function (s) { return s.isFinancial; })
      .sort(function (a, b) { return b.averageTransactionValue - a.averageTransactionValue; })
      .findIndex(function (s) { return s.id === 'qr'; }) + 1;

    var volLeaderName = volumeLeader ? formatServiceShort(volumeLeader.name) : 'other services';

    if (qrAvgRank === 1 && qrVolRank > 1) {
      return qr.name + ' remains smaller in volume compared to ' + volLeaderName + ', but each transaction carries significantly more value — making it a high-potential growth channel for digital merchant payments.';
    }
    if (qrAvgRank === 1 && qrVolRank === 1) {
      return qr.name + ' leads in both volume and average transaction value, demonstrating strong adoption and high-value usage across the payment ecosystem.';
    }
    if (qrAvgRank <= 2) {
      return qr.name + ' shows competitive average transaction value. With growing merchant adoption, it represents a key channel for high-value digital payments.';
    }
    return qr.name + ' has room for growth in average transaction value. Continued merchant onboarding and user education could drive higher-value transactions over time.';
  }

  // ---- RENDER REPORT ----
  function renderReport() {
    reportState = readReportFromForm();
    servicesState = readServicesFromForm();
    var settings = readSettings();
    var calc = calcAll(servicesState);

    // Persist current configuration locally (isolated to the browser,
    // so a server-side persistence layer can be added later).
    saveState({ report: reportState, services: servicesState, settings: settings });

    renderHeader(reportState);
    renderTable(calc, settings);
    renderInsights(calc, reportState, settings);
    renderFooter(reportState);
    validateForm();
    lucide.createIcons();
  }

  function renderHeader(r) {
    setText('r-org', r.organization);
    setText('r-brand', r.brand);
    setText('r-title', r.title);
    setText('r-subtitle', r.subtitle);
    setText('r-date', r.date);
    setText('f-org', r.organization + ' S.C.');
    setText('f-tagline', r.tagline1);
  }

  function renderTable(calc, settings) {
    settings = settings || { showMetricBars: true, autoHighlight: true };
    var tbody = document.getElementById('report-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    var highestAvgId = calc.highestAvg ? calc.highestAvg.id : null;

    function renderRow(s, isTotal) {
      var tr = document.createElement('tr');

      if (isTotal) {
        tr.className = 'total-row';
        tr.appendChild(totalRowCells(calc.total));
        tbody.appendChild(tr);
        return;
      }

      // Auto-highlight the highest-average-value financial service when enabled
      var isSuccessRate = s.isSuccessRate;
      var isHighestAvg = !isSuccessRate && highestAvgId && highestAvgId === s.id;
      var effectiveHighlight = s.highlighted || (settings.autoHighlight && isHighestAvg);
      if (effectiveHighlight) tr.className = 'highlight-row';

      var showBadge = isHighestAvg;

      var bar = function (pct) {
        if (!settings.showMetricBars) return '';
        return '<div class="metric-bar-wrap"><div class="metric-bar"><div class="metric-bar-fill" style="width:' +
          pct + '%"></div></div></div>';
      };

      // Service cell (skip icon row decoration for success-rate rows)
      var iconClass = isSuccessRate ? 'non-financial' : (s.isFinancial ? 'financial' : 'non-financial');
      var tdSvc = '<td><div class="svc-cell">' +
        '<div class="svc-icon ' + iconClass + ' ' + getServiceIconClass(s.name) + '">' +
        '<i data-lucide="' + s.icon + '"></i></div>' +
        '<span class="svc-name">' + escHtml(s.name) + '</span></div></td>';

      // Performance (was Transaction Volume)
      var perfText, perfBar;
      if (isSuccessRate) {
        perfText = (s.transactionVolume > 0 ? fmtDecimal(s.transactionVolume) + '%' : '-');
        perfBar = '';
      } else {
        perfText = (s.transactionVolume > 0 ? fmtNum(s.transactionVolume) : '-');
        perfBar = bar(s.volumePercent);
      }
      var tdPerf = '<td class="num-cell"><div class="num-primary">' + perfText + '</div>' + perfBar + '</td>';

      // Monthly plan (target)
      var targetText;
      if (isSuccessRate) {
        targetText = s.target > 0 ? fmtDecimal(s.target) + '%' : '<span class="num-dash">-</span>';
      } else {
        targetText = s.target > 0 ? fmtNum(s.target) : '<span class="num-dash">-</span>';
      }
      var tdTarget = '<td class="num-cell"><div class="num-primary">' + targetText + '</div></td>';

      // Achievement %
      var achText;
      if (s.achievementPercent === null || s.achievementPercent === undefined || isNaN(s.achievementPercent)) {
        achText = '<span class="num-dash">-</span>';
      } else {
        achText = fmtDecimal(s.achievementPercent) + '%';
      }
      var tdAch = '<td class="num-cell"><div class="num-primary">' + achText + '</div></td>';

      // Total value
      var tdVal = '<td class="num-cell"><div class="num-primary">' +
        (s.isFinancial ? 'ETB ' + fmtDecimal(s.totalValue) : '<span class="num-dash">-</span>') + '</div></td>';

      // Average
      var tdAvg = '<td class="num-cell"><div class="num-primary">' +
        (s.isFinancial ? 'ETB ' + fmtDecimal(s.averageTransactionValue) : '<span class="num-dash">-</span>') + '</div>' +
        (s.isFinancial ? bar(s.avgPercent) : '') + '</td>';

      // Key message
      var msgClass = 'msg-cell';
      if (showBadge) msgClass += ' msg-highlight';
      var tdMsg = '<td class="' + msgClass + '">' + escHtml(s.keyMessage);
      if (showBadge) {
        tdMsg += '<div class="msg-badge"><i data-lucide="star" class="d-inline" style="width:9px;height:9px;display:inline-block;vertical-align:-1px;"></i> HIGHEST AVG VALUE</div>';
      }
      tdMsg += '</td>';

      tr.innerHTML = tdSvc + tdTarget + tdPerf + tdVal + tdAch + tdAvg + tdMsg;
      tbody.appendChild(tr);
    }

    function totalRowCells(tot) {
      var achText = (tot.achievementPercent === null || tot.achievementPercent === undefined || isNaN(tot.achievementPercent))
        ? '<span class="num-dash">-</span>' : fmtDecimal(tot.achievementPercent) + '%';
      return '<td><div class="svc-cell"><span class="svc-name">TOTAL</span></div></td>' +
        '<td class="num-cell"><div class="num-primary num-total">' + fmtNum(tot.target) + '</div></td>' +
        '<td class="num-cell"><div class="num-primary num-total">' + fmtNum(tot.performance) + '</div></td>' +
        '<td class="num-cell"><div class="num-primary num-total">ETB ' + fmtDecimal(tot.totalValue) + '</div></td>' +
        '<td class="num-cell"><div class="num-primary num-total">' + achText + '</div></td>' +
        '<td class="num-cell"><div class="num-dash">-</div></td>' +
        '<td class="msg-cell"></td>';
    }

    calc.services.forEach(function (s) {
      renderRow(s, false);
    });
    renderRow(null, true);
  }

  function renderInsights(calc, report, settings) {
    settings = settings || { takeawayMode: 'auto', takeawayOverride: '' };

    // Volume leader
    if (calc.volumeLeader) {
      setText('vl-service', calc.volumeLeader.name);
      setText('vl-detail', fmtNum(calc.volumeLeader.transactionVolume) + ' transactions drive the ecosystem\'s scale.');
    }

    // QR advantage
    if (calc.qr && calc.qr.isFinancial) {
      setText('qr-avg-value', 'ETB ' + fmtDecimal(calc.qr.averageTransactionValue));

      var compEl = document.getElementById('qr-comparisons');
      if (compEl) {
        compEl.innerHTML = '';
        calc.qrAdvantages.forEach(function (adv) {
          var div = document.createElement('div');
          div.className = 'qr-comp';
          div.innerHTML = '<div class="qr-comp-multiplier">' + adv.label + '</div>' +
            '<div class="qr-comp-label">' + adv.description + '</div>';
          compEl.appendChild(div);
        });
      }
    }

    // Takeaway (use the user's Key Takeaway input when provided)
    var takeaway = calc.takeaway;
    if (settings.takeawayOverride && settings.takeawayOverride.trim()) {
      takeaway = settings.takeawayOverride.trim();
    }
    setText('takeaway-text', takeaway);
  }

  function renderFooter(r) {
    setText('f-org', r.organization + ' S.C.');
    setText('f-tagline', r.tagline1);
  }

  // ---- VALIDATION ----
  function validateForm() {
    var valid = true;
    var cards = document.querySelectorAll('.service-card');

    cards.forEach(function (card) {
      var volInput = card.querySelector('[data-field="transactionVolume"]');
      var valInput = card.querySelector('[data-field="totalValue"]');
      var typeEl = card.querySelector('[data-field="type"]');
      var volMsg = volInput ? volInput.parentElement.querySelector('.validation-msg') : null;
      var valMsg = valInput ? valInput.parentElement.querySelector('.validation-msg') : null;

      if (volInput) {
        var vol = parseNum(volInput.value);
        if (volInput.value.trim() !== '' && (isNaN(parseFloat(volInput.value.replace(/,/g, ''))) || vol < 0)) {
          showFieldError(volInput, volMsg, 'Transaction volume must be a valid positive number.');
          valid = false;
        } else {
          clearFieldError(volInput, volMsg);
        }
      }

      if (valInput && typeEl && typeEl.value !== 'non-financial') {
        var val = parseNum(valInput.value);
        if (valInput.value.trim() !== '' && (isNaN(parseFloat(valInput.value.replace(/,/g, ''))) || val < 0)) {
          showFieldError(valInput, valMsg, 'Total transaction value cannot be negative.');
          valid = false;
        } else {
          clearFieldError(valInput, valMsg);
        }
      }
    });

    var statusEl = document.getElementById('form-status');
    if (statusEl) {
      statusEl.textContent = valid ? 'Valid' : 'Validation errors';
      statusEl.className = valid ? 'text-xs text-green-400' : 'text-xs text-red-400';
    }
    return valid;
  }

  function showFieldError(input, msgEl, msg) {
    input.classList.add('input-error');
    if (msgEl) {
      msgEl.textContent = msg;
      msgEl.classList.remove('hidden');
    }
  }

  function clearFieldError(input, msgEl) {
    input.classList.remove('input-error');
    if (msgEl) msgEl.classList.add('hidden');
  }

  // ---- ACTIONS ----
  window.updatePreview = function () {
    renderReport();
    var statusEl = document.getElementById('form-status');
    if (statusEl) {
      statusEl.textContent = 'Updated';
      statusEl.className = 'text-xs text-green-400';
      setTimeout(function () {
        statusEl.textContent = 'Ready';
        statusEl.className = 'text-xs text-navy-300';
      }, 2000);
    }
  };

  window.resetToDefaults = function () {
    if (!confirm('Reset all fields to default reference data?')) return;
    fetch('/api/default-data')
      .then(function (r) { return r.json(); })
      .then(function (data) {
        restoreFormData(data);
        renderReport();
      });
  };

  function restoreFormData(data) {
    setVal('report-title', data.report.title);
    setVal('report-org', data.report.organization);
    setVal('report-brand', data.report.brand);
    setVal('report-tagline1', data.report.tagline1);
    setVal('report-tagline2', data.report.tagline2);
    setVal('report-subtitle', data.report.subtitle);
    setVal('report-date', data.report.date);

    // Settings
    var s = data.settings || {};
    if (document.getElementById('setting-metric-bars')) document.getElementById('setting-metric-bars').checked = s.showMetricBars !== false;
    if (document.getElementById('setting-auto-highlight')) document.getElementById('setting-auto-highlight').checked = s.autoHighlight !== false;
    if (document.getElementById('setting-takeaway-mode')) document.getElementById('setting-takeaway-mode').value = s.takeawayMode || 'auto';
    if (document.getElementById('setting-takeaway-override')) {
      document.getElementById('setting-takeaway-override').value = s.takeawayOverride || '';
      document.getElementById('setting-takeaway-override').disabled = (s.takeawayMode || 'auto') !== 'custom';
    }

    var container = document.getElementById('services-container');
    if (!container) return;
    container.innerHTML = '';

    data.services.forEach(function (svc, idx) {
      addServiceCard(svc, idx);
    });
    lucide.createIcons();
  }

  window.addService = function () {
    var cards = document.querySelectorAll('.service-card');
    var newSvc = {
      id: 'svc-new-' + Date.now(),
      name: 'NEW SERVICE',
      icon: 'circle-dot',
      type: 'financial',
      transactionVolume: 0,
      totalValue: 0,
      target: 0,
      keyMessage: '',
      highlighted: false
    };
    addServiceCard(newSvc, cards.length);
    lucide.createIcons();
  };

  function addServiceCard(svc, idx) {
    var container = document.getElementById('services-container');
    if (!container) return;

    var div = document.createElement('div');
    div.className = 'service-card border border-gray-200 rounded-lg p-4 bg-gray-50 hover:bg-white transition';
    div.dataset.index = idx;
    div.dataset.id = svc.id;

    div.innerHTML =
      '<div class="flex items-center justify-between mb-3">' +
        '<div class="flex items-center gap-2">' +
          '<span class="text-xs font-bold text-navy-700 bg-navy-100 px-2 py-0.5 rounded">#' + (idx + 1) + '</span>' +
          '<input type="text" value="' + escAttr(svc.name) + '" data-field="name" ' +
            'class="text-sm font-bold text-navy-900 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-brand-500 focus:ring-0 px-1 py-0.5 transition">' +
        '</div>' +
        '<div class="flex items-center gap-2">' +
          '<label class="flex items-center gap-1.5 text-xs text-gray-500">' +
            '<input type="checkbox" data-field="highlighted" ' + (svc.highlighted ? 'checked' : '') +
            ' class="w-3.5 h-3.5 rounded border-gray-300 text-brand-500 focus:ring-brand-500"> Highlight' +
          '</label>' +
          '<button onclick="removeService(' + idx + ')" class="text-gray-400 hover:text-red-500 transition" title="Remove">' +
            '<i data-lucide="trash-2" class="w-4 h-4"></i>' +
          '</button>' +
        '</div>' +
      '</div>' +
      '<div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">' +
        '<div>' +
          '<label class="block text-xs text-gray-500 mb-1">Volume</label>' +
          '<input type="text" value="' + fmtNum(svc.transactionVolume) + '" data-field="transactionVolume" ' +
            'class="w-full px-2.5 py-1.5 text-sm border border-gray-200 rounded focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition volume-input" placeholder="0">' +
          '<span class="validation-msg text-xs text-red-500 hidden"></span>' +
        '</div>' +
        '<div>' +
          '<label class="block text-xs text-gray-500 mb-1">Total Value (ETB)</label>' +
          '<input type="text" value="' + (svc.type === 'non-financial' || svc.type === 'success-rate' ? '' : fmtDecimal(svc.totalValue)) + '" data-field="totalValue" ' +
            'class="w-full px-2.5 py-1.5 text-sm border border-gray-200 rounded focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition total-input" ' +
            'placeholder="0.00" ' + (svc.type === 'non-financial' || svc.type === 'success-rate' ? 'disabled' : '') + '>' +
          '<span class="validation-msg text-xs text-red-500 hidden"></span>' +
        '</div>' +
        '<div>' +
          '<label class="block text-xs text-gray-500 mb-1">Monthly Plan (Target)</label>' +
          '<input type="text" value="' + (svc.type === 'success-rate' ? fmtDecimal(svc.target || 0) : fmtNum(svc.target || 0)) + '" data-field="target" ' +
            'class="w-full px-2.5 py-1.5 text-sm border border-gray-200 rounded focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition" ' +
            'placeholder="0">' +
        '</div>' +
        '<div>' +
          '<label class="block text-xs text-gray-500 mb-1">Key Message</label>' +
          '<input type="text" value="' + escAttr(svc.keyMessage || '') + '" data-field="keyMessage" ' +
            'class="w-full px-2.5 py-1.5 text-sm border border-gray-200 rounded focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition" placeholder="Key insight...">' +
        '</div>' +
        '<div>' +
          '<label class="block text-xs text-gray-500 mb-1">Type</label>' +
          '<select data-field="type" onchange="toggleValueType(this)" ' +
            'class="w-full px-2.5 py-1.5 text-sm border border-gray-200 rounded focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition">' +
            '<option value="financial"' + (svc.type === 'financial' ? ' selected' : '') + '>Financial</option>' +
            '<option value="non-financial"' + (svc.type === 'non-financial' ? ' selected' : '') + '>Non-Financial</option>' +
            '<option value="success-rate"' + (svc.type === 'success-rate' ? ' selected' : '') + '>Success Rate (%)</option>' +
          '</select>' +
        '</div>' +
      '</div>';

    container.appendChild(div);

    // Add event listeners for live update
    div.querySelectorAll('input, select').forEach(function (el) {
      el.addEventListener('input', debounce(function () { renderReport(); }, 300));
      el.addEventListener('change', debounce(function () { renderReport(); }, 300));
    });
  }

  window.removeService = function (idx) {
    var container = document.getElementById('services-container');
    if (!container) return;
    var cards = container.querySelectorAll('.service-card');
    if (cards[idx]) {
      cards[idx].remove();
      // Re-index
      container.querySelectorAll('.service-card').forEach(function (card, i) {
        card.dataset.index = i;
        var badge = card.querySelector('.text-xs.font-bold');
        if (badge) badge.textContent = '#' + (i + 1);
        var removeBtn = card.querySelector('button[onclick]');
        if (removeBtn) removeBtn.setAttribute('onclick', 'removeService(' + i + ')');
      });
      renderReport();
    }
  };

  window.toggleValueType = function (sel) {
    var card = sel.closest('.service-card');
    if (!card) return;
    var valInput = card.querySelector('[data-field="totalValue"]');
    var targetInput = card.querySelector('[data-field="target"]');
    var isRate = sel.value === 'success-rate';
    var isNonFin = sel.value === 'non-financial';

    if (valInput) {
      if (isRate || isNonFin) {
        valInput.disabled = true;
        valInput.value = '';
      } else {
        valInput.disabled = false;
      }
    }
    if (targetInput) {
      targetInput.disabled = false;
    }
    renderReport();
  };

  // ---- EXPORT ----
  window.printReport = function () {
    window.print();
  };

  function captureReport() {
    return new Promise(function (resolve, reject) {
      var el = document.getElementById('report-content');
      if (!el || typeof html2canvas === 'undefined') {
        reject(new Error('html2canvas library is not available.'));
        return;
      }
      document.body.classList.add('export-standalone');
      try {
        html2canvas(el, {
          scale: 2,
          useCORS: true,
          backgroundColor: '#ffffff',
          width: 1038,
          height: 735,
          windowWidth: 1038,
          windowHeight: 735
        }).then(function (canvas) {
          document.body.classList.remove('export-standalone');
          resolve(canvas);
        }).catch(function (err) {
          document.body.classList.remove('export-standalone');
          reject(err);
        });
      } catch (err) {
        document.body.classList.remove('export-standalone');
        reject(err);
      }
    });
  }

  window.downloadPDF = function () {
    showExportOverlay('Generating PDF...');
    captureReport().then(function (canvas) {
      var imgData = canvas.toDataURL('image/png');
      var JsPDF = (typeof jspdf !== 'undefined' && jspdf.jsPDF) || null;
      if (!JsPDF) {
        hideExportOverlay();
        alert('PDF library failed to load. Please check your connection and retry.');
        return;
      }
      var pdf = new JsPDF({
        orientation: 'landscape',
        unit: 'px',
        format: [1038, 735]
      });
      pdf.addImage(imgData, 'PNG', 0, 0, 1038, 735);
      pdf.save('Digital_Transaction_Value_Snapshot.pdf');
      hideExportOverlay();
    }).catch(function (err) {
      hideExportOverlay();
      console.error('PDF export error:', err);
      alert('Failed to generate PDF. Please ensure you are online (html2canvas loads from CDN).');
    });
  };

  window.exportImage = function () {
    showExportOverlay('Generating image...');
    captureReport().then(function (canvas) {
      var link = document.createElement('a');
      link.download = 'Digital_Transaction_Value_Snapshot.png';
      link.href = canvas.toDataURL('image/png');
      link.click();
      hideExportOverlay();
    }).catch(function (err) {
      hideExportOverlay();
      console.error('Image export error:', err);
      alert('Failed to generate image. Please ensure you are online (html2canvas loads from CDN).');
    });
  };

  function showExportOverlay(msg) {
    var div = document.createElement('div');
    div.className = 'export-overlay';
    div.id = 'export-overlay';
    div.innerHTML = '<div class="export-overlay-inner"><div class="spinner"></div><div class="text-sm font-medium text-gray-700">' + msg + '</div></div>';
    document.body.appendChild(div);
  }

  function hideExportOverlay() {
    var el = document.getElementById('export-overlay');
    if (el) el.remove();
  }

  // ---- PERSISTENCE (client-side, isolated) ----
  function saveState(obj) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
    } catch (e) { /* storage unavailable - ignore */ }
  }

  function loadState() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      if (!parsed || !parsed.report || !Array.isArray(parsed.services)) return null;
      return parsed;
    } catch (e) {
      return null;
    }
  }

  function clearSavedState() {
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* noop */ }
  }

  // ---- HELPERS ----
  function setText(id, text) {
    var el = document.getElementById(id);
    if (el) el.textContent = text || '';
  }

  function setVal(id, val) {
    var el = document.getElementById(id);
    if (el) el.value = val || '';
  }

  function escHtml(str) {
    var d = document.createElement('div');
    d.textContent = str || '';
    return d.innerHTML;
  }

  function escAttr(str) {
    return (str || '').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function debounce(fn, ms) {
    var t;
    return function () {
      clearTimeout(t);
      t = setTimeout(fn, ms);
    };
  }
})();
