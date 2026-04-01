const COLOR_PALETTE = [
  '#1d4ed8','#dc2626','#16a34a','#7c3aed','#ea580c','#0891b2','#c026d3','#65a30d','#e11d48','#2563eb',
  '#f59e0b','#059669','#9333ea','#0f766e','#b91c1c','#4f46e5','#84cc16','#db2777','#0284c7','#ca8a04',
  '#22c55e','#8b5cf6','#14b8a6','#f97316','#3b82f6','#ef4444','#10b981','#a855f7','#06b6d4','#eab308',
  '#f43f5e','#6366f1','#2dd4bf','#fb7185','#38bdf8','#a3e635','#c084fc','#34d399','#facc15','#f472b6',
  '#60a5fa','#4ade80','#f87171','#818cf8','#67e8f9','#bef264','#f9a8d4','#93c5fd'
];

function colorDistance(a, b) {
  const rgbA = hexToRgb(a);
  const rgbB = hexToRgb(b);
  if (!rgbA || !rgbB) return 0;
  const dr = rgbA.r - rgbB.r;
  const dg = rgbA.g - rgbB.g;
  const db = rgbA.b - rgbB.b;
  return Math.sqrt(dr * dr + dg * dg + db * db);
}

function pickBestAvailableColor(usedColors = new Set()) {
  const used = [...usedColors].filter(Boolean);
  const unused = COLOR_PALETTE.filter(color => !usedColors.has(color));
  const pool = unused.length ? unused : COLOR_PALETTE;
  if (!used.length) return pool[0] || '#64748b';

  let bestColor = pool[0] || '#64748b';
  let bestScore = -Infinity;

  for (const candidate of pool) {
    const minDistance = used.reduce((min, color) => Math.min(min, colorDistance(candidate, color)), Infinity);
    const avgDistance = used.reduce((sum, color) => sum + colorDistance(candidate, color), 0) / used.length;
    const score = (minDistance * 2.25) + avgDistance;
    if (score > bestScore) {
      bestScore = score;
      bestColor = candidate;
    }
  }

  return bestColor;
}

const COLUMN_ALIASES = {
  latitude: ['latitude','lat','geo_lat','customer_latitude','y','geo_y','customer_y'],
  longitude: ['longitude','lng','lon','geo_longitude','customer_longitude','x','geo_x','customer_x'],
  customerId: ['cust id','customer id','customerid','id','account id','acct id'],
  customerName: ['company','customer name','name','account name','cust name'],
  address: ['address','street address','addr','full address'],
  city: ['city','town','municipality','locality'],
  zip: ['zip','zip code','zipcode','postal code'],
  chain: ['chain','chain name'],
  segment: ['segment','customer segment'],
  premise: ['premise','premise type','on/off premise','premise class'],
  currentRep: ['current rep','rep','sales rep','territory rep','owner rep'],
  assignedRep: ['assigned rep','new rep','territory','route','assigned territory'],
  overallSales: ['overall sales','sales','total sales','revenue','$ revenue','$ vol sept - feb','overall revenue','vol sept feb','dollar vol sept feb'],
  rank: ['rank','class','priority rank'],
  cadence4w: [
    'cadence 4w','cadence_4w','cadence4w','4w','planned 4w','planned_4w','planned4w',
    'calls 4w','call cadence 4w','visit cadence 4w','planned calls 4w','frequency','cadence'
  ],
  protected: ['protected','protected account','locked','do not move','never move']
};

const NONE_SELECTED_TOKEN = '__NONE_SELECTED__';

const state = {
  map: null,
  lightLayer: null,
  darkLayer: null,
  markerLayer: null,
  territoryLayer: null,
  territoryLabelLayer: null,
  drawLayer: null,
  drawControl: null,
  workbook: null,
  workbookSheets: {},
  currentSheetName: '',
  currentHeaderMap: {},
  accounts: [],
  accountById: new Map(),
  neighborMap: new Map(),
  markerById: new Map(),
  markerMetaById: new Map(),
  accountPointById: new Map(),
  filterPassById: new Map(),
  repSummaryCache: new Map(),
  globalStatsCache: null,
  territoryRefreshToken: 0,
  territoryRefreshTimer: null,
  territoryDirty: true,
  selection: new Set(),
  undoStack: [],
  changeLog: [],
  repColors: new Map(),
  allReps: new Set(),
  repFocus: null,
  lockedReps: new Set(),
  theme: 'light',
  loadedFileName: 'territory_export_updated.xlsx',
  lastAction: 'No actions yet',
  uploadStatus: { level: 'neutral', text: 'No file loaded' },
  importSummary: {
    sourceRows: 0, loadedRows: 0, skippedNoCoords: 0,
    duplicateCustomerIds: 0, missingCurrentRep: 0,
    missingAssignedRep: 0, unmappedFields: []
  },
  optimizationSummary: null,
  tableSort: { key: 'rep', dir: 'asc' },
  filters: {
    rep: new Set(), rank: new Set(), chain: new Set(), segment: new Set(),
    premise: 'ALL', protected: 'ALL', moved: 'ALL'
  },
  multiSearch: { rep: '', rank: '', chain: '', segment: '', moved: '' },
  openMultiKey: null
};

const els = {};
let toastTimer = null;

document.addEventListener('DOMContentLoaded', init);

function init() {
  bindElements();
  mountUploadStatusPanelToBody();
  initMap();
  bindEvents();
  initMultiFilters();
  updateLastAction('No actions yet');
  fillSimpleSelect(els.premiseFilter, ['ALL'], 'ALL', v => 'All premises');
  renderMultiFilterOptions();
  renderUploadStatus();
  ensureSummaryCardMounts();
  syncControlState();
  initOptimizerTuningUI();
  updateOptimizerUI();
  requestAnimationFrame(() => { if (state.map) state.map.invalidateSize(); });
}

function rebuildMarkers() { renderMap(); }

function setFieldLabelText(field, text, options = {}) {
  if (!field) return null;
  const { preserveValueId = '', valueText = '', valueClassName = '' } = options;
  const candidates = field.querySelectorAll('label, .field-label, .field-title, .field-head, .control-label');
  for (const node of candidates) {
    if (!node) continue;
    const current = safeString(node.textContent).trim();
    if (!current) continue;
    if (preserveValueId) {
      let valueEl = node.querySelector(`#${preserveValueId}`);
      node.textContent = `${text} `;
      if (!valueEl) {
        valueEl = document.createElement('span');
        valueEl.id = preserveValueId;
      }
      if (valueClassName) valueEl.className = valueClassName;
      valueEl.textContent = valueText;
      node.appendChild(valueEl);
      return node;
    }
    node.textContent = text;
    return node;
  }
  return null;
}

function ensureSummaryCardMounts() {
  const stats = document.querySelector('.stats');
  if (!stats) return;
  const ensureCard = (id, label) => {
    let valueEl = document.getElementById(id);
    if (valueEl) return valueEl;
    const card = document.createElement('div');
    card.className = 'stat stat-compact';
    card.innerHTML = `<div class="k">${escapeHtml(label)}</div><div class="v" id="${id}">0</div>`;
    stats.appendChild(card);
    return card.querySelector(`#${id}`);
  };
  els.globalStopsRange = ensureCard('global-stops-range', 'Stops Range');
  els.globalAvgTotalStops = ensureCard('global-avg-total-stops', 'Avg Total Stops');
}

function ensureOptimizerFeedbackMount() {
  if (els.optimizerFeedback) return els.optimizerFeedback;
  const routesCard = els.repTableBody ? els.repTableBody.closest('.card, .routes-card') : null;
  if (!routesCard) return null;
  const head = routesCard.querySelector('.card-head') || routesCard.firstElementChild;
  if (!head) return null;
  let box = head.querySelector('.routes-head-insights');
  if (!box) {
    box = document.createElement('div');
    box.id = 'optimizer-feedback';
    box.className = 'routes-head-insights';
    head.appendChild(box);
  }
  els.optimizerFeedback = box;
  return box;
}

function mountUploadStatusPanelToBody() {
  if (!els.uploadStatusPanel || !document.body) return;
  if (els.uploadStatusPanel.parentElement !== document.body) {
    document.body.appendChild(els.uploadStatusPanel);
  }
}

function ensureBalanceModeOptions() {
  if (!els.balanceMode) return;
  const selected = els.balanceMode.value || 'hybrid';
  els.balanceMode.innerHTML = `
    <option value="hybrid">Hybrid</option>
    <option value="stops">Stops</option>
    <option value="revenue">Revenue</option>
    <option value="compact">Compact</option>
  `;
  els.balanceMode.value = ['hybrid','stops','revenue','compact'].includes(selected) ? selected : 'hybrid';
}

function getOptimizerMode() {
  const mode = (els.balanceMode && els.balanceMode.value) ? String(els.balanceMode.value).toLowerCase() : 'hybrid';
  return ['hybrid', 'stops', 'revenue', 'compact'].includes(mode) ? mode : 'hybrid';
}

function initOptimizerTuningUI() {
  const disruptionField = els.disruptionSlider ? els.disruptionSlider.closest('.field') : null;
  if (disruptionField) {
    disruptionField.classList.add('field-disruption-enhanced', 'field-disruption-compact');
    setFieldLabelText(disruptionField, 'Customer consistency', {
      preserveValueId: 'disruption-value',
      valueText: String(Number(els.disruptionSlider.value || 100))
    });
    let helper = disruptionField.querySelector('#optimizer-disruption-helper');
    if (!helper) {
      helper = document.createElement('div');
      helper.id = 'optimizer-disruption-helper';
      helper.className = 'optimizer-helper optimizer-disruption-helper';
      disruptionField.appendChild(helper);
    }
    helper.textContent = '';
    helper.hidden = true;
    els.optimizerDisruptionHelper = helper;
  }
  if (els.balanceMode) {
    els.balanceMode.disabled = false;
    els.balanceMode.classList.remove('optimizer-mode-hidden');
    els.balanceMode.removeAttribute('aria-hidden');
    els.balanceMode.tabIndex = 0;
    ensureBalanceModeOptions();
    const balanceField = els.balanceMode.closest('.field');
    if (balanceField) {
      setFieldLabelText(balanceField, 'Optimize weight');
      const wrap = balanceField.querySelector('.optimizer-balance-wrap');
      if (wrap) wrap.remove();
    }
  }
  ensureSummaryCardMounts();
  ensureOptimizerFeedbackMount();
}

function getOptimizerMix() {
  const mode = getOptimizerMode();
  if (mode === 'stops') return { stopsPriority: 1, revenuePriority: 0 };
  if (mode === 'revenue') return { stopsPriority: 0, revenuePriority: 1 };
  if (mode === 'compact') return { stopsPriority: 0.85, revenuePriority: 0.15 };
  return { stopsPriority: 0.7, revenuePriority: 0.3 };
}

function getOptimizerWeightLabel() {
  const mode = getOptimizerMode();
  if (mode === 'stops') return 'Stops';
  if (mode === 'revenue') return 'Revenue';
  if (mode === 'compact') return 'Compact';
  return 'Hybrid';
}

function buildRepLoadOrder(targetRepNames, ctx) {
  return [...targetRepNames].sort((a, b) => {
    const countDiff = ctx.count(a) - ctx.count(b);
    if (countDiff !== 0) return countDiff;
    return ctx.revenue(a) - ctx.revenue(b);
  });
}

function getDisruptionPreset(value = Number(els.disruptionSlider ? els.disruptionSlider.value : 100) || 100) {
  if (value >= 85) return { short: 'Minimum change', detail: 'Strongly favors keeping accounts with their current rep.' };
  if (value >= 65) return { short: 'Continuity first', detail: 'Strongly discourages moving accounts unless geography clearly improves.' };
  if (value >= 40) return { short: 'Balanced', detail: 'Blends continuity with geographic compactness.' };
  if (value >= 20) return { short: 'Geography leaning', detail: 'Allows more reassignment to tighten territory shapes.' };
  return { short: 'Geography first', detail: 'Aggressively prioritizes compact territories over continuity.' };
}

function updateOptimizerUI() {
  if (els.disruptionValue && els.disruptionSlider) {
    const preset = getDisruptionPreset(Number(els.disruptionSlider.value) || 0);
    els.disruptionValue.textContent = `${els.disruptionSlider.value}`;
    const disruptionField = els.disruptionSlider ? els.disruptionSlider.closest('.field') : null;
    if (disruptionField) {
      setFieldLabelText(disruptionField, 'Customer consistency', {
        preserveValueId: 'disruption-value',
        valueText: String(els.disruptionSlider.value)
      });
    }
    if (els.optimizerDisruptionHelper) {
      els.optimizerDisruptionHelper.textContent = '';
      els.optimizerDisruptionHelper.hidden = true;
    }
    els.disruptionSlider.title = preset.detail;
  }
  if (els.balanceMode) {
    ensureBalanceModeOptions();
    if (!['hybrid', 'stops', 'revenue', 'compact'].includes(els.balanceMode.value)) {
      els.balanceMode.value = 'hybrid';
    }
  }
}

function renderOptimizationFeedback() {
  const mount = ensureOptimizerFeedbackMount();
  if (!mount) return;
  const s = state.optimizationSummary;
  if (!s) { mount.hidden = true; mount.innerHTML = ''; return; }
  const stopTone = s.stopRangeDeltaPct > 0 ? 'positive' : (s.stopRangeDeltaPct < 0 ? 'negative' : 'neutral');
  const revenueTone = s.revenueRangeDeltaPct > 0 ? 'positive' : (s.revenueRangeDeltaPct < 0 ? 'negative' : 'neutral');
  const stopLabel = s.stopRangeDeltaPct > 0
    ? `Stop spread improved ${formatNumber(s.stopRangeDeltaPct, 1)}%`
    : (s.stopRangeDeltaPct < 0 ? `Stop spread widened ${formatNumber(Math.abs(s.stopRangeDeltaPct), 1)}%` : 'Stop spread unchanged');
  const revenueLabel = s.revenueRangeDeltaPct > 0
    ? `Revenue spread improved ${formatNumber(s.revenueRangeDeltaPct, 1)}%`
    : (s.revenueRangeDeltaPct < 0 ? `Revenue spread widened ${formatNumber(Math.abs(s.revenueRangeDeltaPct), 1)}%` : 'Revenue spread unchanged');
  mount.innerHTML = `
    <div class="optimizer-feedback-metric ${stopTone}">${escapeHtml(stopLabel)}</div>
    <div class="optimizer-feedback-metric ${revenueTone}">${escapeHtml(revenueLabel)}</div>
  `;
  mount.hidden = false;
}

function refreshMarkerStyles(accountIds = null) { refreshMarkers(accountIds); }
function renderSummary() { updateGlobalStats(); }

function syncSortHeaderIndicators() {
  document.querySelectorAll('th[data-sort]').forEach(th => {
    const active = th.getAttribute('data-sort') === state.tableSort.key;
    th.classList.toggle('is-active', active);
    const indicator = th.querySelector('.sort-indicator');
    if (indicator) indicator.textContent = active ? (state.tableSort.dir === 'asc' ? '▲' : '▼') : '↕';
  });
}

function registerRepNames(reps) {
  if (!state.allReps) state.allReps = new Set();
  for (const rep of reps || []) {
    const value = safeString(rep).trim();
    if (value) state.allReps.add(value);
  }
}

function getAvailableReps() {
  const reps = new Set();
  if (state.allReps) {
    for (const rep of state.allReps) {
      const value = safeString(rep).trim();
      if (value) reps.add(value);
    }
  }
  for (const rep of getAllAssignedReps()) {
    const value = safeString(rep).trim();
    if (value) reps.add(value);
  }
  return [...reps].sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
}

function renderRepControls() {
  const reps = getAvailableReps();
  const currentValue = els.assignRepSelect ? els.assignRepSelect.value : '';
  fillSimpleSelect(els.assignRepSelect, reps, reps.includes(currentValue) ? currentValue : '', v => v, 'Select rep');
}

function markTerritoriesDirty() { state.territoryDirty = true; }

function scheduleTerritoryRefresh(force = false) {
  if (force) state.territoryDirty = true;
  const token = ++state.territoryRefreshToken;
  if (state.territoryRefreshTimer) { clearTimeout(state.territoryRefreshTimer); state.territoryRefreshTimer = null; }
  const delay = force ? 0 : 48;
  state.territoryRefreshTimer = setTimeout(() => {
    state.territoryRefreshTimer = null;
    requestAnimationFrame(() => {
      if (token !== state.territoryRefreshToken) return;
      if (!state.territoryDirty && !force) return;
      refreshTerritories();
    });
  }, delay);
}

function invalidateCaches() {
  state.repSummaryCache = new Map();
  state.globalStatsCache = null;
  markTerritoriesDirty();
}

function createEmptyRepSummaryRow(rep) {
  return { rep, stops: 0, deltaStops: 0, revenue: 0, deltaRevenue: 0, A: 0, B: 0, C: 0, D: 0, planned4W: 0, avgWeekly: 0, protected: 0, movedIn: 0, movedOut: 0 };
}

function computeOriginalRepBaseline(rep) {
  const baseline = { stops: 0, revenue: 0 };
  for (const account of state.accounts) {
    const originalRep = account.originalAssignedRep || 'Unassigned';
    if (originalRep !== rep) continue;
    baseline.stops += 1;
    baseline.revenue += Number(account.overallSales || 0);
  }
  return baseline;
}

function computeRepSummaryRow(rep) {
  const row = createEmptyRepSummaryRow(rep);
  const baseline = computeOriginalRepBaseline(rep);
  for (const account of state.accounts) {
    const assignedRep = account.assignedRep || 'Unassigned';
    const originalRep = account.originalAssignedRep || 'Unassigned';
    if (assignedRep === rep) {
      row.stops += 1;
      row.revenue += Number(account.overallSales || 0);
      row.planned4W += Number(account.cadence4w || 0);
      if (row[account.rank] != null) row[account.rank] += 1;
      if (account.protected) row.protected += 1;
      if (assignedRep !== originalRep) row.movedIn += 1;
    }
    if (originalRep === rep && assignedRep !== rep) row.movedOut += 1;
  }
  row.deltaStops = row.stops - baseline.stops;
  row.deltaRevenue = row.revenue - baseline.revenue;
  row.avgWeekly = row.planned4W / 4;
  return row;
}

function updateRepSummaryCacheForReps(reps) {
  if (!reps || !reps.size) return;
  if (!state.repSummaryCache || !state.repSummaryCache.size) summarizeByRep();
  for (const rep of reps) {
    if (!rep) continue;
    const row = computeRepSummaryRow(rep);
    if (row.stops || row.deltaStops || row.revenue || row.deltaRevenue || row.A || row.B || row.C || row.D || row.planned4W || row.protected || row.movedIn || row.movedOut) {
      state.repSummaryCache.set(rep, row);
    } else {
      state.repSummaryCache.delete(rep);
    }
  }
  state.globalStatsCache = null;
}

function computeFilterPass(account) {
  const repOk = state.filters.rep.has(NONE_SELECTED_TOKEN) ? false : (!state.filters.rep.size || state.filters.rep.has(account.assignedRep));
  const rankOk = state.filters.rank.has(NONE_SELECTED_TOKEN) ? false : (!state.filters.rank.size || state.filters.rank.has(account.rank));
  const chainOk = state.filters.chain.has(NONE_SELECTED_TOKEN) ? false : (!state.filters.chain.size || state.filters.chain.has(account.chain));
  const segmentOk = state.filters.segment.has(NONE_SELECTED_TOKEN) ? false : (!state.filters.segment.size || state.filters.segment.has(account.segment));
  const premiseOk = state.filters.premise === 'ALL' || account.premise === state.filters.premise;
  const protectedOk = state.filters.protected === 'ALL' || (state.filters.protected === 'YES' ? account.protected : !account.protected);
  const moved = account.assignedRep !== account.originalAssignedRep;
  const movedOk = state.filters.moved === 'ALL' || (state.filters.moved === 'MOVED' ? moved : !moved);
  return repOk && rankOk && chainOk && segmentOk && premiseOk && protectedOk && movedOk;
}

function updateFilterPassCache() {
  state.filterPassById.clear();
  for (const account of state.accounts) {
    state.filterPassById.set(account._id, computeFilterPass(account));
  }
}

function getChangedRepNamesFromChanges(changes) {
  const reps = new Set();
  for (const change of changes || []) {
    if (change.from) reps.add(change.from);
    if (change.to) reps.add(change.to);
  }
  return reps;
}

function refreshAfterAssignmentBatch(changes, options = {}) {
  const { repsBefore = null, updateSelection = true, territoryForce = false } = options;
  const dirtyIds = new Set((changes || []).map(change => change.id));
  const touchedReps = getChangedRepNamesFromChanges(changes);
  buildRepColors();
  syncRepFilterSelection(Array.isArray(repsBefore) ? repsBefore : null);
  if (state.repFocus && !getAllAssignedReps().includes(state.repFocus)) state.repFocus = null;
  updateFilterPassCache();
  updateRepSummaryCacheForReps(touchedReps);
  markTerritoriesDirty();
  refreshMarkerStyles(dirtyIds.size ? dirtyIds : null);
  renderRepControls();
  renderRepTable();
  renderSummary();
  renderMovedReview();
  if (updateSelection) renderSelectionPreview();
  renderDetail();
  renderOptimizationFeedback();
  syncControlState();
  scheduleTerritoryRefresh(territoryForce);
}

function refreshSelectionMarkerDiff(previousSelection, nextSelection) {
  const dirty = new Set();
  if (previousSelection) for (const id of previousSelection) dirty.add(id);
  if (nextSelection) for (const id of nextSelection) dirty.add(id);
  refreshMarkerStyles(dirty);
}

function bindElements() {
  [
    'file-input','sheet-select','load-sheet-btn','assign-btn','undo-btn','reset-btn','optimize-btn','export-btn',
    'assign-rep-select','rep-count-input','min-stops-input','max-stops-input','disruption-slider','disruption-value','balance-mode',
    'dim-others-checkbox','show-territory-checkbox','rep-table-body','selection-preview','selection-count',
    'global-accounts','global-revenue','global-protected','global-moved','global-unchanged','global-avg-weekly','global-avg-weekly-per-rep','global-stops-range','global-avg-total-stops',
    'last-action','toast','clear-selection-btn','theme-toggle-check','premise-filter','protected-filter',
    'moved-filter','moved-review-list','moved-review-count','rep-filter-options','rank-filter-options','chain-filter-options',
    'segment-filter-options','rep-filter-summary','rank-filter-summary','chain-filter-summary','segment-filter-summary',
    'routes-table-wrap','moved-search-input','upload-status-pill','upload-status-icon','upload-status-text','upload-status-panel','upload-status-body',
    'detail-panel'
  ].forEach(id => { els[toCamel(id)] = document.getElementById(id); });
}

function bindEvents() {
  els.fileInput.addEventListener('change', onFileChosen);
  els.loadSheetBtn.addEventListener('click', loadSelectedSheet);
  els.assignBtn.addEventListener('click', assignSelectionToRep);
  els.undoBtn.addEventListener('click', undoLastAction);
  els.resetBtn.addEventListener('click', resetAssignments);
  els.optimizeBtn.addEventListener('click', optimizeRoutes);
  els.exportBtn.addEventListener('click', exportWorkbook);
  if (els.clearSelectionBtn) els.clearSelectionBtn.addEventListener('click', clearSelection);
  if (els.detailPanel) {
    const detailCard = els.detailPanel.closest('.detail-card');
    const detailClickTarget = detailCard || els.detailPanel;
    detailClickTarget.addEventListener('click', e => {
      const clearBtn = e.target.closest('[data-clear-detail-selection], [data-clear-detail-selection-static]');
      if (clearBtn) { e.preventDefault(); clearSelection(); }
    });
  }
  if (els.uploadStatusPill) {
    els.uploadStatusPill.addEventListener('click', e => { e.stopPropagation(); toggleUploadStatusPanel(); });
  }
  els.themeToggleCheck.addEventListener('change', toggleTheme);
  els.dimOthersCheckbox.addEventListener('change', refreshUI);
  els.showTerritoryCheckbox.addEventListener('change', () => scheduleTerritoryRefresh(true));
  els.premiseFilter.addEventListener('change', () => { state.filters.premise = els.premiseFilter.value; refreshUI(); });
  els.protectedFilter.addEventListener('change', () => { state.filters.protected = els.protectedFilter.value; refreshUI(); });
  els.movedFilter.addEventListener('change', () => { state.filters.moved = els.movedFilter.value; refreshUI(); });
  els.disruptionSlider.addEventListener('input', updateOptimizerUI);
  if (els.movedSearchInput) {
    els.movedSearchInput.addEventListener('input', e => { state.multiSearch.moved = e.target.value || ''; renderMovedReview(); });
  }
  document.querySelectorAll('th[data-sort]').forEach(th => {
    th.addEventListener('click', () => toggleTableSort(th.getAttribute('data-sort')));
  });
  document.addEventListener('click', handleDocumentClickForPanels);
  window.addEventListener('resize', () => {
    if (state.openMultiKey) positionMultiPanel(state.openMultiKey);
    if (els.uploadStatusPanel && !els.uploadStatusPanel.hidden) positionUploadStatusPanel();
    if (state.map) state.map.invalidateSize();
    refreshTerritories();
  });
  window.addEventListener('scroll', () => {
    if (state.openMultiKey) positionMultiPanel(state.openMultiKey);
    if (els.uploadStatusPanel && !els.uploadStatusPanel.hidden) positionUploadStatusPanel();
  }, true);
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') { closeAllMultiPanels(); closeUploadStatusPanel(); }
  });
}

function initMap() {
  state.map = L.map('map', { preferCanvas: true, renderer: L.canvas({ padding: 0.5 }) }).setView([40.1, -89.2], 7);
  state.lightLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '&copy; OpenStreetMap contributors' });
  state.darkLayer = L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', { maxZoom: 19, attribution: '&copy; OpenStreetMap &copy; CARTO' });
  state.lightLayer.addTo(state.map);
  state.markerLayer = L.layerGroup().addTo(state.map);
  state.territoryLayer = L.layerGroup().addTo(state.map);
  state.territoryLabelLayer = L.layerGroup().addTo(state.map);
  state.drawLayer = new L.FeatureGroup().addTo(state.map);
  state.drawControl = new L.Control.Draw({
    draw: {
      polyline: false, circle: false, circlemarker: false, marker: false,
      polygon: { allowIntersection: false, showArea: true, shapeOptions: { color: '#245fb7', weight: 2 } },
      rectangle: { shapeOptions: { color: '#0e9372', weight: 2 } }
    },
    edit: { featureGroup: state.drawLayer, edit: false, remove: true }
  });
  state.map.addControl(state.drawControl);
  state.map.on(L.Draw.Event.CREATED, handleDrawCreated);
  state.map.on(L.Draw.Event.DELETED, () => { clearSelection(); showToast('Selection cleared.'); });
  state.map.on('zoomend', () => scheduleTerritoryRefresh(true));
  state.map.on('moveend', () => scheduleTerritoryRefresh(true));
}

function initMultiFilters() {
  ['rep','rank','chain','segment'].forEach(key => {
    const trigger = document.querySelector(`[data-multi-trigger="${key}"]`);
    const selectAllBtn = document.querySelector(`[data-select-all="${key}"]`);
    const searchInput = document.querySelector(`[data-search="${key}"]`);
    const actions = trigger ? trigger.closest('.multi')?.querySelector('.multi-actions') : null;
    if (actions && !actions.querySelector(`[data-clear-all="${key}"]`)) {
      const clearBtn = document.createElement('button');
      clearBtn.type = 'button';
      clearBtn.className = 'mini-btn';
      clearBtn.setAttribute('data-clear-all', key);
      clearBtn.textContent = 'Clear';
      if (selectAllBtn) selectAllBtn.insertAdjacentElement('afterend', clearBtn);
      else actions.prepend(clearBtn);
    }
    const clearAllBtn = document.querySelector(`[data-clear-all="${key}"]`);
    if (trigger) trigger.addEventListener('click', e => { e.stopPropagation(); toggleMultiPanel(key); });
    if (selectAllBtn) selectAllBtn.addEventListener('click', e => { e.stopPropagation(); selectAllMulti(key); positionMultiPanel(key); });
    if (clearAllBtn) clearAllBtn.addEventListener('click', e => { e.stopPropagation(); clearAllMulti(key); positionMultiPanel(key); });
    if (searchInput) searchInput.addEventListener('input', e => { state.multiSearch[key] = e.target.value || ''; renderMultiFilterOptions(); positionMultiPanel(key); });
  });
}

function handleDocumentClickForPanels(event) {
  const openMulti = document.querySelector('.multi.open');
  if (openMulti && !openMulti.contains(event.target)) closeAllMultiPanels();
  if (els.uploadStatusPanel && !els.uploadStatusPanel.hidden && !els.uploadStatusPanel.contains(event.target) && !els.uploadStatusPill.contains(event.target)) {
    closeUploadStatusPanel();
  }
}

function toggleMultiPanel(key) {
  const wrap = document.getElementById(`${key}-filter-wrap`);
  if (!wrap) return;
  const alreadyOpen = wrap.classList.contains('open');
  closeAllMultiPanels();
  if (!alreadyOpen) { wrap.classList.add('open'); state.openMultiKey = key; positionMultiPanel(key); }
}

function closeAllMultiPanels() {
  document.querySelectorAll('.multi.open').forEach(el => el.classList.remove('open'));
  state.openMultiKey = null;
}

function positionMultiPanel(key) {
  const wrap = document.getElementById(`${key}-filter-wrap`);
  if (!wrap) return;
  const panel = wrap.querySelector('.multi-panel');
  const trigger = wrap.querySelector('.multi-trigger');
  if (!panel || !trigger || !wrap.classList.contains('open')) return;
  const rect = trigger.getBoundingClientRect();
  const width = 220;
  let left = rect.left;
  if (left + width > window.innerWidth - 12) left = window.innerWidth - width - 12;
  if (left < 10) left = 10;
  let top = rect.bottom + 6;
  const availableBelow = window.innerHeight - top - 12;
  let maxHeight = Math.min(360, Math.max(220, availableBelow));
  if (availableBelow < 220) {
    const desiredAbove = Math.min(360, Math.max(220, rect.top - 16));
    top = Math.max(10, rect.top - desiredAbove - 6);
    maxHeight = desiredAbove;
  }
  panel.style.width = `${width}px`;
  panel.style.maxHeight = `${maxHeight}px`;
  panel.style.left = `${left}px`;
  panel.style.top = `${top}px`;
  const list = panel.querySelector('.multi-list');
  if (list) {
    const searchRow = panel.querySelector('.multi-actions');
    const searchHeight = searchRow ? searchRow.offsetHeight : 46;
    list.style.maxHeight = `${Math.max(140, maxHeight - searchHeight - 8)}px`;
  }
}

function selectAllMulti(key) {
  const options = getFilterOptionsForKey(key);
  const filtered = getVisibleOptionsForKey(key, options);
  const selectedSet = state.filters[key];
  selectedSet.delete(NONE_SELECTED_TOKEN);
  filtered.forEach(v => selectedSet.add(v));
  renderMultiFilterOptions();
  refreshUI();
}

function clearAllMulti(key) {
  const selectedSet = state.filters[key];
  selectedSet.clear();
  selectedSet.add(NONE_SELECTED_TOKEN);
  renderMultiFilterOptions();
  refreshUI();
}

function getFilterOptionsForKey(key) {
  switch (key) {
    case 'rep': return getAvailableReps();
    case 'rank': return getDistinctValues(state.accounts, a => a.rank);
    case 'chain': return getDistinctValues(state.accounts, a => a.chain);
    case 'segment': return getDistinctValues(state.accounts, a => a.segment);
    default: return [];
  }
}

function getVisibleOptionsForKey(key, options) {
  const term = (state.multiSearch[key] || '').trim().toLowerCase();
  if (!term) return options;
  return options.filter(v => String(v).toLowerCase().includes(term));
}

function renderMultiFilterOptions() {
  renderMultiOptionList('rep', els.repFilterOptions, els.repFilterSummary, getAvailableReps(), 'All reps');
  renderMultiOptionList('rank', els.rankFilterOptions, els.rankFilterSummary, getDistinctValues(state.accounts, a => a.rank), 'All ranks');
  renderMultiOptionList('chain', els.chainFilterOptions, els.chainFilterSummary, getDistinctValues(state.accounts, a => a.chain), 'All chains');
  renderMultiOptionList('segment', els.segmentFilterOptions, els.segmentFilterSummary, getDistinctValues(state.accounts, a => a.segment), 'All segments');
  if (state.openMultiKey) positionMultiPanel(state.openMultiKey);
}

function renderMultiOptionList(key, container, summaryEl, options, allLabel) {
  if (!container || !summaryEl) return;
  const selectedSet = state.filters[key];
  const visibleOptions = getVisibleOptionsForKey(key, options);
  const hasNone = selectedSet.has(NONE_SELECTED_TOKEN);
  if (options.length && selectedSet.size === 0) options.forEach(v => selectedSet.add(v));
  container.innerHTML = visibleOptions.length
    ? visibleOptions.map(value => {
        const checked = !hasNone && selectedSet.has(value) ? 'checked' : '';
        return `<div class="multi-option"><label><input type="checkbox" data-multi-check="${escapeHtmlAttr(key)}" value="${escapeHtmlAttr(value)}" ${checked} /><span>${escapeHtml(value)}</span></label></div>`;
      }).join('')
    : '<div class="empty">No matches.</div>';
  container.querySelectorAll('input[data-multi-check]').forEach(input => {
    input.addEventListener('change', e => {
      const value = e.target.value;
      selectedSet.delete(NONE_SELECTED_TOKEN);
      if (e.target.checked) selectedSet.add(value);
      else selectedSet.delete(value);
      if (selectedSet.size === 0) selectedSet.add(NONE_SELECTED_TOKEN);
      renderMultiFilterOptions();
      refreshUI();
    });
  });
  if (!options.length || (!hasNone && selectedSet.size === options.length)) summaryEl.textContent = allLabel;
  else if (hasNone) summaryEl.textContent = 'None selected';
  else if (selectedSet.size === 1) summaryEl.textContent = [...selectedSet][0];
  else summaryEl.textContent = `${selectedSet.size} selected`;
}

function toggleUploadStatusPanel() {
  if (!els.uploadStatusPanel) return;
  const willOpen = els.uploadStatusPanel.hidden;
  if (willOpen) {
    els.uploadStatusPanel.hidden = false;
    els.uploadStatusPill.setAttribute('aria-expanded', 'true');
    renderUploadStatusDetails();
    positionUploadStatusPanel();
  } else { closeUploadStatusPanel(); }
}

function closeUploadStatusPanel() {
  if (!els.uploadStatusPanel) return;
  els.uploadStatusPanel.hidden = true;
  els.uploadStatusPill.setAttribute('aria-expanded', 'false');
}

function positionUploadStatusPanel() {
  if (!els.uploadStatusPanel || els.uploadStatusPanel.hidden || !els.uploadStatusPill) return;
  const rect = els.uploadStatusPill.getBoundingClientRect();
  const panel = els.uploadStatusPanel;
  const width = Math.min(320, window.innerWidth - 20);
  let left = rect.left;
  if (left + width > window.innerWidth - 10) left = window.innerWidth - width - 10;
  if (left < 10) left = 10;
  let top = rect.bottom + 8;
  const estimatedHeight = Math.min(420, panel.offsetHeight || 220);
  if (top + estimatedHeight > window.innerHeight - 10) top = Math.max(10, rect.top - estimatedHeight - 8);
  panel.style.width = `${width}px`;
  panel.style.left = `${left}px`;
  panel.style.top = `${top}px`;
}

function setUploadStatus(level, text) { state.uploadStatus = { level, text }; }

function renderUploadStatus() {
  if (!els.uploadStatusPill) return;
  const level = state.uploadStatus.level || 'neutral';
  const text = state.uploadStatus.text || 'No file loaded';
  els.uploadStatusPill.className = `upload-status-pill upload-status-${level}`;
  els.uploadStatusText.textContent = text;
  els.uploadStatusIcon.textContent = level === 'good' ? '✓' : level === 'warning' ? '!' : level === 'bad' ? '×' : '•';
  renderUploadStatusDetails();
}

function renderUploadStatusDetails() {
  if (!els.uploadStatusBody) return;
  const s = state.importSummary || {};
  const warnings = [];
  if (s.skippedNoCoords) warnings.push(`${formatNumber(s.skippedNoCoords)} row(s) skipped for missing latitude/longitude`);
  if (s.duplicateCustomerIds) warnings.push(`${formatNumber(s.duplicateCustomerIds)} duplicate ID(s) adjusted`);
  if (s.missingCurrentRep) warnings.push(`${formatNumber(s.missingCurrentRep)} row(s) missing current rep`);
  if (s.missingAssignedRep) warnings.push(`${formatNumber(s.missingAssignedRep)} blank assigned rep value(s) defaulted from current rep`);
  els.uploadStatusBody.innerHTML = `
    <div class="upload-diag-summary">${escapeHtml(state.currentSheetName || state.uploadStatus.text || '')}</div>
    <div class="upload-diag-grid">
      <div class="upload-diag-label">Source rows</div><div class="upload-diag-value">${formatNumber(s.sourceRows || 0)}</div>
      <div class="upload-diag-label">Loaded rows</div><div class="upload-diag-value">${formatNumber(s.loadedRows || 0)}</div>
      <div class="upload-diag-label">Skipped rows</div><div class="upload-diag-value">${formatNumber(s.skippedNoCoords || 0)}</div>
      <div class="upload-diag-label">Duplicate IDs adjusted</div><div class="upload-diag-value">${formatNumber(s.duplicateCustomerIds || 0)}</div>
      <div class="upload-diag-label">Blank current rep</div><div class="upload-diag-value">${formatNumber(s.missingCurrentRep || 0)}</div>
      <div class="upload-diag-label">Blank assigned rep</div><div class="upload-diag-value">${formatNumber(s.missingAssignedRep || 0)}</div>
    </div>
    <div class="upload-diag-section">
      <div class="upload-diag-label">Warnings</div>
      ${warnings.length ? `<ul class="upload-diag-list">${warnings.map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul>` : `<div class="upload-diag-empty">No actionable warnings.</div>`}
    </div>
  `;
}

function onFileChosen(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  state.loadedFileName = `${file.name.replace(/\.[^.]+$/, '')}_updated.xlsx`;
  closeUploadStatusPanel();
  setUploadStatus('neutral', 'Loading file...');
  renderUploadStatus();
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const lower = file.name.toLowerCase();
      let workbook;
      if (lower.endsWith('.csv')) workbook = XLSX.read(e.target.result, { type: 'binary' });
      else { const arr = new Uint8Array(e.target.result); workbook = XLSX.read(arr, { type: 'array' }); }
      loadWorkbook(workbook);
      setUploadStatus('good', 'Loaded successfully');
      renderUploadStatus();
      showToast('Loaded successfully.');
    } catch (err) {
      console.error('File read failed:', err);
      setUploadStatus('bad', 'Load failed');
      renderUploadStatus();
      showToast('Could not read that file.');
    }
  };
  reader.onerror = () => { setUploadStatus('bad', 'Load failed'); renderUploadStatus(); showToast('Could not read that file.'); };
  if (file.name.toLowerCase().endsWith('.csv')) reader.readAsBinaryString(file);
  else reader.readAsArrayBuffer(file);
}

function loadWorkbook(workbook) {
  state.workbook = workbook;
  state.workbookSheets = {};
  workbook.SheetNames.forEach(name => { state.workbookSheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '' }); });
  fillSimpleSelect(els.sheetSelect, workbook.SheetNames, workbook.SheetNames[0]);
  const hasMultipleSheets = workbook.SheetNames.length > 1;
  if (els.sheetSelect) { els.sheetSelect.disabled = !hasMultipleSheets; els.sheetSelect.classList.toggle('is-hidden', !hasMultipleSheets); }
  if (els.loadSheetBtn) { els.loadSheetBtn.disabled = !hasMultipleSheets; els.loadSheetBtn.classList.toggle('is-hidden', !hasMultipleSheets); }
  loadSelectedSheet();
}

function loadSelectedSheet() {
  const sheetName = els.sheetSelect.value;
  if (!sheetName || !state.workbookSheets[sheetName]) { showToast('No sheet selected.'); return; }
  state.currentSheetName = sheetName;
  state.selection.clear();
  state.undoStack = [];
  state.changeLog = [];
  state.repFocus = null;
  state.optimizationSummary = null;
  state.multiSearch.moved = '';
  if (els.movedSearchInput) els.movedSearchInput.value = '';
  const rows = state.workbookSheets[sheetName] || [];
  const normalizedResult = normalizeRows(rows);
  const normalized = normalizedResult.accounts;
  state.importSummary = normalizedResult.summary;
  if (!normalized.length) {
    state.accounts = [];
    state.accountById = new Map();
    state.neighborMap = new Map();
    state.markerById = new Map();
    state.currentHeaderMap = normalizedResult.headerMap || {};
    invalidateCaches();
    setUploadStatus('bad', 'Load failed');
    renderUploadStatus();
    renderMap();
    refreshUI();
    showToast('No valid rows found.');
    return;
  }
  state.accounts = normalized;
  state.allReps = new Set();
  registerRepNames(normalized.map(a => a.assignedRep || a.currentRep || a.originalAssignedRep));
  state.accountById = new Map(normalized.map(a => [a._id, a]));
  state.neighborMap = buildNeighborMap(normalized);
  state.currentHeaderMap = normalizedResult.headerMap || {};
  invalidateCaches();
  buildRepColors();
  syncRepFilterSelection();
  fillSimpleSelect(els.premiseFilter, ['ALL', ...getDistinctValues(state.accounts, a => a.premise)], 'ALL', v => v === 'ALL' ? 'All premises' : v);
  renderMultiFilterOptions();
  renderMap();
  refreshUI();
  fitMapToAccounts();
  setUploadStatus(
    state.importSummary.skippedNoCoords || state.importSummary.duplicateCustomerIds || state.importSummary.missingCurrentRep ? 'warning' : 'good',
    state.importSummary.skippedNoCoords || state.importSummary.duplicateCustomerIds || state.importSummary.missingCurrentRep ? 'Loaded with warnings' : 'Loaded successfully'
  );
  renderUploadStatus();
}

function normalizeRows(rows) {
  const headerMap = buildHeaderMap(rows);
  const accounts = [];
  const usedIds = new Set();
  let skippedNoCoords = 0, duplicateCustomerIds = 0, missingCurrentRep = 0, missingAssignedRep = 0;
  const unmappedFields = [];
  const allHeaders = rows.length ? Object.keys(rows[0]) : [];
  allHeaders.forEach(h => {
    const cleaned = cleanHeader(h);
    const matched = Object.values(COLUMN_ALIASES).some(list => list.includes(cleaned));
    if (!matched) unmappedFields.push(h);
  });
  rows.forEach((row, index) => {
    let latitude = toNumber(row[headerMap.latitude]);
    let longitude = toNumber(row[headerMap.longitude]);
    ({ latitude, longitude } = normalizeCoordinates(latitude, longitude));
    if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) { skippedNoCoords += 1; return; }
    const customerIdBase = safeString(row[headerMap.customerId]) || `ROW-${index + 1}`;
    let customerId = customerIdBase;
    if (usedIds.has(customerId)) {
      duplicateCustomerIds += 1;
      let n = 2;
      while (usedIds.has(`${customerIdBase}-${n}`)) n += 1;
      customerId = `${customerIdBase}-${n}`;
    }
    usedIds.add(customerId);
    const currentRep = safeString(row[headerMap.currentRep]);
    const assignedRep = safeString(row[headerMap.assignedRep]) || currentRep;
    if (!currentRep) missingCurrentRep += 1;
    if (!safeString(row[headerMap.assignedRep])) missingAssignedRep += 1;
    const rank = normalizeRank(row[headerMap.rank]);
    accounts.push({
      _id: customerId, customerId,
      customerName: safeString(row[headerMap.customerName]) || customerId,
      address: safeString(row[headerMap.address]),
      city: safeString(row[headerMap.city]),
      zip: safeString(row[headerMap.zip]),
      chain: safeString(row[headerMap.chain]),
      segment: safeString(row[headerMap.segment]),
      premise: safeString(row[headerMap.premise]),
      currentRep: currentRep || 'Unassigned',
      assignedRep: assignedRep || 'Unassigned',
      originalAssignedRep: assignedRep || 'Unassigned',
      overallSales: toNumber(row[headerMap.overallSales]) || 0,
      rank,
      cadence4w: normalizeCadence4W(row[headerMap.cadence4w], rank),
      protected: toBoolean(row[headerMap.protected]),
      latitude: round6(latitude), longitude: round6(longitude),
      sourceRow: row
    });
  });
  return { accounts, headerMap, summary: { sourceRows: rows.length, loadedRows: accounts.length, skippedNoCoords, duplicateCustomerIds, missingCurrentRep, missingAssignedRep, unmappedFields } };
}

function buildHeaderMap(rows) {
  const firstRow = rows[0] || {};
  const map = {};
  const cleanedHeaders = Object.keys(firstRow).map(h => ({ original: h, cleaned: cleanHeader(h) }));
  Object.entries(COLUMN_ALIASES).forEach(([key, aliases]) => {
    const cleanedAliases = aliases.map(a => cleanHeader(a));
    const match = cleanedHeaders.find(h => cleanedAliases.includes(h.cleaned));
    map[key] = match ? match.original : null;
  });
  return map;
}

function normalizeCoordinates(latitude, longitude) {
  if (Number.isFinite(latitude) && Number.isFinite(longitude)) {
    if (Math.abs(latitude) > 90 && Math.abs(longitude) <= 90) return { latitude: longitude, longitude: latitude };
  }
  return { latitude, longitude };
}

function buildNeighborMap(accounts) {
  const map = new Map();
  if (!accounts.length) return map;
  let minLat = Infinity, maxLat = -Infinity, minLng = Infinity, maxLng = -Infinity;
  for (const a of accounts) {
    if (a.latitude < minLat) minLat = a.latitude;
    if (a.latitude > maxLat) maxLat = a.latitude;
    if (a.longitude < minLng) minLng = a.longitude;
    if (a.longitude > maxLng) maxLng = a.longitude;
  }
  const spread = Math.max(maxLng - minLng, maxLat - minLat);
  const densityFactor = Math.sqrt(accounts.length) / 100;
  const cellSize = Math.max(0.05, Math.min(0.5, spread / (40 + densityFactor) || 0.18));
  const grid = new Map();
  const cellKey = (lat, lng) => `${Math.floor(lng / cellSize)}|${Math.floor(lat / cellSize)}`;
  for (const a of accounts) {
    map.set(a._id, new Set());
    const key = cellKey(a.latitude, a.longitude);
    if (!grid.has(key)) grid.set(key, []);
    grid.get(key).push(a);
  }
  for (const a of accounts) {
    const x = Math.floor(a.longitude / cellSize);
    const y = Math.floor(a.latitude / cellSize);
    const candidates = [];
    for (let dx = -1; dx <= 1; dx++) {
      for (let dy = -1; dy <= 1; dy++) {
        const cell = grid.get(`${x + dx}|${y + dy}`);
        if (cell) for (const o of cell) { if (o._id !== a._id) candidates.push(o); }
      }
    }
    if (candidates.length < 10) {
      for (let r = 2; r <= 4 && candidates.length < 20; r++) {
        for (let dx = -r; dx <= r; dx++) {
          for (let dy = -r; dy <= r; dy++) {
            if (Math.abs(dx) !== r && Math.abs(dy) !== r) continue;
            const cell = grid.get(`${x + dx}|${y + dy}`);
            if (cell) for (const o of cell) { if (o._id !== a._id) candidates.push(o); }
          }
        }
      }
    }
    candidates
      .map(o => ({ id: o._id, d: squaredDistance(a.latitude, a.longitude, o.latitude, o.longitude) }))
      .sort((a, b) => a.d - b.d)
      .slice(0, 10)
      .forEach(item => { map.get(a._id).add(item.id); map.get(item.id)?.add(a._id); });
  }
  return map;
}

function renderMap() {
  state.markerLayer.clearLayers();
  state.markerById.clear();
  state.markerMetaById.clear();
  state.accountPointById.clear();
  updateFilterPassCache();
  for (const account of state.accounts) {
    const color = getRepColor(account.assignedRep);
    const marker = L.circleMarker([account.latitude, account.longitude], { radius: 2.8, color, weight: 1, opacity: 0.95, fillColor: color, fillOpacity: 0.88 });
    marker.on('click', e => {
      L.DomEvent.stopPropagation(e);
      const additive = !!(e.originalEvent?.ctrlKey || e.originalEvent?.metaKey);
      toggleSelection(account._id, additive);
      marker.setPopupContent(buildPopupHtml(account));
      marker.openPopup();
    });
    marker.bindPopup(buildPopupHtml(account), { autoPan: true, closeButton: true, offset: [0, -6] });
    state.markerLayer.addLayer(marker);
    state.markerById.set(account._id, marker);
    state.markerMetaById.set(account._id, { color, radius: 2.8, opacity: 0.95, fillOpacity: 0.88, weight: 1, hidden: false, popupKey: `${account.assignedRep}|${account.currentRep}|${account.overallSales}|${account.rank}|${account.protected ? 1 : 0}` });
    state.accountPointById.set(account._id, turf.point([account.longitude, account.latitude]));
  }
  invalidateCaches();
  refreshMarkerStyles();
  scheduleTerritoryRefresh(true);
}

function buildPopupHtml(account) {
  const title = account.customerName || account.customerId;
  const line2 = [account.address, [account.city, account.zip].filter(Boolean).join(' ')].filter(Boolean).join(' • ');
  return `
    <div style="min-width:240px;max-width:280px;">
      <div style="font-size:15px;font-weight:800;line-height:1.2;margin-bottom:4px;">${escapeHtml(title)}</div>
      <div style="font-size:12px;color:#5d7286;line-height:1.35;margin-bottom:8px;">${escapeHtml(line2)}</div>
      <div style="display:grid;grid-template-columns:auto 1fr;gap:4px 8px;font-size:12px;line-height:1.3;">
        <div><strong>Rep</strong></div><div>${escapeHtml(account.assignedRep)}</div>
        <div><strong>Current</strong></div><div>${escapeHtml(account.currentRep)}</div>
        <div><strong>Revenue</strong></div><div>${formatCurrency(account.overallSales)}</div>
        <div><strong>Rank</strong></div><div>${escapeHtml(account.rank)}</div>
        <div><strong>Protected</strong></div><div>${account.protected ? 'Yes' : 'No'}</div>
      </div>
    </div>`;
}

function ensureDetailClearButton() {
  if (!els.detailPanel) return null;
  const detailCard = els.detailPanel.closest('.detail-card');
  if (!detailCard) return null;
  const head = detailCard.querySelector('.card-head');
  if (!head) return null;
  head.style.cssText = 'display:flex;align-items:flex-start;justify-content:space-between;gap:12px;';
  let button = head.querySelector('[data-clear-detail-selection-static]');
  if (!button) {
    button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn btn-subtle';
    button.setAttribute('data-clear-detail-selection-static', 'true');
    button.textContent = 'Clear All';
    head.appendChild(button);
  }
  button.disabled = state.selection.size === 0;
  return button;
}

function renderDetail() {
  if (!els.detailPanel) return;
  ensureDetailClearButton();
  const selectedIds = [...state.selection];
  if (!selectedIds.length) { els.detailPanel.innerHTML = '<div class="empty">No account selected.</div>'; return; }
  const selectedAccounts = selectedIds.map(id => state.accountById.get(id)).filter(Boolean).slice(0, 10);
  const cardsHtml = selectedAccounts.map(account => `
    <div class="selected-item" style="margin-bottom:10px;">
      <div class="selected-item-title">${escapeHtml(account.customerName || account.customerId)}</div>
      <div style="font-size:12px;color:#5d7286;margin-bottom:6px;">${escapeHtml([account.address, [account.city, account.zip].filter(Boolean).join(' ')].filter(Boolean).join(' • '))}</div>
      <div class="transfer-line">
        <span class="rep-chip">${escapeHtml(account.assignedRep)}</span>
        <span class="metric-chip">${formatCurrency(account.overallSales)}</span>
        <span class="metric-chip">Rank ${escapeHtml(account.rank)}</span>
        ${account.protected ? '<span class="metric-chip">Protected</span>' : ''}
      </div>
    </div>`).join('');
  const moreCount = selectedIds.length - selectedAccounts.length;
  els.detailPanel.innerHTML = `
    <div style="font-size:12px;color:#5d7286;font-weight:700;margin-bottom:10px;">${selectedIds.length} selected</div>
    ${cardsHtml}
    ${moreCount > 0 ? `<div class="small muted" style="margin-top:6px;">Showing first ${selectedAccounts.length}. ${moreCount} more selected.</div>` : ''}`;
}

function refreshUI(rebuildMap = false) {
  syncControlState();
  renderRepControls();
  renderUploadStatus();
  updateOptimizerUI();
  if (rebuildMap) rebuildMarkers();
  updateFilterPassCache();
  refreshMarkerStyles();
  renderRepTable();
  renderSelectionPreview();
  renderSummary();
  renderMovedReview();
  renderDetail();
  renderOptimizationFeedback();
  if (state.openMultiKey) positionMultiPanel(state.openMultiKey);
  scheduleTerritoryRefresh();
}

function passesFilters(account) {
  if (!account) return false;
  if (state.filterPassById.has(account._id)) return !!state.filterPassById.get(account._id);
  return computeFilterPass(account);
}

function refreshMarkers(accountIds = null) {
  const dimOthers = !!els.dimOthersCheckbox.checked;
  const ids = accountIds ? Array.from(accountIds) : state.accounts.map(a => a._id);
  for (const id of ids) {
    const account = state.accountById.get(id);
    const marker = state.markerById.get(id);
    if (!account || !marker) continue;
    const pass = passesFilters(account);
    const color = getRepColor(account.assignedRep);
    const selected = state.selection.has(id);
    const focusedOut = state.repFocus && account.assignedRep !== state.repFocus;
    const dimmed = dimOthers && focusedOut && !selected;
    const nextState = {
      color,
      radius: selected ? 4.2 : dimmed ? 1.65 : (state.repFocus && account.assignedRep === state.repFocus ? 3.8 : 2.8),
      opacity: pass ? (dimmed ? 0.02 : 0.98) : 0,
      fillOpacity: pass ? (selected ? 1 : dimmed ? 0.015 : 0.92) : 0,
      weight: selected ? 2 : (dimmed ? 0.5 : 1),
      hidden: !pass
    };
    const prevState = state.markerMetaById.get(id) || {};
    if (prevState.color !== nextState.color || prevState.opacity !== nextState.opacity || prevState.fillOpacity !== nextState.fillOpacity || prevState.weight !== nextState.weight) {
      marker.setStyle({ color: nextState.color, fillColor: nextState.color, opacity: nextState.opacity, fillOpacity: nextState.fillOpacity, weight: nextState.weight });
    }
    if (prevState.radius !== nextState.radius) marker.setRadius(nextState.radius);
    const popupKey = `${account.assignedRep}|${account.currentRep}|${account.overallSales}|${account.rank}|${account.protected ? 1 : 0}`;
    if (prevState.popupKey !== popupKey) { marker.setPopupContent(buildPopupHtml(account)); nextState.popupKey = popupKey; }
    else nextState.popupKey = prevState.popupKey;
    state.markerMetaById.set(id, nextState);
  }
}

function refreshTerritories() {
  state.territoryDirty = false;
  state.territoryLayer.clearLayers();
  state.territoryLabelLayer.clearLayers();
  if (!els.showTerritoryCheckbox.checked) return;
  const membersByRep = new Map();
  for (const a of state.accounts) {
    if (!passesFilters(a)) continue;
    if (!membersByRep.has(a.assignedRep)) membersByRep.set(a.assignedRep, []);
    membersByRep.get(a.assignedRep).push(a);
  }
  for (const [rep, members] of membersByRep.entries()) {
    if (members.length < 3) continue;
    const points = members.map(m => turf.point([m.longitude, m.latitude]));
    let hull = null;
    try { hull = turf.convex(turf.featureCollection(points)); } catch (e) { hull = null; }
    if (!hull) continue;
    const strokeColor = getTerritoryStrokeColor(rep);
    const fillColor = getTerritoryFillColor(rep);
    state.territoryLayer.addLayer(L.geoJSON(hull, { style: { color: strokeColor, weight: 2.15, fillColor, fillOpacity: 0.09, opacity: 0.8, dashArray: '2 4' } }));
    const center = turf.center(hull).geometry.coordinates;
    state.territoryLabelLayer.addLayer(L.marker([center[1], center[0]], {
      interactive: false,
      icon: L.divIcon({ className: 'territory-label', html: `<div style="background:${getTerritoryLabelColor(rep)};color:#fff;border:1px solid rgba(255,255,255,.52);box-shadow:0 6px 14px rgba(30,54,84,.14);border-radius:999px;padding:4px 10px;font-size:11px;font-weight:800;white-space:nowrap;">${escapeHtml(rep)}</div>` })
    }));
  }
}

function handleDrawCreated(event) {
  state.drawLayer.clearLayers();
  const layer = event.layer;
  state.drawLayer.addLayer(layer);
  let polygon = null;
  if (layer instanceof L.Rectangle || layer instanceof L.Polygon) polygon = layer.toGeoJSON();
  if (!polygon) return;
  const bbox = turf.bbox(polygon);
  const nextSelection = new Set();
  for (const a of state.accounts) {
    if (a.longitude < bbox[0] || a.longitude > bbox[2] || a.latitude < bbox[1] || a.latitude > bbox[3]) continue;
    const point = state.accountPointById.get(a._id) || turf.point([a.longitude, a.latitude]);
    if (turf.booleanPointInPolygon(point, polygon)) nextSelection.add(a._id);
  }
  const previousSelection = new Set(state.selection);
  state.selection = nextSelection;
  refreshSelectionMarkerDiff(previousSelection, nextSelection);
  renderSelectionPreview();
  renderDetail();
  syncControlState();
  showToast(`${state.selection.size} account${state.selection.size === 1 ? '' : 's'} selected.`);
}

function toggleSelection(id, additive = false) {
  const previousSelection = new Set(state.selection);
  if (!additive) state.selection.clear();
  if (state.selection.has(id)) state.selection.delete(id);
  else state.selection.add(id);
  refreshSelectionMarkerDiff(previousSelection, state.selection);
  renderSelectionPreview();
  renderDetail();
  syncControlState();
}

function clearSelection() {
  const previousSelection = new Set(state.selection);
  state.selection.clear();
  state.drawLayer.clearLayers();
  refreshSelectionMarkerDiff(previousSelection, state.selection);
  renderSelectionPreview();
  renderDetail();
  syncControlState();
}

function isRepLocked(rep) { return state.lockedReps.has(rep); }
function isAccountLocked(account) { return !!account && isRepLocked(account.assignedRep); }

function toggleRepLock(rep, shouldLock) {
  if (!rep) return;
  if (shouldLock) state.lockedReps.add(rep); else state.lockedReps.delete(rep);
  invalidateCaches();
  refreshMarkerStyles();
  renderRepTable();
  renderSummary();
  scheduleTerritoryRefresh();
  updateLastAction(`${shouldLock ? 'Locked' : 'Unlocked'} territory: ${rep}`);
  showToast(`${shouldLock ? 'Locked' : 'Unlocked'} ${rep}.`);
}

function renderRepTable() {
  let rows = summarizeByRep();
  sortRepRows(rows);
  if (!rows.length) { els.repTableBody.innerHTML = '<tr><td colspan="15" class="empty">Upload a file to begin.</td></tr>'; return; }
  syncSortHeaderIndicators();
  els.repTableBody.innerHTML = rows.map(row => `
    <tr data-rep-row="${encodeURIComponent(row.rep)}" class="${state.repFocus === row.rep ? 'rep-row-active' : ''} ${isRepLocked(row.rep) ? 'rep-row-locked' : ''}">
      <td><div class="rep-cell"><span class="color-dot" style="background:${getRepColor(row.rep)}"></span><span>${escapeHtml(row.rep)}</span></div></td>
      <td class="lock-cell"><label class="lock-toggle" title="Lock this territory"><input type="checkbox" class="rep-lock-checkbox" data-lock-rep="${escapeHtmlAttr(row.rep)}" ${isRepLocked(row.rep) ? 'checked' : ''} /></label></td>
      <td>${formatNumber(row.stops)}</td>
      <td>${renderDeltaCount(row.deltaStops)}</td>
      <td>${formatCurrency(row.revenue)}</td>
      <td>${renderDeltaMoney(row.deltaRevenue)}</td>
      <td>${formatNumber(row.A)}</td>
      <td>${formatNumber(row.B)}</td>
      <td>${formatNumber(row.C)}</td>
      <td>${formatNumber(row.D)}</td>
      <td>${formatNumber(row.planned4W, 2)}</td>
      <td>${formatNumber(row.avgWeekly, 2)}</td>
      <td>${formatNumber(row.protected)}</td>
      <td>${formatNumber(row.movedIn)}</td>
      <td>${formatNumber(row.movedOut)}</td>
    </tr>`).join('');
  els.repTableBody.querySelectorAll('tr[data-rep-row]').forEach(tr => {
    tr.addEventListener('click', e => {
      if (e.target.closest('.rep-lock-checkbox')) return;
      const rep = decodeURIComponent(tr.getAttribute('data-rep-row'));
      state.repFocus = state.repFocus === rep ? null : rep;
      refreshMarkerStyles();
      renderRepTable();
      scheduleTerritoryRefresh();
      if (state.repFocus) zoomToRep(state.repFocus);
    });
  });
  els.repTableBody.querySelectorAll('.rep-lock-checkbox').forEach(input => {
    input.addEventListener('click', e => e.stopPropagation());
    input.addEventListener('change', e => toggleRepLock(e.target.getAttribute('data-lock-rep'), e.target.checked));
  });
}

function summarizeByRep() {
  if (state.repSummaryCache && state.repSummaryCache.size) return [...state.repSummaryCache.values()].map(row => ({ ...row }));
  const map = new Map();
  const originalMap = new Map();
  for (const a of state.accounts) {
    const assignedRep = a.assignedRep || 'Unassigned';
    const originalRep = a.originalAssignedRep || 'Unassigned';
    if (!map.has(assignedRep)) map.set(assignedRep, { rep: assignedRep, stops: 0, deltaStops: 0, revenue: 0, deltaRevenue: 0, A: 0, B: 0, C: 0, D: 0, planned4W: 0, avgWeekly: 0, protected: 0, movedIn: 0, movedOut: 0 });
    if (!originalMap.has(originalRep)) originalMap.set(originalRep, { stops: 0, revenue: 0 });
    const row = map.get(assignedRep);
    row.stops += 1;
    row.revenue += Number(a.overallSales || 0);
    row.planned4W += Number(a.cadence4w || 0);
    if (row[a.rank] != null) row[a.rank] += 1;
    if (a.protected) row.protected += 1;
    if (assignedRep !== originalRep) row.movedIn += 1;
    const orig = originalMap.get(originalRep);
    orig.stops += 1;
    orig.revenue += Number(a.overallSales || 0);
  }
  for (const a of state.accounts) {
    const orig = a.originalAssignedRep || 'Unassigned';
    const assigned = a.assignedRep || 'Unassigned';
    if (orig !== assigned && map.has(orig)) map.get(orig).movedOut += 1;
  }
  for (const row of map.values()) {
    const orig = originalMap.get(row.rep) || { stops: 0, revenue: 0 };
    row.deltaStops = row.stops - orig.stops;
    row.deltaRevenue = row.revenue - orig.revenue;
    row.avgWeekly = row.planned4W / 4;
  }
  state.repSummaryCache = map;
  return [...map.values()].map(row => ({ ...row }));
}

function sortRepRows(rows) {
  const { key, dir } = state.tableSort;
  const factor = dir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    let av = a[key], bv = b[key];
    if (typeof av === 'string' || typeof bv === 'string') return String(av).localeCompare(String(bv), undefined, { numeric: true }) * factor;
    return (Number(av || 0) - Number(bv || 0)) * factor;
  });
}

function toggleTableSort(key) {
  if (state.tableSort.key === key) state.tableSort.dir = state.tableSort.dir === 'asc' ? 'desc' : 'asc';
  else state.tableSort = { key, dir: 'asc' };
  syncSortHeaderIndicators();
  renderRepTable();
}

function renderSelectionPreview() {
  if (!els.selectionPreview || !els.selectionCount) return;
  const ids = [...state.selection];
  els.selectionCount.textContent = ids.length;
  if (!ids.length) { els.selectionPreview.innerHTML = '<div class="empty">No accounts selected.</div>'; return; }
  els.selectionPreview.innerHTML = ids.slice(0, 50).map(id => {
    const a = state.accountById.get(id);
    if (!a) return '';
    return `<div class="selected-item"><div class="selected-item-title">${escapeHtml(a.customerName)}</div><div class="transfer-line"><span class="rep-chip">${escapeHtml(a.assignedRep)}</span><span class="metric-chip">${formatCurrency(a.overallSales)}</span></div></div>`;
  }).join('');
}

function renderMovedReview() {
  if (!els.movedReviewList || !els.movedReviewCount) return;
  const term = (state.multiSearch.moved || '').trim().toLowerCase();
  let moved = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep);
  if (term) moved = moved.filter(a => [a.customerName, a.customerId, a.originalAssignedRep, a.assignedRep].some(v => String(v || '').toLowerCase().includes(term)));
  els.movedReviewCount.textContent = moved.length;
  if (!moved.length) { els.movedReviewList.innerHTML = '<div class="empty">No moved accounts yet.</div>'; return; }
  els.movedReviewList.innerHTML = moved.slice(0, 250).map(a => `
    <div class="moved-item">
      <div class="moved-item-title">${escapeHtml(a.customerName)}</div>
      <div class="transfer-line">
        <span class="metric-chip">${escapeHtml(a.customerId)}</span>
        <span class="rep-chip">${escapeHtml(a.originalAssignedRep)}</span>
        <span class="rep-arrow">→</span>
        <span class="rep-chip">${escapeHtml(a.assignedRep)}</span>
        <span class="metric-chip">${formatCurrency(a.overallSales)}</span>
      </div>
    </div>`).join('');
}

function updateGlobalStats() {
  let stats = state.globalStatsCache;
  if (!stats) {
    let visibleCount = 0, movedCount = 0, totalRevenue = 0, totalProtected = 0, planned4W = 0;
    for (const a of state.accounts) {
      if (a.assignedRep !== a.originalAssignedRep) movedCount += 1;
      if (!passesFilters(a)) continue;
      visibleCount += 1;
      totalRevenue += a.overallSales || 0;
      totalProtected += a.protected ? 1 : 0;
      planned4W += a.cadence4w || 0;
    }
    const reps = getAvailableReps();
    const unchangedPct = state.accounts.length ? ((state.accounts.length - movedCount) / state.accounts.length) * 100 : 0;
    stats = { visibleCount, movedCount, unchangedPct, totalRevenue, totalProtected, planned4W, repCount: reps.length };
    state.globalStatsCache = stats;
  }
  els.globalAccounts.textContent = formatNumber(stats.visibleCount);
  els.globalRevenue.textContent = formatCurrency(stats.totalRevenue);
  els.globalProtected.textContent = formatNumber(stats.totalProtected);
  els.globalMoved.textContent = formatNumber(stats.movedCount);
  els.globalUnchanged.textContent = `${formatNumber(stats.unchangedPct, 1)}%`;
  els.globalAvgWeekly.textContent = formatNumber(stats.planned4W / 4, 1);
  els.globalAvgWeeklyPerRep.textContent = formatNumber((stats.planned4W / 4) / Math.max(1, stats.repCount), 1);
  const rows = summarizeByRep();
  const stopValues = rows.map(r => Number(r.stops || 0)).filter(Number.isFinite);
  if (els.globalStopsRange) els.globalStopsRange.textContent = stopValues.length ? `${formatNumber(Math.min(...stopValues))}-${formatNumber(Math.max(...stopValues))}` : '0-0';
  if (els.globalAvgTotalStops) {
    const avg = rows.length ? stopValues.reduce((s, v) => s + v, 0) / rows.length : 0;
    els.globalAvgTotalStops.textContent = formatNumber(avg, 1);
  }
}

function syncRepFilterSelection(previousAssignedReps = null) {
  const reps = getAvailableReps();
  const currentSelected = state.filters.rep instanceof Set ? new Set(state.filters.rep) : new Set();
  const previousSet = Array.isArray(previousAssignedReps) ? new Set(previousAssignedReps) : null;
  const nextSelected = new Set();
  const hadExplicitNone = currentSelected.has(NONE_SELECTED_TOKEN);
  reps.forEach(rep => {
    if (currentSelected.has(rep)) { nextSelected.add(rep); return; }
    if (previousSet && previousSet.has(rep)) { nextSelected.add(rep); return; }
    if (!hadExplicitNone) nextSelected.add(rep);
  });
  if (!nextSelected.size && hadExplicitNone) nextSelected.add(NONE_SELECTED_TOKEN);
  if (!nextSelected.size) reps.forEach(rep => nextSelected.add(rep));
  state.filters.rep = nextSelected;
  fillSimpleSelect(els.assignRepSelect, reps, '', v => v, 'Select rep');
  renderMultiFilterOptions();
}

function syncControlState() {
  const hasAccounts = state.accounts.length > 0;
  const hasSelection = state.selection.size > 0;
  els.assignBtn.disabled = !hasAccounts || !hasSelection;
  els.undoBtn.disabled = state.undoStack.length === 0;
  els.resetBtn.disabled = !hasAccounts;
  els.optimizeBtn.disabled = !hasAccounts;
  els.exportBtn.disabled = !hasAccounts;
  if (els.clearSelectionBtn) els.clearSelectionBtn.disabled = !hasSelection;
  els.assignRepSelect.disabled = !hasAccounts;
  const reps = getAvailableReps();
  if (document.activeElement !== els.repCountInput) els.repCountInput.value = reps.length || 1;
}

function assignSelectionToRep() {
  const targetRep = els.assignRepSelect.value;
  registerRepNames([targetRep]);
  const selectedIds = [...state.selection];
  if (!selectedIds.length || !targetRep) return;
  if (isRepLocked(targetRep)) { showToast(`"${targetRep}" is locked.`); return; }
  const previousAssignedReps = getAllAssignedReps();
  const changes = [];
  let skippedProtected = 0, skippedLocked = 0;
  ensureRepColor(targetRep);
  for (const id of selectedIds) {
    const account = state.accountById.get(id);
    if (!account) continue;
    if (isAccountLocked(account)) { skippedLocked += 1; continue; }
    if (account.protected && account.assignedRep !== targetRep) { skippedProtected += 1; continue; }
    if (account.assignedRep === targetRep) continue;
    changes.push({ id, from: account.assignedRep, to: targetRep });
  }
  if (!changes.length) {
    if (skippedLocked) { showToast(`${skippedLocked} account(s) belong to locked territories.`); return; }
    showToast(skippedProtected ? `${skippedProtected} protected account(s) were skipped.` : 'No assignment changes to make.');
    return;
  }
  applyChanges(changes, `Assigned ${changes.length} account${changes.length === 1 ? '' : 's'} to ${targetRep}`, previousAssignedReps);
  clearSelection();
}

function applyChanges(changes, label, previousAssignedReps = null) {
  const repsBefore = Array.isArray(previousAssignedReps) ? previousAssignedReps : getAllAssignedReps();
  const appliedChanges = [];
  changes.forEach(change => {
    const account = state.accountById.get(change.id);
    if (!account) return;
    if (isRepLocked(change.from) || isRepLocked(change.to) || isAccountLocked(account)) return;
    if (account.assignedRep === change.to) return;
    ensureRepColor(change.to);
    account.assignedRep = change.to;
    appliedChanges.push({ ...change });
    state.changeLog.push({ timestamp: new Date().toLocaleString(), customerId: account.customerId, customerName: account.customerName, fromRep: change.from, toRep: change.to, protected: account.protected ? 'Yes' : 'No' });
  });
  if (!appliedChanges.length) { showToast('No eligible changes could be applied.'); return; }
  state.undoStack.push({ changes: appliedChanges, label });
  refreshAfterAssignmentBatch(appliedChanges, { repsBefore, updateSelection: true, territoryForce: false });
  updateLastAction(label);
  showToast(label);
}

function undoLastAction() {
  const action = state.undoStack.pop();
  if (!action) return;
  const repsBefore = getAllAssignedReps();
  for (const change of action.changes) { const a = state.accountById.get(change.id); if (a) a.assignedRep = change.from; }
  state.optimizationSummary = null;
  refreshAfterAssignmentBatch(action.changes, { repsBefore, updateSelection: true, territoryForce: false });
  updateLastAction(`Undid: ${action.label}`);
  showToast(`Undid: ${action.label}`);
}

function resetAssignments() {
  registerRepNames(state.accounts.map(a => a.originalAssignedRep || a.currentRep || a.assignedRep));
  const resetChanges = [];
  state.accounts.forEach(a => {
    if (a.assignedRep !== a.originalAssignedRep) {
      resetChanges.push({ id: a._id, from: a.assignedRep, to: a.originalAssignedRep });
      a.assignedRep = a.originalAssignedRep;
    }
  });
  if (!resetChanges.length) { showToast('Nothing to reset.'); return; }
  state.undoStack = [];
  state.changeLog = [];
  state.repFocus = null;
  state.optimizationSummary = null;
  state.multiSearch.moved = '';
  if (els.movedSearchInput) els.movedSearchInput.value = '';
  invalidateCaches();
  refreshAfterAssignmentBatch(resetChanges, { repsBefore: null, updateSelection: true, territoryForce: true });
  fitMapToAccounts();
  updateLastAction('Reset assignments to imported values');
  showToast('Assignments reset to imported values.');
}

// ─── OPTIMIZER ───────────────────────────────────────────────

function runEnclaveCleanupFast(assignments, targetRepNames, minStops, adjacency, ctx) {
  runBorderCleanupFast(assignments, targetRepNames, 0, minStops, adjacency, ctx);
}

function runMajoritySmoothingFast(assignments, targetRepNames, minStops, adjacency, ctx) {
  runBorderCleanupFast(assignments, targetRepNames, 0, minStops, adjacency, ctx);
}

function optimizeRoutes() {
  if (!state.accounts.length) return;
  try {
    const targetCount = Math.max(1, Math.min(100, parseInt(els.repCountInput.value || '1', 10) || 1));
    const minStops = Math.max(1, parseInt(els.minStopsInput.value || '1', 10) || 1);
    const maxStops = Math.max(minStops, parseInt(els.maxStopsInput.value || '999999', 10) || minStops);
    const totalAccounts = state.accounts.length;
    const fixedCount = state.accounts.filter(a => a.protected || isAccountLocked(a)).length;
    const movableCount = totalAccounts - fixedCount;

    if (targetCount > totalAccounts) { showToast(`Target rep count of ${targetCount} exceeds ${totalAccounts} total accounts.`); return; }
    if (targetCount * minStops > totalAccounts) { showToast(`Minimum stops too high. ${targetCount} reps × ${minStops} minimum exceeds ${totalAccounts} total accounts.`); return; }
    if (Math.ceil(totalAccounts / targetCount) > maxStops) { showToast(`Maximum stops too low. ${targetCount} reps cannot cover ${totalAccounts} accounts with a max of ${maxStops} per rep.`); return; }
    if (movableCount === 0) { showToast('All remaining accounts are locked or protected. Nothing can be optimized.'); return; }

    const continuityWeight = Number(els.disruptionSlider.value) / 100;
    const compactMode = getOptimizerMode() === 'compact';
    const optimizerMix = getOptimizerMix();
    const beforeSummary = buildOptimizationSummary();

    const fixedAccounts = state.accounts.filter(a => a.protected || isAccountLocked(a));
    const movableAccounts = state.accounts.filter(a => !a.protected && !isAccountLocked(a));
    const currentReps = getAllAssignedReps().filter(rep => !isRepLocked(rep));
    const targetRepNames = buildTargetRepNames(targetCount, currentReps);
    registerRepNames(targetRepNames);
    const adjacency = state.neighborMap;

    if (!targetRepNames.length) { showToast('No unlocked reps are available for optimization.'); return; }

    targetRepNames.forEach(rep => ensureRepColor(rep));

    const assignments = new Map();
    fixedAccounts.forEach(a => assignments.set(a._id, a.assignedRep));

    const assignmentCtx = createAssignmentContext(targetRepNames, assignments);
    const existingRepNames = new Set(currentReps);

    // Track which reps are new (didn't exist before this optimization)
    const newRepSet = new Set(targetRepNames.filter(r => !existingRepNames.has(r)));

    const seedPlan = compactMode
      ? buildNewRepSeedPlan(movableAccounts, targetRepNames, existingRepNames, minStops, maxStops)
      : createEmptySeedPlan();

    state.optimizerSeedIds = new Set(seedPlan.seededAccountIds || []);

    if (seedPlan.seedAssignments && seedPlan.seedAssignments.size) {
      seedPlan.seedAssignments.forEach((rep, accountId) => {
        const account = state.accountById.get(accountId);
        if (!account) return;
        assignments.set(accountId, rep);
        assignmentCtx.addToRep(rep, account);
      });
    }

    const centroids = initializeCentroidsFast(targetRepNames, assignmentCtx);
    const movableForAssignment = movableAccounts.filter(a => !state.optimizerSeedIds.has(a._id));
    const movableForCleanup = movableForAssignment;

    const orderedMovable = [...movableForAssignment].sort((a, b) => {
      if (a.rank !== b.rank) return rankSortValue(a.rank) - rankSortValue(b.rank);
      if (a.overallSales !== b.overallSales) return b.overallSales - a.overallSales;
      return a.customerName.localeCompare(b.customerName);
    });

    const targetStopsPerRep = Math.max(minStops, Math.min(maxStops, totalAccounts / Math.max(1, targetRepNames.length)));
    const totalRevenuePool = state.accounts.reduce((sum, a) => sum + (a.overallSales || 0), 0);
    const targetRevenuePerRep = totalRevenuePool / Math.max(1, targetRepNames.length);

    let iterationsExecuted = 0;

    for (let iter = 0; iter < 20; iter += 1) {
      iterationsExecuted += 1;
      let changedThisPass = false;
      let repLoadOrder = null;
      let repLoadDirty = true;

      assignmentCtx.clearMovableAssignments(movableForAssignment, assignments);
      refreshCentroidsFromContext(centroids, targetRepNames, assignmentCtx);

      // Initialize repStats from context AFTER clear so fixed accounts count
      const repStats = new Map();
      for (const rep of targetRepNames) {
        repStats.set(rep, {
          rep,
          stops: assignmentCtx.count(rep),
          revenue: assignmentCtx.revenue(rep)
        });
      }

      // New reps get expanded freedoms in early iterations
      const isEarlyIter = iter < 4;

      for (const account of orderedMovable) {
        let bestRep = null;
        let bestScore = Infinity;
        const currentRep = assignments.get(account._id) || account.assignedRep;

        for (const rep of targetRepNames) {
          const isNewRep = newRepSet.has(rep);
          const neighborSupport = countNeighborRepSupport(account, rep, assignments, adjacency);

          // Compact hard constraint — exempt new reps in early iters so they
          // can grow beyond their seed cluster into adjacent territory
          if (compactMode && neighborSupport === 0 && assignmentCtx.count(rep) > 0) {
            if (!(isNewRep && isEarlyIter)) continue;
          }

          const centroid = centroids.get(rep) || averageCentroidForRep(rep, assignmentCtx);
          const compactnessScore = centroid
            ? squaredDistance(account.latitude, account.longitude, centroid.lat, centroid.lng) * ((1 - continuityWeight) * 1.4)
            : 0;

          // New reps have no history — no continuity penalty
          const continuityPenalty = isNewRep ? 0 : (account.currentRep === rep ? 0 : continuityWeight);
          const existingPenalty = account.assignedRep === rep ? -0.15 : 0;

          const stat = repStats.get(rep);
          const nextStops = (stat.stops || 0) + 1;
          const nextRevenue = (stat.revenue || 0) + (account.overallSales || 0);
          const stopDeviation = Math.abs(nextStops - targetStopsPerRep) / Math.max(1, targetStopsPerRep);
          const revenueDeviation = Math.min(Math.abs(nextRevenue - targetRevenuePerRep) / Math.max(1, targetRevenuePerRep || 1), 2.0);
          const balancePenalty = (stopDeviation * optimizerMix.stopsPriority * 1.55) + (revenueDeviation * optimizerMix.revenuePriority * 1.15);

          const underMinBoost = stat.stops < minStops ? -2.2 : 0;
          const overMaxPenalty = nextStops > maxStops ? ((nextStops - maxStops) * 4.5) : 0;
          const localPenaltyBase = localDominancePenalty(account, rep, assignments, adjacency);
          const supportBonus = compactMode
            ? (neighborSupport >= 4 ? -2.4 : neighborSupport >= 3 ? -1.55 : neighborSupport >= 2 ? -0.9 : neighborSupport === 1 ? -0.18 : 1.45)
            : (neighborSupport >= 2 ? -0.25 : 0);
          const fragmentPenalty = fragmentationPenalty(account, rep, assignments, adjacency) * (compactMode ? 2.8 : 0.7);
          const localPenalty = compactMode ? (localPenaltyBase * 2.15) : localPenaltyBase;
          const unsupportedPenalty = (!isNewRep || !isEarlyIter) && compactMode && neighborSupport === 0 && currentRep && currentRep !== rep ? 2.2 : 0;

          const score = compactnessScore + continuityPenalty + existingPenalty + balancePenalty + localPenalty + fragmentPenalty + unsupportedPenalty + underMinBoost + overMaxPenalty + supportBonus;

          if (score < bestScore) { bestScore = score; bestRep = rep; }
        }

        if (!bestRep) {
          if (repLoadDirty || !repLoadOrder) { repLoadOrder = buildRepLoadOrder(targetRepNames, assignmentCtx); repLoadDirty = false; }
          bestRep = repLoadOrder[0] || currentRep || targetRepNames[0];
        }

        const prevRep = assignments.get(account._id);
        assignments.set(account._id, bestRep);
        assignmentCtx.addToRep(bestRep, account);
        const stat = repStats.get(bestRep);
        stat.stops += 1;
        stat.revenue += account.overallSales || 0;

        if (bestRep !== prevRep) {
          changedThisPass = true;
          repLoadDirty = true;
        }
      }

      refreshCentroidsFromContext(centroids, targetRepNames, assignmentCtx);
      if (compactMode && seedPlan.seededReps && seedPlan.seededReps.size) {
        enforceSeedAnchors(assignments, assignmentCtx, seedPlan);
        refreshCentroidsFromContext(centroids, targetRepNames, assignmentCtx);
      }
      if (!changedThisPass) break;
    }

    enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    rebalanceStopTargetsStrict(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    runBorderCleanupFast(assignments, targetRepNames, continuityWeight, minStops, adjacency, assignmentCtx, movableForCleanup);
    performContiguityRefinement(assignments, targetRepNames, minStops, maxStops, adjacency, assignmentCtx, movableForCleanup);

    if (tryBorderSwaps(assignments, targetRepNames, minStops, maxStops, adjacency, assignmentCtx, movableForCleanup)) {
      tryBorderSwaps(assignments, targetRepNames, minStops, maxStops, adjacency, assignmentCtx, movableForCleanup);
    }

    const compactCleanupPasses = compactMode ? 5 : 1;
    for (let compactPass = 0; compactPass < compactCleanupPasses; compactPass += 1) {
      runBorderCleanupFast(assignments, targetRepNames, continuityWeight, minStops, adjacency, assignmentCtx, movableForCleanup);
      if (compactMode) absorbSmallIslandsFast(assignments, targetRepNames, minStops, maxStops, adjacency, assignmentCtx, movableForCleanup);
      performContiguityRefinement(assignments, targetRepNames, minStops, maxStops, adjacency, assignmentCtx, movableForCleanup);
      if (compactMode) {
        tryBorderSwaps(assignments, targetRepNames, minStops, maxStops, adjacency, assignmentCtx, movableForCleanup);
        absorbSmallIslandsFast(assignments, targetRepNames, minStops, maxStops, adjacency, assignmentCtx, movableForCleanup);
      }
      enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
      enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
      rebalanceStopTargetsStrict(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    }

    const finalViolations = targetRepNames.filter(rep => { const c = assignmentCtx.count(rep); return c < minStops || c > maxStops; });
    if (finalViolations.length) showToast('Optimizer could not fully satisfy the stop limits. Try adjusting rep count or stop limits.');

    const changes = [];
    const repsBefore = getAllAssignedReps();
    for (const account of state.accounts) {
      const nextRep = assignments.get(account._id) || account.assignedRep;
      if (nextRep !== account.assignedRep) changes.push({ id: account._id, from: account.assignedRep, to: nextRep });
    }

    if (!changes.length) {
      showToast(iterationsExecuted < 20 ? 'Optimizer converged early — no better assignment found.' : 'Optimizer did not find a better assignment under the current rules.');
      return;
    }

    applyChanges(changes, `Optimized routes to ${targetRepNames.length} reps with minimum ${minStops} stops`, repsBefore);
    state.optimizerSeedIds = new Set();
    state.optimizationSummary = buildOptimizationSummary(beforeSummary, { weightLabel: getOptimizerWeightLabel(), disruptionLabel: getDisruptionPreset().short });
    renderOptimizationFeedback();
    updateLastAction('');

  } catch (err) {
    state.optimizerSeedIds = new Set();
    console.error('Optimize Routes failed:', err);
    showToast('Optimize Routes hit an error. Check the browser console for details.');
  }
}

function buildOptimizationSummary(previousSummary = null, meta = {}) {
  const rows = summarizeByRep();
  const movedCount = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep).length;
  const protectedHeld = state.accounts.filter(a => a.protected && a.assignedRep === a.originalAssignedRep).length;
  const stops = rows.map(r => r.stops);
  const revenue = rows.map(r => r.revenue);
  const stopRange = Math.max(0, Math.max(...stops) - Math.min(...stops));
  const revenueRange = Math.max(0, Math.max(...revenue) - Math.min(...revenue));
  const summary = {
    repCount: rows.length, movedCount, protectedHeld,
    minStops: Math.min(...stops), maxStops: Math.max(...stops),
    minRevenue: Math.min(...revenue), maxRevenue: Math.max(...revenue),
    avgStops: rows.reduce((s, r) => s + r.stops, 0) / Math.max(1, rows.length),
    stopRange, revenueRange, stopRangeDeltaPct: 0, revenueRangeDeltaPct: 0,
    weightLabel: meta.weightLabel || 'Balanced',
    disruptionLabel: meta.disruptionLabel || getDisruptionPreset().short
  };
  if (previousSummary) {
    summary.stopRangeDeltaPct = previousSummary.stopRange > 0 ? ((previousSummary.stopRange - summary.stopRange) / previousSummary.stopRange) * 100 : 0;
    summary.revenueRangeDeltaPct = previousSummary.revenueRange > 0 ? ((previousSummary.revenueRange - summary.revenueRange) / previousSummary.revenueRange) * 100 : 0;
  }
  return summary;
}

function createEmptySeedPlan() {
  return { seedAssignments: new Map(), seededAccountIds: new Set(), seededReps: new Set(), seedTargetByRep: new Map() };
}

function isOptimizerSeedAccount(accountId) {
  return !!(state.optimizerSeedIds && state.optimizerSeedIds.has(accountId));
}

// ── FLOOD-FILL SEED (new) ─────────────────────────────────────
// Grows a contiguous cluster outward through the neighbor graph
// from a starting account, staying within the same donor rep.
// Guarantees the seed cluster is geographically connected.
function floodFillSeed(startAccount, movableIdSet, reservedIds, donorRep, maxSize) {
  const visited = new Set([startAccount._id]);
  const result = [startAccount];
  const queue = [startAccount];

  while (queue.length && result.length < maxSize) {
    const current = queue.shift();
    const neighbors = state.neighborMap.get(current._id);
    if (!neighbors) continue;
    for (const neighborId of neighbors) {
      if (visited.has(neighborId)) continue;
      if (reservedIds.has(neighborId)) continue;
      if (!movableIdSet.has(neighborId)) continue;
      const neighbor = state.accountById.get(neighborId);
      if (!neighbor) continue;
      const neighborRep = neighbor.assignedRep || neighbor.currentRep || 'Unassigned';
      if (neighborRep !== donorRep) continue;
      visited.add(neighborId);
      result.push(neighbor);
      queue.push(neighbor);
      if (result.length >= maxSize) break;
    }
  }
  return result;
}

// ── SEED PLAN ─────────────────────────────────────────────────
// Replaces old proximity-based seeding with flood-fill so new
// rep seed clusters are guaranteed to be contiguous.
function buildNewRepSeedPlan(movableAccounts, targetRepNames, existingRepNames, minStops, maxStops) {
  const plan = createEmptySeedPlan();
  const newReps = targetRepNames.filter(rep => !existingRepNames.has(rep));
  if (!newReps.length || !movableAccounts.length) return plan;

  const movableIdSet = new Set(movableAccounts.map(a => a._id));

  const movableByRep = new Map();
  for (const a of movableAccounts) {
    const rep = a.assignedRep || a.currentRep || 'Unassigned';
    if (!movableByRep.has(rep)) movableByRep.set(rep, []);
    movableByRep.get(rep).push(a);
  }

  const repCentroids = new Map();
  for (const [rep, accounts] of movableByRep) {
    const lat = accounts.reduce((s, a) => s + a.latitude, 0) / accounts.length;
    const lng = accounts.reduce((s, a) => s + a.longitude, 0) / accounts.length;
    repCentroids.set(rep, { lat, lng });
  }

  const reservedIds = new Set();
  const placedCentroids = [];
  const seedTarget = Math.max(minStops, Math.round(minStops * 1.15));

  for (const newRep of newReps) {
    const donors = [...movableByRep.entries()]
      .filter(([, accounts]) => accounts.filter(a => !reservedIds.has(a._id)).length >= Math.max(seedTarget, minStops + 10))
      .sort((a, b) => b[1].filter(x => !reservedIds.has(x._id)).length - a[1].filter(x => !reservedIds.has(x._id)).length);

    let bestSeed = null;
    let bestSeedScore = -Infinity;

    for (const [donorRep, donorAccounts] of donors.slice(0, 6)) {
      const available = donorAccounts.filter(a => !reservedIds.has(a._id));
      if (available.length < seedTarget) continue;
      const centroid = repCentroids.get(donorRep);

      for (const candidate of available) {
        const edgeDist = centroid ? squaredDistance(candidate.latitude, candidate.longitude, centroid.lat, centroid.lng) : 0;
        let minSeedDist = Infinity;
        for (const placed of placedCentroids) minSeedDist = Math.min(minSeedDist, squaredDistance(candidate.latitude, candidate.longitude, placed.lat, placed.lng));
        if (!isFinite(minSeedDist)) minSeedDist = 1;

        // Test if flood-fill from this candidate can reach enough accounts
        const reachable = floodFillSeed(candidate, movableIdSet, reservedIds, donorRep, seedTarget);
        if (reachable.length < Math.round(seedTarget * 0.7)) continue;

        const score = (edgeDist * 180) + (minSeedDist * 220) - (reachable.length < seedTarget ? 500 : 0);
        if (score > bestSeedScore) {
          bestSeedScore = score;
          bestSeed = { donorRep, reachable };
        }
      }
    }

    if (!bestSeed) continue;

    const cluster = bestSeed.reachable.slice(0, seedTarget);
    for (const a of cluster) {
      reservedIds.add(a._id);
      plan.seedAssignments.set(a._id, newRep);
      plan.seededAccountIds.add(a._id);
    }
    plan.seededReps.add(newRep);
    plan.seedTargetByRep.set(newRep, cluster.length);

    const clat = cluster.reduce((s, a) => s + a.latitude, 0) / cluster.length;
    const clng = cluster.reduce((s, a) => s + a.longitude, 0) / cluster.length;
    placedCentroids.push({ lat: clat, lng: clng });
  }

  return plan;
}

function enforceSeedAnchors(assignments, ctx, seedPlan) {
  if (!seedPlan || !seedPlan.seededReps || !seedPlan.seededReps.size) return;
  for (const rep of seedPlan.seededReps) {
    const target = seedPlan.seedTargetByRep.get(rep) || 0;
    if (!target) continue;
    const centroid = averageCentroidForRep(rep, ctx);
    if (!centroid) continue;
    const seedMembers = [...(ctx.members(rep) || [])]
      .filter(id => isOptimizerSeedAccount(id))
      .map(id => state.accountById.get(id))
      .filter(Boolean)
      .sort((a, b) => squaredDistance(a.latitude, a.longitude, centroid.lat, centroid.lng) - squaredDistance(b.latitude, b.longitude, centroid.lat, centroid.lng));
    for (let i = target; i < seedMembers.length; i++) {
      const account = seedMembers[i];
      if ((assignments.get(account._id) || rep) !== rep) continue;
      state.optimizerSeedIds.delete(account._id);
    }
  }
}

function createAssignmentContext(targetRepNames, assignments) {
  const ctx = { reps: new Map() };
  targetRepNames.forEach(rep => ctx.reps.set(rep, { count: 0, revenue: 0, latSum: 0, lngSum: 0, members: new Set() }));
  for (const account of state.accounts) {
    const rep = assignments.get(account._id);
    if (!rep) continue;
    if (!ctx.reps.has(rep)) ctx.reps.set(rep, { count: 0, revenue: 0, latSum: 0, lngSum: 0, members: new Set() });
    addAccountToContext(ctx, rep, account);
  }
  ctx.count = rep => ctx.reps.get(rep)?.count || 0;
  ctx.revenue = rep => ctx.reps.get(rep)?.revenue || 0;
  ctx.members = rep => ctx.reps.get(rep)?.members || new Set();
  ctx.addToRep = (rep, account) => addAccountToContext(ctx, rep, account);
  ctx.removeFromRep = (rep, account) => removeAccountFromContext(ctx, rep, account);
  ctx.clearMovableAssignments = (movable, assignmentMap) => {
    movable.forEach(account => {
      const currentRep = assignmentMap.get(account._id);
      if (!currentRep) return;
      ctx.removeFromRep(currentRep, account);
      assignmentMap.delete(account._id);
    });
  };
  return ctx;
}

function addAccountToContext(ctx, rep, account) {
  if (!ctx.reps.has(rep)) ctx.reps.set(rep, { count: 0, revenue: 0, latSum: 0, lngSum: 0, members: new Set() });
  const entry = ctx.reps.get(rep);
  if (entry.members.has(account._id)) return;
  entry.members.add(account._id);
  entry.count += 1;
  entry.revenue += account.overallSales || 0;
  entry.latSum += account.latitude;
  entry.lngSum += account.longitude;
}

function removeAccountFromContext(ctx, rep, account) {
  const entry = ctx.reps.get(rep);
  if (!entry || !entry.members.has(account._id)) return;
  entry.members.delete(account._id);
  entry.count -= 1;
  entry.revenue -= account.overallSales || 0;
  entry.latSum -= account.latitude;
  entry.lngSum -= account.longitude;
}

function initializeCentroidsFast(targetRepNames, ctx) {
  const centroids = new Map();
  targetRepNames.forEach(rep => centroids.set(rep, averageCentroidForRep(rep, ctx)));
  return centroids;
}

function refreshCentroidsFromContext(centroids, targetRepNames, ctx) {
  targetRepNames.forEach(rep => centroids.set(rep, averageCentroidForRep(rep, ctx)));
}

function averageCentroidForRep(rep, ctx) {
  const entry = ctx.reps.get(rep);
  if (!entry || !entry.count) return null;
  return { lat: entry.latSum / entry.count, lng: entry.lngSum / entry.count };
}

function buildTargetRepNames(targetCount, currentReps) {
  const reps = [...currentReps];
  while (reps.length < targetCount) reps.push(`Rep ${reps.length + 1}`);
  return reps.slice(0, targetCount);
}

function buildFullRepStats(targetRepNames) {
  const map = new Map();
  targetRepNames.forEach(rep => map.set(rep, { rep, stops: 0, revenue: 0 }));
  return map;
}

function localDominancePenalty(account, rep, assignments, adjacency) {
  const neighbors = adjacency.get(account._id);
  if (!neighbors || !neighbors.size) return 0;
  let same = 0, total = 0;
  neighbors.forEach(id => {
    const neighborRep = assignments.get(id) || state.accountById.get(id)?.assignedRep;
    if (!neighborRep) return;
    if (neighborRep === rep) same += 1;
    total += 1;
  });
  if (!total) return 0;
  const ratio = same / total;
  const compactMode = getOptimizerMode() === 'compact';
  if (ratio >= 0.9) return compactMode ? -0.48 : -0.12;
  if (ratio >= 0.75) return compactMode ? -0.18 : -0.03;
  if (ratio >= 0.6) return compactMode ? 0.06 : 0;
  return (0.6 - ratio) * (compactMode ? 4.2 : 2.1);
}

function countNeighborRepSupport(account, rep, assignments, adjacency) {
  const neighbors = adjacency.get(account._id);
  if (!neighbors || !neighbors.size) return 0;
  let support = 0;
  neighbors.forEach(id => { if ((assignments.get(id) || state.accountById.get(id)?.assignedRep) === rep) support += 1; });
  return support;
}

function fragmentationPenalty(account, rep, assignments, adjacency) {
  const neighbors = adjacency.get(account._id);
  if (!neighbors || !neighbors.size) return 0;
  let same = 0, strongestOtherCount = 0;
  const otherCounts = new Map();
  neighbors.forEach(id => {
    const neighborRep = assignments.get(id) || state.accountById.get(id)?.assignedRep;
    if (!neighborRep) return;
    if (neighborRep === rep) { same += 1; return; }
    const next = (otherCounts.get(neighborRep) || 0) + 1;
    otherCounts.set(neighborRep, next);
    if (next > strongestOtherCount) strongestOtherCount = next;
  });
  const compactMode = getOptimizerMode() === 'compact';
  if (same >= 4) return 0;
  if (same >= 3 && strongestOtherCount <= 1) return compactMode ? 0.04 : 0;
  if (strongestOtherCount >= 4 && same <= 1) return compactMode ? 2.4 : 1.2;
  if (strongestOtherCount >= 3 && same <= 1) return compactMode ? 1.8 : 1.2;
  if (strongestOtherCount > same) return compactMode ? 1.1 : 0.55;
  if (same === 0 && strongestOtherCount >= 2) return compactMode ? 1.45 : 0.8;
  return compactMode && same <= 1 ? 0.35 : 0;
}

function tryBorderSwaps(assignments, targetRepNames, minStops, maxStops, adjacency, ctx, movableAccounts = null) {
  const compactMode = getOptimizerMode() === 'compact';
  const candidates = Array.isArray(movableAccounts) && movableAccounts.length ? movableAccounts : state.accounts.filter(a => !a.protected && !isAccountLocked(a));
  let improved = false;
  for (const account of candidates) {
    const currentRep = assignments.get(account._id) || account.assignedRep;
    if (!currentRep) continue;
    const neighbors = adjacency.get(account._id);
    if (!neighbors || neighbors.size < 2) continue;
    if (fragmentationPenalty(account, currentRep, assignments, adjacency) <= 0) continue;
    let bestSwap = null, bestGain = 0;
    neighbors.forEach(id => {
      const neighbor = state.accountById.get(id);
      if (!neighbor || neighbor.protected || isAccountLocked(neighbor)) return;
      const neighborRep = assignments.get(id) || neighbor.assignedRep;
      if (!neighborRep || neighborRep === currentRep) return;
      if (ctx.count(currentRep) <= minStops || ctx.count(neighborRep) <= minStops) return;
      if (ctx.count(currentRep) > maxStops || ctx.count(neighborRep) > maxStops) return;
      const aCent = averageCentroidForRep(currentRep, ctx);
      const nCent = averageCentroidForRep(neighborRep, ctx);
      const aOldDist = aCent ? squaredDistance(account.latitude, account.longitude, aCent.lat, aCent.lng) : 0;
      const nOldDist = nCent ? squaredDistance(neighbor.latitude, neighbor.longitude, nCent.lat, nCent.lng) : 0;
      const aNewDist = nCent ? squaredDistance(account.latitude, account.longitude, nCent.lat, nCent.lng) : aOldDist;
      const nNewDist = aCent ? squaredDistance(neighbor.latitude, neighbor.longitude, aCent.lat, aCent.lng) : nOldDist;
      const fragBefore = fragmentationPenalty(account, currentRep, assignments, adjacency) + fragmentationPenalty(neighbor, neighborRep, assignments, adjacency);
      const suppBefore = countNeighborRepSupport(account, currentRep, assignments, adjacency) + countNeighborRepSupport(neighbor, neighborRep, assignments, adjacency);
      assignments.set(account._id, neighborRep); assignments.set(neighbor._id, currentRep);
      const fragAfter = fragmentationPenalty(account, neighborRep, assignments, adjacency) + fragmentationPenalty(neighbor, currentRep, assignments, adjacency);
      const suppAfter = countNeighborRepSupport(account, neighborRep, assignments, adjacency) + countNeighborRepSupport(neighbor, currentRep, assignments, adjacency);
      assignments.set(account._id, currentRep); assignments.set(neighbor._id, neighborRep);
      const distGain = (aOldDist + nOldDist) - (aNewDist + nNewDist);
      const suppGain = suppAfter - suppBefore;
      const fragGain = fragBefore - fragAfter;
      const totalGain = (fragGain * (compactMode ? 3.2 : 2.4)) + (suppGain * (compactMode ? 1.8 : 1.3)) + (distGain * 0.85);
      if (totalGain > bestGain && (fragGain > 0 || (suppGain >= (compactMode ? 1 : 2) && distGain > 0))) { bestGain = totalGain; bestSwap = { neighbor, neighborRep }; }
    });
    if (bestSwap) {
      ctx.removeFromRep(currentRep, account); ctx.removeFromRep(bestSwap.neighborRep, bestSwap.neighbor);
      ctx.addToRep(bestSwap.neighborRep, account); ctx.addToRep(currentRep, bestSwap.neighbor);
      assignments.set(account._id, bestSwap.neighborRep); assignments.set(bestSwap.neighbor._id, currentRep);
      improved = true;
    }
  }
  return improved;
}

function performContiguityRefinement(assignments, targetRepNames, minStops, maxStops, adjacency, ctx, movableAccounts = null) {
  const compactMode = getOptimizerMode() === 'compact';
  const candidates = Array.isArray(movableAccounts) && movableAccounts.length ? movableAccounts : state.accounts.filter(a => !a.protected && !isAccountLocked(a));
  let changed = true, passes = 0;
  while (changed && passes < (compactMode ? 18 : 10)) {
    changed = false; passes += 1;
    const ordered = [...candidates].sort((a, b) =>
      countNeighborRepSupport(a, assignments.get(a._id) || a.assignedRep, assignments, adjacency) -
      countNeighborRepSupport(b, assignments.get(b._id) || b.assignedRep, assignments, adjacency)
    );
    for (const account of ordered) {
      const currentRep = assignments.get(account._id) || account.assignedRep;
      if (!currentRep || ctx.count(currentRep) <= minStops) continue;
      const neighbors = adjacency.get(account._id);
      if (!neighbors || neighbors.size < 2) continue;
      const currentSupport = countNeighborRepSupport(account, currentRep, assignments, adjacency);
      const oldCentroid = averageCentroidForRep(currentRep, ctx);
      const oldDist = oldCentroid ? squaredDistance(account.latitude, account.longitude, oldCentroid.lat, oldCentroid.lng) : 0;
      const targetCounts = new Map();
      neighbors.forEach(id => {
        const neighborRep = assignments.get(id) || state.accountById.get(id)?.assignedRep;
        if (!neighborRep || neighborRep === currentRep) return;
        targetCounts.set(neighborRep, (targetCounts.get(neighborRep) || 0) + 1);
      });
      let bestRep = null, bestDelta = 0;
      for (const [targetRep, targetSupport] of targetCounts.entries()) {
        if (isRepLocked(targetRep) || ctx.count(targetRep) >= maxStops) continue;
        if (targetSupport < (compactMode ? 3 : 2)) continue;
        const newCentroid = averageCentroidForRep(targetRep, ctx);
        const newDist = newCentroid ? squaredDistance(account.latitude, account.longitude, newCentroid.lat, newCentroid.lng) : oldDist;
        const supportDelta = targetSupport - currentSupport;
        const distanceDelta = newDist - oldDist;
        const fragmentationDelta = fragmentationPenalty(account, currentRep, assignments, adjacency) - fragmentationPenalty(account, targetRep, assignments, adjacency);
        const moveScore = (supportDelta * (compactMode ? 2.55 : 1.55)) + (fragmentationDelta * (compactMode ? 4.3 : 2.2)) - (distanceDelta * (compactMode ? 1.15 : 0.95)) - (ctx.count(targetRep) < minStops ? (compactMode ? 0.5 : 0.35) : 0);
        if (moveScore > bestDelta && (supportDelta >= (compactMode ? 3 : 2) || (supportDelta >= 2 && fragmentationDelta > 0 && compactMode) || (supportDelta >= 1 && distanceDelta <= 0 && !compactMode))) {
          bestDelta = moveScore; bestRep = targetRep;
        }
      }
      if (bestRep) { ctx.removeFromRep(currentRep, account); ctx.addToRep(bestRep, account); assignments.set(account._id, bestRep); changed = true; }
    }
  }
}

function enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx) {
  let guard = 0;
  while (guard < 2000) {
    guard += 1;
    const under = targetRepNames.find(rep => ctx.count(rep) < minStops);
    if (!under) break;
    const donor = targetRepNames.filter(rep => ctx.count(rep) > minStops).sort((a, b) => ctx.count(b) - ctx.count(a))[0];
    if (!donor) break;
    const underCentroid = averageCentroidForRep(under, ctx);
    const candidate = [...ctx.members(donor)].map(id => state.accountById.get(id)).filter(Boolean)
      .sort((a, b) => {
        const aSeed = isOptimizerSeedAccount(a._id) ? 1 : 0;
        const bSeed = isOptimizerSeedAccount(b._id) ? 1 : 0;
        if (aSeed !== bSeed) return aSeed - bSeed;
        const ad = underCentroid ? squaredDistance(a.latitude, a.longitude, underCentroid.lat, underCentroid.lng) : 0;
        const bd = underCentroid ? squaredDistance(b.latitude, b.longitude, underCentroid.lat, underCentroid.lng) : 0;
        return ad - bd;
      })[0];
    if (!candidate) break;
    ctx.removeFromRep(donor, candidate); ctx.addToRep(under, candidate); assignments.set(candidate._id, under);
  }
}

function enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx) {
  let guard = 0;
  while (guard < 2000) {
    guard += 1;
    const over = targetRepNames.find(rep => ctx.count(rep) > maxStops);
    if (!over) break;
    const receiver = targetRepNames.filter(rep => ctx.count(rep) < maxStops).sort((a, b) => ctx.count(a) - ctx.count(b))[0];
    if (!receiver) break;
    const receiverCentroid = averageCentroidForRep(receiver, ctx);
    const candidate = [...ctx.members(over)].map(id => state.accountById.get(id)).filter(Boolean)
      .sort((a, b) => {
        const aSeed = isOptimizerSeedAccount(a._id) ? 1 : 0;
        const bSeed = isOptimizerSeedAccount(b._id) ? 1 : 0;
        if (aSeed !== bSeed) return aSeed - bSeed;
        const ad = receiverCentroid ? squaredDistance(a.latitude, a.longitude, receiverCentroid.lat, receiverCentroid.lng) : 0;
        const bd = receiverCentroid ? squaredDistance(b.latitude, b.longitude, receiverCentroid.lat, receiverCentroid.lng) : 0;
        return ad - bd;
      })[0];
    if (!candidate) break;
    ctx.removeFromRep(over, candidate); ctx.addToRep(receiver, candidate); assignments.set(candidate._id, receiver);
  }
}

function rebalanceStopTargetsStrict(assignments, targetRepNames, minStops, maxStops, ctx) {
  enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx);
  enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx);
}

function runBorderPass(accounts, assignments, minStops, adjacency, ctx) {
  const compactMode = getOptimizerMode() === 'compact';
  for (const account of accounts) {
    const currentRep = assignments.get(account._id) || account.assignedRep;
    const neighbors = adjacency.get(account._id);
    if (!neighbors || !neighbors.size) continue;
    const counts = new Map();
    neighbors.forEach(id => {
      const rep = assignments.get(id) || state.accountById.get(id)?.assignedRep;
      if (!rep) return;
      counts.set(rep, (counts.get(rep) || 0) + 1);
    });
    const ordered = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    const bestNeighborRep = ordered[0]?.[0];
    const bestNeighborCount = ordered[0]?.[1] || 0;
    const currentSupport = counts.get(currentRep) || 0;
    const runnerUpCount = ordered[1]?.[1] || 0;
    if (!bestNeighborRep || bestNeighborRep === currentRep) continue;
    if (ctx.count(currentRep) <= minStops) continue;
    if (isRepLocked(bestNeighborRep)) continue;
    if (compactMode) {
      if (bestNeighborCount < 3) continue;
      if ((bestNeighborCount - currentSupport) < 2) continue;
      if ((bestNeighborCount - runnerUpCount) < 1) continue;
    } else if (bestNeighborCount <= currentSupport) continue;
    ctx.removeFromRep(currentRep, account); ctx.addToRep(bestNeighborRep, account); assignments.set(account._id, bestNeighborRep);
  }
}

function runBorderCleanupFast(assignments, targetRepNames, continuityWeight, minStops, adjacency, ctx, movableAccounts = null) {
  const compactMode = getOptimizerMode() === 'compact';
  const accounts = Array.isArray(movableAccounts) && movableAccounts.length ? movableAccounts : state.accounts.filter(a => !a.protected && !isAccountLocked(a));
  runBorderPass(accounts, assignments, minStops, adjacency, ctx);
  runBorderPass([...accounts].reverse(), assignments, minStops, adjacency, ctx);
  if (compactMode) {
    runBorderPass(accounts, assignments, minStops, adjacency, ctx);
    runBorderPass([...accounts].sort((a, b) =>
      countNeighborRepSupport(a, assignments.get(a._id) || a.assignedRep, assignments, adjacency) -
      countNeighborRepSupport(b, assignments.get(b._id) || b.assignedRep, assignments, adjacency)
    ), assignments, minStops, adjacency, ctx);
  }
}

function absorbSmallIslandsFast(assignments, targetRepNames, minStops, maxStops, adjacency, ctx, movableAccounts = null) {
  const accounts = Array.isArray(movableAccounts) && movableAccounts.length ? movableAccounts : state.accounts.filter(a => !a.protected && !isAccountLocked(a));
  const accountIds = new Set(accounts.map(a => a._id));
  const visited = new Set();
  const accountsByRep = new Map();
  for (const a of accounts) {
    const rep = assignments.get(a._id) || a.assignedRep;
    if (!rep) continue;
    if (!accountsByRep.has(rep)) accountsByRep.set(rep, []);
    accountsByRep.get(rep).push(a);
  }
  for (const rep of targetRepNames) {
    for (const seed of (accountsByRep.get(rep) || [])) {
      if (visited.has(seed._id)) continue;
      const component = [], componentSet = new Set(), stack = [seed._id];
      visited.add(seed._id);
      while (stack.length) {
        const id = stack.pop();
        component.push(id); componentSet.add(id);
        const neighbors = adjacency.get(id);
        if (!neighbors) continue;
        neighbors.forEach(nId => {
          if (visited.has(nId) || !accountIds.has(nId)) return;
          if ((assignments.get(nId) || state.accountById.get(nId)?.assignedRep) !== rep) return;
          visited.add(nId); stack.push(nId);
        });
      }
      if (component.length > 4 || ctx.count(rep) - component.length < minStops) continue;
      const borderCounts = new Map();
      for (const id of component) {
        adjacency.get(id)?.forEach(nId => {
          if (componentSet.has(nId)) return;
          const nRep = assignments.get(nId) || state.accountById.get(nId)?.assignedRep;
          if (!nRep || nRep === rep || isRepLocked(nRep)) return;
          borderCounts.set(nRep, (borderCounts.get(nRep) || 0) + 1);
        });
      }
      const ordered = [...borderCounts.entries()].sort((a, b) => b[1] - a[1]);
      const targetRep = ordered[0]?.[0];
      const targetTouches = ordered[0]?.[1] || 0;
      if (!targetRep || targetTouches < 2 || ctx.count(targetRep) + component.length > maxStops) continue;
      for (const id of component) {
        const a = state.accountById.get(id);
        if (!a) continue;
        ctx.removeFromRep(rep, a); ctx.addToRep(targetRep, a); assignments.set(id, targetRep);
      }
    }
  }
}

// ─── EXPORT ──────────────────────────────────────────────────

async function exportWorkbook() {
  if (!state.accounts.length) { showToast('Nothing to export.'); return; }
  const workbook = new ExcelJS.Workbook();
  const mainSheet = workbook.addWorksheet(state.currentSheetName || 'Sheet1');
  const exportRows = state.accounts.map(account => {
    const row = { ...(account.sourceRow || {}) };
    row[state.currentHeaderMap.assignedRep || 'New Rep'] = account.assignedRep;
    return row;
  });
  if (exportRows.length) {
    const keys = Object.keys(exportRows[0]);
    mainSheet.columns = keys.map(key => ({ header: key, key, width: guessColumnWidth(key, exportRows) }));
    exportRows.forEach(row => mainSheet.addRow(row));
    styleHeaderRow(mainSheet.getRow(1));
    styleDataRows(mainSheet, 2, mainSheet.rowCount);
  }
  const movedSheet = workbook.addWorksheet('Moved Accounts');
  const movedRows = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep).map(a => ({
    Customer_ID: a.customerId, Customer_Name: a.customerName,
    Original_Assigned_Rep: a.originalAssignedRep, Assigned_Rep: a.assignedRep,
    Current_Rep: a.currentRep, Revenue: round2(a.overallSales), Rank: a.rank,
    Protected: a.protected ? 'Yes' : 'No'
  }));
  movedSheet.columns = [
    { header: 'Customer ID', key: 'Customer_ID', width: 16 },
    { header: 'Customer Name', key: 'Customer_Name', width: 28 },
    { header: 'Original Assigned Rep', key: 'Original_Assigned_Rep', width: 20 },
    { header: 'Assigned Rep', key: 'Assigned_Rep', width: 16 },
    { header: 'Current Rep', key: 'Current_Rep', width: 16 },
    { header: 'Revenue', key: 'Revenue', width: 14 },
    { header: 'Rank', key: 'Rank', width: 10 },
    { header: 'Protected', key: 'Protected', width: 12 }
  ];
  if (movedRows.length) {
    movedRows.forEach(row => movedSheet.addRow(row));
    styleHeaderRow(movedSheet.getRow(1));
    styleDataRows(movedSheet, 2, movedSheet.rowCount);
    movedSheet.getColumn('Revenue').numFmt = '$#,##0.00';
  }
  const buffer = await workbook.xlsx.writeBuffer();
  downloadArrayBufferAsFile(buffer, state.loadedFileName || 'territory_export_updated.xlsx');
  showToast('Excel export ready.');
}

function styleHeaderRow(row) {
  row.eachCell(cell => {
    cell.font = { name: 'Tw Cen MT', size: 9, bold: true, color: { argb: 'FF20364F' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF2FB' } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = { top: { style: 'thin', color: { argb: 'FFD3DFEB' } }, bottom: { style: 'thin', color: { argb: 'FFC3D3E2' } } };
  });
}

function styleDataRows(worksheet, fromRow, toRow) {
  for (let r = fromRow; r <= toRow; r++) {
    const row = worksheet.getRow(r);
    row.height = 18;
    row.eachCell(cell => {
      cell.font = { name: 'Tw Cen MT', size: 9, color: { argb: 'FF29415B' } };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFE6EDF5' } } };
    });
  }
}

function guessColumnWidth(key, rows) {
  let maxLen = String(key || '').length;
  rows.slice(0, 250).forEach(row => { const len = String(row[key] == null ? '' : row[key]).length; if (len > maxLen) maxLen = len; });
  return Math.max(10, Math.min(maxLen + 2, 42));
}

function downloadArrayBufferAsFile(buffer, filename) {
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url; link.download = filename;
  document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ─── MAP HELPERS ──────────────────────────────────────────────

function zoomToRep(rep) {
  const points = state.accounts.filter(a => a.assignedRep === rep).map(a => [a.latitude, a.longitude]);
  if (!points.length) return;
  if (points.length === 1) { state.map.setView(points[0], 12); return; }
  const bounds = L.latLngBounds(points);
  const spanLat = Math.abs(bounds.getNorth() - bounds.getSouth());
  const spanLng = Math.abs(bounds.getEast() - bounds.getWest());
  if (spanLat < 0.03 && spanLng < 0.03) { state.map.setView(bounds.getCenter(), 11); return; }
  state.map.fitBounds(bounds, { padding: [35, 35], maxZoom: 11 });
}

function fitMapToAccounts() {
  if (!state.accounts.length) return;
  state.map.fitBounds(state.accounts.map(a => [a.latitude, a.longitude]), { padding: [25, 25] });
}

function toggleTheme() {
  if (els.themeToggleCheck.checked && state.theme === 'light') {
    state.map.removeLayer(state.lightLayer); state.darkLayer.addTo(state.map); state.theme = 'dark';
  } else if (!els.themeToggleCheck.checked && state.theme === 'dark') {
    state.map.removeLayer(state.darkLayer); state.lightLayer.addTo(state.map); state.theme = 'light';
  }
}

// ─── COLOR HELPERS ────────────────────────────────────────────

function buildRepColors() {
  const previous = new Map(state.repColors);
  const reps = getAllKnownReps();
  const usedColors = new Set();
  state.repColors = new Map();
  reps.forEach(rep => {
    const existing = previous.get(rep);
    if (existing && COLOR_PALETTE.includes(existing)) { state.repColors.set(rep, existing); usedColors.add(existing); }
  });
  reps.forEach(rep => {
    if (state.repColors.has(rep)) return;
    const nextColor = pickBestAvailableColor(usedColors);
    state.repColors.set(rep, nextColor);
    usedColors.add(nextColor);
  });
}

function ensureRepColor(rep) {
  if (!rep || state.repColors.has(rep)) return;
  const nextColor = pickBestAvailableColor(new Set(state.repColors.values()));
  state.repColors.set(rep, nextColor);
}

function getRepColor(rep) { ensureRepColor(rep); return state.repColors.get(rep) || '#64748b'; }

function getAllAssignedReps() {
  const set = new Set();
  state.accounts.forEach(a => { if (a.assignedRep) set.add(a.assignedRep); });
  return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function getAllKnownReps() {
  const set = new Set();
  state.accounts.forEach(a => {
    if (a.assignedRep) set.add(a.assignedRep);
    if (a.currentRep) set.add(a.currentRep);
    if (a.originalAssignedRep) set.add(a.originalAssignedRep);
  });
  return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function clampByte(v) { return Math.max(0, Math.min(255, Math.round(v))); }

function hexToRgb(hex) {
  const value = String(hex || '').trim().replace('#', '');
  const normalized = value.length === 3 ? value.split('').map(ch => ch + ch).join('') : value.padEnd(6, '0').slice(0, 6);
  const int = Number.parseInt(normalized, 16);
  if (Number.isNaN(int)) return { r: 64, g: 99, b: 160 };
  return { r: (int >> 16) & 255, g: (int >> 8) & 255, b: int & 255 };
}

function rgbToHex(r, g, b) { return `#${[r, g, b].map(v => clampByte(v).toString(16).padStart(2, '0')).join('')}`; }

function mixHex(colorA, colorB, weight = 0.5) {
  const a = hexToRgb(colorA), b = hexToRgb(colorB);
  const t = Math.max(0, Math.min(1, Number(weight) || 0));
  return rgbToHex(a.r + (b.r - a.r) * t, a.g + (b.g - a.g) * t, a.b + (b.b - a.b) * t);
}

function getTerritoryStrokeColor(rep) { return mixHex(getRepColor(rep), '#24384f', 0.42); }
function getTerritoryFillColor(rep) { return mixHex(getRepColor(rep), '#ffffff', 0.58); }
function getTerritoryLabelColor(rep) { return mixHex(getRepColor(rep), '#22364f', 0.3); }

// ─── UTILITY ──────────────────────────────────────────────────

function fillSimpleSelect(selectEl, values, selectedValue, labelFn = v => v, placeholder = '') {
  if (!selectEl) return;
  const options = [];
  if (placeholder) options.push(`<option value="">${escapeHtml(placeholder)}</option>`);
  values.forEach(v => options.push(`<option value="${escapeHtmlAttr(v)}">${escapeHtml(labelFn(v))}</option>`));
  selectEl.innerHTML = options.join('');
  if (selectedValue != null && values.includes(selectedValue)) selectEl.value = selectedValue;
  else if (placeholder) selectEl.value = '';
}

function updateLastAction(text) { state.lastAction = text; els.lastAction.textContent = text; }

function showToast(message) {
  if (!els.toast) return;
  els.toast.textContent = message;
  els.toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toast.classList.remove('show'), 2200);
}

function getDistinctValues(arr, fn) {
  const set = new Set();
  arr.forEach(item => { const v = fn(item); if (safeString(v)) set.add(v); });
  return [...set].sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
}

function renderDeltaCount(value) {
  if (value > 0) return `<span class="num-pos">+${value}</span>`;
  if (value < 0) return `<span class="num-neg">${value}</span>`;
  return `<span class="num-zero">0</span>`;
}

function renderDeltaMoney(value) {
  if (value > 0) return `<span class="num-pos">+${formatCurrency(value)}</span>`;
  if (value < 0) return `<span class="num-neg">-${formatCurrency(Math.abs(value))}</span>`;
  return `<span class="num-zero">$0</span>`;
}

function normalizeCadence4W(value, rank) {
  const raw = safeString(value);
  if (!raw) { if (rank === 'A') return 4; if (rank === 'B') return 2; if (rank === 'C') return 1; return 0.33; }
  const normalized = raw.toLowerCase();
  const n = toNumber(raw);
  if (Number.isFinite(n) && n >= 0) return n;
  if (normalized.includes('weekly')) return 4;
  if (normalized.includes('biweekly') || normalized.includes('every other')) return 2;
  if (normalized.includes('monthly')) return 1;
  if (normalized.includes('quarter')) return 0.33;
  if (rank === 'A') return 4; if (rank === 'B') return 2; if (rank === 'C') return 1; return 0.33;
}

function rankSortValue(rank) { return { A: 0, B: 1, C: 2, D: 3 }[rank] ?? 9; }
function toCamel(id) { return id.replace(/-([a-z])/g, (_, c) => c.toUpperCase()); }
function cleanHeader(value) { return safeString(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
function safeString(value) { return value == null ? '' : String(value).trim(); }

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
  const raw = String(value ?? '').replace(/[$,%\s,]/g, '').trim();
  if (!raw) return NaN;
  const n = Number(raw);
  return Number.isFinite(n) ? n : NaN;
}

function toBoolean(value) { return ['true','yes','y','1','protected','locked'].includes(safeString(value).toLowerCase()); }
function normalizeRank(value) { const raw = safeString(value).toUpperCase(); return ['A','B','C','D'].includes(raw) ? raw : 'C'; }
function formatCurrency(value) { return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(value || 0); }
function formatNumber(value, digits = 0) { return Number(value || 0).toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits }); }
function round2(v) { return Math.round((v || 0) * 100) / 100; }
function round6(v) { return Math.round((v || 0) * 1000000) / 1000000; }
function squaredDistance(lat1, lng1, lat2, lng2) { const dx = lng1 - lng2, dy = lat1 - lat2; return dx * dx + dy * dy; }
function escapeHtml(text) { return String(text ?? '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
function escapeHtmlAttr(text) { return escapeHtml(text).replace(/"/g, '&quot;'); }
