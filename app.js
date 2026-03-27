const COLOR_PALETTE = [
  '#1f77b4','#d62728','#2ca02c','#9467bd','#ff7f0e','#17becf','#8c564b','#e377c2','#7f7f7f','#bcbd22',
  '#0b7285','#c92a2a','#2b8a3e','#5f3dc4','#e67700','#087f5b','#364fc7','#a61e4d','#495057','#2f9e44',
  '#f03e3e','#3b5bdb','#e8590c','#1098ad','#9c36b5','#5c940d','#d9480f','#1864ab','#c2255c','#12b886'
];

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
  repFocus: null,
  lockedReps: new Set(),
  theme: 'light',
  loadedFileName: 'territory_export_updated.xlsx',
  lastAction: 'No actions yet',
  uploadStatus: {
    level: 'neutral',
    text: 'No file loaded'
  },
  importSummary: {
    sourceRows: 0,
    loadedRows: 0,
    skippedNoCoords: 0,
    duplicateCustomerIds: 0,
    missingCurrentRep: 0,
    missingAssignedRep: 0,
    unmappedFields: []
  },
  optimizationSummary: null,
  tableSort: {
    key: 'rep',
    dir: 'asc'
  },
  filters: {
    rep: new Set(),
    rank: new Set(),
    chain: new Set(),
    segment: new Set(),
    premise: 'ALL',
    protected: 'ALL',
    moved: 'ALL'
  },
  multiSearch: {
    rep: '',
    rank: '',
    chain: '',
    segment: '',
    moved: ''
  },
  openMultiKey: null
};

const els = {};
let toastTimer = null;

document.addEventListener('DOMContentLoaded', init);

function init() {
  bindElements();
  initMap();
  bindEvents();
  initMultiFilters();
  updateLastAction('No actions yet');
  fillSimpleSelect(els.premiseFilter, ['ALL'], 'ALL', v => 'All premises');
  renderMultiFilterOptions();
  renderUploadStatus();
  syncControlState();
  initOptimizerTuningUI();
  updateOptimizerUI();

  requestAnimationFrame(() => {
    if (state.map) state.map.invalidateSize();
  });
}


function rebuildMarkers() {
  renderMap();
}

function setFieldLabelText(field, text) {
  if (!field) return;
  const candidates = field.querySelectorAll('label, .field-label, .field-title, .field-head, .control-label');
  for (const node of candidates) {
    if (!node) continue;
    const current = safeString(node.textContent).trim();
    if (!current) continue;
    node.textContent = text;
    return;
  }
}

function ensureOptimizerFeedbackMount() {
  if (els.optimizerFeedback) return els.optimizerFeedback;
  const routesCard = els.repTableBody ? els.repTableBody.closest('.routes-card') : null;
  if (!routesCard) return null;
  const host = routesCard.querySelector('.routes-table-wrap') || els.routesTableWrap || routesCard;
  if (!host) return null;
  const box = document.createElement('div');
  box.id = 'optimizer-feedback';
  box.className = 'optimizer-feedback';
  box.hidden = true;
  host.parentNode.insertBefore(box, host);
  els.optimizerFeedback = box;
  return box;
}

function initOptimizerTuningUI() {
  const disruptionField = els.disruptionSlider ? els.disruptionSlider.closest('.field') : null;
  if (disruptionField) {
    disruptionField.classList.add('field-disruption-enhanced');
    if (!els.optimizerDisruptionHelper) {
      const helper = document.createElement('div');
      helper.id = 'optimizer-disruption-helper';
      helper.className = 'optimizer-helper';
      disruptionField.appendChild(helper);
      els.optimizerDisruptionHelper = helper;
    }
  }

  const balanceField = els.balanceMode ? els.balanceMode.closest('.field') : null;
  if (balanceField) {
    balanceField.classList.add('field-optimizer-balance');
    setFieldLabelText(balanceField, 'Optimize weight');

    if (els.balanceMode) {
      els.balanceMode.value = 'hybrid';
      els.balanceMode.dataset.lockedMode = 'hybrid';
      els.balanceMode.disabled = true;
      els.balanceMode.classList.add('optimizer-mode-hidden');
      els.balanceMode.setAttribute('aria-hidden', 'true');
      els.balanceMode.tabIndex = -1;
      const selectWrap = els.balanceMode.closest('.field-control, .input-wrap, .select-wrap') || els.balanceMode.parentElement;
      if (selectWrap) selectWrap.classList.add('optimizer-mode-hidden-wrap');
    }

    if (!balanceField.querySelector('.optimizer-balance-wrap')) {
      const wrap = document.createElement('div');
      wrap.className = 'optimizer-balance-wrap';
      wrap.innerHTML = `
        <div class="optimizer-mini-head">
          <span>Optimize weight</span>
          <span id="optimizer-balance-value">Balanced</span>
        </div>
        <input id="optimizer-balance-slider" type="range" min="0" max="100" value="50" step="5" />
        <div class="optimizer-mini-scale">
          <span>Revenue</span>
          <span>Balanced</span>
          <span>Stops</span>
        </div>
        <div id="optimizer-balance-helper" class="optimizer-balance-helper"></div>
      `;
      balanceField.appendChild(wrap);
      els.optimizerBalanceSlider = wrap.querySelector('#optimizer-balance-slider');
      els.optimizerBalanceValue = wrap.querySelector('#optimizer-balance-value');
      els.optimizerBalanceHelper = wrap.querySelector('#optimizer-balance-helper');
      els.optimizerBalanceSlider.addEventListener('input', updateOptimizerUI);
    }
  }

  ensureOptimizerFeedbackMount();
}

function getOptimizerMix() {
  const stopsPriority = Math.max(0, Math.min(100, Number(els.optimizerBalanceSlider ? els.optimizerBalanceSlider.value : 50) || 50)) / 100;
  return {
    stopsPriority,
    revenuePriority: 1 - stopsPriority
  };
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
    els.disruptionValue.textContent = `${els.disruptionSlider.value} • ${preset.short}`;
    if (els.optimizerDisruptionHelper) {
      els.optimizerDisruptionHelper.textContent = preset.detail;
    }
  }

  if (els.balanceMode) {
    els.balanceMode.value = 'hybrid';
  }

  if (els.optimizerBalanceSlider && els.optimizerBalanceValue) {
    const { stopsPriority, revenuePriority } = getOptimizerMix();
    const stopPct = Math.round(stopsPriority * 100);
    const revenuePct = Math.round(revenuePriority * 100);
    let label = 'Balanced';
    if (stopPct >= 65) label = 'Stops first';
    else if (revenuePct >= 65) label = 'Revenue first';
    els.optimizerBalanceValue.textContent = label;
    if (els.optimizerBalanceHelper) {
      els.optimizerBalanceHelper.textContent = `Stops ${stopPct}% • Revenue ${revenuePct}%`;
    }
  }
}

function renderOptimizationFeedback() {
  const mount = ensureOptimizerFeedbackMount();
  if (!mount) return;
  const s = state.optimizationSummary;
  if (!s) {
    mount.hidden = true;
    mount.innerHTML = '';
    return;
  }

  const stopTone = s.stopRangeDeltaPct > 0 ? 'positive' : (s.stopRangeDeltaPct < 0 ? 'negative' : 'neutral');
  const revenueTone = s.revenueRangeDeltaPct > 0 ? 'positive' : (s.revenueRangeDeltaPct < 0 ? 'negative' : 'neutral');
  const stopLabel = s.stopRangeDeltaPct > 0
    ? `Stop spread improved ${formatNumber(s.stopRangeDeltaPct, 1)}%`
    : (s.stopRangeDeltaPct < 0 ? `Stop spread widened ${formatNumber(Math.abs(s.stopRangeDeltaPct), 1)}%` : 'Stop spread unchanged');
  const revenueLabel = s.revenueRangeDeltaPct > 0
    ? `Revenue spread improved ${formatNumber(s.revenueRangeDeltaPct, 1)}%`
    : (s.revenueRangeDeltaPct < 0 ? `Revenue spread widened ${formatNumber(Math.abs(s.revenueRangeDeltaPct), 1)}%` : 'Revenue spread unchanged');

  mount.innerHTML = `
    <div class="optimizer-feedback-card">
      <div class="optimizer-feedback-title-row">
        <div>
          <div class="optimizer-feedback-title">Optimization feedback</div>
          <div class="optimizer-feedback-subtitle">Latest run summary</div>
        </div>
        <div class="optimizer-feedback-run">${escapeHtml(s.weightLabel)} • ${escapeHtml(s.disruptionLabel)}</div>
      </div>
      <div class="optimizer-feedback-grid">
        <div class="optimizer-feedback-chip"><span class="k">Accounts moved</span><strong>${formatNumber(s.movedCount)}</strong></div>
        <div class="optimizer-feedback-chip"><span class="k">Protected held</span><strong>${formatNumber(s.protectedHeld)}</strong></div>
        <div class="optimizer-feedback-chip"><span class="k">Stops range</span><strong>${formatNumber(s.minStops)}-${formatNumber(s.maxStops)}</strong></div>
        <div class="optimizer-feedback-chip"><span class="k">Avg stops</span><strong>${formatNumber(s.avgStops, 1)}</strong></div>
      </div>
      <div class="optimizer-feedback-metrics">
        <div class="optimizer-feedback-metric ${stopTone}">${escapeHtml(stopLabel)}</div>
        <div class="optimizer-feedback-metric ${revenueTone}">${escapeHtml(revenueLabel)}</div>
      </div>
    </div>
  `;
  mount.hidden = false;
}

function refreshMarkerStyles(accountIds = null) {
  refreshMarkers(accountIds);
}
function refreshMarkerStyles(accountIds = null) {
  refreshMarkers(accountIds);
}

function renderSummary() {
  updateGlobalStats();
}

function syncSortHeaderIndicators() {
  document.querySelectorAll('th[data-sort]').forEach(th => {
    const active = th.getAttribute('data-sort') === state.tableSort.key;
    th.classList.toggle('is-active', active);
    const indicator = th.querySelector('.sort-indicator');
    if (indicator) indicator.textContent = active ? (state.tableSort.dir === 'asc' ? '▲' : '▼') : '↕';
  });
}

function renderRepControls() {
  const reps = getAllAssignedReps();
  const currentValue = els.assignRepSelect ? els.assignRepSelect.value : '';
  fillSimpleSelect(els.assignRepSelect, reps, reps.includes(currentValue) ? currentValue : '', v => v, 'Select rep');
}

function markTerritoriesDirty() {
  state.territoryDirty = true;
}

function scheduleTerritoryRefresh(force = false) {
  if (force) state.territoryDirty = true;
  const token = ++state.territoryRefreshToken;

  if (state.territoryRefreshTimer) {
    clearTimeout(state.territoryRefreshTimer);
    state.territoryRefreshTimer = null;
  }

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
  return {
    rep,
    stops: 0,
    deltaStops: 0,
    revenue: 0,
    deltaRevenue: 0,
    A: 0,
    B: 0,
    C: 0,
    D: 0,
    planned4W: 0,
    avgWeekly: 0,
    protected: 0,
    movedIn: 0,
    movedOut: 0
  };
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

    if (originalRep === rep && assignedRep !== rep) {
      row.movedOut += 1;
    }
  }

  row.deltaStops = row.stops - baseline.stops;
  row.deltaRevenue = row.revenue - baseline.revenue;
  row.avgWeekly = row.planned4W / 4;
  return row;
}

function updateRepSummaryCacheForReps(reps) {
  if (!reps || !reps.size) return;
  if (!state.repSummaryCache || !state.repSummaryCache.size) {
    summarizeByRep();
  }

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
  const repOk = !state.filters.rep.size || state.filters.rep.has(account.assignedRep);
  const rankOk = !state.filters.rank.size || state.filters.rank.has(account.rank);
  const chainOk = !state.filters.chain.size || state.filters.chain.has(account.chain);
  const segmentOk = !state.filters.segment.size || state.filters.segment.has(account.segment);
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

  if (state.repFocus && !getAllAssignedReps().includes(state.repFocus)) {
    state.repFocus = null;
  }

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
    'global-accounts','global-revenue','global-protected','global-moved','global-unchanged','global-avg-weekly','global-avg-weekly-per-rep',
    'last-action','toast','clear-selection-btn','theme-toggle-check','premise-filter','protected-filter',
    'moved-filter','moved-review-list','moved-review-count','rep-filter-options','rank-filter-options','chain-filter-options',
    'segment-filter-options','rep-filter-summary','rank-filter-summary','chain-filter-summary','segment-filter-summary',
    'routes-table-wrap','moved-search-input','upload-status-pill','upload-status-icon','upload-status-text','upload-status-panel','upload-status-body',
    'detail-panel'
  ].forEach(id => {
    els[toCamel(id)] = document.getElementById(id);
  });
}

function bindEvents() {
  els.fileInput.addEventListener('change', onFileChosen);
  els.loadSheetBtn.addEventListener('click', loadSelectedSheet);
  els.assignBtn.addEventListener('click', assignSelectionToRep);
  els.undoBtn.addEventListener('click', undoLastAction);
  els.resetBtn.addEventListener('click', resetAssignments);
  els.optimizeBtn.addEventListener('click', optimizeRoutes);
  els.exportBtn.addEventListener('click', exportWorkbook);
  els.clearSelectionBtn.addEventListener('click', clearSelection);

  if (els.detailPanel) {
    const detailCard = els.detailPanel.closest('.detail-card');
    const detailClickTarget = detailCard || els.detailPanel;
    detailClickTarget.addEventListener('click', e => {
      const clearBtn = e.target.closest('[data-clear-detail-selection], [data-clear-detail-selection-static]');
      if (clearBtn) {
        e.preventDefault();
        clearSelection();
      }
    });
  }

  if (els.uploadStatusPill) {
    els.uploadStatusPill.addEventListener('click', e => {
      e.stopPropagation();
      toggleUploadStatusPanel();
    });
  }

  els.themeToggleCheck.addEventListener('change', toggleTheme);
  els.dimOthersCheckbox.addEventListener('change', refreshUI);
  els.showTerritoryCheckbox.addEventListener('change', () => scheduleTerritoryRefresh(true));

  els.premiseFilter.addEventListener('change', () => {
    state.filters.premise = els.premiseFilter.value;
    refreshUI();
  });

  els.protectedFilter.addEventListener('change', () => {
    state.filters.protected = els.protectedFilter.value;
    refreshUI();
  });

  els.movedFilter.addEventListener('change', () => {
    state.filters.moved = els.movedFilter.value;
    refreshUI();
  });

  els.disruptionSlider.addEventListener('input', updateOptimizerUI);

  if (els.movedSearchInput) {
    els.movedSearchInput.addEventListener('input', e => {
      state.multiSearch.moved = e.target.value || '';
      renderMovedReview();
    });
  }

  document.querySelectorAll('th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      toggleTableSort(th.getAttribute('data-sort'));
    });
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
    if (e.key === 'Escape') {
      closeAllMultiPanels();
      closeUploadStatusPanel();
    }
  });
}

function initMap() {
  state.map = L.map('map', { preferCanvas: true }).setView([40.1, -89.2], 7);

  state.lightLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap contributors'
  });

  state.darkLayer = L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap &copy; CARTO'
  });

  state.lightLayer.addTo(state.map);

  state.markerLayer = L.layerGroup().addTo(state.map);
  state.territoryLayer = L.layerGroup().addTo(state.map);
  state.territoryLabelLayer = L.layerGroup().addTo(state.map);
  state.drawLayer = new L.FeatureGroup().addTo(state.map);

  state.drawControl = new L.Control.Draw({
    draw: {
      polyline: false,
      circle: false,
      circlemarker: false,
      marker: false,
      polygon: {
        allowIntersection: false,
        showArea: true,
        shapeOptions: { color: '#245fb7', weight: 2 }
      },
      rectangle: {
        shapeOptions: { color: '#0e9372', weight: 2 }
      }
    },
    edit: {
      featureGroup: state.drawLayer,
      edit: false,
      remove: true
    }
  });

  state.map.addControl(state.drawControl);

  state.map.on(L.Draw.Event.CREATED, handleDrawCreated);
  state.map.on(L.Draw.Event.DELETED, () => {
    clearSelection();
    showToast('Selection cleared.');
  });
  state.map.on('zoomend', () => scheduleTerritoryRefresh(true));
  state.map.on('moveend', () => scheduleTerritoryRefresh(true));
}

function initMultiFilters() {
  ['rep','rank','chain','segment'].forEach(key => {
    const trigger = document.querySelector(`[data-multi-trigger="${key}"]`);
    const selectAllBtn = document.querySelector(`[data-select-all="${key}"]`);
    const searchInput = document.querySelector(`[data-search="${key}"]`);

    if (trigger) {
      trigger.addEventListener('click', e => {
        e.stopPropagation();
        toggleMultiPanel(key);
      });
    }

    if (selectAllBtn) {
      selectAllBtn.addEventListener('click', e => {
        e.stopPropagation();
        toggleSelectAllMulti(key);
        positionMultiPanel(key);
      });
    }

    if (searchInput) {
      searchInput.addEventListener('input', e => {
        state.multiSearch[key] = e.target.value || '';
        renderMultiFilterOptions();
        positionMultiPanel(key);
      });
    }
  });
}

function handleDocumentClickForPanels(event) {
  const openMulti = document.querySelector('.multi.open');
  if (openMulti && !openMulti.contains(event.target)) closeAllMultiPanels();

  if (
    els.uploadStatusPanel &&
    !els.uploadStatusPanel.hidden &&
    !els.uploadStatusPanel.contains(event.target) &&
    !els.uploadStatusPill.contains(event.target)
  ) {
    closeUploadStatusPanel();
  }
}

function toggleMultiPanel(key) {
  const wrap = document.getElementById(`${key}-filter-wrap`);
  if (!wrap) return;

  const alreadyOpen = wrap.classList.contains('open');
  closeAllMultiPanels();

  if (!alreadyOpen) {
    wrap.classList.add('open');
    state.openMultiKey = key;
    positionMultiPanel(key);
  }
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

function toggleSelectAllMulti(key) {
  const options = getFilterOptionsForKey(key);
  const filtered = getVisibleOptionsForKey(key, options);
  const selectedSet = state.filters[key];
  const allVisibleSelected = filtered.every(v => selectedSet.has(v));

  if (allVisibleSelected) filtered.forEach(v => selectedSet.delete(v));
  else filtered.forEach(v => selectedSet.add(v));

  if (selectedSet.size === 0) options.forEach(v => selectedSet.add(v));

  renderMultiFilterOptions();
  refreshUI();
}

function getFilterOptionsForKey(key) {
  switch (key) {
    case 'rep': return getAllAssignedReps();
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
  renderMultiOptionList('rep', els.repFilterOptions, els.repFilterSummary, getAllAssignedReps(), 'All reps');
  renderMultiOptionList('rank', els.rankFilterOptions, els.rankFilterSummary, getDistinctValues(state.accounts, a => a.rank), 'All ranks');
  renderMultiOptionList('chain', els.chainFilterOptions, els.chainFilterSummary, getDistinctValues(state.accounts, a => a.chain), 'All chains');
  renderMultiOptionList('segment', els.segmentFilterOptions, els.segmentFilterSummary, getDistinctValues(state.accounts, a => a.segment), 'All segments');

  if (state.openMultiKey) positionMultiPanel(state.openMultiKey);
}

function renderMultiOptionList(key, container, summaryEl, options, allLabel) {
  if (!container || !summaryEl) return;

  const selectedSet = state.filters[key];
  const visibleOptions = getVisibleOptionsForKey(key, options);

  if (options.length && selectedSet.size === 0) options.forEach(v => selectedSet.add(v));

  container.innerHTML = visibleOptions.length
    ? visibleOptions.map(value => {
        const checked = selectedSet.has(value) ? 'checked' : '';
        return `
          <div class="multi-option">
            <label>
              <input type="checkbox" data-multi-check="${escapeHtmlAttr(key)}" value="${escapeHtmlAttr(value)}" ${checked} />
              <span>${escapeHtml(value)}</span>
            </label>
          </div>
        `;
      }).join('')
    : '<div class="empty">No matches.</div>';

  container.querySelectorAll('input[data-multi-check]').forEach(input => {
    input.addEventListener('change', e => {
      const value = e.target.value;
      if (e.target.checked) selectedSet.add(value);
      else selectedSet.delete(value);
      if (selectedSet.size === 0) options.forEach(v => selectedSet.add(v));
      renderMultiFilterOptions();
      refreshUI();
    });
  });

  if (!options.length || selectedSet.size === options.length) {
    summaryEl.textContent = allLabel;
  } else if (selectedSet.size === 1) {
    summaryEl.textContent = [...selectedSet][0];
  } else {
    summaryEl.textContent = `${selectedSet.size} selected`;
  }
}

function toggleUploadStatusPanel() {
  if (!els.uploadStatusPanel) return;

  const willOpen = els.uploadStatusPanel.hidden;
  if (willOpen) {
    els.uploadStatusPanel.hidden = false;
    els.uploadStatusPill.setAttribute('aria-expanded', 'true');
    renderUploadStatusDetails();
    positionUploadStatusPanel();
  } else {
    closeUploadStatusPanel();
  }
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
  if (top + estimatedHeight > window.innerHeight - 10) {
    top = Math.max(10, rect.top - estimatedHeight - 8);
  }

  panel.style.width = `${width}px`;
  panel.style.left = `${left}px`;
  panel.style.top = `${top}px`;
}

function setUploadStatus(level, text) {
  state.uploadStatus = { level, text };
}

function renderUploadStatus() {
  if (!els.uploadStatusPill) return;

  const level = state.uploadStatus.level || 'neutral';
  const text = state.uploadStatus.text || 'No file loaded';

  els.uploadStatusPill.className = `upload-status-pill upload-status-${level}`;
  els.uploadStatusText.textContent = text;
  els.uploadStatusIcon.textContent =
    level === 'good' ? '✓' :
    level === 'warning' ? '!' :
    level === 'bad' ? '×' : '•';

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

      if (lower.endsWith('.csv')) {
        workbook = XLSX.read(e.target.result, { type: 'binary' });
      } else {
        const arr = new Uint8Array(e.target.result);
        workbook = XLSX.read(arr, { type: 'array' });
      }

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

  reader.onerror = () => {
    setUploadStatus('bad', 'Load failed');
    renderUploadStatus();
    showToast('Could not read that file.');
  };

  if (file.name.toLowerCase().endsWith('.csv')) {
    reader.readAsBinaryString(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function loadWorkbook(workbook) {
  state.workbook = workbook;
  state.workbookSheets = {};

  workbook.SheetNames.forEach(name => {
    state.workbookSheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '' });
  });

  fillSimpleSelect(els.sheetSelect, workbook.SheetNames, workbook.SheetNames[0]);
  els.sheetSelect.disabled = false;
  els.loadSheetBtn.disabled = false;

  loadSelectedSheet();
}

function loadSelectedSheet() {
  const sheetName = els.sheetSelect.value;
  if (!sheetName || !state.workbookSheets[sheetName]) {
    showToast('No sheet selected.');
    return;
  }

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
  state.accountById = new Map(normalized.map(a => [a._id, a]));
  state.neighborMap = buildNeighborMap(normalized);
  state.currentHeaderMap = normalizedResult.headerMap || {};
  invalidateCaches();

  buildRepColors();
  syncRepFilterSelection();
  fillSimpleSelect(
    els.premiseFilter,
    ['ALL', ...getDistinctValues(state.accounts, a => a.premise)],
    'ALL',
    v => v === 'ALL' ? 'All premises' : v
  );

  renderMultiFilterOptions();
  renderMap();
  refreshUI();
  fitMapToAccounts();

  setUploadStatus(
    state.importSummary.skippedNoCoords || state.importSummary.duplicateCustomerIds || state.importSummary.missingCurrentRep
      ? 'warning'
      : 'good',
    state.importSummary.skippedNoCoords || state.importSummary.duplicateCustomerIds || state.importSummary.missingCurrentRep
      ? 'Loaded with warnings'
      : 'Loaded successfully'
  );
  renderUploadStatus();
}

function normalizeRows(rows) {
  const headerMap = buildHeaderMap(rows);
  const accounts = [];
  const usedIds = new Set();
  let skippedNoCoords = 0;
  let duplicateCustomerIds = 0;
  let missingCurrentRep = 0;
  let missingAssignedRep = 0;
  const unmappedFields = [];

  const allHeaders = rows.length ? Object.keys(rows[0]) : [];
  allHeaders.forEach(h => {
    const cleaned = cleanHeader(h);
    const matched = Object.values(COLUMN_ALIASES).some(list => list.includes(cleaned));
    if (!matched) unmappedFields.push(h);
  });

  rows.forEach((row, index) => {
    const latitudeRaw = row[headerMap.latitude];
    const longitudeRaw = row[headerMap.longitude];

    let latitude = toNumber(latitudeRaw);
    let longitude = toNumber(longitudeRaw);
    ({ latitude, longitude } = normalizeCoordinates(latitude, longitude));

    if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
      skippedNoCoords += 1;
      return;
    }

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
      _id: customerId,
      customerId,
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
      latitude: round6(latitude),
      longitude: round6(longitude),
      sourceRow: row
    });
  });

  return {
    accounts,
    headerMap,
    summary: {
      sourceRows: rows.length,
      loadedRows: accounts.length,
      skippedNoCoords,
      duplicateCustomerIds,
      missingCurrentRep,
      missingAssignedRep,
      unmappedFields
    }
  };
}

function buildHeaderMap(rows) {
  const firstRow = rows[0] || {};
  const map = {};
  const cleanedHeaders = Object.keys(firstRow).map(h => ({
    original: h,
    cleaned: cleanHeader(h)
  }));

  Object.entries(COLUMN_ALIASES).forEach(([key, aliases]) => {
    const cleanedAliases = aliases.map(a => cleanHeader(a));
    const match = cleanedHeaders.find(h => cleanedAliases.includes(h.cleaned));
    map[key] = match ? match.original : null;
  });

  return map;
}

function normalizeCoordinates(latitude, longitude) {
  if (Number.isFinite(latitude) && Number.isFinite(longitude)) {
    if (Math.abs(latitude) > 90 && Math.abs(longitude) <= 90) {
      return { latitude: longitude, longitude: latitude };
    }
  }
  return { latitude, longitude };
}

function buildNeighborMap(accounts) {
  const map = new Map();
  if (!accounts.length) return map;

  const cellSize = 0.18;
  const grid = new Map();

  function cellKey(lat, lng) {
    const x = Math.floor(lng / cellSize);
    const y = Math.floor(lat / cellSize);
    return `${x}|${y}`;
  }

  for (const account of accounts) {
    map.set(account._id, new Set());
    const key = cellKey(account.latitude, account.longitude);
    if (!grid.has(key)) grid.set(key, []);
    grid.get(key).push(account);
  }

  for (const account of accounts) {
    const x = Math.floor(account.longitude / cellSize);
    const y = Math.floor(account.latitude / cellSize);
    const candidates = [];

    for (let dx = -1; dx <= 1; dx += 1) {
      for (let dy = -1; dy <= 1; dy += 1) {
        const cellAccounts = grid.get(`${x + dx}|${y + dy}`);
        if (!cellAccounts) continue;
        for (const other of cellAccounts) {
          if (other._id === account._id) continue;
          candidates.push(other);
        }
      }
    }

    if (candidates.length < 10) {
      for (let radius = 2; radius <= 4 && candidates.length < 20; radius += 1) {
        for (let dx = -radius; dx <= radius; dx += 1) {
          for (let dy = -radius; dy <= radius; dy += 1) {
            if (Math.abs(dx) !== radius && Math.abs(dy) !== radius) continue;
            const cellAccounts = grid.get(`${x + dx}|${y + dy}`);
            if (!cellAccounts) continue;
            for (const other of cellAccounts) {
              if (other._id === account._id) continue;
              candidates.push(other);
            }
          }
        }
      }
    }

    candidates
      .map(other => ({
        id: other._id,
        d: squaredDistance(account.latitude, account.longitude, other.latitude, other.longitude)
      }))
      .sort((a, b) => a.d - b.d)
      .slice(0, 10)
      .forEach(item => {
        map.get(account._id).add(item.id);
        map.get(item.id)?.add(account._id);
      });
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

    const marker = L.circleMarker([account.latitude, account.longitude], {
      radius: 2.8,
      color,
      weight: 1,
      opacity: 0.95,
      fillColor: color,
      fillOpacity: 0.88
    });

    marker.on('click', e => {
      L.DomEvent.stopPropagation(e);
      const additive = !!(e.originalEvent?.ctrlKey || e.originalEvent?.metaKey);
      toggleSelection(account._id, additive);
      marker.setPopupContent(buildPopupHtml(account));
      marker.openPopup();
    });

    marker.bindPopup(buildPopupHtml(account), {
      autoPan: true,
      closeButton: true,
      offset: [0, -6]
    });

    state.markerLayer.addLayer(marker);
    state.markerById.set(account._id, marker);
    state.markerMetaById.set(account._id, {
      color,
      radius: 2.8,
      opacity: 0.95,
      fillOpacity: 0.88,
      weight: 1,
      hidden: false,
      popupKey: `${account.assignedRep}|${account.currentRep}|${account.overallSales}|${account.rank}|${account.protected ? 1 : 0}`
    });
    state.accountPointById.set(account._id, turf.point([account.longitude, account.latitude]));
  }

  invalidateCaches();
  refreshMarkerStyles();
  scheduleTerritoryRefresh(true);
}

function buildPopupHtml(account) {
  const title = account.customerName || account.customerId;
  const line2 = [
    account.address,
    [account.city, account.zip].filter(Boolean).join(' ')
  ].filter(Boolean).join(' • ');

  return `
    <div style="min-width:240px;max-width:280px;">
      <div style="font-size:15px;font-weight:800;line-height:1.2;margin-bottom:4px;">
        ${escapeHtml(title)}
      </div>

      <div style="font-size:12px;color:#5d7286;line-height:1.35;margin-bottom:8px;">
        ${escapeHtml(line2)}
      </div>

      <div style="display:grid;grid-template-columns:auto 1fr;gap:4px 8px;font-size:12px;line-height:1.3;">
        <div><strong>Rep</strong></div><div>${escapeHtml(account.assignedRep)}</div>
        <div><strong>Current</strong></div><div>${escapeHtml(account.currentRep)}</div>
        <div><strong>Revenue</strong></div><div>${formatCurrency(account.overallSales)}</div>
        <div><strong>Rank</strong></div><div>${escapeHtml(account.rank)}</div>
        <div><strong>Protected</strong></div><div>${account.protected ? 'Yes' : 'No'}</div>
      </div>
    </div>
  `;
}

function ensureDetailClearButton() {
  if (!els.detailPanel) return null;
  const detailCard = els.detailPanel.closest('.detail-card');
  if (!detailCard) return null;

  const head = detailCard.querySelector('.card-head');
  if (!head) return null;

  head.style.display = 'flex';
  head.style.alignItems = 'flex-start';
  head.style.justifyContent = 'space-between';
  head.style.gap = '12px';

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
  if (!selectedIds.length) {
    els.detailPanel.innerHTML = '<div class="empty">No account selected.</div>';
    return;
  }

  const selectedAccounts = selectedIds
    .map(id => state.accountById.get(id))
    .filter(Boolean)
    .slice(0, 10);

  const cardsHtml = selectedAccounts.map(account => `
    <div class="selected-item" style="margin-bottom:10px;">
      <div class="selected-item-title">${escapeHtml(account.customerName || account.customerId)}</div>
      <div style="font-size:12px;color:#5d7286;margin-bottom:6px;">
        ${escapeHtml([account.address, [account.city, account.zip].filter(Boolean).join(' ')].filter(Boolean).join(' • '))}
      </div>
      <div class="transfer-line">
        <span class="rep-chip">${escapeHtml(account.assignedRep)}</span>
        <span class="metric-chip">${formatCurrency(account.overallSales)}</span>
        <span class="metric-chip">Rank ${escapeHtml(account.rank)}</span>
        ${account.protected ? '<span class="metric-chip">Protected</span>' : ''}
      </div>
    </div>
  `).join('');

  const metaHtml = `
    <div style="font-size:12px;color:#5d7286;font-weight:700;margin-bottom:10px;">${selectedIds.length} selected</div>
  `;

  const moreCount = selectedIds.length - selectedAccounts.length;
  const moreHtml = moreCount > 0
    ? `<div class="small muted" style="margin-top:6px;">Showing first ${selectedAccounts.length} selected accounts. ${moreCount} more selected.</div>`
    : '';

  els.detailPanel.innerHTML = `${metaHtml}${cardsHtml}${moreHtml}`;
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
  const ids = accountIds
    ? Array.from(accountIds)
    : state.accounts.map(account => account._id);

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
      radius: selected
        ? 4.2
        : dimmed
          ? 1.65
          : (state.repFocus && account.assignedRep === state.repFocus ? 3.8 : 2.8),
      opacity: pass ? (dimmed ? 0.02 : 0.98) : 0,
      fillOpacity: pass ? (selected ? 1 : dimmed ? 0.015 : 0.92) : 0,
      weight: selected ? 2 : (dimmed ? 0.5 : 1),
      hidden: !pass
    };

    const prevState = state.markerMetaById.get(id) || {};
    if (
      prevState.color !== nextState.color ||
      prevState.opacity !== nextState.opacity ||
      prevState.fillOpacity !== nextState.fillOpacity ||
      prevState.weight !== nextState.weight
    ) {
      marker.setStyle({
        color: nextState.color,
        fillColor: nextState.color,
        opacity: nextState.opacity,
        fillOpacity: nextState.fillOpacity,
        weight: nextState.weight
      });
    }

    if (prevState.radius !== nextState.radius) {
      marker.setRadius(nextState.radius);
    }

    const popupKey = `${account.assignedRep}|${account.currentRep}|${account.overallSales}|${account.rank}|${account.protected ? 1 : 0}`;
    if (prevState.popupKey !== popupKey) {
      marker.setPopupContent(buildPopupHtml(account));
      nextState.popupKey = popupKey;
    } else {
      nextState.popupKey = prevState.popupKey;
    }

    state.markerMetaById.set(id, nextState);
  }
}

function refreshTerritories() {
  state.territoryDirty = false;
  state.territoryLayer.clearLayers();
  state.territoryLabelLayer.clearLayers();

  if (!els.showTerritoryCheckbox.checked) return;

  const membersByRep = new Map();
  for (const account of state.accounts) {
    if (!passesFilters(account)) continue;
    const rep = account.assignedRep;
    if (!membersByRep.has(rep)) membersByRep.set(rep, []);
    membersByRep.get(rep).push(account);
  }

  for (const [rep, members] of membersByRep.entries()) {
    if (members.length < 3) continue;

    const points = new Array(members.length);
    for (let i = 0; i < members.length; i += 1) {
      points[i] = turf.point([members[i].longitude, members[i].latitude]);
    }

    let hull = null;
    try {
      hull = turf.convex(turf.featureCollection(points));
    } catch (err) {
      hull = null;
    }

    if (!hull) continue;

    const color = getRepColor(rep);
    const strokeColor = getTerritoryStrokeColor(rep);
    const fillColor = getTerritoryFillColor(rep);
    const polygon = L.geoJSON(hull, {
      style: {
        color: strokeColor,
        weight: 2.15,
        fillColor,
        fillOpacity: 0.09,
        opacity: 0.8,
        dashArray: '2 4'
      }
    });

    state.territoryLayer.addLayer(polygon);

    const center = turf.center(hull).geometry.coordinates;
    const label = L.marker([center[1], center[0]], {
      interactive: false,
      icon: L.divIcon({
        className: 'territory-label',
        html: `<div style="background:${getTerritoryLabelColor(rep)};color:#fff;border:1px solid rgba(255,255,255,.52);box-shadow:0 6px 14px rgba(30,54,84,.14);border-radius:999px;padding:4px 10px;font-size:11px;font-weight:800;white-space:nowrap;letter-spacing:.1px;">${escapeHtml(rep)}</div>`
      })
    });

    state.territoryLabelLayer.addLayer(label);
  }
}

function handleDrawCreated(event) {
  state.drawLayer.clearLayers();
  const layer = event.layer;
  state.drawLayer.addLayer(layer);

  let polygon = null;
  if (layer instanceof L.Rectangle || layer instanceof L.Polygon) {
    polygon = layer.toGeoJSON();
  }
  if (!polygon) return;

  const bbox = turf.bbox(polygon);
  const nextSelection = new Set();

  for (const account of state.accounts) {
    if (
      account.longitude < bbox[0] || account.longitude > bbox[2] ||
      account.latitude < bbox[1] || account.latitude > bbox[3]
    ) {
      continue;
    }

    const point = state.accountPointById.get(account._id) || turf.point([account.longitude, account.latitude]);
    if (turf.booleanPointInPolygon(point, polygon)) {
      nextSelection.add(account._id);
    }
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

function isRepLocked(rep) {
  return state.lockedReps.has(rep);
}

function isAccountLocked(account) {
  return !!account && isRepLocked(account.assignedRep);
}

function toggleRepLock(rep, shouldLock) {
  if (!rep) return;
  if (shouldLock) state.lockedReps.add(rep);
  else state.lockedReps.delete(rep);
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

  if (!rows.length) {
    els.repTableBody.innerHTML = '<tr><td colspan="15" class="empty">Upload a file to begin.</td></tr>';
    return;
  }

  syncSortHeaderIndicators();

  els.repTableBody.innerHTML = rows.map(row => `
    <tr
      data-rep-row="${encodeURIComponent(row.rep)}"
      class="${state.repFocus === row.rep ? 'rep-row-active' : ''} ${isRepLocked(row.rep) ? 'rep-row-locked' : ''}"
    >
      <td>
        <div class="rep-cell">
          <span class="color-dot" style="background:${getRepColor(row.rep)}"></span>
          <span>${escapeHtml(row.rep)}</span>
        </div>
      </td>
      <td class="lock-cell">
        <label class="lock-toggle" title="Lock this territory">
          <input
            type="checkbox"
            class="rep-lock-checkbox"
            data-lock-rep="${escapeHtmlAttr(row.rep)}"
            ${isRepLocked(row.rep) ? 'checked' : ''}
          />
        </label>
      </td>
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
    </tr>
  `).join('');

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
    input.addEventListener('change', e => {
      const rep = e.target.getAttribute('data-lock-rep');
      toggleRepLock(rep, e.target.checked);
    });
  });
}

function summarizeByRep() {
  if (state.repSummaryCache && state.repSummaryCache.size) {
    return [...state.repSummaryCache.values()].map(row => ({ ...row }));
  }

  const map = new Map();
  const originalMap = new Map();

  for (const account of state.accounts) {
    const assignedRep = account.assignedRep || 'Unassigned';
    const originalRep = account.originalAssignedRep || 'Unassigned';

    if (!map.has(assignedRep)) {
      map.set(assignedRep, {
        rep: assignedRep,
        stops: 0,
        deltaStops: 0,
        revenue: 0,
        deltaRevenue: 0,
        A: 0,
        B: 0,
        C: 0,
        D: 0,
        planned4W: 0,
        avgWeekly: 0,
        protected: 0,
        movedIn: 0,
        movedOut: 0
      });
    }

    if (!originalMap.has(originalRep)) {
      originalMap.set(originalRep, { stops: 0, revenue: 0 });
    }

    const row = map.get(assignedRep);
    const orig = originalMap.get(originalRep);

    row.stops += 1;
    row.revenue += Number(account.overallSales || 0);
    row.planned4W += Number(account.cadence4w || 0);
    if (row[account.rank] != null) row[account.rank] += 1;
    if (account.protected) row.protected += 1;
    if (assignedRep !== originalRep) row.movedIn += 1;

    orig.stops += 1;
    orig.revenue += Number(account.overallSales || 0);
  }

  for (const account of state.accounts) {
    const originalRep = account.originalAssignedRep || 'Unassigned';
    const assignedRep = account.assignedRep || 'Unassigned';
    if (originalRep !== assignedRep && map.has(originalRep)) {
      map.get(originalRep).movedOut += 1;
    }
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
    let av = a[key];
    let bv = b[key];
    if (typeof av === 'string' || typeof bv === 'string') {
      return String(av).localeCompare(String(bv), undefined, { numeric: true }) * factor;
    }
    av = Number(av || 0);
    bv = Number(bv || 0);
    return (av - bv) * factor;
  });
}

function toggleTableSort(key) {
  if (state.tableSort.key === key) {
    state.tableSort.dir = state.tableSort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    state.tableSort = { key, dir: 'asc' };
  }

  document.querySelectorAll('th[data-sort]').forEach(th => {
    const active = th.getAttribute('data-sort') === state.tableSort.key;
    th.classList.toggle('is-active', active);
    const indicator = th.querySelector('.sort-indicator');
    if (indicator) indicator.textContent = active ? (state.tableSort.dir === 'asc' ? '▲' : '▼') : '↕';
  });

  renderRepTable();
}

function renderSelectionPreview() {
  const ids = [...state.selection];
  els.selectionCount.textContent = ids.length;

  if (!ids.length) {
    els.selectionPreview.innerHTML = '<div class="empty">No accounts selected.</div>';
    return;
  }

  els.selectionPreview.innerHTML = ids.slice(0, 50).map(id => {
    const a = state.accountById.get(id);
    if (!a) return '';
    return `
      <div class="selected-item">
        <div class="selected-item-title">${escapeHtml(a.customerName)}</div>
        <div class="transfer-line">
          <span class="rep-chip">${escapeHtml(a.assignedRep)}</span>
          <span class="metric-chip">${formatCurrency(a.overallSales)}</span>
        </div>
      </div>
    `;
  }).join('');
}

function renderMovedReview() {
  const term = (state.multiSearch.moved || '').trim().toLowerCase();
  let moved = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep);

  if (term) {
    moved = moved.filter(a =>
      [a.customerName, a.customerId, a.originalAssignedRep, a.assignedRep]
        .some(v => String(v || '').toLowerCase().includes(term))
    );
  }

  els.movedReviewCount.textContent = moved.length;

  if (!moved.length) {
    els.movedReviewList.innerHTML = '<div class="empty">No moved accounts yet.</div>';
    return;
  }

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
    </div>
  `).join('');
}

function updateGlobalStats() {
  let stats = state.globalStatsCache;

  if (!stats) {
    let visibleCount = 0;
    let movedCount = 0;
    let totalRevenue = 0;
    let totalProtected = 0;
    let planned4W = 0;

    for (const account of state.accounts) {
      const moved = account.assignedRep !== account.originalAssignedRep;
      if (moved) movedCount += 1;
      if (!passesFilters(account)) continue;
      visibleCount += 1;
      totalRevenue += account.overallSales || 0;
      totalProtected += account.protected ? 1 : 0;
      planned4W += account.cadence4w || 0;
    }

    const reps = getAllAssignedReps();
    const unchangedPct = state.accounts.length
      ? ((state.accounts.length - movedCount) / state.accounts.length) * 100
      : 0;

    stats = {
      visibleCount,
      movedCount,
      unchangedPct,
      totalRevenue,
      totalProtected,
      planned4W,
      repCount: reps.length
    };
    state.globalStatsCache = stats;
  }

  els.globalAccounts.textContent = formatNumber(stats.visibleCount);
  els.globalRevenue.textContent = formatCurrency(stats.totalRevenue);
  els.globalProtected.textContent = formatNumber(stats.totalProtected);
  els.globalMoved.textContent = formatNumber(stats.movedCount);
  els.globalUnchanged.textContent = `${formatNumber(stats.unchangedPct, 1)}%`;
  els.globalAvgWeekly.textContent = formatNumber(stats.planned4W / 4, 1);
  els.globalAvgWeeklyPerRep.textContent = formatNumber((stats.planned4W / 4) / Math.max(1, stats.repCount), 1);
}

function syncRepFilterSelection(previousAssignedReps = null) {
  const reps = getAllAssignedReps();
  const prev = Array.isArray(previousAssignedReps) ? new Set(previousAssignedReps) : state.filters.rep;

  state.filters.rep = new Set(reps.filter(rep => prev.has(rep) || !previousAssignedReps));
  if (!state.filters.rep.size) reps.forEach(rep => state.filters.rep.add(rep));

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
  els.clearSelectionBtn.disabled = !hasSelection;
  els.assignRepSelect.disabled = !hasAccounts;

  const reps = getAllAssignedReps();
  if (document.activeElement !== els.repCountInput) {
    els.repCountInput.value = reps.length || 1;
  }
}

function assignSelectionToRep() {
  const targetRep = els.assignRepSelect.value;
  const selectedIds = [...state.selection];
  if (!selectedIds.length || !targetRep) return;

  if (isRepLocked(targetRep)) {
    showToast(`"${targetRep}" is locked.`);
    return;
  }

  const previousAssignedReps = getAllAssignedReps();
  const changes = [];
  let skippedProtected = 0;
  let skippedLocked = 0;

  ensureRepColor(targetRep);

  for (const id of selectedIds) {
    const account = state.accountById.get(id);
    if (!account) continue;

    if (isAccountLocked(account)) {
      skippedLocked += 1;
      continue;
    }

    if (account.protected && account.assignedRep !== targetRep) {
      skippedProtected += 1;
      continue;
    }

    if (account.assignedRep === targetRep) continue;

    changes.push({
      id,
      from: account.assignedRep,
      to: targetRep
    });
  }

  if (!changes.length) {
    if (skippedLocked) {
      showToast(`${skippedLocked} account(s) belong to locked territories.`);
      return;
    }

    showToast(skippedProtected ? `${skippedProtected} protected account(s) were skipped.` : 'No assignment changes to make.');
    return;
  }

  applyChanges(
    changes,
    `Assigned ${changes.length} account${changes.length === 1 ? '' : 's'} to ${targetRep}`,
    previousAssignedReps
  );

  clearSelection();
}

function applyChanges(changes, label, previousAssignedReps = null) {
  const repsBefore = Array.isArray(previousAssignedReps) ? previousAssignedReps : getAllAssignedReps();
  const appliedChanges = [];

  changes.forEach(change => {
    const account = state.accountById.get(change.id);
    if (!account) return;

    if (isRepLocked(change.from) || isRepLocked(change.to) || isAccountLocked(account)) {
      return;
    }

    if (account.assignedRep === change.to) return;

    ensureRepColor(change.to);
    account.assignedRep = change.to;
    appliedChanges.push({ ...change });

    state.changeLog.push({
      timestamp: new Date().toLocaleString(),
      customerId: account.customerId,
      customerName: account.customerName,
      fromRep: change.from,
      toRep: change.to,
      protected: account.protected ? 'Yes' : 'No'
    });
  });

  if (!appliedChanges.length) {
    showToast('No eligible changes could be applied.');
    return;
  }

  state.undoStack.push({
    changes: appliedChanges,
    label
  });

  refreshAfterAssignmentBatch(appliedChanges, {
    repsBefore,
    updateSelection: true,
    territoryForce: false
  });

  updateLastAction(label);
  showToast(label);
}

function undoLastAction() {
  const action = state.undoStack.pop();
  if (!action) return;

  const repsBefore = getAllAssignedReps();

  for (const change of action.changes) {
    const account = state.accountById.get(change.id);
    if (account) account.assignedRep = change.from;
  }

  state.optimizationSummary = null;
  refreshAfterAssignmentBatch(action.changes, {
    repsBefore,
    updateSelection: true,
    territoryForce: false
  });

  updateLastAction(`Undid: ${action.label}`);
  showToast(`Undid: ${action.label}`);
}

function resetAssignments() {
  let resetCount = 0;
  const resetChanges = [];

  state.accounts.forEach(account => {
    if (account.assignedRep !== account.originalAssignedRep) {
      resetChanges.push({
        id: account._id,
        from: account.assignedRep,
        to: account.originalAssignedRep
      });
      account.assignedRep = account.originalAssignedRep;
      resetCount += 1;
    }
  });

  if (!resetCount) {
    showToast('Nothing to reset.');
    return;
  }

  state.undoStack = [];
  state.changeLog = [];
  state.repFocus = null;
  state.optimizationSummary = null;
  state.multiSearch.moved = '';
  if (els.movedSearchInput) els.movedSearchInput.value = '';

  invalidateCaches();
  refreshAfterAssignmentBatch(resetChanges, {
    repsBefore: null,
    updateSelection: true,
    territoryForce: true
  });
  fitMapToAccounts();
  updateLastAction('Reset assignments to imported values');
  showToast('Assignments reset to imported values.');
}

function optimizeRoutes() {
  if (!state.accounts.length) return;

  try {
    const targetCountRaw = parseInt(els.repCountInput.value || '1', 10);
    const minStopsRaw = parseInt(els.minStopsInput.value || '1', 10);
    const maxStopsRaw = parseInt(els.maxStopsInput.value || '999999', 10);

    const targetCount = Math.max(1, Math.min(100, targetCountRaw || 1));
    const minStops = Math.max(1, minStopsRaw || 1);
    const maxStops = Math.max(minStops, maxStopsRaw || minStops);
    const totalAccounts = state.accounts.length;

    const fixedCount = state.accounts.filter(a => a.protected || isAccountLocked(a)).length;
    const movableCount = totalAccounts - fixedCount;

    if (!Number.isFinite(targetCount) || !Number.isFinite(minStops) || !Number.isFinite(maxStops)) {
      showToast('Optimizer inputs are invalid. Check rep count and stop limits.');
      return;
    }

    if (targetCount > totalAccounts) {
      showToast(`Target rep count of ${targetCount} exceeds ${totalAccounts} total accounts.`);
      return;
    }

    if (targetCount * minStops > totalAccounts) {
      showToast(`Minimum stops too high. ${targetCount} reps × ${minStops} minimum exceeds ${totalAccounts} total accounts.`);
      return;
    }

    if (Math.ceil(totalAccounts / targetCount) > maxStops) {
      showToast(`Maximum stops too low. ${targetCount} reps cannot cover ${totalAccounts} accounts with a max of ${maxStops} per rep.`);
      return;
    }

    if (movableCount === 0) {
      showToast('All remaining accounts are locked or protected. Nothing can be optimized.');
      return;
    }

    const continuityWeight = Number(els.disruptionSlider.value) / 100;
    const geographyWeight = 1 - continuityWeight;
    const balanceMode = 'hybrid';
    const optimizerMix = getOptimizerMix();
    const beforeSummary = buildOptimizationSummary();

    const fixedAccounts = state.accounts.filter(a => a.protected || isAccountLocked(a));
    const movableAccounts = state.accounts.filter(a => !a.protected && !isAccountLocked(a));

    const currentReps = getAllAssignedReps().filter(rep => !isRepLocked(rep));
    const targetRepNames = buildTargetRepNames(targetCount, currentReps);
    const adjacency = state.neighborMap;

    if (!targetRepNames.length) {
      showToast('No unlocked reps are available for optimization.');
      return;
    }

    targetRepNames.forEach(rep => ensureRepColor(rep));

    const assignments = new Map();
    fixedAccounts.forEach(a => assignments.set(a._id, a.assignedRep));

    const assignmentCtx = createAssignmentContext(targetRepNames, assignments);
    const centroids = initializeCentroidsFast(targetRepNames, assignmentCtx);

    const orderedMovable = [...movableAccounts].sort((a, b) => {
      if (a.rank !== b.rank) return rankSortValue(a.rank) - rankSortValue(b.rank);
      if (a.overallSales !== b.overallSales) return b.overallSales - a.overallSales;
      return a.customerName.localeCompare(b.customerName);
    });

    const targetStopsPerRep = Math.max(minStops, Math.min(maxStops, totalAccounts / Math.max(1, targetRepNames.length)));
    const totalRevenuePool = state.accounts.reduce((sum, account) => sum + (account.overallSales || 0), 0);
    const targetRevenuePerRep = totalRevenuePool / Math.max(1, targetRepNames.length);

    for (let iter = 0; iter < 6; iter += 1) {
      assignmentCtx.clearMovableAssignments(orderedMovable, assignments);
      resetCentroidsFromContext(centroids, targetRepNames, assignmentCtx);

      const repStats = buildFullRepStats(targetRepNames);

      for (const rep of targetRepNames) {
        repStats.set(rep, {
          rep,
          stops: assignmentCtx.count(rep),
          revenue: assignmentCtx.revenue(rep)
        });
      }

      for (const account of orderedMovable) {
        let bestRep = null;
        let bestScore = Infinity;

        for (const rep of targetRepNames) {
          const centroid = centroids.get(rep) || averageCentroidForRep(rep, assignmentCtx);
          const compactnessScore = centroid
            ? squaredDistance(account.latitude, account.longitude, centroid.lat, centroid.lng) * geographyWeight
            : 0;

          const continuityPenalty = account.currentRep === rep ? 0 : continuityWeight;
          const existingPenalty = account.assignedRep === rep ? -0.15 : 0;

          let balancePenalty = 0;
          const stat = repStats.get(rep);

          const nextStops = (stat.stops || 0) + 1;
          const nextRevenue = (stat.revenue || 0) + (account.overallSales || 0);
          const stopDeviation = Math.abs(nextStops - targetStopsPerRep) / Math.max(1, targetStopsPerRep);
          const revenueDeviation = Math.abs(nextRevenue - targetRevenuePerRep) / Math.max(1, targetRevenuePerRep || 1);
          balancePenalty = (stopDeviation * optimizerMix.stopsPriority * 1.55) + (revenueDeviation * optimizerMix.revenuePriority * 1.15);

          const underMinBoost = stat.stops < minStops ? -2.2 : 0;
          const localPenalty = localDominancePenalty(account, rep, assignments, adjacency);

          const score =
            compactnessScore +
            continuityPenalty +
            existingPenalty +
            balancePenalty +
            localPenalty +
            underMinBoost +
            overMaxPenalty;

          if (score < bestScore) {
            bestScore = score;
            bestRep = rep;
          }
        }

        if (!bestRep) {
          bestRep = targetRepNames
            .slice()
            .sort((a, b) => {
              const ac = assignmentCtx.count(a);
              const bc = assignmentCtx.count(b);
              if (ac !== bc) return ac - bc;
              return (assignmentCtx.revenue(a) || 0) - (assignmentCtx.revenue(b) || 0);
            })[0];
        }

        assignments.set(account._id, bestRep);
        assignmentCtx.addToRep(bestRep, account);
        const stat = repStats.get(bestRep);
        stat.stops += 1;
        stat.revenue += account.overallSales || 0;
      }

      refreshCentroidsFromContext(centroids, targetRepNames, assignmentCtx);
    }

    enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    rebalanceStopTargetsStrict(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    runBorderCleanupFast(assignments, targetRepNames, continuityWeight, minStops, adjacency, assignmentCtx);
    runEnclaveCleanupFast(assignments, targetRepNames, minStops, adjacency, assignmentCtx);
    runMajoritySmoothingFast(assignments, targetRepNames, minStops, adjacency, assignmentCtx);
    enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx);
    rebalanceStopTargetsStrict(assignments, targetRepNames, minStops, maxStops, assignmentCtx);

    const finalViolations = targetRepNames.filter(rep => {
      const count = assignmentCtx.count(rep);
      return count < minStops || count > maxStops;
    });

    if (finalViolations.length) {
      showToast('Optimizer could not fully satisfy the stop limits with the current constraints. Try adjusting rep count or stop limits.');
    }

    const changes = [];
    const repsBefore = getAllAssignedReps();

    for (const account of state.accounts) {
      const nextRep = assignments.get(account._id) || account.assignedRep;
      if (nextRep !== account.assignedRep) {
        changes.push({
          id: account._id,
          from: account.assignedRep,
          to: nextRep
        });
      }
    }

    if (!changes.length) {
      showToast('Optimizer did not find a better assignment under the current rules.');
      return;
    }

    const disruptionPreset = getDisruptionPreset();
    const optimizeLabel = `Optimized routes to ${targetRepNames.length} reps with minimum ${minStops} stops`;
    applyChanges(
      changes,
      optimizeLabel,
      repsBefore
    );

    state.optimizationSummary = buildOptimizationSummary(beforeSummary, {
      weightLabel: els.optimizerBalanceValue ? els.optimizerBalanceValue.textContent : 'Balanced',
      disruptionLabel: disruptionPreset.short
    });
    renderOptimizationFeedback();
    updateLastActionWithOptimization(`${optimizeLabel} • ${disruptionPreset.short}`);

  } catch (err) {
    console.error('Optimize Routes failed:', err);
    showToast('Optimize Routes hit an error. Send me the first red error line from the browser console.');
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
    repCount: rows.length,
    movedCount,
    protectedHeld,
    minStops: Math.min(...stops),
    maxStops: Math.max(...stops),
    minRevenue: Math.min(...revenue),
    maxRevenue: Math.max(...revenue),
    avgStops: rows.reduce((sum, r) => sum + r.stops, 0) / Math.max(1, rows.length),
    stopRange,
    revenueRange,
    stopRangeDeltaPct: 0,
    revenueRangeDeltaPct: 0,
    weightLabel: meta.weightLabel || 'Balanced',
    disruptionLabel: meta.disruptionLabel || getDisruptionPreset().short
  };

  if (previousSummary) {
    summary.stopRangeDeltaPct = previousSummary.stopRange > 0
      ? ((previousSummary.stopRange - summary.stopRange) / previousSummary.stopRange) * 100
      : 0;
    summary.revenueRangeDeltaPct = previousSummary.revenueRange > 0
      ? ((previousSummary.revenueRange - summary.revenueRange) / previousSummary.revenueRange) * 100
      : 0;
  }

  return summary;
}

function updateLastActionWithOptimization(baseText) {
  const s = state.optimizationSummary;
  if (!s) {
    updateLastAction(baseText);
    return;
  }

  updateLastAction(`${baseText} • ${formatNumber(s.movedCount)} moved • Stops range ${s.minStops}-${s.maxStops} • Avg stops ${formatNumber(s.avgStops, 1)}`);
}

function createAssignmentContext(targetRepNames, assignments) {
  const ctx = { reps: new Map() };

  targetRepNames.forEach(rep => {
    ctx.reps.set(rep, {
      count: 0,
      revenue: 0,
      latSum: 0,
      lngSum: 0,
      members: new Set()
    });
  });

  for (const account of state.accounts) {
    const rep = assignments.get(account._id);
    if (!rep) continue;
    if (!ctx.reps.has(rep)) {
      ctx.reps.set(rep, {
        count: 0,
        revenue: 0,
        latSum: 0,
        lngSum: 0,
        members: new Set()
      });
    }
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
  if (!ctx.reps.has(rep)) {
    ctx.reps.set(rep, {
      count: 0,
      revenue: 0,
      latSum: 0,
      lngSum: 0,
      members: new Set()
    });
  }
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
  targetRepNames.forEach(rep => {
    centroids.set(rep, averageCentroidForRep(rep, ctx));
  });
  return centroids;
}

function resetCentroidsFromContext(centroids, targetRepNames, ctx) {
  targetRepNames.forEach(rep => {
    centroids.set(rep, averageCentroidForRep(rep, ctx));
  });
}

function refreshCentroidsFromContext(centroids, targetRepNames, ctx) {
  targetRepNames.forEach(rep => {
    centroids.set(rep, averageCentroidForRep(rep, ctx));
  });
}

function averageCentroidForRep(rep, ctx) {
  const entry = ctx.reps.get(rep);
  if (!entry || !entry.count) return null;
  return {
    lat: entry.latSum / entry.count,
    lng: entry.lngSum / entry.count
  };
}

function buildTargetRepNames(targetCount, currentReps) {
  const reps = [...currentReps];
  while (reps.length < targetCount) reps.push(`Rep ${reps.length + 1}`);
  return reps.slice(0, targetCount);
}

function buildFullRepStats(targetRepNames) {
  const map = new Map();
  targetRepNames.forEach(rep => {
    map.set(rep, { rep, stops: 0, revenue: 0 });
  });
  return map;
}

function localDominancePenalty(account, rep, assignments, adjacency) {
  const neighbors = adjacency.get(account._id);
  if (!neighbors || !neighbors.size) return 0;

  let same = 0;
  let diff = 0;

  neighbors.forEach(id => {
    const neighborRep = assignments.get(id) || state.accountById.get(id)?.assignedRep;
    if (!neighborRep) return;
    if (neighborRep === rep) same += 1;
    else diff += 1;
  });

  return diff > same ? 0.3 : 0;
}

function enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx) {
  let guard = 0;

  while (guard < 2000) {
    guard += 1;
    const under = targetRepNames.find(rep => ctx.count(rep) < minStops);
    if (!under) break;

    const donor = targetRepNames
      .filter(rep => ctx.count(rep) > minStops)
      .sort((a, b) => ctx.count(b) - ctx.count(a))[0];

    if (!donor) break;

    const underCentroid = averageCentroidForRep(under, ctx);
    const candidate = [...ctx.members(donor)]
      .map(id => state.accountById.get(id))
      .filter(Boolean)
      .sort((a, b) => {
        const ad = underCentroid ? squaredDistance(a.latitude, a.longitude, underCentroid.lat, underCentroid.lng) : 0;
        const bd = underCentroid ? squaredDistance(b.latitude, b.longitude, underCentroid.lat, underCentroid.lng) : 0;
        return ad - bd;
      })[0];

    if (!candidate) break;

    ctx.removeFromRep(donor, candidate);
    ctx.addToRep(under, candidate);
    assignments.set(candidate._id, under);
  }
}

function enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx) {
  let guard = 0;

  while (guard < 2000) {
    guard += 1;
    const over = targetRepNames.find(rep => ctx.count(rep) > maxStops);
    if (!over) break;

    const receiver = targetRepNames
      .filter(rep => ctx.count(rep) < maxStops)
      .sort((a, b) => ctx.count(a) - ctx.count(b))[0];

    if (!receiver) break;

    const receiverCentroid = averageCentroidForRep(receiver, ctx);
    const candidate = [...ctx.members(over)]
      .map(id => state.accountById.get(id))
      .filter(Boolean)
      .sort((a, b) => {
        const ad = receiverCentroid ? squaredDistance(a.latitude, a.longitude, receiverCentroid.lat, receiverCentroid.lng) : 0;
        const bd = receiverCentroid ? squaredDistance(b.latitude, b.longitude, receiverCentroid.lat, receiverCentroid.lng) : 0;
        return ad - bd;
      })[0];

    if (!candidate) break;

    ctx.removeFromRep(over, candidate);
    ctx.addToRep(receiver, candidate);
    assignments.set(candidate._id, receiver);
  }
}

function rebalanceStopTargetsStrict(assignments, targetRepNames, minStops, maxStops, ctx) {
  enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx);
  enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx);
}

function runBorderCleanupFast(assignments, targetRepNames, continuityWeight, minStops, adjacency, ctx) {
  for (const account of state.accounts) {
    if (account.protected || isAccountLocked(account)) continue;

    const currentRep = assignments.get(account._id) || account.assignedRep;
    const neighbors = adjacency.get(account._id);
    if (!neighbors || !neighbors.size) continue;

    const counts = new Map();
    neighbors.forEach(id => {
      const rep = assignments.get(id) || state.accountById.get(id)?.assignedRep;
      if (!rep) return;
      counts.set(rep, (counts.get(rep) || 0) + 1);
    });

    const bestNeighborRep = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
    if (!bestNeighborRep || bestNeighborRep === currentRep) continue;
    if (ctx.count(currentRep) <= minStops) continue;
    if (isRepLocked(bestNeighborRep)) continue;

    ctx.removeFromRep(currentRep, account);
    ctx.addToRep(bestNeighborRep, account);
    assignments.set(account._id, bestNeighborRep);
  }
}

function runEnclaveCleanupFast(assignments, targetRepNames, minStops, adjacency, ctx) {
  runBorderCleanupFast(assignments, targetRepNames, 0, minStops, adjacency, ctx);
}

function runMajoritySmoothingFast(assignments, targetRepNames, minStops, adjacency, ctx) {
  runBorderCleanupFast(assignments, targetRepNames, 0, minStops, adjacency, ctx);
}

async function exportWorkbook() {
  if (!state.accounts.length) {
    showToast('Nothing to export.');
    return;
  }

  const workbook = new ExcelJS.Workbook();

  const mainSheet = workbook.addWorksheet(state.currentSheetName || 'Sheet1');
  const exportRows = state.accounts.map(account => {
    const row = { ...(account.sourceRow || {}) };
    const assignedHeader = state.currentHeaderMap.assignedRep || 'New Rep';
    row[assignedHeader] = account.assignedRep;
    return row;
  });

  if (exportRows.length) {
    const keys = Object.keys(exportRows[0]);
    mainSheet.columns = keys.map(key => ({
      header: key,
      key,
      width: guessColumnWidth(key, exportRows)
    }));
    exportRows.forEach(row => mainSheet.addRow(row));
    styleHeaderRow(mainSheet.getRow(1));
    styleDataRows(mainSheet, 2, mainSheet.rowCount);
  }

  const movedSheet = workbook.addWorksheet('Moved Accounts');
  const movedRows = state.accounts
    .filter(a => a.assignedRep !== a.originalAssignedRep)
    .map(a => ({
      Customer_ID: a.customerId,
      Customer_Name: a.customerName,
      Original_Assigned_Rep: a.originalAssignedRep,
      Assigned_Rep: a.assignedRep,
      Current_Rep: a.currentRep,
      Revenue: round2(a.overallSales),
      Rank: a.rank,
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
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD3DFEB' } },
      bottom: { style: 'thin', color: { argb: 'FFC3D3E2' } }
    };
  });
}

function styleDataRows(worksheet, fromRow, toRow) {
  for (let r = fromRow; r <= toRow; r += 1) {
    const row = worksheet.getRow(r);
    row.height = 18;
    row.eachCell(cell => {
      cell.font = { name: 'Tw Cen MT', size: 9, color: { argb: 'FF29415B' } };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFE6EDF5' } }
      };
    });
  }
}

function guessColumnWidth(key, rows) {
  let maxLen = String(key || '').length;
  rows.slice(0, 250).forEach(row => {
    const value = row[key];
    const len = String(value == null ? '' : value).length;
    if (len > maxLen) maxLen = len;
  });
  return Math.max(10, Math.min(maxLen + 2, 42));
}

function downloadArrayBufferAsFile(buffer, filename) {
  const blob = new Blob(
    [buffer],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
  );

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function zoomToRep(rep) {
  const members = state.accounts.filter(a => a.assignedRep === rep);
  const points = members.map(a => [a.latitude, a.longitude]);

  if (!points.length) return;

  if (points.length === 1) {
    state.map.setView(points[0], 12);
    return;
  }

  const bounds = L.latLngBounds(points);
  const spanLat = Math.abs(bounds.getNorth() - bounds.getSouth());
  const spanLng = Math.abs(bounds.getEast() - bounds.getWest());

  if (spanLat < 0.03 && spanLng < 0.03) {
    state.map.setView(bounds.getCenter(), 11);
    return;
  }

  state.map.fitBounds(bounds, { padding: [35, 35], maxZoom: 11 });
}

function fitMapToAccounts() {
  if (!state.accounts.length) return;
  const latlngs = state.accounts.map(a => [a.latitude, a.longitude]);
  state.map.fitBounds(latlngs, { padding: [25, 25] });
}

function toggleTheme() {
  if (els.themeToggleCheck.checked && state.theme === 'light') {
    state.map.removeLayer(state.lightLayer);
    state.darkLayer.addTo(state.map);
    state.theme = 'dark';
  } else if (!els.themeToggleCheck.checked && state.theme === 'dark') {
    state.map.removeLayer(state.darkLayer);
    state.lightLayer.addTo(state.map);
    state.theme = 'light';
  }
}

function buildRepColors() {
  const previous = new Map(state.repColors);
  const reps = getAllKnownReps();
  const usedColors = new Set();

  state.repColors = new Map();

  reps.forEach(rep => {
    const existing = previous.get(rep);
    if (existing) {
      state.repColors.set(rep, existing);
      usedColors.add(existing);
    }
  });

  reps.forEach(rep => {
    if (state.repColors.has(rep)) return;
    const nextColor = COLOR_PALETTE.find(c => !usedColors.has(c)) || COLOR_PALETTE[state.repColors.size % COLOR_PALETTE.length];
    state.repColors.set(rep, nextColor);
    usedColors.add(nextColor);
  });
}

function ensureRepColor(rep) {
  if (!rep) return;
  if (state.repColors.has(rep)) return;

  const usedColors = new Set(state.repColors.values());
  const nextColor = COLOR_PALETTE.find(c => !usedColors.has(c)) || COLOR_PALETTE[state.repColors.size % COLOR_PALETTE.length];
  state.repColors.set(rep, nextColor);
}

function getRepColor(rep) {
  ensureRepColor(rep);
  return state.repColors.get(rep) || '#64748b';
}

function getAllAssignedReps() {
  const set = new Set();
  state.accounts.forEach(a => {
    if (a.assignedRep) set.add(a.assignedRep);
  });
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

function fillSimpleSelect(selectEl, values, selectedValue, labelFn = v => v, placeholder = '') {
  if (!selectEl) return;

  const options = [];

  if (placeholder) {
    options.push(`<option value="">${escapeHtml(placeholder)}</option>`);
  }

  values.forEach(v => {
    options.push(`<option value="${escapeHtmlAttr(v)}">${escapeHtml(labelFn(v))}</option>`);
  });

  selectEl.innerHTML = options.join('');

  if (selectedValue != null && values.includes(selectedValue)) {
    selectEl.value = selectedValue;
  } else if (placeholder) {
    selectEl.value = '';
  }
}

function updateLastAction(text) {
  state.lastAction = text;
  els.lastAction.textContent = text;
}

function showToast(message) {
  if (!els.toast) return;
  els.toast.textContent = message;
  els.toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toast.classList.remove('show'), 2200);
}

function getDistinctValues(arr, fn) {
  const set = new Set();
  arr.forEach(item => {
    const v = fn(item);
    if (safeString(v)) set.add(v);
  });
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

  if (!raw) {
    if (rank === 'A') return 4;
    if (rank === 'B') return 2;
    if (rank === 'C') return 1;
    return 0.33;
  }

  const normalized = raw.toLowerCase();
  const n = toNumber(raw);
  if (Number.isFinite(n) && n >= 0) return n;

  if (normalized.includes('weekly')) return 4;
  if (normalized.includes('biweekly') || normalized.includes('every other')) return 2;
  if (normalized.includes('monthly')) return 1;
  if (normalized.includes('quarter')) return 0.33;

  if (rank === 'A') return 4;
  if (rank === 'B') return 2;
  if (rank === 'C') return 1;
  return 0.33;
}

function rankSortValue(rank) {
  return { A: 0, B: 1, C: 2, D: 3 }[rank] ?? 9;
}

function toCamel(id) {
  return id.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
}


function clampByte(value) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function hexToRgb(hex) {
  const value = String(hex || '').trim().replace('#', '');
  const normalized = value.length === 3
    ? value.split('').map(ch => ch + ch).join('')
    : value.padEnd(6, '0').slice(0, 6);

  const int = Number.parseInt(normalized, 16);
  if (Number.isNaN(int)) return { r: 64, g: 99, b: 160 };

  return {
    r: (int >> 16) & 255,
    g: (int >> 8) & 255,
    b: int & 255
  };
}

function rgbToHex(r, g, b) {
  return `#${[r, g, b].map(v => clampByte(v).toString(16).padStart(2, '0')).join('')}`;
}

function mixHex(colorA, colorB, weight = 0.5) {
  const a = hexToRgb(colorA);
  const b = hexToRgb(colorB);
  const t = Math.max(0, Math.min(1, Number(weight) || 0));
  return rgbToHex(
    a.r + (b.r - a.r) * t,
    a.g + (b.g - a.g) * t,
    a.b + (b.b - a.b) * t
  );
}

function getTerritoryStrokeColor(rep) {
  return mixHex(getRepColor(rep), '#24384f', 0.42);
}

function getTerritoryFillColor(rep) {
  return mixHex(getRepColor(rep), '#ffffff', 0.58);
}

function getTerritoryLabelColor(rep) {
  return mixHex(getRepColor(rep), '#22364f', 0.3);
}

function cleanHeader(value) {
  return safeString(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function safeString(value) {
  return value == null ? '' : String(value).trim();
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
  const raw = String(value ?? '').replace(/[$,%\s,]/g, '').trim();
  if (!raw) return NaN;
  const n = Number(raw);
  return Number.isFinite(n) ? n : NaN;
}

function toBoolean(value) {
  const v = safeString(value).toLowerCase();
  return ['true', 'yes', 'y', '1', 'protected', 'locked'].includes(v);
}

function normalizeRank(value) {
  const raw = safeString(value).toUpperCase();
  return ['A', 'B', 'C', 'D'].includes(raw) ? raw : 'C';
}

function formatCurrency(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0
  }).format(value || 0);
}

function formatNumber(value, digits = 0) {
  return Number(value || 0).toLocaleString('en-US', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function round2(v) {
  return Math.round((v || 0) * 100) / 100;
}

function round6(v) {
  return Math.round((v || 0) * 1000000) / 1000000;
}

function squaredDistance(lat1, lng1, lat2, lng2) {
  const dx = lng1 - lng2;
  const dy = lat1 - lat2;
  return dx * dx + dy * dy;
}

function escapeHtml(text) {
  return String(text ?? '').replace(/[&<>"']/g, m => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[m]));
}

function escapeHtmlAttr(text) {
  return escapeHtml(text).replace(/"/g, '&quot;');
}
