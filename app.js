const COLOR_PALETTE = [
  '#1f77b4','#d62728','#2ca02c','#9467bd','#ff7f0e','#17becf','#8c564b','#e377c2','#7f7f7f','#bcbd22',
  '#0b7285','#c92a2a','#2b8a3e','#5f3dc4','#e67700','#087f5b','#364fc7','#a61e4d','#495057','#2f9e44',
  '#f03e3e','#3b5bdb','#e8590c','#1098ad','#9c36b5','#5c940d','#d9480f','#1864ab','#c2255c','#12b886'
];

const COLUMN_ALIASES = {
  latitude: ['latitude','lat','y','geo_lat','customer_latitude'],
  longitude: ['longitude','lng','lon','x','geo_longitude','customer_longitude'],
  customerId: ['cust id','customer id','customerid','id','account id','acct id'],
  customerName: ['company','customer name','name','account name','cust name'],
  address: ['address','street address','addr','full address'],
  zip: ['zip','zip code','zipcode','postal code'],
  chain: ['chain','chain name'],
  segment: ['segment','customer segment'],
  premise: ['premise','premise type','on/off premise','premise class'],
  currentRep: ['current rep','rep','sales rep','territory rep','owner rep'],
  assignedRep: ['assigned rep','new rep','territory','route','assigned territory'],
  overallSales: ['overall sales','sales','total sales','revenue','$ revenue','$ vol sept - feb','overall revenue'],
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
  accounts: [],
  accountById: new Map(),
  neighborMap: new Map(),
  markerById: new Map(),
  selection: new Set(),
  undoStack: [],
  changeLog: [],
  repColors: new Map(),
  repFocus: null,
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
  openMultiKey: null,
  lockedReps: {}
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

  requestAnimationFrame(() => {
    if (state.map) state.map.invalidateSize();
  });
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
    'routes-table-wrap','moved-search-input','upload-status-pill','upload-status-icon','upload-status-text','upload-status-panel','upload-status-body'
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

  if (els.uploadStatusPill) {
    els.uploadStatusPill.addEventListener('click', e => {
      e.stopPropagation();
      toggleUploadStatusPanel();
    });
  }

  els.themeToggleCheck.addEventListener('change', toggleTheme);
  els.dimOthersCheckbox.addEventListener('change', refreshUI);
  els.showTerritoryCheckbox.addEventListener('change', refreshTerritories);

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

  els.disruptionSlider.addEventListener('input', () => {
    els.disruptionValue.textContent = els.disruptionSlider.value;
  });

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
  state.map.on('zoomend', refreshTerritories);
  state.map.on('moveend', refreshTerritories);
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
      });
    }

    if (searchInput) {
      searchInput.addEventListener('input', e => {
        state.multiSearch[key] = e.target.value || '';
        renderMultiFilterOptions(key);
      });

      searchInput.addEventListener('click', e => e.stopPropagation());
    }
  });
}

function handleDocumentClickForPanels(event) {
  const multiTrigger = event.target.closest('.multi-trigger');
  const multiPanel = event.target.closest('.multi-panel');
  const uploadPill = event.target.closest('#upload-status-pill');
  const uploadPanel = event.target.closest('#upload-status-panel');

  if (!multiTrigger && !multiPanel) closeAllMultiPanels();
  if (!uploadPill && !uploadPanel) closeUploadStatusPanel();
}

function toggleMultiPanel(key) {
  if (!key) return;

  if (state.openMultiKey === key) {
    closeAllMultiPanels();
    return;
  }

  closeAllMultiPanels();
  state.openMultiKey = key;

  const panel = document.querySelector(`[data-multi-panel="${key}"]`);
  if (!panel) return;

  panel.hidden = false;
  renderMultiFilterOptions(key);
  positionMultiPanel(key);

  const input = panel.querySelector(`[data-search="${key}"]`);
  if (input) setTimeout(() => input.focus(), 0);
}

function closeAllMultiPanels() {
  state.openMultiKey = null;
  document.querySelectorAll('.multi-panel').forEach(panel => {
    panel.hidden = true;
    panel.style.top = '-9999px';
    panel.style.left = '-9999px';
  });
}

function positionMultiPanel(key) {
  const trigger = document.querySelector(`[data-multi-trigger="${key}"]`);
  const panel = document.querySelector(`[data-multi-panel="${key}"]`);
  if (!trigger || !panel) return;

  const rect = trigger.getBoundingClientRect();
  const panelWidth = Math.min(320, window.innerWidth - 20);
  let left = rect.left;
  let top = rect.bottom + 6;

  if (left + panelWidth > window.innerWidth - 10) left = window.innerWidth - panelWidth - 10;
  if (left < 10) left = 10;

  panel.style.width = `${panelWidth}px`;
  panel.style.left = `${left}px`;
  panel.style.top = `${top}px`;
}

function toggleUploadStatusPanel() {
  if (!els.uploadStatusPanel) return;

  if (els.uploadStatusPanel.hidden) {
    els.uploadStatusPanel.hidden = false;
    positionUploadStatusPanel();
  } else {
    closeUploadStatusPanel();
  }
}

function closeUploadStatusPanel() {
  if (!els.uploadStatusPanel) return;
  els.uploadStatusPanel.hidden = true;
  els.uploadStatusPanel.style.top = '-9999px';
  els.uploadStatusPanel.style.left = '-9999px';
}

function positionUploadStatusPanel() {
  if (!els.uploadStatusPill || !els.uploadStatusPanel || els.uploadStatusPanel.hidden) return;

  const rect = els.uploadStatusPill.getBoundingClientRect();
  const panelWidth = Math.min(360, window.innerWidth - 20);
  let left = rect.left;
  let top = rect.bottom + 6;

  if (left + panelWidth > window.innerWidth - 10) left = window.innerWidth - panelWidth - 10;
  if (left < 10) left = 10;

  els.uploadStatusPanel.style.width = `${panelWidth}px`;
  els.uploadStatusPanel.style.left = `${left}px`;
  els.uploadStatusPanel.style.top = `${top}px`;
}

function renderUploadStatus() {
  if (!els.uploadStatusPill) return;

  const status = state.uploadStatus || { level:'neutral', text:'No file loaded' };
  const klass = {
    neutral: 'upload-status-neutral',
    good: 'upload-status-good',
    warning: 'upload-status-warning',
    bad: 'upload-status-bad'
  }[status.level] || 'upload-status-neutral';

  els.uploadStatusPill.className = `upload-status-pill ${klass}`;
  els.uploadStatusText.textContent = status.text || 'No file loaded';

  if (els.uploadStatusIcon) {
    els.uploadStatusIcon.textContent =
      status.level === 'good' ? '✓' :
      status.level === 'warning' ? '!' :
      status.level === 'bad' ? '×' : 'i';
  }

  const s = state.importSummary || {};
  if (els.uploadStatusBody) {
    const issueLines = [];
    if (s.skippedNoCoords) issueLines.push(`<li>${s.skippedNoCoords.toLocaleString()} row(s) skipped due to missing coordinates</li>`);
    if (s.duplicateCustomerIds) issueLines.push(`<li>${s.duplicateCustomerIds.toLocaleString()} duplicate customer ID(s) adjusted</li>`);
    if (s.missingCurrentRep) issueLines.push(`<li>${s.missingCurrentRep.toLocaleString()} row(s) missing current rep</li>`);
    if (s.missingAssignedRep) issueLines.push(`<li>${s.missingAssignedRep.toLocaleString()} row(s) missing assigned rep</li>`);
    if ((s.unmappedFields || []).filter(key => isWarningField(key)).length) {
      issueLines.push(`<li>Some optional columns were not mapped cleanly</li>`);
    }

    const summaryText =
      status.level === 'bad' && !s.loadedRows
        ? 'No valid rows are loaded right now.'
        : `${Number(s.loadedRows || 0).toLocaleString()} valid row(s) are currently loaded.`;

    els.uploadStatusBody.innerHTML = `
      <div class="upload-diag-summary">${summaryText}</div>

      <div class="upload-diag-grid">
        <div class="upload-diag-label">Source rows</div>
        <div class="upload-diag-value">${Number(s.sourceRows || 0).toLocaleString()}</div>

        <div class="upload-diag-label">Loaded rows</div>
        <div class="upload-diag-value">${Number(s.loadedRows || 0).toLocaleString()}</div>

        <div class="upload-diag-label">Skipped for missing coords</div>
        <div class="upload-diag-value">${Number(s.skippedNoCoords || 0).toLocaleString()}</div>

        <div class="upload-diag-label">Duplicate customer IDs</div>
        <div class="upload-diag-value">${Number(s.duplicateCustomerIds || 0).toLocaleString()}</div>

        <div class="upload-diag-label">Missing current rep</div>
        <div class="upload-diag-value">${Number(s.missingCurrentRep || 0).toLocaleString()}</div>

        <div class="upload-diag-label">Missing assigned rep</div>
        <div class="upload-diag-value">${Number(s.missingAssignedRep || 0).toLocaleString()}</div>
      </div>

      ${issueLines.length ? `<ul class="upload-diag-list">${issueLines.join('')}</ul>` : ''}
    `;
  }
}

function isWarningField(key) {
  return !['latitude','longitude','customerId','customerName','currentRep','assignedRep'].includes(String(key || ''));
}

function onFileChosen(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  state.loadedFileName = file.name || 'territory_export_updated.xlsx';

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const result = e.target.result;
      const workbook = XLSX.read(result, { type: 'binary' });
      state.workbook = workbook;
      state.workbookSheets = {};

      workbook.SheetNames.forEach(name => {
        const ws = workbook.Sheets[name];
        state.workbookSheets[name] = XLSX.utils.sheet_to_json(ws, { defval: '' });
      });

      populateSheetSelect(workbook.SheetNames);
      showToast(`Loaded file: ${state.loadedFileName}`);
    } catch (err) {
      console.error('File load failed:', err);
      showToast('Could not read that file. Please try another workbook or CSV.');
    }
  };

  reader.readAsBinaryString(file);
}

function populateSheetSelect(sheetNames) {
  els.sheetSelect.innerHTML = '<option value="">Sheet</option>';

  sheetNames.forEach((sheetName, idx) => {
    const option = document.createElement('option');
    option.value = sheetName;
    option.textContent = sheetName;
    if (idx === 0) option.selected = true;
    els.sheetSelect.appendChild(option);
  });

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

  const rows = state.workbookSheets[sheetName];
  const normalizedResult = normalizeRows(rows);
  const normalized = normalizedResult.mapped;

  state.importSummary = normalizedResult.summary;
  state.optimizationSummary = null;
  closeUploadStatusPanel();

  if (!normalized.length) {
    state.accounts = [];
    state.accountById = new Map();
    state.neighborMap = new Map();
    state.selection.clear();
    state.undoStack = [];
    state.changeLog = [];
    state.repFocus = null;
    state.lockedReps = {};
    state.multiSearch.moved = '';
    if (els.movedSearchInput) els.movedSearchInput.value = '';

    setUploadStatus('bad', 'No valid rows found');
    refreshUI(true);
    showToast('No valid rows found in that sheet.');
    return;
  }

  state.accounts = normalized;
  state.accountById = new Map(state.accounts.map(a => [a._id, a]));
  state.neighborMap = buildNeighborMap(state.accounts);
  state.selection.clear();
  state.undoStack = [];
  state.changeLog = [];
  state.repFocus = null;
  state.lockedReps = {};
  state.multiSearch.moved = '';
  if (els.movedSearchInput) els.movedSearchInput.value = '';

  buildRepColors();
  seedFiltersFromData();
  updateUploadStatusFromSummary();
  refreshUI(true);
  fitMapToAccounts();

  if (!els.repCountInput.dataset.userTouched) {
    els.repCountInput.value = getAllAssignedReps().length || 1;
  }

  requestAnimationFrame(() => {
    if (state.map) state.map.invalidateSize();
  });
}

function normalizeRows(rows) {
  const normalized = [];
  const summary = {
    sourceRows: rows.length,
    loadedRows: 0,
    skippedNoCoords: 0,
    duplicateCustomerIds: 0,
    missingCurrentRep: 0,
    missingAssignedRep: 0,
    unmappedFields: []
  };

  if (!rows.length) return { mapped: normalized, summary };

  const headerMap = mapHeaders(rows[0]);
  summary.unmappedFields = Object.keys(COLUMN_ALIASES).filter(key => !headerMap[key]);

  const usedIds = new Set();

  rows.forEach((row, idx) => {
    const latitude = toNumber(getMappedValue(row, headerMap.latitude));
    const longitude = toNumber(getMappedValue(row, headerMap.longitude));

    if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
      summary.skippedNoCoords += 1;
      return;
    }

    let customerId = String(getMappedValue(row, headerMap.customerId) || '').trim();
    if (!customerId) customerId = `ROW-${idx + 1}`;

    if (usedIds.has(customerId)) {
      summary.duplicateCustomerIds += 1;
      customerId = `${customerId}-${idx + 1}`;
    }
    usedIds.add(customerId);

    const currentRep = cleanRepName(getMappedValue(row, headerMap.currentRep)) || 'Unassigned';
    const assignedRep = cleanRepName(getMappedValue(row, headerMap.assignedRep)) || currentRep || 'Unassigned';

    if (!cleanRepName(getMappedValue(row, headerMap.currentRep))) summary.missingCurrentRep += 1;
    if (!cleanRepName(getMappedValue(row, headerMap.assignedRep))) summary.missingAssignedRep += 1;

    const account = {
      _id: customerId,
      customerId,
      customerName: String(getMappedValue(row, headerMap.customerName) || customerId).trim(),
      address: String(getMappedValue(row, headerMap.address) || '').trim(),
      zip: String(getMappedValue(row, headerMap.zip) || '').trim(),
      chain: String(getMappedValue(row, headerMap.chain) || 'Unknown').trim(),
      segment: String(getMappedValue(row, headerMap.segment) || 'Unknown').trim(),
      premise: String(getMappedValue(row, headerMap.premise) || 'Unknown').trim(),
      currentRep,
      assignedRep,
      originalAssignedRep: assignedRep,
      overallSales: toNumber(getMappedValue(row, headerMap.overallSales)),
      rank: normalizeRank(getMappedValue(row, headerMap.rank)),
      cadence4w: toNumber(getMappedValue(row, headerMap.cadence4w)),
      protected: toBoolean(getMappedValue(row, headerMap.protected)),
      latitude,
      longitude,
      _raw: row
    };

    normalized.push(account);
  });

  summary.loadedRows = normalized.length;

  return { mapped: normalized, summary };
}

function mapHeaders(firstRow) {
  const keys = Object.keys(firstRow || {});
  const normalizedKeys = keys.map(k => ({
    original: k,
    normalized: normalizeHeader(k)
  }));

  const headerMap = {};

  Object.entries(COLUMN_ALIASES).forEach(([field, aliases]) => {
    const match = normalizedKeys.find(k => aliases.includes(k.normalized));
    if (match) headerMap[field] = match.original;
  });

  return headerMap;
}

function getMappedValue(row, key) {
  if (!row || !key) return '';
  return row[key];
}

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s\-_\/\\]+/g, ' ')
    .replace(/[^\w$ ]+/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanRepName(value) {
  return String(value || '').trim();
}

function normalizeRank(value) {
  const raw = String(value || '').trim().toUpperCase();
  if (!raw) return 'D';
  if (['A','B','C','D'].includes(raw)) return raw;
  return 'D';
}

function toBoolean(value) {
  const raw = String(value || '').trim().toLowerCase();
  return ['yes','y','true','1','protected','lock','locked'].includes(raw);
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  const cleaned = String(value)
    .replace(/[$,%\s,]/g, '')
    .trim();

  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function buildNeighborMap(accounts) {
  const map = new Map();
  const points = accounts.map(a => turf.point([a.longitude, a.latitude], { id: a._id }));
  const featureCollection = turf.featureCollection(points);

  accounts.forEach(account => {
    const nearest = turf.nearestPoint(accountPoint(account), featureCollection);
    map.set(account._id, new Set([nearest.properties.id]));
  });

  const sorted = [...accounts].sort((a,b) => a.latitude - b.latitude || a.longitude - b.longitude);

  for (let i = 0; i < sorted.length; i += 1) {
    const a = sorted[i];
    if (!map.has(a._id)) map.set(a._id, new Set());

    for (let j = Math.max(0, i - 8); j < Math.min(sorted.length, i + 9); j += 1) {
      if (i === j) continue;
      const b = sorted[j];
      if (squaredDistance(a.latitude, a.longitude, b.latitude, b.longitude) <= 0.09) {
        map.get(a._id).add(b._id);
      }
    }
  }

  return map;
}

function accountPoint(account) {
  return turf.point([account.longitude, account.latitude], { id: account._id });
}

function buildRepColors() {
  const reps = getAllKnownReps();
  reps.forEach(rep => ensureRepColor(rep));
}

function ensureRepColor(rep) {
  if (!rep) return '#7f8c9b';
  if (!state.repColors.has(rep)) {
    const next = COLOR_PALETTE[state.repColors.size % COLOR_PALETTE.length];
    state.repColors.set(rep, next);
  }
  return state.repColors.get(rep);
}

function getRepColor(rep) {
  return ensureRepColor(rep);
}

function getAllAssignedReps() {
  return [...new Set(state.accounts.map(a => a.assignedRep).filter(Boolean))].sort((a,b) =>
    a.localeCompare(b, undefined, { numeric:true })
  );
}

function getAllKnownReps() {
  return [...new Set(
    state.accounts
      .flatMap(a => [a.currentRep, a.assignedRep, a.originalAssignedRep])
      .filter(Boolean)
  )].sort((a,b) => a.localeCompare(b, undefined, { numeric:true }));
}

function seedFiltersFromData() {
  state.filters.rep = new Set(getAllAssignedReps());
  state.filters.rank = new Set(getDistinctValues(state.accounts, a => a.rank));
  state.filters.chain = new Set(getDistinctValues(state.accounts, a => a.chain));
  state.filters.segment = new Set(getDistinctValues(state.accounts, a => a.segment));
  state.filters.premise = 'ALL';
  state.filters.protected = 'ALL';
  state.filters.moved = 'ALL';
  state.multiSearch.rep = '';
  state.multiSearch.rank = '';
  state.multiSearch.chain = '';
  state.multiSearch.segment = '';
  renderMultiFilterOptions();
  fillSimpleSelect(els.premiseFilter, ['ALL', ...getDistinctValues(state.accounts, a => a.premise)], 'ALL', v => v === 'ALL' ? 'All premises' : v);
}

function getDistinctValues(list, accessor) {
  return [...new Set(list.map(accessor).filter(v => String(v || '').trim()))]
    .sort((a,b) => String(a).localeCompare(String(b), undefined, { numeric:true }));
}

function fillSimpleSelect(select, values, selectedValue, labelFn = v => v) {
  if (!select) return;
  select.innerHTML = '';

  values.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = labelFn(value);
    if (value === selectedValue) option.selected = true;
    select.appendChild(option);
  });
}

function renderMultiFilterOptions(specificKey = null) {
  ['rep','rank','chain','segment'].forEach(key => {
    if (specificKey && specificKey !== key) return;

    const container = els[`${key}FilterOptions`];
    if (!container) return;

    const values = getMultiValuesForKey(key);
    const search = (state.multiSearch[key] || '').toLowerCase().trim();
    const selected = state.filters[key] || new Set();

    const filteredValues = search
      ? values.filter(v => String(v).toLowerCase().includes(search))
      : values;

    container.innerHTML = filteredValues.length
      ? filteredValues.map(value => `
          <label class="multi-option">
            <input type="checkbox" data-multi-value="${escapeHtmlAttr(value)}" ${selected.has(value) ? 'checked' : ''} />
            <span>${escapeHtml(value)}</span>
          </label>
        `).join('')
      : '<div class="empty">No matches</div>';

    container.querySelectorAll('input[data-multi-value]').forEach(input => {
      input.addEventListener('click', e => e.stopPropagation());
      input.addEventListener('change', e => {
        const value = e.target.getAttribute('data-multi-value') || '';
        if (selected.has(value)) selected.delete(value);
        else selected.add(value);

        updateMultiSummary(key);
        refreshUI();
      });
    });

    updateMultiSummary(key);
  });
}

function updateMultiSummary(key) {
  const values = getMultiValuesForKey(key);
  const selected = state.filters[key] || new Set();
  const el = els[`${key}FilterSummary`];
  if (!el) return;

  if (!values.length || selected.size === values.length) {
    el.textContent =
      key === 'rep' ? 'All reps' :
      key === 'rank' ? 'All ranks' :
      key === 'chain' ? 'All chains' :
      'All segments';
    return;
  }

  if (selected.size === 0) {
    el.textContent = 'None';
    return;
  }

  if (selected.size <= 2) {
    el.textContent = [...selected].join(', ');
    return;
  }

  el.textContent = `${selected.size} selected`;
}

function toggleSelectAllMulti(key) {
  const values = getMultiValuesForKey(key);
  const selected = state.filters[key];
  const isAllSelected = selected.size === values.length && values.every(v => selected.has(v));

  selected.clear();
  if (!isAllSelected) values.forEach(v => selected.add(v));

  renderMultiFilterOptions();
  refreshUI();
}

function getMultiValuesForKey(key) {
  if (key === 'rep') return getAllAssignedReps();
  if (key === 'rank') return getDistinctValues(state.accounts, a => a.rank);
  if (key === 'chain') return getDistinctValues(state.accounts, a => a.chain);
  if (key === 'segment') return getDistinctValues(state.accounts, a => a.segment);
  return [];
}

function refreshUI(rebuildMap = false) {
  syncControlState();
  renderRepControls();
  renderUploadStatus();
  if (rebuildMap) rebuildMarkers();
  refreshMarkerStyles();
  renderRepTable();
  renderSelectionPreview();
  renderSummary();
  renderMovedReview();
  refreshTerritories();
  if (state.openMultiKey) positionMultiPanel(state.openMultiKey);
}

function syncControlState() {
  const hasData = state.accounts.length > 0;
  els.assignBtn.disabled = !hasData || state.selection.size === 0;
  els.undoBtn.disabled = !hasData || state.undoStack.length === 0;
  els.resetBtn.disabled = !hasData;
  els.optimizeBtn.disabled = !hasData;
  els.exportBtn.disabled = !hasData;
  els.clearSelectionBtn.disabled = !hasData || state.selection.size === 0;
  els.assignRepSelect.disabled = !hasData;
  els.repCountInput.disabled = !hasData;
  els.minStopsInput.disabled = !hasData;
  els.maxStopsInput.disabled = !hasData;
  els.balanceMode.disabled = !hasData;
  els.disruptionSlider.disabled = !hasData;
  els.premiseFilter.disabled = !hasData;
  els.protectedFilter.disabled = !hasData;
  els.movedFilter.disabled = !hasData;
  els.dimOthersCheckbox.disabled = !hasData;
  els.showTerritoryCheckbox.disabled = !hasData;
  if (els.movedSearchInput) els.movedSearchInput.disabled = !hasData;
}

function renderRepControls() {
  const reps = getAllAssignedReps();
  fillSimpleSelect(els.assignRepSelect, reps, reps[0] || '');

  if (!els.repCountInput.dataset.userTouched) {
    els.repCountInput.value = Math.max(1, reps.length || 1);
  }

  els.repCountInput.oninput = () => {
    els.repCountInput.dataset.userTouched = '1';
  };
}

function setUploadStatus(level, text) {
  state.uploadStatus = { level, text };
}

function updateUploadStatusFromSummary() {
  const s = state.importSummary || {};
  const issueCount =
    (s.skippedNoCoords || 0) +
    (s.duplicateCustomerIds || 0) +
    (s.missingCurrentRep || 0) +
    ((s.unmappedFields || []).filter(key => isWarningField(key)).length ? 1 : 0);

  if ((s.loadedRows || 0) === 0) {
    setUploadStatus('bad', 'No valid rows loaded');
    return;
  }

  if (!issueCount) {
    setUploadStatus('good', `${(s.loadedRows || 0).toLocaleString()} loaded clean`);
    return;
  }

  const parts = [];
  if (s.skippedNoCoords) parts.push(`${s.skippedNoCoords} no coords`);
  if (s.duplicateCustomerIds) parts.push(`${s.duplicateCustomerIds} duplicate IDs`);
  if (s.missingCurrentRep) parts.push(`${s.missingCurrentRep} missing current rep`);
  if ((s.unmappedFields || []).filter(key => isWarningField(key)).length) parts.push('some unmapped fields');

  setUploadStatus('warning', `${(s.loadedRows || 0).toLocaleString()} loaded • ${parts.join(' • ')}`);
}

function rebuildMarkers() {
  state.markerLayer.clearLayers();
  state.markerById.clear();

  state.accounts.forEach(account => {
    const color = getRepColor(account.assignedRep);
    const marker = L.circleMarker([account.latitude, account.longitude], {
      radius: 4.2,
      weight: 1.2,
      color,
      fillColor: color,
      fillOpacity: 0.92,
      opacity: 0.95
    });

    marker.on('click', e => {
      const additive = e.originalEvent && (e.originalEvent.metaKey || e.originalEvent.ctrlKey || e.originalEvent.shiftKey);
      toggleSelection(account._id, additive);
      syncMarkerPopupContent(marker, account._id);
      marker.openPopup();
    });

    marker.bindPopup(buildPopupHtml(account), { maxWidth: 320 });
    marker.addTo(state.markerLayer);
    state.markerById.set(account._id, marker);
  });
}

function syncMarkerPopupContent(marker, id) {
  const account = state.accountById.get(id);
  if (!account || !marker) return;
  marker.setPopupContent(buildPopupHtml(account));
}

function refreshMarkerStyles() {
  const dimOthers = !!els.dimOthersCheckbox.checked;

  state.accounts.forEach(account => {
    const marker = state.markerById.get(account._id);
    if (!marker) return;

    const isSelected = state.selection.has(account._id);
    const isFocusedRep = !state.repFocus || account.assignedRep === state.repFocus;
    const color = getRepColor(account.assignedRep);

    let fillOpacity = isSelected ? 1 : 0.92;
    let opacity = isSelected ? 1 : 0.95;
    let weight = isSelected ? 2.4 : 1.2;
    let radius = isSelected ? 7 : (state.repFocus ? 3.8 : 4.2);

    if (state.repFocus && account.assignedRep !== state.repFocus) {
      if (dimOthers) {
        fillOpacity = 0.14;
        opacity = 0.18;
      } else {
        fillOpacity = 0.55;
        opacity = 0.62;
      }
      weight = 0.8;
      radius = 3.3;
    }

    if (!passesFilters(account)) {
      fillOpacity = 0.07;
      opacity = 0.08;
      weight = 0.6;
      radius = 3;
    }

    marker.setStyle({
      color,
      fillColor: color,
      fillOpacity,
      opacity,
      weight,
      radius
    });
  });
}

function passesFilters(account) {
  const repPass = state.filters.rep.size === 0 || state.filters.rep.has(account.assignedRep);
  const rankPass = state.filters.rank.size === 0 || state.filters.rank.has(account.rank);
  const chainPass = state.filters.chain.size === 0 || state.filters.chain.has(account.chain);
  const segmentPass = state.filters.segment.size === 0 || state.filters.segment.has(account.segment);

  const premisePass =
    state.filters.premise === 'ALL' ||
    account.premise === state.filters.premise;

  const protectedPass =
    state.filters.protected === 'ALL' ||
    (state.filters.protected === 'YES' && account.protected) ||
    (state.filters.protected === 'NO' && !account.protected);

  const movedPass =
    state.filters.moved === 'ALL' ||
    (state.filters.moved === 'MOVED' && account.assignedRep !== account.originalAssignedRep) ||
    (state.filters.moved === 'UNCHANGED' && account.assignedRep === account.originalAssignedRep);

  return repPass && rankPass && chainPass && segmentPass && premisePass && protectedPass && movedPass;
}

function buildPopupHtml(account) {
  const addressLine = [account.address, account.zip].filter(Boolean).join(' ');

  return `
    <div style="min-width:250px;">
      <div style="font-size:15px;font-weight:800;">${escapeHtml(account.customerName || account.customerId)}</div>
      <div style="margin-top:5px;color:#5d7286;font-size:12px;">${escapeHtml(addressLine)}</div>
      <div style="margin-top:8px;"><strong>Premise:</strong> ${escapeHtml(account.premise)}</div>
      <div><strong>Segment:</strong> ${escapeHtml(account.segment)}</div>
      <div><strong>Chain:</strong> ${escapeHtml(account.chain)}</div>
      <div><strong>Assigned Rep:</strong> ${escapeHtml(account.assignedRep)}</div>
      <div><strong>Current Rep:</strong> ${escapeHtml(account.currentRep)}</div>
      <div><strong>Revenue:</strong> ${formatCurrency(account.overallSales)}</div>
      <div><strong>Rank:</strong> ${escapeHtml(account.rank)}</div>
      <div><strong>Cadence 4W:</strong> ${formatNumber(account.cadence4w, 2)}</div>
      <div><strong>Protected:</strong> ${account.protected ? 'Yes' : 'No'}</div>
    </div>
  `;
}

function renderRepTable() {
  let rows = summarizeByRep();
  rows = sortRepRows(rows);

  if (!rows.length) {
    els.repTableBody.innerHTML = '<tr><td colspan="15" class="empty">Upload a file to begin.</td></tr>';
    return;
  }

  syncSortHeaderIndicators();

  els.repTableBody.innerHTML = rows.map(row => `
    <tr data-rep-row="${encodeURIComponent(row.rep)}" class="${state.repFocus === row.rep ? 'rep-row-active' : ''} ${state.lockedReps[row.rep] ? 'rep-row-locked' : ''}">
      <td>
        <div class="rep-cell">
          <span class="color-dot" style="background:${escapeHtmlAttr(row.color)}"></span>
          <span>${escapeHtml(row.rep)}</span>
        </div>
      </td>
      <td>${row.stops}</td>
      <td>${renderDeltaCount(row.deltaStops)}</td>
      <td>${formatCurrency(row.revenue)}</td>
      <td>${renderDeltaMoney(row.deltaRevenue)}</td>
      <td>${row.A}</td>
      <td>${row.B}</td>
      <td>${row.C}</td>
      <td>${row.D}</td>
      <td>${formatNumber(row.planned4W, 2)}</td>
      <td>${formatNumber(row.avgWeekly, 2)}</td>
      <td>${row.protected}</td>
      <td>${row.movedIn}</td>
      <td>${row.movedOut}</td>
      <td class="lock-cell">
        <label class="lock-check-wrap" title="Lock this territory for optimization and manual reassignment">
          <input
            type="checkbox"
            class="lock-checkbox"
            data-lock-rep="${escapeHtmlAttr(encodeURIComponent(row.rep))}"
            ${state.lockedReps[row.rep] ? 'checked' : ''}
          />
          <span>Lock</span>
        </label>
      </td>
    </tr>
  `).join('');

  [...els.repTableBody.querySelectorAll('.lock-check-wrap')].forEach(label => {
    label.addEventListener('click', e => e.stopPropagation());
  });

  [...els.repTableBody.querySelectorAll('input[data-lock-rep]')].forEach(input => {
    input.addEventListener('click', e => e.stopPropagation());
    input.addEventListener('change', e => {
      e.stopPropagation();
      const rep = decodeURIComponent(e.currentTarget.getAttribute('data-lock-rep') || '');
      if (!rep) return;
      toggleRepLock(rep, e.currentTarget.checked);
    });
  });

  [...els.repTableBody.querySelectorAll('tr[data-rep-row]')].forEach(tr => {
    tr.addEventListener('click', () => {
      const rep = decodeURIComponent(tr.getAttribute('data-rep-row') || '');

      if (!rep) return;

      if (state.repFocus === rep) {
        state.repFocus = null;
        refreshUI(false);
        if (state.accounts.length) fitMapToAccounts();
      } else {
        state.repFocus = rep;
        refreshUI(false);
        zoomToRep(rep);
      }
    });
  });
}

function toggleTableSort(key) {
  if (!key) return;

  if (state.tableSort.key === key) {
    state.tableSort.dir = state.tableSort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    state.tableSort.key = key;
    state.tableSort.dir = key === 'rep' ? 'asc' : 'desc';
  }

  renderRepTable();
}

function sortRepRows(rows) {
  const key = state.tableSort.key;
  const dir = state.tableSort.dir === 'asc' ? 1 : -1;

  return [...rows].sort((a, b) => {
    const av = a[key];
    const bv = b[key];

    if (typeof av === 'string' || typeof bv === 'string') {
      return String(av).localeCompare(String(bv), undefined, { numeric: true }) * dir;
    }

    return ((av || 0) - (bv || 0)) * dir;
  });
}

function syncSortHeaderIndicators() {
  document.querySelectorAll('th[data-sort]').forEach(th => {
    const key = th.getAttribute('data-sort');
    const indicator = th.querySelector('.sort-indicator');
    th.classList.remove('is-active', 'asc', 'desc');

    if (key === state.tableSort.key) {
      th.classList.add('is-active', state.tableSort.dir);
      if (indicator) indicator.textContent = state.tableSort.dir === 'asc' ? '▲' : '▼';
    } else if (indicator) {
      indicator.textContent = '↕';
    }
  });
}

function renderSelectionPreview() {
  const selected = state.accounts.filter(a => state.selection.has(a._id)).slice(0, 250);
  els.selectionCount.textContent = String(state.selection.size);

  if (!selected.length) {
    els.selectionPreview.innerHTML = '<div class="empty">No accounts selected.</div>';
    return;
  }

  els.selectionPreview.innerHTML = selected.map(a => `
    <div class="selected-item">
      <div class="selected-item-title">${escapeHtml(a.customerName)}</div>
      <div class="transfer-line">
        <span class="rep-chip">${escapeHtml(a.assignedRep)}</span>
        <span class="metric-chip">${escapeHtml(a.premise)}</span>
        <span class="metric-chip">${escapeHtml(a.segment)}</span>
        <span class="metric-chip">${formatCurrency(a.overallSales)}</span>
        <span class="metric-chip">Rank ${escapeHtml(a.rank)}</span>
        <span class="metric-chip">4W ${formatNumber(a.cadence4w, 2)}</span>
        ${a.protected ? '<span class="metric-chip">Protected</span>' : ''}
      </div>
    </div>
  `).join('');
}

function renderMovedReview() {
  const search = (state.multiSearch.moved || '').toLowerCase().trim();

  let moved = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep);

  if (search) {
    moved = moved.filter(a => {
      const haystack = [
        a.customerName,
        a.customerId,
        a.originalAssignedRep,
        a.assignedRep
      ].map(v => String(v || '').toLowerCase()).join(' | ');

      return haystack.includes(search);
    });
  }

  const limited = moved.slice(0, 100);
  els.movedReviewCount.textContent = String(state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep).length);

  if (!limited.length) {
    els.movedReviewList.innerHTML = `<div class="empty">${search ? 'No moved accounts match your search.' : 'No moved accounts yet.'}</div>`;
    return;
  }

  els.movedReviewList.innerHTML = limited.map(a => `
    <div class="moved-item" data-account-id="${escapeHtmlAttr(a._id)}" title="Zoom to account on map">
      <div class="moved-item-title">${escapeHtml(a.customerName)}</div>
      <div class="transfer-line">
        <span class="metric-chip">${escapeHtml(a.customerId)}</span>
        <span class="rep-chip">${escapeHtml(a.originalAssignedRep)}</span>
        <span class="rep-arrow">→</span>
        <span class="rep-chip">${escapeHtml(a.assignedRep)}</span>
        <span class="metric-chip">${formatCurrency(a.overallSales)}</span>
        <span class="metric-chip">Rank ${escapeHtml(a.rank)}</span>
        ${a.protected ? '<span class="metric-chip">Protected</span>' : ''}
      </div>
    </div>
  `).join('');

  els.movedReviewList.querySelectorAll('.moved-item[data-account-id]').forEach(item => {
    item.addEventListener('click', () => {
      const id = item.getAttribute('data-account-id');
      if (id) zoomToAccount(id, true);
    });
  });
}

function renderSummary() {
  const totalAccounts = state.accounts.length;
  const totalRevenue = state.accounts.reduce((sum, a) => sum + (a.overallSales || 0), 0);
  const protectedCount = state.accounts.filter(a => a.protected).length;
  const movedCount = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep).length;
  const unchangedPct = totalAccounts ? ((totalAccounts - movedCount) / totalAccounts) * 100 : 0;
  const totalPlanned4W = state.accounts.reduce((sum, a) => sum + (a.cadence4w || 0), 0);
  const avgWeekly = totalPlanned4W / 4;
  const repCount = Math.max(1, getAllAssignedReps().length);
  const avgWeeklyPerRep = avgWeekly / repCount;

  els.globalAccounts.textContent = totalAccounts.toLocaleString();
  els.globalRevenue.textContent = formatCurrency(totalRevenue);
  els.globalProtected.textContent = protectedCount.toLocaleString();
  els.globalMoved.textContent = movedCount.toLocaleString();
  els.globalUnchanged.textContent = `${unchangedPct.toFixed(1)}%`;
  els.globalAvgWeekly.textContent = formatNumber(avgWeekly, 1);
  els.globalAvgWeeklyPerRep.textContent = formatNumber(avgWeeklyPerRep, 1);
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

  state.selection.clear();

  for (const account of state.accounts) {
    const point = turf.point([account.longitude, account.latitude]);
    if (turf.booleanPointInPolygon(point, polygon)) state.selection.add(account._id);
  }

  refreshUI();
  showToast(`${state.selection.size} account${state.selection.size === 1 ? '' : 's'} selected.`);
}

function toggleSelection(id, additive = false) {
  if (!additive) state.selection.clear();
  if (state.selection.has(id)) state.selection.delete(id);
  else state.selection.add(id);
  refreshUI();
}

function clearSelection() {
  state.selection.clear();
  state.drawLayer.clearLayers();
  refreshUI();
}

function zoomToAccount(id, openPopup = true) {
  const account = state.accountById.get(id);
  const marker = state.markerById.get(id);
  if (!account || !marker) return;

  state.repFocus = account.assignedRep;
  state.selection.clear();
  state.selection.add(id);
  refreshUI(false);

  state.map.setView([account.latitude, account.longitude], 12, { animate: true });

  if (openPopup) {
    setTimeout(() => {
      syncMarkerPopupContent(marker, id);
      marker.openPopup();
    }, 120);
  }
}

function isRepLocked(rep) {
  return !!state.lockedReps[rep];
}

function getLockedRepNames() {
  return Object.keys(state.lockedReps).filter(rep => state.lockedReps[rep]);
}

function toggleRepLock(rep, locked) {
  if (!rep) return;

  const current = !!state.lockedReps[rep];
  const next = !!locked;

  if (current === next) return;

  state.undoStack.push({
    type: 'rep-lock',
    rep,
    from: current,
    to: next,
    label: `${next ? 'Locked' : 'Unlocked'} ${rep}`
  });

  state.lockedReps[rep] = next;
  state.optimizationSummary = null;
  refreshUI(false);
  updateLastAction(`${next ? 'Locked' : 'Unlocked'} ${rep}`);
  showToast(`${next ? 'Locked' : 'Unlocked'} ${rep}.`);
}

function assignSelectionToRep() {
  const targetRep = els.assignRepSelect.value;
  const selectedIds = [...state.selection];
  if (!selectedIds.length || !targetRep) return;

  const previousAssignedReps = getAllAssignedReps();
  const changes = [];
  let skippedProtected = 0;
  let skippedLockedFrom = 0;
  let skippedLockedTo = 0;

  ensureRepColor(targetRep);

  for (const id of selectedIds) {
    const account = state.accountById.get(id);
    if (!account) continue;

    if (account.protected && account.assignedRep !== targetRep) {
      skippedProtected += 1;
      continue;
    }

    if (isRepLocked(account.assignedRep) && account.assignedRep !== targetRep) {
      skippedLockedFrom += 1;
      continue;
    }

    if (isRepLocked(targetRep) && account.assignedRep !== targetRep) {
      skippedLockedTo += 1;
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
    const parts = [];
    if (skippedProtected) parts.push(`${skippedProtected} protected`);
    if (skippedLockedFrom) parts.push(`${skippedLockedFrom} locked-from`);
    if (skippedLockedTo) parts.push(`${skippedLockedTo} locked-to`);

    showToast(parts.length ? `${parts.join(', ')} account(s) were skipped.` : 'No assignment changes to make.');
    return;
  }

  applyChanges(changes, `Assigned ${changes.length} account${changes.length === 1 ? '' : 's'} to ${targetRep}`, previousAssignedReps);
  clearSelection();

  const notes = [];
  if (skippedProtected) notes.push(`${skippedProtected} protected account(s) stayed put`);
  if (skippedLockedFrom) notes.push(`${skippedLockedFrom} locked account(s) could not be moved out`);
  if (skippedLockedTo) notes.push(`${skippedLockedTo} account(s) could not be moved into locked rep ${targetRep}`);

  if (notes.length) {
    showToast(`${changes.length} reassigned to ${targetRep}. ${notes.join('. ')}.`);
  }
}

function applyChanges(changes, label, previousAssignedReps = null) {
  const repsBefore = Array.isArray(previousAssignedReps) ? previousAssignedReps : getAllAssignedReps();

  changes.forEach(change => {
    const account = state.accountById.get(change.id);
    if (!account) return;

    ensureRepColor(change.to);
    account.assignedRep = change.to;

    state.changeLog.push({
      timestamp: new Date().toISOString(),
      customerId: account.customerId,
      customerName: account.customerName,
      fromRep: change.from,
      toRep: change.to,
      protected: account.protected ? 'Yes' : 'No'
    });
  });

  state.undoStack.push({
    changes: changes.map(c => ({ ...c })),
    label
  });

  buildRepColors();
  syncRepFilterSelection(repsBefore);

  if (state.repFocus && !getAllAssignedReps().includes(state.repFocus)) {
    state.repFocus = null;
  }

  refreshUI();
  updateLastAction(label);
  showToast(label);
}

function undoLastAction() {
  const action = state.undoStack.pop();
  if (!action) return;

  if (action.type === 'rep-lock') {
    state.lockedReps[action.rep] = action.from;
    state.optimizationSummary = null;
    refreshUI(false);
    updateLastAction(`Undid: ${action.label}`);
    showToast(`Undid: ${action.label}`);
    return;
  }

  const repsBefore = getAllAssignedReps();

  for (const change of action.changes) {
    const account = state.accountById.get(change.id);
    if (account) account.assignedRep = change.from;
  }

  buildRepColors();
  syncRepFilterSelection(repsBefore);

  if (state.repFocus && !getAllAssignedReps().includes(state.repFocus)) {
    state.repFocus = null;
  }

  state.optimizationSummary = null;
  refreshUI();
  updateLastAction(`Undid: ${action.label}`);
  showToast(`Undid: ${action.label}`);
}

function resetAssignments() {
  const changes = [];
  const repsBefore = getAllAssignedReps();

  for (const account of state.accounts) {
    if (account.assignedRep !== account.originalAssignedRep) {
      changes.push({
        id: account._id,
        from: account.assignedRep,
        to: account.originalAssignedRep
      });
    }
  }

  if (!changes.length) {
    showToast('Nothing to reset.');
    return;
  }

  changes.forEach(change => {
    const account = state.accountById.get(change.id);
    if (!account) return;
    account.assignedRep = change.to;
  });

  state.undoStack = [];
  state.changeLog = [];
  state.repFocus = null;
  state.optimizationSummary = null;
  state.multiSearch.moved = '';
  if (els.movedSearchInput) els.movedSearchInput.value = '';

  buildRepColors();
  syncRepFilterSelection(repsBefore);
  refreshUI();
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
    const protectedCount = state.accounts.filter(a => a.protected).length;
    const movableCount = totalAccounts - protectedCount;
    const lockedRepNames = getLockedRepNames();
    const lockedAccounts = state.accounts.filter(a => lockedRepNames.includes(a.assignedRep));
    const unlockedAccounts = state.accounts.filter(a => !lockedRepNames.includes(a.assignedRep));
    const unlockedProtectedCount = unlockedAccounts.filter(a => a.protected).length;
    const unlockedMovableCount = unlockedAccounts.length - unlockedProtectedCount;
    const targetUnlockedCount = targetCount - lockedRepNames.length;

    if (!Number.isFinite(targetCount) || !Number.isFinite(minStops) || !Number.isFinite(maxStops)) {
      showToast('Optimizer inputs are invalid. Check rep count and stop limits.');
      return;
    }

    if (targetCount > totalAccounts) {
      showToast(`Target rep count of ${targetCount} exceeds ${totalAccounts} total accounts.`);
      return;
    }

    if (lockedRepNames.length >= targetCount) {
      showToast(`Target rep count must be greater than the ${lockedRepNames.length} locked rep${lockedRepNames.length === 1 ? '' : 's'}.`);
      return;
    }

    if (targetUnlockedCount <= 0) {
      showToast('All target reps are already locked. Nothing is left to optimize.');
      return;
    }

    if (!unlockedAccounts.length) {
      showToast('All accounts belong to locked reps. Nothing is left to optimize.');
      return;
    }

    if (targetUnlockedCount > unlockedAccounts.length) {
      showToast(`Only ${unlockedAccounts.length} unlocked accounts remain. Reduce target rep count or unlock territories.`);
      return;
    }

    if (targetUnlockedCount * minStops > unlockedAccounts.length) {
      showToast(`Minimum stops too high for the unlocked pool. ${targetUnlockedCount} unlocked reps × ${minStops} minimum exceeds ${unlockedAccounts.length} unlocked accounts.`);
      return;
    }

    if (Math.ceil(unlockedAccounts.length / targetUnlockedCount) > maxStops) {
      showToast(`Maximum stops too low for the unlocked pool. ${targetUnlockedCount} unlocked reps cannot cover ${unlockedAccounts.length} unlocked accounts with a max of ${maxStops} per rep.`);
      return;
    }

    if (movableCount === 0 || unlockedMovableCount === 0) {
      showToast('No unlocked, movable accounts are available to optimize.');
      return;
    }

    const continuityWeight = Number(els.disruptionSlider.value) / 100;
    const geographyWeight = 1 - continuityWeight;
    const balanceMode = els.balanceMode.value;

    const protectedAccounts = unlockedAccounts.filter(a => a.protected);
    const movableAccounts = unlockedAccounts.filter(a => !a.protected);
    const currentUnlockedReps = getAllAssignedReps().filter(rep => !isRepLocked(rep));
    const targetRepNames = buildTargetRepNames(targetUnlockedCount, currentUnlockedReps);
    const adjacency = buildNeighborMap(unlockedAccounts);

    targetRepNames.forEach(rep => ensureRepColor(rep));

    const assignments = new Map();

    lockedAccounts.forEach(a => assignments.set(a._id, a.assignedRep));
    protectedAccounts.forEach(a => assignments.set(a._id, a.assignedRep));

    const assignmentCtx = createAssignmentContext(targetRepNames, assignments);
    const centroids = initializeCentroidsFast(targetRepNames, assignmentCtx);

    const orderedMovable = [...movableAccounts].sort((a, b) => {
      if (a.rank !== b.rank) return rankSortValue(a.rank) - rankSortValue(b.rank);
      if (a.overallSales !== b.overallSales) return b.overallSales - a.overallSales;
      return a.customerName.localeCompare(b.customerName);
    });

    for (let iter = 0; iter < 6; iter += 1) {
      assignmentCtx.clearMovableAssignments(orderedMovable, assignments);
      resetCentroidsFromContext(centroids, targetRepNames, assignmentCtx);

      const repStats = buildFullRepStats(targetRepNames);
      for (const rep of targetRepNames) {
        const repCtx = assignmentCtx.reps.get(rep);
        repStats.set(rep, { rep, stops: repCtx.count, revenue: repCtx.revenue });
      }

      for (const account of orderedMovable) {
        let bestRep = null;
        let bestScore = Infinity;

        for (const rep of targetRepNames) {
          const centroid = centroids.get(rep);
          const repStat = repStats.get(rep);

          if (repStat.stops >= maxStops) continue;

          const dist = squaredDistance(account.latitude, account.longitude, centroid.lat, centroid.lng);

          const compactnessScore = dist * (1.08 + geographyWeight * 1.15);
          const continuityPenalty = account.currentRep === rep ? 0 : (continuityWeight * 0.95);
          const existingPenalty = account.assignedRep === rep ? 0 : (continuityWeight * 0.28);

          let balancePenalty = 0;
          if (balanceMode === 'stops') {
            balancePenalty = repStat.stops * 0.04;
          } else if (balanceMode === 'revenue') {
            balancePenalty = repStat.revenue * 0.0000013;
          } else {
            balancePenalty = repStat.stops * 0.028 + repStat.revenue * 0.00000065;
          }

          const underMinBoost = repStat.stops < minStops ? -2.2 : 0;
          const localPenalty = localDominancePenalty(account, rep, assignments, adjacency);

          const score =
            compactnessScore +
            continuityPenalty +
            existingPenalty +
            balancePenalty +
            localPenalty +
            underMinBoost;

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
      showToast('Optimizer could not fully satisfy the stop limits with the current unlocked pool. Try adjusting rep count, stop limits, or unlocking territories.');
    }

    const changes = [];
    const repsBefore = getAllAssignedReps();

    for (const account of unlockedAccounts) {
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

    const lockedLabel = lockedRepNames.length ? ` (${lockedRepNames.length} locked)` : '';
    applyChanges(changes, `Optimized routes to ${targetCount} reps with minimum ${minStops} stops${lockedLabel}`, repsBefore);
    state.optimizationSummary = buildOptimizationSummary();
    updateLastActionWithOptimization(`Optimized routes to ${targetCount} reps with minimum ${minStops} stops${lockedLabel}`);
    refreshUI(false);
  } catch (err) {
    console.error('Optimize Routes failed:', err);
    showToast('Optimize Routes hit an error. Send me the first red error line from the browser console.');
  }
}

function buildOptimizationSummary() {
  const rows = summarizeByRep();
  const movedCount = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep).length;
  const protectedHeld = state.accounts.filter(a => a.protected && a.assignedRep === a.originalAssignedRep).length;

  const stops = rows.map(r => r.stops);
  const revenue = rows.map(r => r.revenue);

  return {
    repCount: rows.length,
    movedCount,
    protectedHeld,
    minStops: Math.min(...stops),
    maxStops: Math.max(...stops),
    minRevenue: Math.min(...revenue),
    maxRevenue: Math.max(...revenue),
    avgStops: rows.reduce((sum, r) => sum + r.stops, 0) / Math.max(1, rows.length)
  };
}

function updateLastActionWithOptimization(baseText) {
  const s = state.optimizationSummary;
  if (!s) {
    updateLastAction(baseText);
    return;
  }

  updateLastAction(
    `${baseText} • Stops range ${s.minStops}-${s.maxStops} • Avg stops ${formatNumber(s.avgStops, 1)}`
  );
}

function createAssignmentContext(targetRepNames, assignments) {
  const ctx = {
    reps: new Map()
  };

  targetRepNames.forEach(rep => {
    ctx.reps.set(rep, {
      rep,
      count: 0,
      revenue: 0,
      latSum: 0,
      lngSum: 0,
      ids: new Set()
    });
  });

  for (const account of state.accounts) {
    const rep = assignments.get(account._id);
    if (!rep || !ctx.reps.has(rep)) continue;

    const bucket = ctx.reps.get(rep);
    bucket.count += 1;
    bucket.revenue += account.overallSales || 0;
    bucket.latSum += account.latitude;
    bucket.lngSum += account.longitude;
    bucket.ids.add(account._id);
  }

  ctx.count = rep => (ctx.reps.get(rep)?.count || 0);
  ctx.revenue = rep => (ctx.reps.get(rep)?.revenue || 0);

  ctx.addToRep = (rep, account) => {
    const bucket = ctx.reps.get(rep);
    if (!bucket || bucket.ids.has(account._id)) return;
    bucket.count += 1;
    bucket.revenue += account.overallSales || 0;
    bucket.latSum += account.latitude;
    bucket.lngSum += account.longitude;
    bucket.ids.add(account._id);
  };

  ctx.removeFromRep = (rep, account) => {
    const bucket = ctx.reps.get(rep);
    if (!bucket || !bucket.ids.has(account._id)) return;
    bucket.count -= 1;
    bucket.revenue -= account.overallSales || 0;
    bucket.latSum -= account.latitude;
    bucket.lngSum -= account.longitude;
    bucket.ids.delete(account._id);
  };

  ctx.move = (account, fromRep, toRep) => {
    if (fromRep === toRep) return;
    ctx.removeFromRep(fromRep, account);
    ctx.addToRep(toRep, account);
    assignments.set(account._id, toRep);
  };

  ctx.clearMovableAssignments = (movableAccounts, assignmentsMap) => {
    movableAccounts.forEach(account => {
      const current = assignmentsMap.get(account._id);
      if (!current) return;
      ctx.removeFromRep(current, account);
      assignmentsMap.delete(account._id);
    });
  };

  return ctx;
}

function initializeCentroidsFast(targetRepNames, assignmentCtx) {
  const centroids = new Map();

  targetRepNames.forEach(rep => {
    const bucket = assignmentCtx.reps.get(rep);
    if (bucket && bucket.count) {
      centroids.set(rep, {
        lat: bucket.latSum / bucket.count,
        lng: bucket.lngSum / bucket.count
      });
      return;
    }

    const repAccounts = state.accounts.filter(a => a.assignedRep === rep);
    if (repAccounts.length) {
      centroids.set(rep, {
        lat: repAccounts.reduce((sum,a) => sum + a.latitude, 0) / repAccounts.length,
        lng: repAccounts.reduce((sum,a) => sum + a.longitude, 0) / repAccounts.length
      });
      return;
    }

    const fallback = state.accounts[Math.floor(Math.random() * state.accounts.length)];
    centroids.set(rep, { lat: fallback.latitude, lng: fallback.longitude });
  });

  return centroids;
}

function resetCentroidsFromContext(centroids, targetRepNames, assignmentCtx) {
  targetRepNames.forEach(rep => {
    const bucket = assignmentCtx.reps.get(rep);
    if (bucket && bucket.count) {
      centroids.set(rep, {
        lat: bucket.latSum / bucket.count,
        lng: bucket.lngSum / bucket.count
      });
    }
  });
}

function refreshCentroidsFromContext(centroids, targetRepNames, assignmentCtx) {
  resetCentroidsFromContext(centroids, targetRepNames, assignmentCtx);
}

function buildTargetRepNames(targetCount, existingReps) {
  const unique = [...new Set(existingReps)].slice(0, targetCount);

  while (unique.length < targetCount) {
    const generated = `Rep ${unique.length + 1}`;
    if (!unique.includes(generated)) unique.push(generated);
  }

  return unique;
}

function buildFullRepStats(targetRepNames) {
  const stats = new Map();
  targetRepNames.forEach(rep => {
    stats.set(rep, { rep, stops: 0, revenue: 0 });
  });
  return stats;
}

function localDominancePenalty(account, rep, assignments, adjacency) {
  const neighbors = adjacency.get(account._id);
  if (!neighbors || !neighbors.size) return 0;

  let same = 0;
  let diff = 0;

  neighbors.forEach(neighborId => {
    const neighborRep = assignments.get(neighborId);
    if (!neighborRep) return;
    if (neighborRep === rep) same += 1;
    else diff += 1;
  });

  if (!same && diff) return 0.9;
  return diff > same ? 0.35 : 0;
}

function enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx) {
  const accounts = state.accounts.filter(a => assignments.has(a._id));

  let changed = true;
  let safety = 0;

  while (changed && safety < 500) {
    changed = false;
    safety += 1;

    const under = targetRepNames.filter(rep => ctx.count(rep) < minStops).sort((a,b) => ctx.count(a) - ctx.count(b));
    const over = targetRepNames.filter(rep => ctx.count(rep) > minStops).sort((a,b) => ctx.count(b) - ctx.count(a));

    if (!under.length || !over.length) break;

    for (const needRep of under) {
      const donor = over.find(rep => ctx.count(rep) > minStops);
      if (!donor) break;

      const candidates = accounts
        .filter(a => assignments.get(a._id) === donor && !a.protected)
        .sort((a,b) => squaredDistance(a.latitude, a.longitude, getRepCentroidLat(needRep, ctx), getRepCentroidLng(needRep, ctx)) -
                       squaredDistance(b.latitude, b.longitude, getRepCentroidLat(needRep, ctx), getRepCentroidLng(needRep, ctx)));

      const chosen = candidates[0];
      if (!chosen) continue;

      ctx.move(chosen, donor, needRep);
      changed = true;
    }
  }
}

function enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, ctx) {
  const accounts = state.accounts.filter(a => assignments.has(a._id));

  let changed = true;
  let safety = 0;

  while (changed && safety < 500) {
    changed = false;
    safety += 1;

    const over = targetRepNames.filter(rep => ctx.count(rep) > maxStops).sort((a,b) => ctx.count(b) - ctx.count(a));
    const under = targetRepNames.filter(rep => ctx.count(rep) < maxStops).sort((a,b) => ctx.count(a) - ctx.count(b));

    if (!over.length || !under.length) break;

    for (const donor of over) {
      const receiver = under.find(rep => rep !== donor && ctx.count(rep) < maxStops);
      if (!receiver) break;

      const candidates = accounts
        .filter(a => assignments.get(a._id) === donor && !a.protected)
        .sort((a,b) => squaredDistance(a.latitude, a.longitude, getRepCentroidLat(receiver, ctx), getRepCentroidLng(receiver, ctx)) -
                       squaredDistance(b.latitude, b.longitude, getRepCentroidLat(receiver, ctx), getRepCentroidLng(receiver, ctx)));

      const chosen = candidates[0];
      if (!chosen) continue;

      ctx.move(chosen, donor, receiver);
      changed =
