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
  const trigger = wrap?.querySelector('.multi-trigger');
  const panel = wrap?.querySelector('.multi-panel');
  if (!wrap || !trigger || !panel || !wrap.classList.contains('open')) return;

  const rect = trigger.getBoundingClientRect();
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;

  const width = 220;
  let left = rect.left;
  if (left + width > viewportWidth - 12) left = viewportWidth - width - 12;
  if (left < 10) left = 10;

  let top = rect.bottom + 6;
  const availableBelow = viewportHeight - top - 12;
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

function toggleUploadStatusPanel() {
  if (!els.uploadStatusPanel) return;

  const isOpen = !els.uploadStatusPanel.hidden;
  if (isOpen) {
    closeUploadStatusPanel();
    return;
  }

  els.uploadStatusBody.innerHTML = buildUploadStatusDetailsHtml();
  els.uploadStatusPanel.hidden = false;
  els.uploadStatusPill.setAttribute('aria-expanded', 'true');
  positionUploadStatusPanel();
}

function closeUploadStatusPanel() {
  if (!els.uploadStatusPanel) return;
  els.uploadStatusPanel.hidden = true;
  els.uploadStatusPill.setAttribute('aria-expanded', 'false');
}

function positionUploadStatusPanel() {
  if (!els.uploadStatusPanel || els.uploadStatusPanel.hidden || !els.uploadStatusPill) return;

  const rect = els.uploadStatusPill.getBoundingClientRect();
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const width = Math.min(320, viewportWidth - 20);

  let left = rect.left;
  if (left + width > viewportWidth - 10) left = viewportWidth - width - 10;
  if (left < 10) left = 10;

  let top = rect.bottom + 8;
  const estimatedHeight = Math.min(420, els.uploadStatusPanel.offsetHeight || 220);

  if (top + estimatedHeight > viewportHeight - 10) {
    top = Math.max(10, rect.top - estimatedHeight - 8);
  }

  els.uploadStatusPanel.style.width = `${width}px`;
  els.uploadStatusPanel.style.left = `${left}px`;
  els.uploadStatusPanel.style.top = `${top}px`;
}

function buildUploadStatusDetailsHtml() {
  const status = state.uploadStatus || { level: 'neutral', text: 'No file loaded' };
  const s = state.importSummary || {};
  const warningFields = (s.unmappedFields || []).filter(key => isWarningField(key));
  const infoFields = (s.unmappedFields || []).filter(key => !isWarningField(key));

  if (status.level === 'neutral') {
    return `
      <div class="upload-diag-summary">No file loaded yet.</div>
      <div class="upload-diag-empty">Upload an Excel or CSV file to view import diagnostics.</div>
    `;
  }

  if (status.level === 'bad' && !s.loadedRows) {
    return `
      <div class="upload-diag-summary">${escapeHtml(status.text || 'Import failed.')}</div>
      <div class="upload-diag-empty">No valid account rows were loaded.</div>
    `;
  }

  const warnings = [];
  if (s.skippedNoCoords) warnings.push(`${s.skippedNoCoords.toLocaleString()} row(s) skipped for missing latitude/longitude`);
  if (s.duplicateCustomerIds) warnings.push(`${s.duplicateCustomerIds.toLocaleString()} duplicate customer ID(s) adjusted`);
  if (s.missingCurrentRep) warnings.push(`${s.missingCurrentRep.toLocaleString()} row(s) missing current rep`);
  if (warningFields.length) warnings.push(`Unmapped warning field(s): ${warningFields.map(prettyFieldName).join(', ')}`);

  const info = [];
  if (s.missingAssignedRep) info.push(`${s.missingAssignedRep.toLocaleString()} row(s) had blank assigned rep and defaulted from current rep`);
  if (infoFields.length) info.push(`Optional unmapped field(s): ${infoFields.map(prettyFieldName).join(', ')}`);
  if (!headerMapped('cadence4w')) info.push('Cadence 4W not mapped; cadence defaults from rank');

  return `
    <div class="upload-diag-summary">${escapeHtml(status.text || '')}</div>

    <div class="upload-diag-grid">
      <div class="upload-diag-label">Source rows</div>
      <div class="upload-diag-value">${Number(s.sourceRows || 0).toLocaleString()}</div>

      <div class="upload-diag-label">Loaded rows</div>
      <div class="upload-diag-value">${Number(s.loadedRows || 0).toLocaleString()}</div>

      <div class="upload-diag-label">Skipped rows</div>
      <div class="upload-diag-value">${Number(s.skippedNoCoords || 0).toLocaleString()}</div>

      <div class="upload-diag-label">Duplicate IDs adjusted</div>
      <div class="upload-diag-value">${Number(s.duplicateCustomerIds || 0).toLocaleString()}</div>

      <div class="upload-diag-label">Blank current rep</div>
      <div class="upload-diag-value">${Number(s.missingCurrentRep || 0).toLocaleString()}</div>

      <div class="upload-diag-label">Blank assigned rep</div>
      <div class="upload-diag-value">${Number(s.missingAssignedRep || 0).toLocaleString()}</div>
    </div>

    <div class="upload-diag-section">
      <div class="upload-diag-label">Warnings</div>
      ${
        warnings.length
          ? `<ul class="upload-diag-list">${warnings.map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul>`
          : `<div class="upload-diag-empty">No actionable warnings.</div>`
      }
    </div>

    <div class="upload-diag-section">
      <div class="upload-diag-label">Info</div>
      ${
        info.length
          ? `<ul class="upload-diag-list">${info.map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul>`
          : `<div class="upload-diag-empty">No additional notes.</div>`
      }
    </div>
  `;
}

function isWarningField(key) {
  return ['latitude','longitude','customerId','customerName','currentRep','overallSales','rank','protected'].includes(key);
}

function prettyFieldName(key) {
  const map = {
    latitude: 'Latitude',
    longitude: 'Longitude',
    customerId: 'Customer ID',
    customerName: 'Customer Name',
    currentRep: 'Current Rep',
    assignedRep: 'Assigned Rep',
    overallSales: 'Overall Sales',
    rank: 'Rank',
    cadence4w: 'Cadence 4W',
    protected: 'Protected'
  };
  return map[key] || key;
}

function headerMapped(key) {
  const fields = state.importSummary?.unmappedFields || [];
  return !fields.includes(key);
}

function onFileChosen(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  state.loadedFileName = file.name.replace(/\.[^.]+$/, '') + '_updated.xlsx';
  closeUploadStatusPanel();
  setUploadStatus('neutral', `Reading ${file.name}...`);
  renderUploadStatus();

  const reader = new FileReader();

  reader.onload = e => {
    const data = e.target.result;

    try {
      if (file.name.toLowerCase().endsWith('.csv')) {
        const workbook = XLSX.read(data, { type: 'binary' });
        loadWorkbook(workbook);
      } else {
        const arr = new Uint8Array(data);
        const workbook = XLSX.read(arr, { type: 'array' });
        loadWorkbook(workbook);
      }
    } catch (err) {
      console.error(err);
      setUploadStatus('bad', 'File could not be read');
      renderUploadStatus();
      showToast('File could not be read. Check the format and try again.');
    }
  };

  reader.onerror = () => {
    setUploadStatus('bad', 'File could not be read');
    renderUploadStatus();
    showToast('File could not be read. Check the format and try again.');
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
  els.sheetSelect.innerHTML = '';

  workbook.SheetNames.forEach((sheetName, idx) => {
    state.workbookSheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

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

  showToast(`Loaded ${state.accounts.length} accounts from ${sheetName}.`);
  updateLastAction(`Loaded ${state.accounts.length} accounts from ${sheetName}`);
}

function normalizeRows(rows) {
  if (!rows?.length) {
    return {
      mapped: [],
      summary: {
        sourceRows: 0,
        loadedRows: 0,
        skippedNoCoords: 0,
        duplicateCustomerIds: 0,
        missingCurrentRep: 0,
        missingAssignedRep: 0,
        unmappedFields: []
      }
    };
  }

  const mapped = [];
  const headers = Object.keys(rows[0] || {});
  const headerMap = {};
  const cleanedHeaderLookup = new Map(headers.map(h => [cleanHeader(h), h]));
  const idCounts = new Map();

  Object.entries(COLUMN_ALIASES).forEach(([key, aliases]) => {
    let match = null;

    for (const alias of aliases) {
      const cleaned = cleanHeader(alias);

      if (cleanedHeaderLookup.has(cleaned)) {
        match = cleanedHeaderLookup.get(cleaned);
        break;
      }

      for (const original of headers) {
        const c = cleanHeader(original);
        if (c === cleaned || c.includes(cleaned) || cleaned.includes(c)) {
          match = original;
          break;
        }
      }

      if (match) break;
    }

    headerMap[key] = match;
  });

  let generatedId = 1;
  let skippedNoCoords = 0;
  let duplicateCustomerIds = 0;
  let missingCurrentRep = 0;
  let missingAssignedRep = 0;

  const recommendedFields = [
    'latitude','longitude','customerId','customerName','currentRep','overallSales','rank','protected'
  ];

  const unmappedFields = recommendedFields.filter(key => !headerMap[key]);

  for (const row of rows) {
    const lat = toNumber(row[headerMap.latitude]);
    const lng = toNumber(row[headerMap.longitude]);

    if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
      skippedNoCoords += 1;
      continue;
    }

    const rawCustomerId = safeString(row[headerMap.customerId]) || `AUTO_${generatedId++}`;
    const seen = idCounts.get(rawCustomerId) || 0;
    idCounts.set(rawCustomerId, seen + 1);

    let uniqueCustomerId = rawCustomerId;
    if (seen > 0) {
      duplicateCustomerIds += 1;
      uniqueCustomerId = `${rawCustomerId}__${seen + 1}`;
    }

    const customerName = safeString(row[headerMap.customerName]) || rawCustomerId;

    const rawCurrentRep = safeString(row[headerMap.currentRep]);
    const rawAssignedRep = safeString(row[headerMap.assignedRep]);

    if (!rawCurrentRep) missingCurrentRep += 1;
    if (headerMap.assignedRep && !rawAssignedRep) missingAssignedRep += 1;

    const currentRep = rawCurrentRep || 'Unassigned';
    const assignedRep = rawAssignedRep || currentRep || 'Unassigned';

    const rank = normalizeRank(row[headerMap.rank]);
    const premise = safeString(row[headerMap.premise]) || 'Unknown';
    const cadence4w = normalizeCadence4W(row[headerMap.cadence4w], rank);
    const overallSales = toNumber(row[headerMap.overallSales]);

    mapped.push({
      _id: uniqueCustomerId,
      latitude: lat,
      longitude: lng,
      customerId: rawCustomerId,
      customerName,
      address: safeString(row[headerMap.address]),
      zip: safeString(row[headerMap.zip]),
      chain: safeString(row[headerMap.chain]) || 'Independent',
      segment: safeString(row[headerMap.segment]) || 'Unknown',
      premise,
      currentRep,
      assignedRep,
      originalAssignedRep: assignedRep,
      overallSales: Number.isFinite(overallSales) ? overallSales : 0,
      rank,
      cadence4w,
      protected: toBoolean(row[headerMap.protected]),
      raw: row
    });
  }

  return {
    mapped,
    summary: {
      sourceRows: rows.length,
      loadedRows: mapped.length,
      skippedNoCoords,
      duplicateCustomerIds,
      missingCurrentRep,
      missingAssignedRep,
      unmappedFields
    }
  };
}

function seedFiltersFromData() {
  state.filters.rep = new Set(getAllAssignedReps());
  state.filters.rank = new Set(getDistinctValues(state.accounts, a => a.rank));
  state.filters.chain = new Set(getDistinctValues(state.accounts, a => a.chain));
  state.filters.segment = new Set(getDistinctValues(state.accounts, a => a.segment));
  state.filters.premise = 'ALL';
  state.filters.protected = 'ALL';
  state.filters.moved = 'ALL';

  fillSimpleSelect(
    els.premiseFilter,
    ['ALL', ...getDistinctValues(state.accounts, a => a.premise)],
    'ALL',
    v => v === 'ALL' ? 'All premises' : v
  );

  renderMultiFilterOptions();
}

function syncRepFilterSelection(previousAssignedReps = []) {
  const currentReps = getAllAssignedReps();
  const selected = state.filters.rep instanceof Set ? state.filters.rep : new Set();
  const prevSet = new Set(previousAssignedReps);

  for (const rep of [...selected]) {
    if (!currentReps.includes(rep)) selected.delete(rep);
  }

  for (const rep of currentReps) {
    if (!prevSet.has(rep)) selected.add(rep);
  }

  if (selected.size === 0) {
    currentReps.forEach(rep => selected.add(rep));
  }

  state.filters.rep = selected;
}

function renderMultiFilterOptions() {
  renderSingleMultiFilter('rep', getAllAssignedReps(), els.repFilterOptions, els.repFilterSummary, 'All reps');
  renderSingleMultiFilter('rank', getDistinctValues(state.accounts, a => a.rank), els.rankFilterOptions, els.rankFilterSummary, 'All ranks');
  renderSingleMultiFilter('chain', getDistinctValues(state.accounts, a => a.chain), els.chainFilterOptions, els.chainFilterSummary, 'All chains');
  renderSingleMultiFilter('segment', getDistinctValues(state.accounts, a => a.segment), els.segmentFilterOptions, els.segmentFilterSummary, 'All segments');

  if (state.openMultiKey) positionMultiPanel(state.openMultiKey);
}

function renderSingleMultiFilter(key, values, container, summaryEl, allLabel) {
  const selected = state.filters[key];
  const search = (state.multiSearch[key] || '').toLowerCase().trim();
  const visibleValues = values.filter(v => String(v).toLowerCase().includes(search));

  container.innerHTML = visibleValues.length
    ? visibleValues.map(value => {
        const checked = selected.has(value) ? 'checked' : '';
        return `
          <label class="multi-item">
            <input type="checkbox" data-multi-key="${key}" value="${escapeHtmlAttr(value)}" ${checked}/>
            <span title="${escapeHtmlAttr(value)}">${escapeHtml(value)}</span>
          </label>
        `;
      }).join('')
    : `<div class="empty">No matches.</div>`;

  container.querySelectorAll(`input[data-multi-key="${key}"]`).forEach(input => {
    input.addEventListener('change', e => {
      const val = e.target.value;

      if (e.target.checked) selected.add(val);
      else selected.delete(val);

      if (selected.size === 0) {
        values.forEach(v => selected.add(v));
      }

      updateMultiFilterSummary(key, values, summaryEl, allLabel);
      refreshUI();
      renderSingleMultiFilter(key, values, container, summaryEl, allLabel);
      positionMultiPanel(key);
    });
  });

  updateMultiFilterSummary(key, values, summaryEl, allLabel);
}

function updateMultiFilterSummary(key, values, summaryEl, allLabel) {
  const selected = state.filters[key];
  const total = values.length;
  const count = selected.size;

  if (!total || count === total) {
    summaryEl.textContent = allLabel;
    return;
  }

  if (count === 1) {
    summaryEl.textContent = [...selected][0];
    return;
  }

  summaryEl.textContent = `${count} selected`;
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
    setUploadStatus('bad', 'No valid rows found');
    return;
  }

  if (issueCount === 0) {
    setUploadStatus('good', `${(s.loadedRows || 0).toLocaleString()} loaded clean`);
    return;
  }

  const parts = [];
  if (s.skippedNoCoords) parts.push(`${s.skippedNoCoords} skipped`);
  if (s.duplicateCustomerIds) parts.push(`${s.duplicateCustomerIds} duplicate ID${s.duplicateCustomerIds === 1 ? '' : 's'}`);
  if (s.missingCurrentRep) parts.push(`${s.missingCurrentRep} blank current rep`);
  if ((s.unmappedFields || []).filter(key => isWarningField(key)).length) parts.push('unmapped fields');

  setUploadStatus('warning', `${(s.loadedRows || 0).toLocaleString()} loaded • ${parts.join(' • ')}`);
}

function renderUploadStatus() {
  const status = state.uploadStatus || { level: 'neutral', text: 'No file loaded' };

  els.uploadStatusPill.classList.remove(
    'upload-status-neutral',
    'upload-status-good',
    'upload-status-warning',
    'upload-status-bad'
  );

  if (status.level === 'good') {
    els.uploadStatusPill.classList.add('upload-status-good');
    els.uploadStatusIcon.textContent = '✓';
    els.uploadStatusText.textContent = status.text;
  } else if (status.level === 'warning') {
    els.uploadStatusPill.classList.add('upload-status-warning');
    els.uploadStatusIcon.textContent = '!';
    els.uploadStatusText.textContent = status.text;
  } else if (status.level === 'bad') {
    els.uploadStatusPill.classList.add('upload-status-bad');
    els.uploadStatusIcon.textContent = '✕';
    els.uploadStatusText.textContent = status.text;
  } else {
    els.uploadStatusPill.classList.add('upload-status-neutral');
    els.uploadStatusIcon.textContent = '•';
    els.uploadStatusText.textContent = status.text || 'No file loaded';
  }

  if (els.uploadStatusPanel && !els.uploadStatusPanel.hidden) {
    els.uploadStatusBody.innerHTML = buildUploadStatusDetailsHtml();
    positionUploadStatusPanel();
  }
}

function rebuildMarkers() {
  state.markerLayer.clearLayers();
  state.markerById.clear();

  for (const account of state.accounts) {
    const marker = L.circleMarker(
      [account.latitude, account.longitude],
      markerStyleForAccount(account)
    );

    marker.__accountId = account._id;
    syncMarkerPopupContent(marker, account._id);

    marker.on('click', () => {
      const id = marker.__accountId;
      syncMarkerPopupContent(marker, id);
      toggleSelection(id, true);
    });

    marker.on('popupopen', () => {
      syncMarkerPopupContent(marker, marker.__accountId);
    });

    state.markerLayer.addLayer(marker);
    state.markerById.set(account._id, marker);
  }

  syncMarkerZOrder();
}

function refreshMarkerStyles() {
  for (const account of state.accounts) {
    const marker = state.markerById.get(account._id);
    if (!marker) continue;

    marker.setStyle(markerStyleForAccount(account));
    marker.setRadius(state.selection.has(account._id) ? 8 : (state.repFocus && account.assignedRep === state.repFocus ? 7 : 6));
    syncMarkerPopupContent(marker, account._id);
  }

  syncMarkerZOrder();
}

function syncMarkerPopupContent(marker, accountId) {
  const account = state.accountById.get(accountId);
  if (!marker || !account) return;

  const popup = marker.getPopup();
  const html = buildPopupHtml(account);

  if (popup) {
    popup.setContent(html);
  } else {
    marker.bindPopup(html);
  }
}

function syncMarkerZOrder() {
  if (!state.markerById.size) return;

  for (const account of state.accounts) {
    const marker = state.markerById.get(account._id);
    if (!marker || typeof marker.bringToBack !== 'function') continue;

    if (!state.selection.has(account._id) && !(state.repFocus && account.assignedRep === state.repFocus)) {
      marker.bringToBack();
    }
  }

  for (const account of state.accounts) {
    const marker = state.markerById.get(account._id);
    if (!marker || typeof marker.bringToFront !== 'function') continue;

    if (state.repFocus && account.assignedRep === state.repFocus && !state.selection.has(account._id)) {
      marker.bringToFront();
    }
  }

  for (const id of state.selection) {
    const marker = state.markerById.get(id);
    if (marker && typeof marker.bringToFront === 'function') {
      marker.bringToFront();
    }
  }
}

function markerStyleForAccount(account) {
  const matches = passesFilters(account);
  const dimOthers = els.dimOthersCheckbox.checked;
  const isSelected = state.selection.has(account._id);
  const isRepFocused = !!state.repFocus;
  const isFocusedRep = isRepFocused && account.assignedRep === state.repFocus;
  const color = getRepColor(account.assignedRep);

  let opacity = matches ? 0.96 : (dimOthers ? 0.05 : 0.16);
  let fillOpacity = dimOthers && !matches ? 0.03 : (account.protected ? 0.88 : 0.74);

  if (isRepFocused) {
    if (isFocusedRep) {
      opacity = matches ? 1 : 0.4;
      fillOpacity = account.protected ? 0.95 : 0.88;
    } else {
      opacity = 0.1;
      fillOpacity = 0.05;
    }
  }

  return {
    radius: isSelected ? 8 : (isFocusedRep ? 7 : 6),
    color: isSelected ? '#1e293b' : color,
    weight: isSelected ? 2.5 : (isFocusedRep ? 2.2 : 1.2),
    opacity,
    fillColor: color,
    fillOpacity
  };
}

function passesFilters(account) {
  const repPass = state.filters.rep.size === 0 || state.filters.rep.has(account.assignedRep);
  const rankPass = state.filters.rank.size === 0 || state.filters.rank.has(account.rank);
  const chainPass = state.filters.chain.size === 0 || state.filters.chain.has(account.chain);
  const segmentPass = state.filters.segment.size === 0 || state.filters.segment.has(account.segment);
  const premisePass = state.filters.premise === 'ALL' || account.premise === state.filters.premise;

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
    els.repTableBody.innerHTML = '<tr><td colspan="14" class="empty">Upload a file to begin.</td></tr>';
    return;
  }

  syncSortHeaderIndicators();

  els.repTableBody.innerHTML = rows.map(row => `
    <tr data-rep-row="${encodeURIComponent(row.rep)}" class="${state.repFocus === row.rep ? 'rep-row-active' : ''}">
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
    </tr>
  `).join('');

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

function assignSelectionToRep() {
  const targetRep = els.assignRepSelect.value;
  const selectedIds = [...state.selection];
  if (!selectedIds.length || !targetRep) return;

  const previousAssignedReps = getAllAssignedReps();
  const changes = [];
  let skippedProtected = 0;

  ensureRepColor(targetRep);

  for (const id of selectedIds) {
    const account = state.accountById.get(id);
    if (!account) continue;

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
    showToast(skippedProtected ? `${skippedProtected} protected account(s) were skipped.` : 'No assignment changes to make.');
    return;
  }

  applyChanges(changes, `Assigned ${changes.length} account${changes.length === 1 ? '' : 's'} to ${targetRep}`, previousAssignedReps);
  clearSelection();

  if (skippedProtected) {
    showToast(`${changes.length} reassigned to ${targetRep}. ${skippedProtected} protected account(s) stayed put.`);
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
      showToast('All accounts are protected. Nothing can be optimized.');
      return;
    }

    const continuityWeight = Number(els.disruptionSlider.value) / 100;
    const geographyWeight = 1 - continuityWeight;
    const balanceMode = els.balanceMode.value;

    const protectedAccounts = state.accounts.filter(a => a.protected);
    const movableAccounts = state.accounts.filter(a => !a.protected);
    const currentReps = getAllAssignedReps();
    const targetRepNames = buildTargetRepNames(targetCount, currentReps);
    const adjacency = state.neighborMap;

    targetRepNames.forEach(rep => ensureRepColor(rep));

    const assignments = new Map();
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

    applyChanges(changes, `Optimized routes to ${targetRepNames.length} reps with minimum ${minStops} stops`, repsBefore);
    state.optimizationSummary = buildOptimizationSummary();
    updateLastActionWithOptimization(`Optimized routes to ${targetRepNames.length} reps with minimum ${minStops} stops`);
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

  ctx.addToRep = (rep, account) => addAccountToContext(ctx, rep, account);
  ctx.removeFromRep = (rep, account) => removeAccountFromContext(ctx, rep, account);
  ctx.moveAccount = (account, fromRep, toRep) => moveAccountInContext(ctx, account, fromRep, toRep);
  ctx.count = rep => ctx.reps.get(rep)?.count || 0;
  ctx.revenue = rep => ctx.reps.get(rep)?.revenue || 0;
  ctx.membersArray = rep => Array.from(ctx.reps.get(rep)?.members || []);
  ctx.centroid = rep => centroidFromContext(ctx, rep);

  ctx.clearMovableAssignments = (movableAccounts, assignmentMap) => {
    for (const account of movableAccounts) {
      const rep = assignmentMap.get(account._id);
      if (!rep) continue;
      removeAccountFromContext(ctx, rep, account);
      assignmentMap.delete(account._id);
    }
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

  const repCtx = ctx.reps.get(rep);
  if (repCtx.members.has(account._id)) return;

  repCtx.members.add(account._id);
  repCtx.count += 1;
  repCtx.revenue += account.overallSales || 0;
  repCtx.latSum += account.latitude;
  repCtx.lngSum += account.longitude;
}

function removeAccountFromContext(ctx, rep, account) {
  const repCtx = ctx.reps.get(rep);
  if (!repCtx || !repCtx.members.has(account._id)) return;

  repCtx.members.delete(account._id);
  repCtx.count -= 1;
  repCtx.revenue -= account.overallSales || 0;
  repCtx.latSum -= account.latitude;
  repCtx.lngSum -= account.longitude;
}

function moveAccountInContext(ctx, account, fromRep, toRep) {
  if (fromRep === toRep) return;
  removeAccountFromContext(ctx, fromRep, account);
  addAccountToContext(ctx, toRep, account);
}

function centroidFromContext(ctx, rep) {
  const repCtx = ctx.reps.get(rep);
  if (repCtx && repCtx.count > 0) {
    return {
      lat: repCtx.latSum / repCtx.count,
      lng: repCtx.lngSum / repCtx.count
    };
  }

  const fallback = state.accounts[0];
  return { lat: fallback?.latitude || 0, lng: fallback?.longitude || 0 };
}

function initializeCentroidsFast(targetRepNames, assignmentCtx) {
  const centroids = new Map();

  targetRepNames.forEach((rep, idx) => {
    const repCtx = assignmentCtx.reps.get(rep);
    if (repCtx && repCtx.count > 0) {
      centroids.set(rep, centroidFromContext(assignmentCtx, rep));
      return;
    }

    const fallback = state.accounts[Math.floor(idx * state.accounts.length / Math.max(1, targetRepNames.length))] || state.accounts[0];
    centroids.set(rep, { lat: fallback.latitude, lng: fallback.longitude });
  });

  return centroids;
}

function resetCentroidsFromContext(centroids, targetRepNames, assignmentCtx) {
  for (const rep of targetRepNames) {
    const repCtx = assignmentCtx.reps.get(rep);
    if (repCtx && repCtx.count > 0) {
      centroids.set(rep, centroidFromContext(assignmentCtx, rep));
    }
  }
}

function refreshCentroidsFromContext(centroids, targetRepNames, assignmentCtx) {
  for (const rep of targetRepNames) {
    const repCtx = assignmentCtx.reps.get(rep);
    if (!repCtx || !repCtx.count) continue;
    centroids.set(rep, {
      lat: repCtx.latSum / repCtx.count,
      lng: repCtx.lngSum / repCtx.count
    });
  }
}

function moveAssignmentFast(assignments, assignmentCtx, account, toRep) {
  const fromRep = assignments.get(account._id);
  if (fromRep === toRep) return false;
  assignments.set(account._id, toRep);
  assignmentCtx.moveAccount(account, fromRep, toRep);
  return true;
}

function runBorderCleanupFast(assignments, targetRepNames, continuityWeight, minStops, neighborMap, assignmentCtx) {
  let moved = true;
  let pass = 0;

  while (moved && pass < 3) {
    pass += 1;
    moved = false;

    const borderAccounts = getBorderAccounts(assignments, neighborMap)
      .filter(a => !a.protected)
      .map(a => ({ account: a, detail: dominantNeighborRep(a, assignments, neighborMap) }))
      .filter(x => x.detail && x.detail.borderStrength >= 0.6 && x.detail.rep !== assignments.get(x.account._id))
      .sort((a, b) => b.detail.borderStrength - a.detail.borderStrength);

    for (const { account, detail } of borderAccounts) {
      const from = assignments.get(account._id);
      const to = detail.rep;
      if (!to || from === to) continue;
      if (assignmentCtx.count(from) <= Math.max(1, minStops)) continue;

      const currentScore = borderCleanupScoreFast(account, from, assignments, neighborMap, continuityWeight, assignmentCtx);
      const nextScore = borderCleanupScoreFast(account, to, assignments, neighborMap, continuityWeight, assignmentCtx);

      if (nextScore + 0.18 < currentScore) {
        if (moveAssignmentFast(assignments, assignmentCtx, account, to)) moved = true;
      }
    }
  }
}

function runEnclaveCleanupFast(assignments, targetRepNames, minStops, neighborMap, assignmentCtx) {
  for (const account of getBorderAccounts(assignments, neighborMap).filter(a => !a.protected)) {
    const own = assignments.get(account._id);
    const neighbors = neighborMap.get(account._id) || [];
    if (!neighbors.length) continue;

    const repCounts = new Map();
    neighbors.forEach(nid => {
      const rep = assignments.get(nid);
      if (!rep) return;
      repCounts.set(rep, (repCounts.get(rep) || 0) + 1);
    });

    const sorted = [...repCounts.entries()].sort((a,b) => b[1] - a[1]);
    const topRep = sorted[0]?.[0];
    const topCount = sorted[0]?.[1] || 0;
    const ownCount = repCounts.get(own) || 0;

    if (!topRep || topRep === own) continue;
    if (topCount < 4 || ownCount > 1) continue;
    if (assignmentCtx.count(own) <= Math.max(1, minStops)) continue;

    moveAssignmentFast(assignments, assignmentCtx, account, topRep);
  }
}

function runMajoritySmoothingFast(assignments, targetRepNames, minStops, neighborMap, assignmentCtx) {
  const moves = [];

  for (const account of getBorderAccounts(assignments, neighborMap).filter(a => !a.protected)) {
    const own = assignments.get(account._id);
    const neighbors = neighborMap.get(account._id) || [];
    if (neighbors.length < 3) continue;
    if (assignmentCtx.count(own) <= Math.max(1, minStops)) continue;

    const repCounts = new Map();
    neighbors.forEach(nid => {
      const rep = assignments.get(nid);
      if (!rep) return;
      repCounts.set(rep, (repCounts.get(rep) || 0) + 1);
    });

    const sorted = [...repCounts.entries()].sort((a,b) => b[1] - a[1]);
    const topRep = sorted[0]?.[0];
    const topCount = sorted[0]?.[1] || 0;
    const ownCount = repCounts.get(own) || 0;

    if (topRep && topRep !== own && topCount >= 5 && ownCount <= 1) {
      moves.push({ account, from: own, to: topRep });
    }
  }

  for (const move of moves) {
    if (assignmentCtx.count(move.from) <= Math.max(1, minStops)) continue;
    moveAssignmentFast(assignments, assignmentCtx, move.account, move.to);
  }
}

function borderCleanupScoreFast(account, rep, assignments, neighborMap, continuityWeight, assignmentCtx) {
  const neighbors = neighborMap.get(account._id) || [];
  const centroid = assignmentCtx.centroid(rep);
  const dist = squaredDistance(account.latitude, account.longitude, centroid.lat, centroid.lng);

  let same = 0;
  let other = 0;
  neighbors.forEach(nid => {
    if (assignments.get(nid) === rep) same += 1;
    else other += 1;
  });

  const continuityPenalty = account.currentRep === rep ? 0 : continuityWeight * 0.35;
  return dist * 1.35 - same * 0.45 + other * 0.12 + continuityPenalty;
}

function localDominancePenalty(account, rep, assignments, neighborMap) {
  const neighbors = neighborMap.get(account._id) || [];
  if (!neighbors.length) return 0;

  let same = 0;
  let other = 0;
  const repCounts = new Map();

  neighbors.forEach(nid => {
    const nrep = assignments.get(nid);
    if (!nrep) return;
    repCounts.set(nrep, (repCounts.get(nrep) || 0) + 1);
    if (nrep === rep) same += 1;
    else other += 1;
  });

  const sorted = [...repCounts.entries()].sort((a,b) => b[1] - a[1]);
  const dominantRep = sorted[0]?.[0];
  const dominantCount = sorted[0]?.[1] || 0;

  let penalty = 0;
  if (dominantRep && dominantRep !== rep) penalty += dominantCount * 0.48;
  if (same === 0) penalty += 2.4;
  if (same <= 1 && neighbors.length >= 4) penalty += 1.25;
  if (other > same) penalty += (other - same) * 0.28;
  return penalty;
}

function dominantNeighborRep(account, assignments, neighborMap) {
  const neighbors = neighborMap.get(account._id) || [];
  if (!neighbors.length) return null;

  const counts = new Map();
  neighbors.forEach(id => {
    const rep = assignments.get(id);
    if (!rep) return;
    counts.set(rep, (counts.get(rep) || 0) + 1);
  });

  const sorted = [...counts.entries()].sort((a,b) => b[1] - a[1]);
  if (!sorted.length) return null;

  return {
    rep: sorted[0][0],
    borderStrength: sorted[0][1] / neighbors.length
  };
}

function getBorderAccounts(assignments, neighborMap) {
  const out = [];

  for (const account of state.accounts) {
    const neighbors = neighborMap.get(account._id) || [];
    if (!neighbors.length) continue;

    const own = assignments.get(account._id);
    let mixed = false;

    for (const nid of neighbors) {
      const rep = assignments.get(nid);
      if (rep && rep !== own) {
        mixed = true;
        break;
      }
    }

    if (mixed) out.push(account);
  }

  return out;
}

function enforceMinimumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx) {
  let pass = 0;
  const maxPasses = 1000;

  while (pass < maxPasses) {
    pass += 1;

    const underfilled = targetRepNames
      .filter(rep => assignmentCtx.count(rep) < minStops)
      .sort((a,b) => assignmentCtx.count(a) - assignmentCtx.count(b));

    if (!underfilled.length) break;

    let movedThisPass = false;

    for (const needyRep of underfilled) {
      const needyCentroid = assignmentCtx.centroid(needyRep);

      const donorCandidates = targetRepNames
        .filter(rep => rep !== needyRep && assignmentCtx.count(rep) > minStops)
        .sort((a,b) => assignmentCtx.count(b) - assignmentCtx.count(a));

      let bestMove = null;

      for (const donorRep of donorCandidates) {
        const donorMemberIds = assignmentCtx.membersArray(donorRep);

        for (const accountId of donorMemberIds) {
          const account = state.accountById.get(accountId);
          if (!account || account.protected) continue;
          if (assignmentCtx.count(needyRep) >= maxStops) continue;
          if (assignmentCtx.count(donorRep) <= minStops) continue;

          const dist = squaredDistance(account.latitude, account.longitude, needyCentroid.lat, needyCentroid.lng);
          const continuityPenalty = account.currentRep === donorRep ? 0.22 : 0;
          const score = dist + continuityPenalty;

          if (!bestMove || score < bestMove.score) {
            bestMove = { account, from: donorRep, to: needyRep, score };
          }
        }
      }

      if (bestMove) {
        if (moveAssignmentFast(assignments, assignmentCtx, bestMove.account, bestMove.to)) {
          movedThisPass = true;
        }
      }
    }

    if (!movedThisPass) break;
  }
}

function enforceMaximumStopsFast(assignments, targetRepNames, minStops, maxStops, assignmentCtx) {
  let pass = 0;
  const maxPasses = 1000;

  while (pass < maxPasses) {
    pass += 1;

    const overfilled = targetRepNames
      .filter(rep => assignmentCtx.count(rep) > maxStops)
      .sort((a, b) => assignmentCtx.count(b) - assignmentCtx.count(a));

    if (!overfilled.length) break;

    let movedThisPass = false;

    for (const donorRep of overfilled) {
      const donorCentroid = assignmentCtx.centroid(donorRep);
      const donorAccounts = assignmentCtx.membersArray(donorRep)
        .map(id => state.accountById.get(id))
        .filter(a => a && !a.protected)
        .sort((a, b) => {
          const ad = squaredDistance(a.latitude, a.longitude, donorCentroid.lat, donorCentroid.lng);
          const bd = squaredDistance(b.latitude, b.longitude, donorCentroid.lat, donorCentroid.lng);
          return bd - ad;
        });

      let bestMove = null;

      for (const account of donorAccounts) {
        for (const targetRep of targetRepNames) {
          if (targetRep === donorRep) continue;
          if (assignmentCtx.count(targetRep) >= maxStops) continue;

          const targetCentroid = assignmentCtx.centroid(targetRep);
          const score =
            squaredDistance(account.latitude, account.longitude, targetCentroid.lat, targetCentroid.lng) +
            (account.currentRep === targetRep ? 0 : 0.25) +
            (assignmentCtx.count(targetRep) < minStops ? -1.5 : 0);

          if (!bestMove || score < bestMove.score) {
            bestMove = {
              account,
              from: donorRep,
              to: targetRep,
              score
            };
          }
        }
      }

      if (bestMove) {
        if (moveAssignmentFast(assignments, assignmentCtx, bestMove.account, bestMove.to)) {
          movedThisPass = true;
        }
      }
    }

    if (!movedThisPass) break;
  }
}

function rebalanceStopTargetsStrict(assignments, targetRepNames, minStops, maxStops, assignmentCtx) {
  let moved = true;
  let safety = 0;

  while (moved && safety < 2000) {
    safety += 1;
    moved = false;

    const underfilled = targetRepNames
      .filter(rep => assignmentCtx.count(rep) < minStops)
      .sort((a, b) => assignmentCtx.count(a) - assignmentCtx.count(b));

    const overfilled = targetRepNames
      .filter(rep => assignmentCtx.count(rep) > maxStops)
      .sort((a, b) => assignmentCtx.count(b) - assignmentCtx.count(a));

    if (!underfilled.length && !overfilled.length) break;

    for (const needyRep of underfilled) {
      let bestMove = null;
      const needyCentroid = assignmentCtx.centroid(needyRep);

      const donorReps = targetRepNames
        .filter(rep => rep !== needyRep && assignmentCtx.count(rep) > minStops)
        .sort((a, b) => assignmentCtx.count(b) - assignmentCtx.count(a));

      for (const donorRep of donorReps) {
        const donorIds = assignmentCtx.membersArray(donorRep);

        for (const accountId of donorIds) {
          const account = state.accountById.get(accountId);
          if (!account || account.protected) continue;
          if (assignmentCtx.count(donorRep) <= minStops) continue;
          if (assignmentCtx.count(needyRep) >= maxStops) continue;

          const score =
            squaredDistance(account.latitude, account.longitude, needyCentroid.lat, needyCentroid.lng) +
            (account.currentRep === needyRep ? -0.15 : 0) +
            (account.currentRep === donorRep ? 0.15 : 0);

          if (!bestMove || score < bestMove.score) {
            bestMove = {
              account,
              from: donorRep,
              to: needyRep,
              score
            };
          }
        }
      }

      if (bestMove) {
        if (moveAssignmentFast(assignments, assignmentCtx, bestMove.account, bestMove.to)) {
          moved = true;
        }
      }
    }

    for (const donorRep of overfilled) {
      let bestMove = null;

      const donorIds = assignmentCtx.membersArray(donorRep);
      for (const accountId of donorIds) {
        const account = state.accountById.get(accountId);
        if (!account || account.protected) continue;
        if (assignmentCtx.count(donorRep) <= maxStops) continue;

        for (const targetRep of targetRepNames) {
          if (targetRep === donorRep) continue;
          if (assignmentCtx.count(targetRep) >= maxStops) continue;

          const targetCentroid = assignmentCtx.centroid(targetRep);
          const score =
            squaredDistance(account.latitude, account.longitude, targetCentroid.lat, targetCentroid.lng) +
            (assignmentCtx.count(targetRep) < minStops ? -2.5 : 0) +
            (account.currentRep === targetRep ? -0.15 : 0);

          if (!bestMove || score < bestMove.score) {
            bestMove = {
              account,
              from: donorRep,
              to: targetRep,
              score
            };
          }
        }
      }

      if (bestMove) {
        if (moveAssignmentFast(assignments, assignmentCtx, bestMove.account, bestMove.to)) {
          moved = true;
        }
      }
    }
  }
}

function buildNeighborMap(accounts) {
  const map = new Map();

  for (let i = 0; i < accounts.length; i += 1) {
    const account = accounts[i];
    const nearest = [];

    for (let j = 0; j < accounts.length; j += 1) {
      if (i === j) continue;
      const other = accounts[j];
      const d = squaredDistance(account.latitude, account.longitude, other.latitude, other.longitude);

      if (nearest.length < 8) {
        nearest.push({ id: other._id, d });
        nearest.sort((a, b) => a.d - b.d);
      } else if (d < nearest[nearest.length - 1].d) {
        nearest[nearest.length - 1] = { id: other._id, d };
        nearest.sort((a, b) => a.d - b.d);
      }
    }

    map.set(account._id, nearest.map(x => x.id));
  }

  return map;
}

function buildTargetRepNames(targetCount, currentReps) {
  const reps = [...currentReps];
  while (reps.length < targetCount) reps.push(`Rep ${reps.length + 1}`);
  if (reps.length > targetCount) return reps.slice(0, targetCount);
  return reps.length ? reps : ['Rep 1'];
}

function buildFullRepStats(reps) {
  const map = new Map();
  reps.forEach(rep => map.set(rep, { rep, stops: 0, revenue: 0 }));
  return map;
}

function refreshTerritories() {
  if (state.territoryLayer) state.territoryLayer.clearLayers();
  if (state.territoryLabelLayer) state.territoryLabelLayer.clearLayers();
  if (!els.showTerritoryCheckbox.checked || !state.accounts.length || !state.map) return;

  const reps = getAllAssignedReps().filter(rep => state.filters.rep.has(rep));
  const zoom = state.map.getZoom();
  const allowAllLabels = zoom >= 9;
  const allowFocusedOnly = !!state.repFocus && zoom >= 7;

  reps.forEach(rep => {
    const members = state.accounts.filter(a => a.assignedRep === rep);
    if (members.length < 3) return;

    const territoryFeature = buildTerritoryFeatureForRep(rep, members);
    if (!territoryFeature) return;

    const isFocusedRep = state.repFocus && rep === state.repFocus;
    addTerritoryFeatureToMap(territoryFeature, rep, isFocusedRep);

    const shouldLabel = allowAllLabels || (allowFocusedOnly && isFocusedRep);
    if (!shouldLabel) return;

    const rings = getFeatureOuterRings(territoryFeature);
    if (!rings.length) return;

    const labelRing = pickBestLabelRing(rings);
    if (!labelRing) return;
    if (!territoryHasEnoughScreenRoom(labelRing, isFocusedRep)) return;

    const centroid = getHullLabelLatLng(territoryFeature, labelRing);
    if (!centroid) return;

    const labelHtml = `
      <div class="territory-label-inner ${isFocusedRep ? 'territory-label-focused' : ''}">
        <div class="territory-label-name">${escapeHtml(rep)}</div>
      </div>
    `;

    const marker = L.marker(centroid, {
      interactive: false,
      keyboard: false,
      zIndexOffset: isFocusedRep ? 800 : 500,
      icon: L.divIcon({
        className: 'territory-label',
        html: labelHtml,
        iconSize: null
      })
    });

    marker.addTo(state.territoryLabelLayer);
  });
}

function buildTerritoryFeatureForRep(rep, members) {
  const trimmedMembers = trimTerritoryOutliers(members);
  const points = trimmedMembers.map(a => [a.longitude, a.latitude]);
  if (points.length < 3) return null;

  const featureCollection = turf.featureCollection(points.map(p => turf.point(p)));

  let feature = null;

  try {
    feature = turf.concave(featureCollection, {
      maxEdge: getConcaveMaxEdgeKm(trimmedMembers),
      units: 'kilometers'
    });
  } catch (e) {
    feature = null;
  }

  if (!isUsableTerritoryFeature(feature)) {
    try {
      feature = turf.convex(featureCollection);
    } catch (e) {
      feature = null;
    }
  }

  if (!isUsableTerritoryFeature(feature)) return null;
  return feature;
}

function trimTerritoryOutliers(members) {
  if (!members || members.length <= 6) return members;

  const centroid = members.reduce((acc, a) => {
    acc.lat += a.latitude;
    acc.lng += a.longitude;
    return acc;
  }, { lat: 0, lng: 0 });

  centroid.lat /= members.length;
  centroid.lng /= members.length;

  const distances = members.map(account => ({
    account,
    d: squaredDistance(account.latitude, account.longitude, centroid.lat, centroid.lng)
  })).sort((a, b) => a.d - b.d);

  const maxTrimCount = Math.min(2, Math.floor(members.length * 0.08));
  if (maxTrimCount <= 0) return members;

  const keepCount = Math.max(4, distances.length - maxTrimCount);
  const kept = distances.slice(0, keepCount);

  const keptMax = kept[kept.length - 1]?.d || 0;
  const trimmed = distances.slice(keepCount);

  if (!trimmed.length) return members;

  const shouldTrim = trimmed.every(x => x.d > keptMax * 1.35);
  return shouldTrim ? kept.map(x => x.account) : members;
}

function getConcaveMaxEdgeKm(members) {
  if (!members || members.length < 2) return 20;

  let minLat = Infinity;
  let maxLat = -Infinity;
  let minLng = Infinity;
  let maxLng = -Infinity;

  members.forEach(a => {
    if (a.latitude < minLat) minLat = a.latitude;
    if (a.latitude > maxLat) maxLat = a.latitude;
    if (a.longitude < minLng) minLng = a.longitude;
    if (a.longitude > maxLng) maxLng = a.longitude;
  });

  const latKm = Math.abs(maxLat - minLat) * 111;
  const avgLatRad = ((minLat + maxLat) / 2) * Math.PI / 180;
  const lngKm = Math.abs(maxLng - minLng) * 111 * Math.max(0.25, Math.cos(avgLatRad));
  const spanKm = Math.sqrt((latKm * latKm) + (lngKm * lngKm));

  return Math.max(4, Math.min(28, spanKm * 0.42));
}

function isUsableTerritoryFeature(feature) {
  if (!feature || !feature.geometry) return false;
  return feature.geometry.type === 'Polygon' || feature.geometry.type === 'MultiPolygon';
}

function addTerritoryFeatureToMap(feature, rep, isFocusedRep) {
  const color = getRepColor(rep);

  const style = {
    color,
    weight: isFocusedRep ? 3 : 2,
    fillOpacity: isFocusedRep ? 0.1 : 0.05,
    opacity: state.repFocus ? (isFocusedRep ? 0.95 : 0.18) : 0.75,
    interactive: false
  };

  L.geoJSON(feature, { style }).addTo(state.territoryLayer);
}

function getFeatureOuterRings(feature) {
  if (!feature?.geometry) return [];

  if (feature.geometry.type === 'Polygon') {
    const outer = feature.geometry.coordinates?.[0];
    return outer ? [outer.map(([lng, lat]) => [lat, lng])] : [];
  }

  if (feature.geometry.type === 'MultiPolygon') {
    return feature.geometry.coordinates
      .map(poly => poly?.[0] ? poly[0].map(([lng, lat]) => [lat, lng]) : null)
      .filter(Boolean);
  }

  return [];
}

function pickBestLabelRing(rings) {
  if (!rings?.length) return null;
  return [...rings].sort((a, b) => polygonRingAreaEstimate(b) - polygonRingAreaEstimate(a))[0] || null;
}

function polygonRingAreaEstimate(coords) {
  if (!coords || coords.length < 3) return 0;

  let minLat = Infinity;
  let maxLat = -Infinity;
  let minLng = Infinity;
  let maxLng = -Infinity;

  coords.forEach(([lat, lng]) => {
    if (lat < minLat) minLat = lat;
    if (lat > maxLat) maxLat = lat;
    if (lng < minLng) minLng = lng;
    if (lng > maxLng) maxLng = lng;
  });

  return Math.abs((maxLat - minLat) * (maxLng - minLng));
}

function territoryHasEnoughScreenRoom(coords, isFocusedRep) {
  if (!state.map || !coords?.length) return false;

  const bounds = L.latLngBounds(coords);
  const nw = state.map.latLngToContainerPoint(bounds.getNorthWest());
  const se = state.map.latLngToContainerPoint(bounds.getSouthEast());

  const width = Math.abs(se.x - nw.x);
  const height = Math.abs(se.y - nw.y);

  if (isFocusedRep) return width >= 70 && height >= 26;
  return width >= 95 && height >= 34;
}

function getHullLabelLatLng(feature, coords) {
  try {
    const centroid = turf.centroid(feature);
    const [lng, lat] = centroid.geometry.coordinates;
    if (Number.isFinite(lat) && Number.isFinite(lng)) return [lat, lng];
  } catch (e) {}

  try {
    const bounds = L.latLngBounds(coords);
    const center = bounds.getCenter();
    return [center.lat, center.lng];
  } catch (e) {}

  return null;
}

function summarizeByRep() {
  const currentMap = new Map();
  const baselineMap = new Map();
  const reps = getAllKnownReps();

  reps.forEach(rep => {
    currentMap.set(rep, makeEmptyRepSummary(rep));
    baselineMap.set(rep, makeEmptyRepSummary(rep));
  });

  for (const account of state.accounts) {
    if (!currentMap.has(account.assignedRep)) currentMap.set(account.assignedRep, makeEmptyRepSummary(account.assignedRep));
    if (!baselineMap.has(account.originalAssignedRep)) baselineMap.set(account.originalAssignedRep, makeEmptyRepSummary(account.originalAssignedRep));

    const cur = currentMap.get(account.assignedRep);
    cur.stops += 1;
    cur.revenue += account.overallSales;
    cur[account.rank] += 1;
    cur.planned4W += Number(account.cadence4w || 0);
    cur.avgWeekly = cur.planned4W / 4;
    if (account.protected) cur.protected += 1;
    if (account.assignedRep !== account.originalAssignedRep) cur.movedIn += 1;

    const base = baselineMap.get(account.originalAssignedRep);
    base.stops += 1;
    base.revenue += account.overallSales;
  }

  for (const account of state.accounts) {
    if (account.assignedRep !== account.originalAssignedRep) {
      if (!currentMap.has(account.originalAssignedRep)) currentMap.set(account.originalAssignedRep, makeEmptyRepSummary(account.originalAssignedRep));
      currentMap.get(account.originalAssignedRep).movedOut += 1;
    }
  }

  const rows = [...currentMap.values()].map(row => {
    const base = baselineMap.get(row.rep) || makeEmptyRepSummary(row.rep);
    row.avgWeekly = row.planned4W / 4;

    return {
      ...row,
      color: getRepColor(row.rep),
      deltaStops: row.stops - base.stops,
      deltaRevenue: row.revenue - base.revenue
    };
  });

  return rows.sort((a,b) => a.rep.localeCompare(b.rep, undefined, { numeric:true }));
}

function makeEmptyRepSummary(rep) {
  return {
    rep,
    color: getRepColor(rep),
    stops: 0,
    revenue: 0,
    A: 0,
    B: 0,
    C: 0,
    D: 0,
    planned4W: 0,
    avgWeekly: 0,
    protected: 0,
    movedIn: 0,
    movedOut: 0,
    deltaStops: 0,
    deltaRevenue: 0
  };
}

async function exportWorkbook() {
  if (!state.accounts.length) return;

  if (typeof ExcelJS === 'undefined') {
    showToast('Export styling library is not available. Refresh and try again.');
    return;
  }

  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'ChatGPT';
    workbook.created = new Date();
    workbook.modified = new Date();

    const timestamp = buildTimestampForFileName();
    const outputName = (state.loadedFileName || 'territory_export_updated.xlsx').replace(/\.xlsx$/i, `_${timestamp}.xlsx`);

    const assignments = buildAssignmentsExportRows();
    const repSummary = buildRepSummaryExportRows();
    const movedAccounts = buildMovedAccountsExportRows();
    const changeLog = buildChangeLogExportRows();
    const runSettings = buildRunSettingsExportRows();

    addStyledWorksheet(workbook, {
      name: 'Assignments',
      rows: assignments,
      keyOrder: [
        'Customer_ID','Customer_Name','Address','ZIP','Chain','Segment','Premise',
        'Current_Rep','Assigned_Rep','Original_Assigned_Rep','Protected','Moved',
        'Rank','Latitude','Longitude','Overall_Sales','Cadence_4W'
      ],
      currencyColumns: ['Overall_Sales'],
      decimalColumns: ['Latitude','Longitude','Cadence_4W'],
      freezeTopRow: true,
      autofilter: true
    });

    addStyledWorksheet(workbook, {
      name: 'Rep Summary',
      rows: repSummary,
      keyOrder: [
        'Rep','Stops','Delta_Stops','Revenue','Delta_Revenue','A','B','C','D',
        'Planned_4W','Avg_Weekly','Protected','Moved_In','Moved_Out'
      ],
      currencyColumns: ['Revenue','Delta_Revenue'],
      integerColumns: ['Stops','Delta_Stops','A','B','C','D','Protected','Moved_In','Moved_Out'],
      decimalColumns: ['Planned_4W','Avg_Weekly'],
      freezeTopRow: true,
      autofilter: true
    });

    addStyledWorksheet(workbook, {
      name: 'Moved Accounts',
      rows: movedAccounts.length ? movedAccounts : [{}],
      keyOrder: [
        'Customer_ID','Customer_Name','Original_Assigned_Rep','Assigned_Rep','Current_Rep','Revenue','Rank','Protected'
      ],
      currencyColumns: ['Revenue'],
      freezeTopRow: true,
      autofilter: true
    });

    addStyledWorksheet(workbook, {
      name: 'Run Settings',
      rows: runSettings,
      keyOrder: ['Setting','Value'],
      freezeTopRow: true,
      autofilter: true
    });

    addStyledWorksheet(workbook, {
      name: 'Change Log',
      rows: changeLog.length ? changeLog : [{}],
      keyOrder: ['timestamp','customerId','customerName','fromRep','toRep','protected'],
      freezeTopRow: true,
      autofilter: true
    });

    const buffer = await workbook.xlsx.writeBuffer();
    downloadArrayBufferAsFile(buffer, outputName);

    showToast('Styled Excel export created.');
    updateLastAction('Exported styled workbook');
  } catch (err) {
    console.error('Styled export failed:', err);
    showToast('Export hit an error. Open the browser console and send the first red line.');
  }
}

function buildAssignmentsExportRows() {
  return state.accounts.map(a => ({
    Customer_ID: a.customerId,
    Customer_Name: a.customerName,
    Address: a.address,
    ZIP: a.zip,
    Chain: a.chain,
    Segment: a.segment,
    Premise: a.premise,
    Current_Rep: a.currentRep,
    Assigned_Rep: a.assignedRep,
    Original_Assigned_Rep: a.originalAssignedRep,
    Protected: a.protected ? 'Yes' : 'No',
    Moved: a.assignedRep !== a.originalAssignedRep ? 'Yes' : 'No',
    Rank: a.rank,
    Latitude: round6(a.latitude),
    Longitude: round6(a.longitude),
    Overall_Sales: round2(a.overallSales),
    Cadence_4W: round2(a.cadence4w)
  }));
}

function buildRepSummaryExportRows() {
  return summarizeByRep().map(r => ({
    Rep: r.rep,
    Stops: r.stops,
    Delta_Stops: r.deltaStops,
    Revenue: round2(r.revenue),
    Delta_Revenue: round2(r.deltaRevenue),
    A: r.A,
    B: r.B,
    C: r.C,
    D: r.D,
    Planned_4W: round2(r.planned4W),
    Avg_Weekly: round2(r.avgWeekly),
    Protected: r.protected,
    Moved_In: r.movedIn,
    Moved_Out: r.movedOut
  }));
}

function buildMovedAccountsExportRows() {
  return state.accounts
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
}

function buildRunSettingsExportRows() {
  const s = state.importSummary || {};
  const o = state.optimizationSummary || null;

  return [
    { Setting: 'Export Timestamp', Value: new Date().toLocaleString() },
    { Setting: 'Loaded File Name', Value: state.loadedFileName || '' },
    { Setting: 'Sheet Name', Value: state.currentSheetName || '' },
    { Setting: 'Target Rep Count', Value: els.repCountInput.value || '' },
    { Setting: 'Minimum Stops / Rep', Value: els.minStopsInput.value || '' },
    { Setting: 'Maximum Stops / Rep', Value: els.maxStopsInput.value || '' },
    { Setting: 'Optimize By', Value: els.balanceMode.value || '' },
    { Setting: 'Customer Disruption', Value: els.disruptionSlider.value || '' },
    { Setting: 'Source Rows', Value: s.sourceRows || 0 },
    { Setting: 'Loaded Rows', Value: s.loadedRows || 0 },
    { Setting: 'Skipped Missing Lat/Long', Value: s.skippedNoCoords || 0 },
    { Setting: 'Duplicate IDs Adjusted', Value: s.duplicateCustomerIds || 0 },
    { Setting: 'Missing Current Rep', Value: s.missingCurrentRep || 0 },
    { Setting: 'Missing Assigned Rep', Value: s.missingAssignedRep || 0 },
    { Setting: 'Optimization Summary', Value: o ? `${o.repCount} reps / ${o.movedCount} moved / stops ${o.minStops}-${o.maxStops} / avg ${formatNumber(o.avgStops, 1)}` : 'No optimization run' }
  ];
}

function buildChangeLogExportRows() {
  return state.changeLog.length
    ? state.changeLog.map(x => ({ ...x }))
    : [{
        timestamp: '',
        customerId: '',
        customerName: '',
        fromRep: '',
        toRep: '',
        protected: ''
      }];
}

function addStyledWorksheet(workbook, options) {
  const {
    name,
    rows,
    keyOrder = [],
    currencyColumns = [],
    decimalColumns = [],
    integerColumns = [],
    freezeTopRow = true,
    autofilter = true
  } = options;

  const worksheet = workbook.addWorksheet(name, {
    views: freezeTopRow ? [{ state: 'frozen', ySplit: 1 }] : []
  });

  const normalizedRows = Array.isArray(rows) && rows.length ? rows : [{}];
  const keys = keyOrder.length ? keyOrder : Object.keys(normalizedRows[0] || {});

  worksheet.columns = keys.map(key => ({
    header: prettifyHeader(key),
    key,
    width: guessColumnWidth(key, normalizedRows)
  }));

  normalizedRows.forEach(row => {
    const cleanRow = {};
    keys.forEach(key => {
      cleanRow[key] = row[key] ?? '';
    });
    worksheet.addRow(cleanRow);
  });

  styleWorksheetBase(worksheet);
  styleHeaderRow(worksheet.getRow(1));
  applyColumnFormats(worksheet, { currencyColumns, decimalColumns, integerColumns });

  if (autofilter && worksheet.rowCount >= 1 && worksheet.columnCount >= 1) {
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: worksheet.columnCount }
    };
  }

  worksheet.eachRow((row, rowNumber) => {
    row.height = rowNumber === 1 ? 18 : 15;
  });
}

function styleWorksheetBase(worksheet) {
  worksheet.eachRow({ includeEmpty: true }, row => {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font = {
        name: 'Tw Cen MT',
        size: 9,
        bold: false,
        color: { argb: 'FF000000' }
      };

      cell.alignment = {
        vertical: 'middle',
        horizontal: 'left'
      };

      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFE6EDF5' } }
      };
    });
  });
}

function styleHeaderRow(row) {
  row.eachCell(cell => {
    cell.font = {
      name: 'Tw Cen MT',
      size: 9,
      bold: true,
      color: { argb: 'FF20364F' }
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFEAF2FB' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'left'
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD3DFEB' } },
      bottom: { style: 'thin', color: { argb: 'FFC3D3E2' } }
    };
  });
}

function applyColumnFormats(worksheet, formatConfig) {
  const {
    currencyColumns = [],
    decimalColumns = [],
    integerColumns = []
  } = formatConfig;

  worksheet.columns.forEach(column => {
    if (!column || !column.key) return;

    if (currencyColumns.includes(column.key)) {
      column.numFmt = '$#,##0;[Red]($#,##0)';
    } else if (decimalColumns.includes(column.key)) {
      column.numFmt = '0.00';
    } else if (integerColumns.includes(column.key)) {
      column.numFmt = '0';
    }
  });
}

function guessColumnWidth(key, rows) {
  const pretty = prettifyHeader(key);
  let maxLen = pretty.length;

  rows.slice(0, 250).forEach(row => {
    const value = row[key];
    const len = String(value == null ? '' : value).length;
    if (len > maxLen) maxLen = len;
  });

  const widthOverrides = {
    Customer_ID: 16,
    Customer_Name: 28,
    Address: 30,
    ZIP: 12,
    Chain: 18,
    Segment: 16,
    Premise: 14,
    Current_Rep: 16,
    Assigned_Rep: 16,
    Original_Assigned_Rep: 20,
    Protected: 11,
    Moved: 10,
    Rank: 8,
    Latitude: 12,
    Longitude: 12,
    Overall_Sales: 14,
    Cadence_4W: 12,
    Revenue: 14,
    Delta_Revenue: 14,
    Planned_4W: 12,
    Avg_Weekly: 12,
    timestamp: 24,
    customerId: 16,
    customerName: 28,
    fromRep: 16,
    toRep: 16,
    protected: 11,
    Setting: 28,
    Value: 42
  };

  if (widthOverrides[key]) return widthOverrides[key];
  return Math.max(10, Math.min(maxLen + 2, 42));
}

function prettifyHeader(key) {
  return String(key || '')
    .replace(/_/g, ' ')
    .replace(/\b\w/g, m => m.toUpperCase());
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

function buildTimestampForFileName() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  return `${yyyy}${mm}${dd}_${hh}${mi}`;
}

function fitMapToAccounts() {
  if (!state.accounts.length) return;
  const latlngs = state.accounts.map(a => [a.latitude, a.longitude]);
  state.map.fitBounds(latlngs, { padding: [25, 25] });
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

  state.map.fitBounds(bounds, {
    padding: [35, 35],
    maxZoom: 11
  });
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
  return [...set].sort((a,b) => a.localeCompare(b, undefined, { numeric:true }));
}

function getAllKnownReps() {
  const set = new Set();
  state.accounts.forEach(a => {
    if (a.assignedRep) set.add(a.assignedRep);
    if (a.currentRep) set.add(a.currentRep);
    if (a.originalAssignedRep) set.add(a.originalAssignedRep);
  });
  return [...set].sort((a,b) => a.localeCompare(b, undefined, { numeric:true }));
}

function fillSimpleSelect(selectEl, values, selectedValue, labelFn = v => v) {
  selectEl.innerHTML = values.map(v => `<option value="${escapeHtmlAttr(v)}">${escapeHtml(labelFn(v))}</option>`).join('');
  if (selectedValue != null && values.includes(selectedValue)) selectEl.value = selectedValue;
}

function updateLastAction(text) {
  state.lastAction = text;
  els.lastAction.textContent = text;
}

function showToast(message) {
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
  return [...set].sort((a,b) => String(a).localeCompare(String(b), undefined, { numeric:true }));
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

  if (normalized === 'weekly' || normalized === 'wkly' || normalized === 'week') return 4;
  if (normalized === 'biweekly' || normalized === 'every other week' || normalized === 'eow' || normalized === 'bi-weekly') return 2;
  if (normalized === 'monthly' || normalized === 'month') return 1;
  if (normalized === 'quarterly' || normalized === 'qtr' || normalized === 'quarter') return 0.33;

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
  return { A:0, B:1, C:2, D:3 }[rank] ?? 9;
}

function toCamel(id) {
  return id.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
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
  return ['true','yes','y','1','protected','locked'].includes(v);
}

function normalizeRank(value) {
  const raw = safeString(value).toUpperCase();
  return ['A','B','C','D'].includes(raw) ? raw : 'C';
}

function formatCurrency(value) {
  return new Intl.NumberFormat('en-US', {
    style:'currency',
    currency:'USD',
    maximumFractionDigits:0
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
    '&':'&amp;',
    '<':'&lt;',
    '>':'&gt;',
    '"':'&quot;',
    "'":'&#39;'
  }[m]));
}

function escapeHtmlAttr(text) {
  return escapeHtml(text).replace(/"/g, '&quot;');
}
