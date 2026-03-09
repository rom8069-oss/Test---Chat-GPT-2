const COLOR_PALETTE = [
  '#1f77b4','#d62728','#2ca02c','#9467bd','#ff7f0e','#17becf','#8c564b','#e377c2','#7f7f7f','#bcbd22',
  '#0b7285','#c92a2a','#2b8a3e','#5f3dc4','#e67700','#087f5b','#364fc7','#a61e4d','#495057','#2f9e44',
  '#f03e3e','#3b5bdb','#e8590c','#1098ad','#9c36b5','#5c940d','#d9480f','#1864ab','#c2255c','#12b886'
];

const RANK_WEIGHTS = { A: 4, B: 2, C: 1, D: 0.33 };

const COLUMN_ALIASES = {
  latitude: ['latitude','lat','y','geo_lat','customer_latitude'],
  longitude: ['longitude','lng','lon','x','geo_longitude','customer_longitude'],
  customerId: ['cust id','customer id','customerid','id','account id','acct id'],
  customerName: ['customer name','name','account name','cust name'],
  address: ['address','street address','addr','full address'],
  zip: ['zip','zip code','zipcode','postal code'],
  chain: ['chain','chain name'],
  segment: ['segment','customer segment'],
  currentRep: ['current rep','rep','sales rep','territory rep','owner rep'],
  assignedRep: ['assigned rep','new rep','territory','route','assigned territory'],
  overallSales: ['overall sales','sales','total sales','revenue','$ revenue','$ vol sept - feb','overall revenue'],
  wineSales: ['wine sales','wine','$ wine','wine revenue'],
  spiritsSales: ['spirits sales','spirits','$ spirits','spirits revenue'],
  thcSales: ['thc sales','thc','$ thc','thc revenue'],
  rank: ['rank','class','priority rank','visit class','frequency','rotation','abcdrank','route class'],
  protected: ['protected','protected account','locked','do not move','never move']
};

const state = {
  map: null,
  lightLayer: null,
  darkLayer: null,
  markerLayer: null,
  territoryLayer: null,
  drawLayer: null,
  drawControl: null,
  workbook: null,
  workbookSheets: {},
  accounts: [],
  markerById: new Map(),
  selection: new Set(),
  undoStack: [],
  changeLog: [],
  repColors: new Map(),
  repFocus: null,
  theme: 'light',
  loadedFileName: 'territory_export.xlsx',
  lastAction: 'No actions yet'
};

const els = {};

document.addEventListener('DOMContentLoaded', init);

function init() {
  bindElements();
  initMap();
  bindEvents();
  updateLastAction('No actions yet');
}

function bindElements() {
  [
    'file-input','sheet-select','load-sheet-btn','assign-btn','undo-btn','reset-btn','optimize-btn','export-btn','theme-toggle',
    'assign-rep-select','rep-filter','rep-count-input','min-stops-input','disruption-slider','disruption-value','balance-mode','rank-filter',
    'dim-others-checkbox','show-territory-checkbox','rep-table-body','selection-preview','selection-count','clear-selection-btn',
    'global-accounts','global-revenue','global-protected','global-moved','global-unchanged','global-workload','map-legend','last-action','toast'
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
  els.themeToggle.addEventListener('click', toggleTheme);

  els.repFilter.addEventListener('change', () => {
    state.repFocus = null;
    refreshUI();
  });

  els.rankFilter.addEventListener('change', refreshUI);
  els.dimOthersCheckbox.addEventListener('change', refreshUI);
  els.showTerritoryCheckbox.addEventListener('change', refreshTerritories);
  els.clearSelectionBtn.addEventListener('click', clearSelection);

  els.disruptionSlider.addEventListener('input', () => {
    els.disruptionValue.textContent = els.disruptionSlider.value;
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
}

function onFileChosen(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  state.loadedFileName = file.name.replace(/\.[^.]+$/, '') + '_updated.xlsx';
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
      showToast('File could not be read. Check the format and try again.');
    }
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
  enableTopControls();
  loadSelectedSheet();
}

function loadSelectedSheet() {
  const sheetName = els.sheetSelect.value;
  if (!sheetName || !state.workbookSheets[sheetName]) {
    showToast('No sheet selected.');
    return;
  }

  const rows = state.workbookSheets[sheetName];
  const normalized = normalizeRows(rows);

  if (!normalized.length) {
    state.accounts = [];
    refreshUI();
    showToast('No valid rows found in that sheet.');
    return;
  }

  state.accounts = normalized;
  state.undoStack = [];
  state.changeLog = [];
  state.selection.clear();
  state.repFocus = null;

  buildRepColors();
  refreshUI(true);
  fitMapToAccounts();
  showToast(`Loaded ${state.accounts.length} accounts from ${sheetName}.`);
  updateLastAction(`Loaded ${state.accounts.length} accounts from ${sheetName}`);
}

function normalizeRows(rows) {
  if (!rows?.length) return [];

  const mapped = [];
  const headers = Object.keys(rows[0]);
  const cleanedHeaderLookup = new Map(headers.map(h => [cleanHeader(h), h]));
  const headerMap = {};

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
        if (c.includes(cleaned) || cleaned.includes(c)) {
          match = original;
          break;
        }
      }

      if (match) break;
    }

    headerMap[key] = match;
  });

  let generatedId = 1;

  for (const row of rows) {
    const lat = toNumber(row[headerMap.latitude]);
    const lng = toNumber(row[headerMap.longitude]);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) continue;

    const customerId = safeString(row[headerMap.customerId]) || `AUTO_${generatedId++}`;
    const customerName = safeString(row[headerMap.customerName]) || customerId;
    const currentRep = safeString(row[headerMap.currentRep]) || 'Unassigned';
    const assignedRep = safeString(row[headerMap.assignedRep]) || currentRep || 'Unassigned';
    const rank = normalizeRank(row[headerMap.rank]);
    const overallSales = toNumber(row[headerMap.overallSales]);
    const wineSales = toNumber(row[headerMap.wineSales]);
    const spiritsSales = toNumber(row[headerMap.spiritsSales]);
    const thcSales = toNumber(row[headerMap.thcSales]);

    mapped.push({
      _id: customerId,
      latitude: lat,
      longitude: lng,
      customerId,
      customerName,
      address: safeString(row[headerMap.address]),
      zip: safeString(row[headerMap.zip]),
      chain: safeString(row[headerMap.chain]),
      segment: safeString(row[headerMap.segment]),
      currentRep,
      assignedRep,
      originalAssignedRep: assignedRep,
      overallSales: Number.isFinite(overallSales) ? overallSales : 0,
      wineSales: Number.isFinite(wineSales) ? wineSales : 0,
      spiritsSales: Number.isFinite(spiritsSales) ? spiritsSales : 0,
      thcSales: Number.isFinite(thcSales) ? thcSales : 0,
      rank,
      protected: toBoolean(row[headerMap.protected]),
      raw: row
    });
  }

  return mapped;
}

function buildRepColors() {
  state.repColors.clear();
  const reps = getAllReps();
  reps.forEach((rep, idx) => {
    state.repColors.set(rep, COLOR_PALETTE[idx % COLOR_PALETTE.length]);
  });
}

function refreshUI(rebuildMap = false) {
  syncControlState();
  renderLegend();
  renderRepControls();

  if (rebuildMap) rebuildMarkers();

  refreshMarkerStyles();
  renderRepTable();
  renderSelectionPreview();
  renderSummary();
  refreshTerritories();
}

function syncControlState() {
  const hasData = state.accounts.length > 0;

  [
    els.assignBtn, els.undoBtn, els.resetBtn, els.optimizeBtn, els.exportBtn,
    els.assignRepSelect, els.repFilter, els.repCountInput, els.minStopsInput,
    els.disruptionSlider, els.balanceMode, els.rankFilter,
    els.dimOthersCheckbox, els.showTerritoryCheckbox, els.clearSelectionBtn
  ].forEach(el => {
    if (!el) return;
    el.disabled = !hasData || (el === els.assignBtn && state.selection.size === 0);
  });

  els.undoBtn.disabled = !hasData || state.undoStack.length === 0;
  els.clearSelectionBtn.disabled = !hasData || state.selection.size === 0;
}

function renderRepControls() {
  const reps = getAllReps();

  fillSelect(els.assignRepSelect, reps, reps[0]);

  fillSelect(
    els.repFilter,
    ['ALL', ...reps],
    els.repFilter.value && ['ALL', ...reps].includes(els.repFilter.value) ? els.repFilter.value : 'ALL',
    rep => rep === 'ALL' ? 'All reps' : rep
  );

  const suggestedCount = reps.length || 1;
  if (!els.repCountInput.dataset.userTouched) {
    els.repCountInput.value = suggestedCount;
  }

  els.repCountInput.oninput = () => {
    els.repCountInput.dataset.userTouched = '1';
  };
}

function rebuildMarkers() {
  state.markerLayer.clearLayers();
  state.markerById.clear();

  for (const account of state.accounts) {
    const marker = L.circleMarker([account.latitude, account.longitude], markerStyleForAccount(account))
      .bindPopup(buildPopupHtml(account));

    marker.on('click', () => {
      toggleSelection(account._id, true);
    });

    state.markerLayer.addLayer(marker);
    state.markerById.set(account._id, marker);
  }
}

function refreshMarkerStyles() {
  for (const account of state.accounts) {
    const marker = state.markerById.get(account._id);
    if (!marker) continue;
    marker.setStyle(markerStyleForAccount(account));
    marker.setRadius(state.selection.has(account._id) ? 8 : 6);
    marker.getPopup()?.setContent(buildPopupHtml(account));
  }
}

function markerStyleForAccount(account) {
  const repFilter = els.repFilter.value || 'ALL';
  const rankFilter = els.rankFilter.value || 'ALL';
  const dimOthers = els.dimOthersCheckbox.checked;
  const isSelected = state.selection.has(account._id);
  const color = state.repColors.get(account.assignedRep) || '#4a5568';

  const repPass = repFilter === 'ALL' || account.assignedRep === repFilter;
  const rankPass = rankFilter === 'ALL' || account.rank === rankFilter;
  const visible = repPass && rankPass;
  const focusDim = dimOthers && !visible;

  return {
    radius: isSelected ? 8 : 6,
    color: isSelected ? '#1e293b' : color,
    weight: isSelected ? 2.4 : 1.2,
    opacity: visible ? 0.95 : 0.18,
    fillColor: color,
    fillOpacity: focusDim ? 0.10 : (account.protected ? 0.85 : 0.72)
  };
}

function buildPopupHtml(account) {
  return `
    <div style="min-width:240px;">
      <div style="font-size:16px;font-weight:700;">${escapeHtml(account.customerName)}</div>
      <div style="font-size:12px;color:#5d7286;margin:4px 0 8px;">${escapeHtml(account.address || '')}${account.address && account.zip ? ' • ' : ''}${escapeHtml(account.zip || '')}</div>
      <div><strong>Assigned Rep:</strong> ${escapeHtml(account.assignedRep)}</div>
      <div><strong>Current Rep:</strong> ${escapeHtml(account.currentRep)}</div>
      <div><strong>Revenue:</strong> ${formatCurrency(account.overallSales)}</div>
      <div><strong>Rank:</strong> ${escapeHtml(account.rank)}</div>
      <div><strong>Protected:</strong> ${account.protected ? 'Yes' : 'No'}</div>
      <div style="margin-top:6px;color:#5d7286;font-size:12px;">ID: ${escapeHtml(account.customerId)}</div>
    </div>
  `;
}

function renderLegend() {
  const reps = getAllReps();
  els.mapLegend.innerHTML = reps.slice(0, 10).map(rep => `
    <div class="legend-chip">
      <span class="color-dot" style="background:${state.repColors.get(rep)}"></span>
      <span>${escapeHtml(rep)}</span>
    </div>
  `).join('');
}

function renderRepTable() {
  const rows = summarizeByRep();

  if (!rows.length) {
    els.repTableBody.innerHTML = '<tr><td colspan="14" class="empty-state">Upload a file to begin.</td></tr>';
    return;
  }

  els.repTableBody.innerHTML = rows.map(row => {
    const focused = state.repFocus === row.rep ? 'rep-focus-row' : '';
    return `
      <tr class="${focused}" data-rep-row="${escapeHtmlAttr(row.rep)}">
        <td>
          <div class="rep-cell">
            <span class="color-dot" style="background:${row.color}"></span>
            <span>${escapeHtml(row.rep)}</span>
          </div>
        </td>
        <td>${row.stops}</td>
        <td>${formatCurrency(row.revenue)}</td>
        <td>${formatCurrency(row.wine)}</td>
        <td>${formatCurrency(row.spirits)}</td>
        <td>${formatCurrency(row.thc)}</td>
        <td>${row.A}</td>
        <td>${row.B}</td>
        <td>${row.C}</td>
        <td>${row.D}</td>
        <td>${row.protected}</td>
        <td>${row.movedIn}</td>
        <td>${row.movedOut}</td>
        <td>${row.workload.toFixed(2)}</td>
      </tr>
    `;
  }).join('');

  [...els.repTableBody.querySelectorAll('tr[data-rep-row]')].forEach(tr => {
    tr.addEventListener('click', () => {
      const rep = tr.getAttribute('data-rep-row');
      state.repFocus = state.repFocus === rep ? null : rep;
      els.repFilter.value = state.repFocus || 'ALL';
      refreshUI();
      zoomToRep(rep);
    });
  });
}

function renderSelectionPreview() {
  const selected = state.accounts.filter(a => state.selection.has(a._id)).slice(0, 40);
  els.selectionCount.textContent = String(state.selection.size);

  if (!selected.length) {
    els.selectionPreview.textContent = 'No accounts selected.';
    syncControlState();
    return;
  }

  els.selectionPreview.innerHTML = selected.map(a => `
    <div class="preview-item">
      <div class="preview-title">${escapeHtml(a.customerName)}</div>
      <div class="preview-meta">${escapeHtml(a.assignedRep)} • ${formatCurrency(a.overallSales)} • Rank ${escapeHtml(a.rank)}${a.protected ? ' • Protected' : ''}</div>
    </div>
  `).join('');

  syncControlState();
}

function renderSummary() {
  const totalAccounts = state.accounts.length;
  const totalRevenue = state.accounts.reduce((sum, a) => sum + a.overallSales, 0);
  const protectedCount = state.accounts.filter(a => a.protected).length;
  const movedCount = state.accounts.filter(a => a.assignedRep !== a.originalAssignedRep).length;
  const unchangedPct = totalAccounts ? ((totalAccounts - movedCount) / totalAccounts * 100) : 0;
  const workload = state.accounts.reduce((sum, a) => sum + (RANK_WEIGHTS[a.rank] || 1), 0);

  els.globalAccounts.textContent = totalAccounts.toLocaleString();
  els.globalRevenue.textContent = formatCurrency(totalRevenue);
  els.globalProtected.textContent = protectedCount.toLocaleString();
  els.globalMoved.textContent = movedCount.toLocaleString();
  els.globalUnchanged.textContent = `${unchangedPct.toFixed(1)}%`;
  els.globalWorkload.textContent = workload.toFixed(1);
}

function handleDrawCreated(event) {
  state.drawLayer.clearLayers();
  const layer = event.layer;
  state.drawLayer.addLayer(layer);

  let polygon;
  if (layer instanceof L.Rectangle || layer instanceof L.Polygon) {
    polygon = layer.toGeoJSON();
  } else {
    return;
  }

  state.selection.clear();

  for (const account of state.accounts) {
    const point = turf.point([account.longitude, account.latitude]);
    if (turf.booleanPointInPolygon(point, polygon)) {
      state.selection.add(account._id);
    }
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

function assignSelectionToRep() {
  const targetRep = els.assignRepSelect.value;
  const selectedIds = [...state.selection];
  if (!selectedIds.length || !targetRep) return;

  const changes = [];
  let skippedProtected = 0;

  for (const id of selectedIds) {
    const account = state.accounts.find(a => a._id === id);
    if (!account) continue;

    if (account.protected && account.assignedRep !== targetRep) {
      skippedProtected += 1;
      continue;
    }

    if (account.assignedRep === targetRep) continue;
    changes.push({ id, from: account.assignedRep, to: targetRep });
  }

  if (!changes.length) {
    showToast(skippedProtected ? `${skippedProtected} protected account(s) were skipped.` : 'No assignment changes to make.');
    return;
  }

  applyChanges(changes, `Assigned ${changes.length} account${changes.length === 1 ? '' : 's'} to ${targetRep}`);
  clearSelection();

  if (skippedProtected) {
    showToast(`${changes.length} changed. ${skippedProtected} protected account(s) skipped.`);
  }
}

function applyChanges(changes, label) {
  changes.forEach(change => {
    const account = state.accounts.find(a => a._id === change.id);
    if (!account) return;

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
  refreshUI();
  updateLastAction(label);
  showToast(label);
}

function undoLastAction() {
  const action = state.undoStack.pop();
  if (!action) return;

  for (const change of action.changes) {
    const account = state.accounts.find(a => a._id === change.id);
    if (account) account.assignedRep = change.from;
  }

  buildRepColors();
  refreshUI();
  updateLastAction(`Undid: ${action.label}`);
  showToast(`Undid: ${action.label}`);
}

function resetAssignments() {
  const changes = [];

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

  applyChanges(changes, 'Reset assignments to imported values');
  state.undoStack = [];
  updateLastAction('Reset assignments to imported values');
  refreshUI();
}

function optimizeRoutes() {
  if (!state.accounts.length) return;

  const targetCount = Math.max(1, Math.min(100, parseInt(els.repCountInput.value || '1', 10)));
  const minimumStops = Math.max(0, parseInt(els.minStopsInput.value || '0', 10));
  const disruptionWeight = Number(els.disruptionSlider.value) / 100;
  const balanceMode = els.balanceMode.value;

  const totalAccounts = state.accounts.length;
  if (minimumStops > 0 && targetCount * minimumStops > totalAccounts) {
    showToast(`Minimum stops too high for ${targetCount} reps. Max feasible minimum is ${Math.floor(totalAccounts / targetCount)}.`);
    return;
  }

  const currentReps = getAllReps();
  const targetRepNames = buildTargetRepNames(targetCount, currentReps);

  targetRepNames.forEach(rep => {
    if (!state.repColors.has(rep)) {
      state.repColors.set(rep, COLOR_PALETTE[state.repColors.size % COLOR_PALETTE.length]);
    }
  });

  const lockedAccounts = state.accounts.filter(a => a.protected);
  const movableAccounts = state.accounts.filter(a => !a.protected);

  const seedGroups = buildSeedGroups(targetRepNames, minimumStops);
  const assignments = new Map();

  for (const account of state.accounts) {
    if (account.protected) {
      assignments.set(account._id, account.assignedRep);
    }
  }

  const initialCentroids = computeSeedCentroids(seedGroups, targetRepNames);
  let centroids = initialCentroids;

  for (let iter = 0; iter < 6; iter += 1) {
    const repStats = buildDetailedRepStats(targetRepNames, assignments);
    const movableSorted = [...movableAccounts].sort((a, b) => b.overallSales - a.overallSales);

    for (const account of movableSorted) {
      let bestRep = null;
      let bestScore = Infinity;

      for (const rep of targetRepNames) {
        const centroid = centroids.get(rep) || { lat: account.latitude, lng: account.longitude };
        const dist = squaredDistance(account.latitude, account.longitude, centroid.lat, centroid.lng);
        const repStat = repStats.get(rep);

        let disruptionPenalty = 0;
        if (account.currentRep !== rep) disruptionPenalty += disruptionWeight * 0.75;
        if (account.assignedRep !== rep) disruptionPenalty += disruptionWeight * 0.25;

        const avgStopsTarget = totalAccounts / targetRepNames.length;
        const projectedStops = repStat.stops + 1;
        const stopBalancePenalty = balanceMode === 'revenue'
          ? 0
          : Math.pow((projectedStops - avgStopsTarget) / Math.max(1, avgStopsTarget), 2) * 0.35;

        const avgRevenueTarget = state.accounts.reduce((s, a) => s + a.overallSales, 0) / targetRepNames.length;
        const projectedRevenue = repStat.revenue + account.overallSales;
        const revenueBalancePenalty = balanceMode === 'stops'
          ? 0
          : Math.pow((projectedRevenue - avgRevenueTarget) / Math.max(1, avgRevenueTarget), 2) * 0.22;

        const underMinBonus = minimumStops > 0 && repStat.stops < minimumStops ? -0.30 : 0;

        const score = dist + disruptionPenalty + stopBalancePenalty + revenueBalancePenalty + underMinBonus;

        if (score < bestScore) {
          bestScore = score;
          bestRep = rep;
        }
      }

      assignments.set(account._id, bestRep);
      const repStat = repStats.get(bestRep);
      repStat.stops += 1;
      repStat.revenue += account.overallSales;
    }

    enforceMinimumStops(assignments, targetRepNames, minimumStops);
    centroids = recomputeCentroids(assignments, targetRepNames);
  }

  enforceMinimumStops(assignments, targetRepNames, minimumStops);

  const changes = [];
  for (const account of state.accounts) {
    const nextRep = assignments.get(account._id) || account.assignedRep;
    if (nextRep !== account.assignedRep) {
      changes.push({ id: account._id, from: account.assignedRep, to: nextRep });
    }
  }

  if (!changes.length) {
    showToast('Optimizer did not find a better assignment under the current rules.');
    return;
  }

  applyChanges(changes, `Optimized routes to ${targetRepNames.length} reps with minimum ${minimumStops} stops`);
}

function buildSeedGroups(targetRepNames, minimumStops) {
  const movable = state.accounts.filter(a => !a.protected);
  const sorted = [...movable].sort((a, b) => {
    if (a.latitude !== b.latitude) return a.latitude - b.latitude;
    return a.longitude - b.longitude;
  });

  const groups = new Map();
  targetRepNames.forEach(rep => groups.set(rep, []));

  if (!sorted.length) return groups;

  const sliceSize = Math.max(1, Math.floor(sorted.length / targetRepNames.length));
  let index = 0;

  for (const rep of targetRepNames) {
    const take = Math.max(1, minimumStops > 0 ? Math.min(minimumStops, sliceSize) : sliceSize);
    groups.set(rep, sorted.slice(index, index + take));
    index += take;
    if (index >= sorted.length) break;
  }

  targetRepNames.forEach((rep, i) => {
    if (!groups.get(rep)?.length) {
      const fallback = sorted[Math.min(i, sorted.length - 1)];
      groups.set(rep, fallback ? [fallback] : []);
    }
  });

  return groups;
}

function computeSeedCentroids(seedGroups, targetRepNames) {
  const centroids = new Map();

  targetRepNames.forEach((rep, idx) => {
    const members = seedGroups.get(rep) || [];
    if (members.length) {
      centroids.set(rep, {
        lat: avg(members.map(a => a.latitude)),
        lng: avg(members.map(a => a.longitude))
      });
    } else {
      const fallback = state.accounts[Math.min(idx, state.accounts.length - 1)];
      centroids.set(rep, {
        lat: fallback ? fallback.latitude : 40,
        lng: fallback ? fallback.longitude : -89
      });
    }
  });

  return centroids;
}

function buildDetailedRepStats(targetRepNames, assignments) {
  const stats = new Map();
  targetRepNames.forEach(rep => {
    stats.set(rep, { rep, stops: 0, revenue: 0 });
  });

  for (const account of state.accounts) {
    const rep = assignments.get(account._id);
    if (!rep || !stats.has(rep)) continue;
    const stat = stats.get(rep);
    stat.stops += 1;
    stat.revenue += account.overallSales;
  }

  return stats;
}

function enforceMinimumStops(assignments, targetRepNames, minimumStops) {
  if (minimumStops <= 0) return;

  const byRep = new Map();
  targetRepNames.forEach(rep => byRep.set(rep, []));

  for (const account of state.accounts) {
    const rep = assignments.get(account._id);
    if (!rep || !byRep.has(rep)) continue;
    byRep.get(rep).push(account);
  }

  const underfilled = targetRepNames.filter(rep => byRep.get(rep).length < minimumStops);
  const donorCandidates = () => targetRepNames
    .filter(rep => byRep.get(rep).length > minimumStops)
    .sort((a, b) => byRep.get(b).length - byRep.get(a).length);

  for (const needyRep of underfilled) {
    while (byRep.get(needyRep).length < minimumStops) {
      const donors = donorCandidates();
      if (!donors.length) break;

      let moved = false;
      const needyCentroid = computeCentroidForAccounts(byRep.get(needyRep));

      for (const donorRep of donors) {
        const donorAccounts = byRep.get(donorRep).filter(a => !a.protected);
        if (!donorAccounts.length) continue;

        donorAccounts.sort((a, b) => {
          const da = squaredDistance(a.latitude, a.longitude, needyCentroid.lat, needyCentroid.lng);
          const db = squaredDistance(b.latitude, b.longitude, needyCentroid.lat, needyCentroid.lng);
          return da - db;
        });

        const candidate = donorAccounts[0];
        if (!candidate) continue;

        assignments.set(candidate._id, needyRep);
        byRep.set(donorRep, byRep.get(donorRep).filter(a => a._id !== candidate._id));
        byRep.get(needyRep).push(candidate);
        moved = true;
        break;
      }

      if (!moved) break;
    }
  }
}

function recomputeCentroids(assignments, targetRepNames) {
  const centroids = new Map();

  targetRepNames.forEach(rep => {
    const members = state.accounts.filter(a => assignments.get(a._id) === rep);
    if (members.length) {
      centroids.set(rep, {
        lat: avg(members.map(a => a.latitude)),
        lng: avg(members.map(a => a.longitude))
      });
    } else {
      const fallback = state.accounts[0];
      centroids.set(rep, {
        lat: fallback ? fallback.latitude : 40,
        lng: fallback ? fallback.longitude : -89
      });
    }
  });

  return centroids;
}

function computeCentroidForAccounts(accounts) {
  if (!accounts.length) {
    return { lat: 40, lng: -89 };
  }

  return {
    lat: avg(accounts.map(a => a.latitude)),
    lng: avg(accounts.map(a => a.longitude))
  };
}

function buildTargetRepNames(targetCount, currentReps) {
  const reps = [...currentReps];
  while (reps.length < targetCount) {
    reps.push(`Rep ${reps.length + 1}`);
  }
  return reps.slice(0, targetCount);
}

function refreshTerritories() {
  state.territoryLayer.clearLayers();
  if (!els.showTerritoryCheckbox.checked || !state.accounts.length) return;

  const repFilter = els.repFilter.value || 'ALL';
  const rankFilter = els.rankFilter.value || 'ALL';
  const reps = getAllReps().filter(rep => repFilter === 'ALL' || rep === repFilter);

  reps.forEach(rep => {
    const members = state.accounts.filter(a => {
      return a.assignedRep === rep && (rankFilter === 'ALL' || a.rank === rankFilter);
    });

    if (members.length < 3) return;

    const points = members.map(a => [a.longitude, a.latitude]);
    const featureCollection = turf.featureCollection(points.map(p => turf.point(p)));

    let hull = null;
    try {
      hull = turf.convex(featureCollection);
    } catch (e) {
      hull = null;
    }

    if (!hull) return;

    const coords = hull.geometry.coordinates[0].map(([lng, lat]) => [lat, lng]);
    const poly = L.polygon(coords, {
      color: state.repColors.get(rep) || '#666',
      weight: 2,
      fillOpacity: 0.06,
      opacity: 0.7,
      interactive: false
    });

    poly.addTo(state.territoryLayer);
  });
}

function summarizeByRep() {
  const map = new Map();

  getAllReps().forEach(rep => {
    map.set(rep, {
      rep,
      color: state.repColors.get(rep) || '#4a5568',
      stops: 0,
      revenue: 0,
      wine: 0,
      spirits: 0,
      thc: 0,
      A: 0,
      B: 0,
      C: 0,
      D: 0,
      protected: 0,
      movedIn: 0,
      movedOut: 0,
      workload: 0
    });
  });

  for (const account of state.accounts) {
    if (!map.has(account.assignedRep)) {
      map.set(account.assignedRep, {
        rep: account.assignedRep,
        color: state.repColors.get(account.assignedRep) || '#4a5568',
        stops: 0,
        revenue: 0,
        wine: 0,
        spirits: 0,
        thc: 0,
        A: 0,
        B: 0,
        C: 0,
        D: 0,
        protected: 0,
        movedIn: 0,
        movedOut: 0,
        workload: 0
      });
    }

    const row = map.get(account.assignedRep);
    row.stops += 1;
    row.revenue += account.overallSales;
    row.wine += account.wineSales;
    row.spirits += account.spiritsSales;
    row.thc += account.thcSales;
    row[account.rank] = (row[account.rank] || 0) + 1;
    row.workload += RANK_WEIGHTS[account.rank] || 1;
    if (account.protected) row.protected += 1;
    if (account.assignedRep !== account.currentRep) row.movedIn += 1;
  }

  for (const account of state.accounts) {
    if (map.has(account.currentRep) && account.currentRep !== account.assignedRep) {
      map.get(account.currentRep).movedOut += 1;
    }
  }

  return [...map.values()].sort((a, b) => a.rep.localeCompare(b.rep, undefined, { numeric: true }));
}

function exportWorkbook() {
  if (!state.accounts.length) return;

  const wb = XLSX.utils.book_new();

  const assignments = state.accounts.map(a => ({
    ...a.raw,
    Customer_ID: a.customerId,
    Customer_Name: a.customerName,
    Current_Rep: a.currentRep,
    Assigned_Rep: a.assignedRep,
    Original_Assigned_Rep: a.originalAssignedRep,
    Protected: a.protected ? 'Yes' : 'No',
    Rank: a.rank,
    Latitude: a.latitude,
    Longitude: a.longitude,
    Overall_Sales: a.overallSales,
    Wine_Sales: a.wineSales,
    Spirits_Sales: a.spiritsSales,
    THC_Sales: a.thcSales,
    Moved: a.assignedRep !== a.originalAssignedRep ? 'Yes' : 'No'
  }));

  const summary = summarizeByRep().map(r => ({
    Rep: r.rep,
    Stops: r.stops,
    Revenue: round2(r.revenue),
    Wine: round2(r.wine),
    Spirits: round2(r.spirits),
    THC: round2(r.thc),
    A: r.A,
    B: r.B,
    C: r.C,
    D: r.D,
    Protected: r.protected,
    Moved_In: r.movedIn,
    Moved_Out: r.movedOut,
    Four_Week_Load: round2(r.workload)
  }));

  const changeLog = state.changeLog.length ? state.changeLog : [{
    timestamp: '',
    customerId: '',
    customerName: '',
    fromRep: '',
    toRep: '',
    protected: ''
  }];

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(assignments), 'Assignments');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'Rep Summary');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(changeLog), 'Change Log');

  XLSX.writeFile(wb, state.loadedFileName || 'territory_export.xlsx');
  showToast('Excel export created.');
  updateLastAction('Exported workbook');
}

function fitMapToAccounts() {
  if (!state.accounts.length) return;
  const latlngs = state.accounts.map(a => [a.latitude, a.longitude]);
  state.map.fitBounds(latlngs, { padding: [25, 25] });
}

function zoomToRep(rep) {
  const points = state.accounts.filter(a => a.assignedRep === rep).map(a => [a.latitude, a.longitude]);
  if (points.length) {
    state.map.fitBounds(points, { padding: [30, 30] });
  }
}

function toggleTheme() {
  if (state.theme === 'light') {
    state.map.removeLayer(state.lightLayer);
    state.darkLayer.addTo(state.map);
    state.theme = 'dark';
    els.themeToggle.textContent = 'Light Map';
  } else {
    state.map.removeLayer(state.darkLayer);
    state.lightLayer.addTo(state.map);
    state.theme = 'light';
    els.themeToggle.textContent = 'Dark Map';
  }
}

function enableTopControls() {
  [els.sheetSelect, els.loadSheetBtn].forEach(el => {
    el.disabled = false;
  });
}

function getAllReps() {
  const set = new Set();
  state.accounts.forEach(a => {
    if (a.assignedRep) set.add(a.assignedRep);
    if (a.currentRep) set.add(a.currentRep);
  });
  return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function fillSelect(selectEl, values, selectedValue, labelFn = v => v) {
  selectEl.innerHTML = values.map(v => `<option value="${escapeHtmlAttr(v)}">${escapeHtml(labelFn(v))}</option>`).join('');
  if (selectedValue && values.includes(selectedValue)) {
    selectEl.value = selectedValue;
  }
}

function updateLastAction(text) {
  state.lastAction = text;
  els.lastAction.textContent = text;
}

let toastTimer = null;
function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    els.toast.classList.remove('show');
  }, 2400);
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
  if (typeof value === 'number') return value;
  const n = Number(String(value ?? '').replace(/[$,%\s,]/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

function toBoolean(value) {
  const v = safeString(value).toLowerCase();
  return ['true','yes','y','1','protected','locked'].includes(v);
}

function normalizeRank(rankValue) {
  const raw = safeString(rankValue).toUpperCase();
  if (['A','B','C','D'].includes(raw)) return raw;
  if (raw.includes('WEEK')) return 'A';
  if (raw.includes('BI')) return 'B';
  if (raw.includes('MONTH')) return 'C';
  if (raw.includes('QUART')) return 'D';
  return 'C';
}

function formatCurrency(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0
  }).format(value || 0);
}

function avg(values) {
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
}

function squaredDistance(lat1, lng1, lat2, lng2) {
  const dx = lng1 - lng2;
  const dy = lat1 - lat2;
  return dx * dx + dy * dy;
}

function round2(v) {
  return Math.round((v || 0) * 100) / 100;
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
