(() => {
  const KEY = 'starish.loading-log.v1';
  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
  const todayISO = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
  const DEFAULT_INCHARGES = ['AJAY', 'DILIP', 'MANTU', 'VAMSI'];
  const displayIncharge = name => String(name ?? '').trim().toUpperCase() === 'AJEET' ? 'AJAY' : String(name ?? '').trim();
  const uid = () => crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  let state;
  try { state = JSON.parse(localStorage.getItem(KEY)) || { headers: [], items: [] }; } catch { state = { headers: [], items: [] }; }
  if (!Array.isArray(state.headers) || !Array.isArray(state.items)) state = { headers: [], items: [], stageRows: [] };
  if (!Array.isArray(state.stageRows)) state.stageRows = [];
  if (!Array.isArray(state.inchargeNames)) state.inchargeNames = [];
  const apiUrl = String(window.LOADING_API_URL || '').trim();
  let currentDetailLoadingId = null;
  let editingItemId = null;
  let editingHeaderId = null;

  function persist() { localStorage.setItem(KEY, JSON.stringify(state)); }
  function inchargeNames() {
    const all = [...DEFAULT_INCHARGES, ...state.inchargeNames, ...state.headers.map(h => h.incharge).filter(Boolean)];
    const unique = new Map();
    all.forEach(name => { const clean = String(name || '').trim(); if (clean && clean.toUpperCase() !== 'AJEET' && !unique.has(clean.toUpperCase())) unique.set(clean.toUpperCase(), clean.toUpperCase()); });
    return [...unique.values()].sort((a,b) => a.localeCompare(b));
  }
  function refreshInchargeOptions() {
    $('#inchargeOptions').innerHTML = inchargeNames().map(name => `<option value="${escapeHtml(name)}"></option>`).join('');
  }
  function setStorageStatus(title, detail) {
    const note = $('#storageNote');
    if (note) note.innerHTML = `<span>ⓘ</span><div><b>${escapeHtml(title)}</b><small>${escapeHtml(detail)}</small></div>`;
  }
  function loadSheetData() {
    if (!apiUrl) return Promise.reject(new Error('Google Sheet connector URL is not configured.'));
    return new Promise((resolve, reject) => {
      const callback = `loadingSheetCallback_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      const script = document.createElement('script');
      const timer = setTimeout(() => finish(new Error('Timed out contacting the Google Sheet.')), 20000);
      function finish(error, data) {
        clearTimeout(timer); delete window[callback]; script.remove();
        error ? reject(error) : data && data.ok ? resolve(data) : reject(new Error(data && data.error || 'Google Sheet did not return data.'));
      }
      window[callback] = data => finish(null, data);
      script.onerror = () => finish(new Error('Google Sheet connection failed. Check sign-in, sharing access, and the deployment URL.'));
      script.src = `${apiUrl}${apiUrl.includes('?') ? '&' : '?'}callback=${encodeURIComponent(callback)}&t=${Date.now()}`;
      document.head.append(script);
    });
  }
  async function syncFromSheet() {
    setStorageStatus('Connecting to Google Sheet…', 'Loading LOADING_HEADER and LOADING_ITEMS.');
    try {
      const data = await loadSheetData();
      const savedNames = state.inchargeNames || [];
      state = { headers: data.headers || [], items: data.items || [], stageRows: data.stageRows || [], inchargeNames: savedNames, activeShift: state.activeShift || null };
      persist(); renderHome();
      refreshInchargeOptions();
      setStorageStatus('Connected to Google Sheet.', 'Loading entries sync through the source sheet across authorized devices.');
    } catch (error) {
      setStorageStatus('Google Sheet sync is unavailable.', `${error.message} Entries currently stay in this browser.`);
    }
  }
  function hoursBetween(start, end) {
    if (!start || !end) return null;
    const [sh, sm] = start.split(':').map(Number), [eh, em] = end.split(':').map(Number);
    let minutes = eh * 60 + em - (sh * 60 + sm);
    if (minutes < 0) minutes += 1440;
    return minutes / 60;
  }
  function durationText(hours) {
    if (hours === null || !Number.isFinite(hours)) return '—';
    const minutes = Math.round(hours * 60);
    return `${Math.floor(minutes / 60)}h ${String(minutes % 60).padStart(2, '0')}m`;
  }
  function showToast(message) { const el = $('#toast'); el.textContent = message; el.classList.add('show'); setTimeout(() => el.classList.remove('show'), 2600); }
  function navigate(view) {
    $$('.view').forEach(el => el.classList.toggle('active', el.id === `${view}View`));
    $$('.nav-item').forEach(el => el.classList.toggle('selected', el.dataset.view === view));
    if (view === 'home') renderHome();
    if (view === 'reports') renderReports();
    if (view === 'gaps') renderGaps();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  function addVehicleRow(data = {}) {
    const index = $('#vehicleRows').children.length + 1;
    const row = document.createElement('div');
    row.className = 'vehicle-row';
    row.innerHTML = `<div class="vehicle-row-head"><strong>Vehicle ${index}</strong><button class="remove-row" type="button">Remove</button></div>
      <div class="vehicle-fields">
        <label>Customer<input class="customer" list="customerOptions" placeholder="Customer name" value="${escapeHtml(data.customer)}" required /></label>
        <label>Vehicle number<input class="vehicle" placeholder="KA 00 AA 0000" value="${escapeHtml(data.vehicleNo)}" required /></label>
        <label>Vehicle feet / type<input class="feet" list="feetOptions" placeholder="22 FEET" value="${escapeHtml(data.vehicleFeet)}" /></label>
        <label>Loading start<input class="start" type="time" value="${escapeHtml(data.start)}" required /></label>
        <label>Loading end<input class="end" type="time" value="${escapeHtml(data.end)}" required /></label>
        <div class="duration-cell"><small>Total loading time</small><b class="duration">${durationText(hoursBetween(data.start, data.end))}</b></div>
        <label>Remarks<input class="remarks" placeholder="Optional" value="${escapeHtml(data.remarks)}" /></label>
      </div>`;
    $('#vehicleRows').append(row);
    row.addEventListener('input', e => { if (e.target.matches('.start,.end')) $('.duration', row).textContent = durationText(hoursBetween($('.start', row).value, $('.end', row).value)); });
    $('.remove-row', row).addEventListener('click', () => { if ($('#vehicleRows').children.length === 1) return showToast('At least one vehicle row is required.'); row.remove(); renumberRows(); });
  }
  function renumberRows() { $$('.vehicle-row-head strong').forEach((el, i) => el.textContent = `Vehicle ${i + 1}`); }
  function openNew() {
    $('#loadingForm').reset(); $('#vehicleRows').replaceChildren(); addVehicleRow();
    refreshInchargeOptions();
    if (state.activeShift) {
      setActiveShiftMode(true);
    } else {
      $('#loadingDate').value = todayISO();
      setActiveShiftMode(false);
    }
    navigate('new');
  }
  function setActiveShiftMode(active) {
    const shift = state.activeShift;
    $('#shiftDetailsCard').classList.toggle('hidden', active);
    $('#activeShiftNote').classList.toggle('hidden', !active);
    $('#addVehicleButton').classList.toggle('hidden', active);
    $('#finishShiftFormButton').classList.toggle('hidden', !active);
    $('#formHeading').textContent = active ? 'Add vehicle' : 'New loading entry';
    $('#saveLoadingButton').textContent = active ? 'Save vehicle & add next' : 'Save shift & vehicles';
    if (active && shift) {
      $('#activeShiftNote').textContent = `Adding vehicles to ${displayIncharge(shift.incharge)} · ${String(shift.shift).toUpperCase()} SHIFT · ${fmtDate(shift.date)} · ${shift.helperCount} helpers`;
    } else {
      $('#activeShiftNote').textContent = '';
    }
  }
  function finishActiveShift() {
    if (!state.activeShift) return;
    const name = displayIncharge(state.activeShift.incharge);
    if (!confirm(`Finish ${name}'s shift? Saved vehicle entries will remain in the report. The next loading entry will ask for shift details.`)) return;
    state.activeShift = null;
    persist();
    setActiveShiftMode(false);
    renderHome();
    navigate('home');
    showToast('Shift finished. Start a new shift to enter details again.');
  }
  function allRows() {
    const headers = new Map(state.headers.map(h => [h.id, h]));
    return state.items.map(item => ({ ...item, header: headers.get(item.loadingId) })).filter(row => row.header);
  }
  function fmtDate(value) { if (!value) return '—'; return new Date(`${value}T00:00:00`).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }); }
  function renderHome() {
    $('#todayLabel').textContent = new Date().toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).toUpperCase();
    const today = todayISO(), todays = allRows().filter(r => r.header.date === today), todayHours = todays.reduce((a, r) => a + (Number(r.totalHours) || 0), 0);
    const metrics = [
      ['Vehicles today', todays.length, 'Vehicles logged', '▣'],
      ['Loading hours today', durationText(todayHours), 'Total recorded time', '◷'],
      ['Shift entries today', new Set(todays.map(r => r.loadingId)).size, 'Day and night shifts', '◫'],
      ['All-time entries', state.items.length, 'Vehicle records saved', '↗']
    ];
    $('#metrics').innerHTML = metrics.map(([label, value, hint, icon]) => `<article class="metric-card"><div class="metric-top">${label}<span class="metric-icon">${icon}</span></div><div class="metric-value">${value}</div><div class="metric-hint">${hint}</div></article>`).join('');
    const activeCard = $('#activeShiftCard');
    if (state.activeShift) {
      const shift = state.activeShift, count = state.items.filter(item => item.loadingId === shift.id).length;
      $('#activeShiftLabel').textContent = `${displayIncharge(shift.incharge)} · ${String(shift.shift).toUpperCase()} SHIFT · ${fmtDate(shift.date)}`;
      $('#activeShiftCount').textContent = `${count} vehicle${count === 1 ? '' : 's'} saved · ${shift.helperCount} helpers`;
      activeCard.classList.remove('hidden');
    } else activeCard.classList.add('hidden');
    const latest = [...state.headers].sort((a, b) => `${b.date}${b.createdAt}`.localeCompare(`${a.date}${a.createdAt}`)).slice(0, 6);
    const list = $('#recentList');
    if (!latest.length) { list.innerHTML = '<div class="empty-state">No loading entries yet. Tap <b>New loading</b> to record the first shift.</div>'; return; }
    list.innerHTML = latest.map(h => {
      const items = state.items.filter(i => i.loadingId === h.id), total = items.reduce((a, i) => a + (Number(i.totalHours) || 0), 0);
      return `<article class="record-card" data-loading-id="${escapeHtml(h.id)}" role="button" tabindex="0" aria-label="View ${items.length} vehicles for ${escapeHtml(displayIncharge(h.incharge))}, ${fmtDate(h.date)}"><div class="record-date">${fmtDate(h.date)}<small>${escapeHtml(h.shift)} SHIFT · ${items.length} vehicle${items.length === 1 ? '' : 's'}</small></div><div class="record-meta">${escapeHtml(displayIncharge(h.incharge))}<small>Loading incharge</small></div><div class="record-meta">${h.helperCount} helpers<small>${durationText(total)} loading time</small></div><div class="record-count">${items.length}<small>VEHICLES · TAP FOR DETAILS</small></div></article>`;
    }).join('');
    $$('.record-card[data-loading-id]', list).forEach(card => {
      const open = () => showLoadingDetails(card.dataset.loadingId);
      card.addEventListener('click', open);
      card.addEventListener('keydown', event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); open(); } });
    });
  }
  function showLoadingDetails(id) {
    const header = state.headers.find(row => row.id === id);
    if (!header) return;
    const vehicles = state.items.filter(item => item.loadingId === id);
    const total = vehicles.reduce((sum, item) => sum + (Number(item.totalHours) || 0), 0);
    $('#loadingDetailsTitle').textContent = `${displayIncharge(header.incharge)} · ${fmtDate(header.date)}`;
    $('#loadingDetailsSummary').textContent = `${String(header.shift || '').toUpperCase()} SHIFT · ${vehicles.length} vehicles · ${durationText(total)} total loading time`;
    currentDetailLoadingId = id;
    $('#addVehicleToShiftButton').textContent = state.activeShift?.id === id ? '＋ Add another vehicle' : '＋ Add vehicle to this shift';
    $('#loadingDetailsList').innerHTML = vehicles.length ? vehicles.map((item, index) => `<article class="loading-detail-card"><div class="loading-detail-head"><strong>Vehicle ${index + 1}</strong><b>${escapeHtml(item.vehicleNo || '—')}</b></div><div class="loading-detail-customer">${escapeHtml(item.customer || '—')}</div><div class="loading-detail-grid"><span>Type<strong>${escapeHtml(item.vehicleFeet || '—')}</strong></span><span>Loading start<strong>${escapeHtml(item.start || '—')}</strong></span><span>Loading end<strong>${escapeHtml(item.end || '—')}</strong></span><span>Duration<strong>${durationText(Number(item.totalHours))}</strong></span></div>${item.remarks ? `<p class="loading-detail-remarks">Remarks: ${escapeHtml(item.remarks)}</p>` : ''}<button class="secondary-button edit-vehicle-button" type="button" data-item-id="${escapeHtml(item.id)}">Edit this vehicle</button></article>`).join('') : '<div class="empty-state">No vehicle details saved for this shift.</div>';
    $$('.edit-vehicle-button', $('#loadingDetailsList')).forEach(button => button.addEventListener('click', () => openEditVehicle(button.dataset.itemId)));
    $('#loadingDetailsDialog').showModal();
  }
  function filteredRows() {
    const from = $('#filterFrom').value, to = $('#filterTo').value, shift = $('#filterShift').value, incharge = $('#filterIncharge').value.trim().toLowerCase();
    return allRows().filter(r => (!from || r.header.date >= from) && (!to || r.header.date <= to) && (!shift || r.header.shift === shift) && (!incharge || displayIncharge(r.header.incharge).toLowerCase().includes(incharge))).sort((a,b) => `${b.header.date}${b.start}`.localeCompare(`${a.header.date}${a.start}`));
  }
  function renderReports() {
    const rows = filteredRows();
    const hours = rows.reduce((a, r) => a + (Number(r.totalHours) || 0), 0);
    $('#reportSummary').textContent = `${rows.length} vehicle${rows.length === 1 ? '' : 's'} · ${new Set(rows.map(r => r.loadingId)).size} shift entries · ${durationText(hours)} total loading time`;
    $('#reportRows').innerHTML = rows.map(r => `<tr><td>${fmtDate(r.header.date)}<br><span class="muted">${escapeHtml(r.header.shift)}</span></td><td>${escapeHtml(displayIncharge(r.header.incharge))}</td><td>${escapeHtml(r.customer)}</td><td>${escapeHtml(r.start)}</td><td>${escapeHtml(r.end)}</td><td>${durationText(Number(r.totalHours))}</td><td>${escapeHtml(r.vehicleNo)}</td><td>${escapeHtml(r.vehicleFeet || '—')}</td><td>${r.header.helperCount}</td><td>${escapeHtml(r.remarks || '—')}</td><td><button class="table-edit-button" type="button" data-loading-id="${escapeHtml(r.loadingId)}">Details / edit</button></td></tr>`).join('');
    $$('.table-edit-button', $('#reportRows')).forEach(button => button.addEventListener('click', () => showLoadingDetails(button.dataset.loadingId)));
    $('#reportEmpty').classList.toggle('hidden', rows.length > 0);
    $('.table-wrap').classList.toggle('hidden', rows.length === 0);
    return rows;
  }
  function openEditVehicle(id) {
    const item = state.items.find(row => row.id === id);
    if (!item) return showToast('Vehicle entry nahi mili.');
    editingItemId = id;
    $('#editCustomer').value = item.customer || '';
    $('#editVehicleNo').value = item.vehicleNo || '';
    $('#editVehicleFeet').value = item.vehicleFeet || '';
    $('#editStart').value = item.start || '';
    $('#editEnd').value = item.end || '';
    $('#editRemarks').value = item.remarks || '';
    $('#editVehicleStatus').textContent = apiUrl ? 'Changes will be saved to the connected sheet.' : 'Changes will be saved on this device only.';
    $('#loadingDetailsDialog').close();
    $('#editVehicleDialog').showModal();
  }
  function openEditShift() {
    const header = state.headers.find(row => row.id === currentDetailLoadingId);
    if (!header) return showToast('Shift entry nahi mili.');
    editingHeaderId = header.id;
    $('#editShiftDate').value = header.date || '';
    $('#editShiftName').value = String(header.shift || 'DAY').toUpperCase();
    $('#editShiftIncharge').value = displayIncharge(header.incharge);
    $('#editShiftHelpers').value = Number(header.helperCount) || 0;
    $('#editShiftStatus').textContent = apiUrl ? 'Changes will be saved to the connected sheet.' : 'Changes will be saved on this device only.';
    $('#loadingDetailsDialog').close();
    $('#editShiftDialog').showModal();
  }
  async function saveEditedVehicle(event) {
    event.preventDefault();
    const index = state.items.findIndex(row => row.id === editingItemId), oldItem = state.items[index];
    if (!oldItem) return showToast('Vehicle entry nahi mili.');
    const updated = { ...oldItem, customer: $('#editCustomer').value.trim(), vehicleNo: $('#editVehicleNo').value.trim().toUpperCase(), vehicleFeet: $('#editVehicleFeet').value.trim().toUpperCase(), start: $('#editStart').value, end: $('#editEnd').value, remarks: $('#editRemarks').value.trim() };
    if (!updated.customer || !updated.vehicleNo || !updated.start || !updated.end) return showToast('Required fields fill karein.');
    updated.totalHours = Number(hoursBetween(updated.start, updated.end).toFixed(2));
    const submit = $('#editVehicleForm button[type="submit"]'); submit.disabled = true;
    let synced = !apiUrl;
    try {
      if (apiUrl) {
        const header = state.headers.find(row => row.id === updated.loadingId);
        await fetch(apiUrl, { method: 'POST', mode: 'no-cors', credentials: 'include', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify({ action: 'save', header, items: [updated] }) });
        const remote = await loadSheetData();
        const saved = (remote.items || []).find(row => row.id === updated.id);
        if (!saved || saved.vehicleNo !== updated.vehicleNo || saved.start !== updated.start || saved.end !== updated.end || saved.customer !== updated.customer) throw new Error('Edited entry could not be confirmed in the Sheet.');
        state = { ...state, headers: remote.headers || [], items: remote.items || [], stageRows: remote.stageRows || [] };
        synced = true;
      } else state.items[index] = updated;
    } catch (error) {
      state.items[index] = updated;
      setStorageStatus('Edit could not be confirmed in Google Sheet.', `${error.message} The edited entry is saved only in this browser.`);
    }
    persist(); renderHome(); renderReports();
    $('#editVehicleDialog').close();
    if (currentDetailLoadingId) showLoadingDetails(currentDetailLoadingId);
    showToast(synced ? 'Vehicle entry updated.' : 'Updated on this device only; Sheet sync needs checking.');
    submit.disabled = false;
  }
  async function saveEditedShift(event) {
    event.preventDefault();
    const header = state.headers.find(row => row.id === editingHeaderId);
    if (!header) return showToast('Shift entry nahi mili.');
    const updated = { ...header, date: $('#editShiftDate').value, shift: $('#editShiftName').value, incharge: $('#editShiftIncharge').value.trim(), helperCount: Number($('#editShiftHelpers').value) };
    if (!updated.date || !updated.shift || !updated.incharge || updated.helperCount < 0) return showToast('Shift ki required details fill karein.');
    const submit = $('#editShiftForm button[type="submit"]'); submit.disabled = true;
    let synced = !apiUrl;
    try {
      if (apiUrl) {
        await fetch(apiUrl, { method: 'POST', mode: 'no-cors', credentials: 'include', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify({ action: 'save', header: updated, items: [] }) });
        const remote = await loadSheetData();
        const saved = (remote.headers || []).find(row => row.id === updated.id);
        if (!saved || saved.date !== updated.date || String(saved.shift).toUpperCase() !== updated.shift || displayIncharge(saved.incharge).toUpperCase() !== displayIncharge(updated.incharge).toUpperCase() || Number(saved.helperCount) !== updated.helperCount) throw new Error('Edited shift could not be confirmed in the Sheet.');
        state.headers = remote.headers || []; state.items = remote.items || []; state.stageRows = remote.stageRows || [];
        synced = true;
      } else state.headers[state.headers.findIndex(row => row.id === updated.id)] = updated;
    } catch (error) {
      state.headers[state.headers.findIndex(row => row.id === updated.id)] = updated;
      setStorageStatus('Edit could not be confirmed in Google Sheet.', `${error.message} Shift changes are saved only in this browser.`);
    }
    if (state.activeShift?.id === updated.id) state.activeShift = updated;
    state.inchargeNames = [...new Set([...(state.inchargeNames || []), displayIncharge(updated.incharge)])];
    persist(); refreshInchargeOptions(); renderHome(); renderReports();
    $('#editShiftDialog').close();
    showLoadingDetails(updated.id);
    showToast(synced ? 'Shift details updated.' : 'Updated on this device only; Sheet sync needs checking.');
    submit.disabled = false;
  }
  function filteredStageRows() {
    const from = $('#gapFrom').value, to = $('#gapTo').value, query = $('#gapSearch').value.trim().toLowerCase();
    return state.stageRows.filter(r => (!from || r.date >= from) && (!to || r.date <= to) && (!query || `${r.customerName} ${r.invoiceNumber} ${r.itemName} ${r.itemCode}`.toLowerCase().includes(query)))
      .sort((a, b) => `${b.date}${b.invoiceNumber}`.localeCompare(`${a.date}${a.invoiceNumber}`));
  }
  function prettyStamp(value) {
    if (!value) return '—';
    const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})(?:\s+(\d{2}:\d{2})(?::\d{2})?)?$/);
    if (!match) return String(value);
    return `${match[3]}-${match[2]}-${match[1]}${match[4] ? ` ${match[4]}` : ''}`;
  }
  function renderGaps() {
    const rows = filteredStageRows();
    $('#gapSummary').textContent = `${rows.length} product rows · ${new Set(rows.map(r => r.invoiceNumber).filter(Boolean)).size} invoices`;
    const from = $('#gapFrom').value, to = $('#gapTo').value;
    $('#printPeriod').textContent = from || to ? `Period: ${from ? fmtDate(from) : 'All dates'} – ${to ? fmtDate(to) : 'All dates'}` : 'Period: All dates';
    $('#gapRows').innerHTML = rows.map(r => `<tr><td>${fmtDate(r.date)}</td><td>${escapeHtml(r.customerName)}</td><td>${escapeHtml(r.itemName)}</td><td>${escapeHtml(r.itemCode)}</td><td>${escapeHtml(r.invoiceNumber)}</td><td>${escapeHtml(prettyStamp(r.oqcEnd))}</td><td class="print-hide">${escapeHtml(prettyStamp(r.loadingEnd))}</td><td class="print-hide">${escapeHtml(prettyStamp(r.invoiceTime))}</td><td class="print-hide">${escapeHtml(prettyStamp(r.gateOutTime))}</td><td>${escapeHtml(r.gap1 || '—')}</td><td>${escapeHtml(r.gap2 || '—')}</td><td>${escapeHtml(r.gap3 || '—')}</td></tr>`).join('');
    $('#gapEmpty').classList.toggle('hidden', rows.length > 0);
    $('.gap-report .table-wrap').classList.toggle('hidden', rows.length === 0);
    return rows;
  }
  function csvCell(value) { return `"${String(value ?? '').replaceAll('"', '""')}"`; }
  function download(filename, text, type) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type })); a.download = filename; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000); }
  function downloadCSV() {
    const rows = filteredRows(), headings = ['DATE','SHIFT','LOADING INCHARGE','HELPER COUNT','CUSTOMER','LOADING START','LOADING END','TOTAL HRS','VEHICLE NO','VEHICLE FEET','REMARKS'];
    const data = rows.map(r => [r.header.date,r.header.shift,displayIncharge(r.header.incharge),r.header.helperCount,r.customer,r.start,r.end,Number(r.totalHours).toFixed(2),r.vehicleNo,r.vehicleFeet,r.remarks]);
    download(`loading-report-${todayISO()}.csv`, [headings,...data].map(row => row.map(csvCell).join(',')).join('\r\n'), 'text/csv;charset=utf-8');
  }
  async function saveForm(event) {
    event.preventDefault();
    const continuing = Boolean(state.activeShift);
    const date = continuing ? state.activeShift.date : $('#loadingDate').value;
    const shift = continuing ? state.activeShift.shift : $('#shift').value;
    const incharge = continuing ? state.activeShift.incharge : $('#incharge').value.trim();
    const helperCount = continuing ? state.activeShift.helperCount : Number($('#helperCount').value);
    const rows = $$('.vehicle-row').map(row => ({ customer: $('.customer',row).value.trim(), vehicleNo: $('.vehicle',row).value.trim().toUpperCase(), vehicleFeet: $('.feet',row).value.trim().toUpperCase(), start: $('.start',row).value, end: $('.end',row).value, remarks: $('.remarks',row).value.trim() }));
    if ((!continuing && (!date || !shift || !incharge || helperCount < 0)) || rows.some(r => !r.customer || !r.vehicleNo || !r.start || !r.end)) return showToast('Please complete all required fields.');
    const header = state.activeShift || { id: uid(), date, shift, incharge, helperCount, createdAt: new Date().toISOString() };
    const items = rows.map(r => ({ id: uid(), loadingId: header.id, ...r, totalHours: Number(hoursBetween(r.start,r.end).toFixed(2)) }));
    const saveButton = $('#loadingForm button[type="submit"]');
    saveButton.disabled = true;
    if (apiUrl) {
      setStorageStatus('Saving to Google Sheet…', 'Waiting for the sheet to confirm the entry.');
      try {
        await fetch(apiUrl, { method: 'POST', mode: 'no-cors', credentials: 'include', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify({ action: 'save', header, items }) });
        const remote = await loadSheetData();
        const savedItems = new Set((remote.items || []).map(row => row.id));
        if (!remote.headers.some(row => row.id === header.id) || items.some(item => !savedItems.has(item.id))) throw new Error('The vehicle entry was not found in the sheet after saving.');
        state = { headers: remote.headers || [], items: remote.items || [], stageRows: remote.stageRows || [], inchargeNames: [...new Set([...(state.inchargeNames || []), incharge])], activeShift: header };
        persist(); renderHome();
        refreshInchargeOptions();
        setStorageStatus('Connected to Google Sheet.', 'Loading entries sync through the source sheet across authorized devices.');
        $('#vehicleRows').replaceChildren(); addVehicleRow(); setActiveShiftMode(true); showToast('Vehicle saved to the Google Sheet. Add the next vehicle.');
      } catch (error) {
        if (!state.headers.some(row => row.id === header.id)) state.headers.push(header);
        const existingIds = new Set(state.items.map(row => row.id)); state.items.push(...items.filter(item => !existingIds.has(item.id)));
        state.inchargeNames = [...new Set([...(state.inchargeNames || []), incharge])]; state.activeShift = header; persist(); renderHome();
        setStorageStatus('Google Sheet save could not be confirmed.', `${error.message} This entry is saved only in this browser. Download a backup and check the connection.`);
        $('#vehicleRows').replaceChildren(); addVehicleRow(); setActiveShiftMode(true); showToast('Saved on this device only. Continue adding vehicles; check Sheet sync later.');
      } finally { saveButton.disabled = false; }
      return;
    }
    if (!state.headers.some(row => row.id === header.id)) state.headers.push(header);
    state.items.push(...items); state.inchargeNames = [...new Set([...(state.inchargeNames || []), incharge])]; state.activeShift = header; persist(); renderHome();
    $('#vehicleRows').replaceChildren(); addVehicleRow(); setActiveShiftMode(true); showToast('Vehicle saved. Add the next vehicle; shift details are saved.');
    saveButton.disabled = false;
  }
  $('#newLoadingButton').addEventListener('click', openNew);
  $('#continueShiftButton').addEventListener('click', openNew);
  $('#finishShiftHomeButton').addEventListener('click', finishActiveShift);
  $('#finishShiftFormButton').addEventListener('click', finishActiveShift);
  $('#addVehicleToShiftButton').addEventListener('click', () => {
    const header = state.headers.find(row => row.id === currentDetailLoadingId);
    if (!header) return showToast('Shift entry nahi mili.');
    if (state.activeShift && state.activeShift.id !== header.id && !confirm('Aap ek purani shift mein vehicle add kar rahe hain. Maujooda active shift ki entries saved rahengi; continue karein?')) return;
    state.activeShift = header; persist(); $('#loadingDetailsDialog').close(); openNew();
  });
  $('#editShiftDetailsButton').addEventListener('click', openEditShift);
  $('#editVehicleForm').addEventListener('submit', saveEditedVehicle);
  $('#editShiftForm').addEventListener('submit', saveEditedShift);
  $('#cancelEditVehicle').addEventListener('click', () => $('#editVehicleDialog').close());
  $('#cancelEditShift').addEventListener('click', () => $('#editShiftDialog').close());
  $('#addInchargeButton').addEventListener('click', () => {
    const entered = prompt('Naye loading incharge ka naam likhen:');
    const name = String(entered || '').trim().toUpperCase();
    if (!name) return;
    if (name === 'AJEET') return showToast('AJAY select karein.');
    if (!inchargeNames().includes(name)) state.inchargeNames.push(name);
    persist(); refreshInchargeOptions(); $('#incharge').value = name;
    showToast('Incharge added. It will sync to other devices after a loading entry is saved to the Sheet.');
  });
  $('#addVehicleButton').addEventListener('click', () => addVehicleRow());
  $('#loadingForm').addEventListener('submit', saveForm);
  $$('[data-view]').forEach(el => el.addEventListener('click', () => { const view = el.dataset.view; if (view === 'new') openNew(); else navigate(view); }));
  ['filterFrom','filterTo','filterShift','filterIncharge'].forEach(id => $(`#${id}`).addEventListener('input', renderReports));
  $('#clearFilters').addEventListener('click', () => { ['filterFrom','filterTo','filterShift','filterIncharge'].forEach(id => $(`#${id}`).value = ''); renderReports(); });
  $('#csvButton').addEventListener('click', downloadCSV);
  ['gapFrom','gapTo','gapSearch'].forEach(id => $(`#${id}`).addEventListener('input', renderGaps));
  $('#clearGapFilters').addEventListener('click', () => { ['gapFrom','gapTo','gapSearch'].forEach(id => $(`#${id}`).value = ''); renderGaps(); });
  $('#printGapReport').addEventListener('click', () => {
    if (!renderGaps().length) return showToast('No product rows to include in this PDF.');
    const oldTitle = document.title; document.title = `Stage-Gap-Report-${todayISO()}`;
    window.addEventListener('afterprint', () => { document.title = oldTitle; }, { once: true });
    window.print();
  });
  const dialog = $('#backupDialog');
  $('#backupButton').addEventListener('click', () => dialog.showModal());
  $('#downloadBackup').addEventListener('click', () => download(`loading-log-backup-${todayISO()}.json`, JSON.stringify({ format: 'starish-loading-log-v1', ...state }, null, 2), 'application/json'));
  $('#restoreFile').addEventListener('change', async event => {
    const file = event.target.files[0]; if (!file) return;
    try {
      const imported = JSON.parse(await file.text());
      if (!Array.isArray(imported.headers) || !Array.isArray(imported.items) || imported.headers.some(h => !h.id || !h.date || !h.shift) || imported.items.some(i => !i.id || !i.loadingId || !i.vehicleNo)) throw new Error('The file does not have the expected loading log format.');
      const replace = confirm(`Replace the ${state.items.length} records saved in this browser with ${imported.items.length} records from the backup?`);
      if (!replace) return;
      state = { headers: imported.headers, items: imported.items, stageRows: imported.stageRows || [], inchargeNames: imported.inchargeNames || [] }; persist(); renderHome(); $('#backupStatus').textContent = 'Backup restored.'; showToast('Backup restored successfully.');
    } catch (error) { $('#backupStatus').textContent = error.message || 'Could not read backup.'; }
    event.target.value = '';
  });
  $('#loadingDate').value = todayISO(); addVehicleRow(); renderHome();
  if (apiUrl) syncFromSheet(); else setStorageStatus('Google Sheet sync is not configured.', 'Records are saved in this browser until the sheet connector is set up.');
})();
