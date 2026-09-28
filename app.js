(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const SESSION_KEY = "fmAssetSession";
  const LOCAL_ANNOTATIONS_KEY = "fmDailyReportAnnotationsV3";
  const DISCARD_STORAGE_KEY = "fmDailyReportDiscardReasons";
  const state = {
    accessToken: null,
    idTokenPayload: null,
    inventory: [],
    transactions: [],
    routines: [],
    scheduleTracking: [],
    notableTracking: [],
    assets: [],
    auditVisibleRows: [],
    scheduleDrafts: new Map(),
    notableDrafts: new Map(),
    reportBlobUrl: "",
    saveBusy: false,
    dataReady: false,
    sourceMode: "none"
  };

  let tokenRequestPromise = null;
  let saveDebounce = null;

  const $ = id => document.getElementById(id);

  const londonDateFormatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/London", year: "numeric", month: "2-digit", day: "2-digit"
  });
  const londonFormatter = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23"
  });

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    const now = new Date();
    const today = londonDateFormatter.format(now);
    $("reportDate").value = today;
    $("signoffDate").value = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
    $("shiftSelect").value = "day";
    buildInventoryForm([]);
    loadDiscardReasons();
    bindEvents();
    updateShiftWindow();
    renderStaticEmptyStates();
    renderSlackSummary();
    waitForGoogle();
  }

  function bindEvents() {
    $("reportDate")?.addEventListener("change", handleContextChange);
    $("shiftSelect")?.addEventListener("change", handleContextChange);
    $("generateBtn")?.addEventListener("click", generateReport);
    $("copySummaryBtn")?.addEventListener("click", copySummary);
    $("downloadSummaryBtn")?.addEventListener("click", downloadSummary);
    $("saveAnnotationsBtn")?.addEventListener("click", () => saveAnnotations(true));
    $("refreshBtn")?.addEventListener("click", () => {
      if (state.sourceMode === "upload") return renderEverything();
      return loadSheetsData();
    });
    $("googleConnectBtn")?.addEventListener("click", () => requestSheetAccess(false));
    $("xlsxFile")?.addEventListener("change", event => {
      const file = event.target.files?.[0];
      if (file) loadWorkbookFromUpload(file);
    });
    $("signOutBtn")?.addEventListener("click", signOut);
    $("routineBody")?.addEventListener("input", onScheduleEdit);
    $("routineBody")?.addEventListener("change", onScheduleEdit);
    $("routineBody")?.addEventListener("click", event => {
      const button = event.target.closest("[data-complete-routine]");
      if (!button) return;
      const key = button.dataset.completeRoutine;
      const draft = state.scheduleDrafts.get(key) || {};
      draft.status = "Completed";
      if (draft.actualQty === "" || draft.actualQty === undefined || draft.actualQty === null) {
        const routine = getScheduleRows().find(item => item.key === key)?.routine;
        if (routine) draft.actualQty = String(routine.quantity);
      }
      state.scheduleDrafts.set(key, draft);
      renderScheduledMovements();
      renderSlackSummary();
      queueSaveAnnotations();
    });
    $("notableBody")?.addEventListener("change", onNotableEdit);
    $("notableBody")?.addEventListener("input", onNotableEdit);
    $("transactionModal")?.addEventListener("click", event => {
      if (event.target.matches("[data-close-modal]")) closeTransactionModal();
    });
    document.addEventListener("input", event => {
      if (event.target.matches("input[data-discard-reason]")) {
        const key = event.target.dataset.key;
        saveDiscardReason(key, event.target.value);
      }
      if (event.target.closest(".form-shell") && !event.target.closest("#routineBody") && !event.target.closest("#notableBody")) {
        renderSlackSummary();
      }
    });
    document.addEventListener("change", event => {
      if (event.target.closest(".form-shell")) renderSlackSummary();
    });
  }

  function renderStaticEmptyStates() {
    $("inventorySummaryPreview").innerHTML = emptyPanel("Connect or upload the ledger to load asset inventory.");
    $("clientMovementPreview").innerHTML = emptyPanel("Connect or upload the ledger to load shift movements.");
    $("discardedAssetsPreview").innerHTML = emptyPanel("Connect or upload the ledger to load discarded assets.");
    $("routineBody").innerHTML = `<div class="empty-panel">Connect or upload the ledger to load scheduled movements.</div>`;
    $("notableBody").innerHTML = `<div class="empty-panel">Connect or upload the ledger to load transactions for this shift.</div>`;
  }

  function emptyPanel(text) {
    return `<div class="empty-panel">${escapeHtml(text)}</div>`;
  }

  function waitForGoogle() {
    if (window.google?.accounts?.oauth2?.initTokenClient) {
      initializeGoogle();
      return;
    }
    let finished = false;
    const started = Date.now();
    const onLoaded = () => {
      if (finished) return;
      finished = true;
      window.removeEventListener("fm-google-loaded", onLoaded);
      initializeGoogle();
    };
    window.addEventListener("fm-google-loaded", onLoaded, { once: true });
    const timer = setInterval(() => {
      if (finished) { clearInterval(timer); return; }
      if (window.google?.accounts?.oauth2?.initTokenClient) {
        clearInterval(timer);
        finished = true;
        window.removeEventListener("fm-google-loaded", onLoaded);
        initializeGoogle();
        return;
      }
      if (Date.now() - started >= 8000) {
        clearInterval(timer);
        finished = true;
        window.removeEventListener("fm-google-loaded", onLoaded);
        // XLSX upload remains available even when Google cannot be reached.
        setAuthStatus("");
      }
    }, 50);
  }

  function initializeGoogle() {
    if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes("PASTE_YOUR")) {
      setAuthStatus("");
      return;
    }

    const saved = readSavedSession();
    const cached = window.FM_AUTH_CACHE?.read?.(saved?.email);
    if (cached?.token) {
      state.accessToken = cached.token;
      state.idTokenPayload = saved || (cached.email ? { email: cached.email } : null);
      loadSheetsData();
    }
  }

  function requestSheetAccess(silent = false) {
    return acquireAccessToken(silent ? "none" : "select_account")
      .then(async () => {
        hideConnection();
        await loadSheetsData();
        return state.accessToken;
      })
      .catch(error => {
        console.warn("Google sign-in failed:", error);
        if (!silent) {
          showConnection();
          setAuthStatus("");
          showToast("Google sign-in unavailable. You can upload the ledger instead.");
        }
        throw error;
      });
  }

  function acquireAccessToken(prompt = "none", email) {
    const saved = readSavedSession();
    const cached = window.FM_AUTH_CACHE?.read?.(email || saved?.email);
    if (cached?.token) {
      state.accessToken = cached.token;
      state.idTokenPayload = saved || (cached.email ? { email: cached.email } : state.idTokenPayload);
      return Promise.resolve(cached.token);
    }
    if (tokenRequestPromise) return tokenRequestPromise;
    if (!window.google?.accounts?.oauth2?.initTokenClient) {
      return Promise.reject(new Error("Google sign-in is not ready yet."));
    }

    tokenRequestPromise = new Promise((resolve, reject) => {
      let settled = false;
      let timeoutId = null;
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        tokenRequestPromise = null;
        if (timeoutId) clearTimeout(timeoutId);
        fn(value);
      };
      const timeoutMs = prompt === "none" ? 7000 : 20000;
      timeoutId = setTimeout(() => finish(reject, new Error("Google authorization timed out.")), timeoutMs);

      try {
        const client = google.accounts.oauth2.initTokenClient({
          client_id: CONFIG.GOOGLE_CLIENT_ID,
          scope: "https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/drive.readonly",
          callback: response => {
            if (response.error) {
              finish(reject, new Error(response.error_description || response.error));
              return;
            }
            state.accessToken = response.access_token;
            const accountEmail = email || saved?.email || "";
            state.idTokenPayload = state.idTokenPayload || (accountEmail ? { email: accountEmail } : null);
            window.FM_AUTH_CACHE?.write?.(response.access_token, response.expires_in, accountEmail);
            if (state.idTokenPayload) saveSession();
            finish(resolve, response.access_token);
          }
        });
        const options = { prompt };
        const loginHint = email || saved?.email;
        if (loginHint) options.login_hint = loginHint;
        client.requestAccessToken(options);
      } catch (error) {
        finish(reject, error);
      }
    });
    return tokenRequestPromise;
  }

  async function loadSheetsData() {
    if (!state.accessToken) return;
    setSyncStatus("Syncing live Sheets data…");
    try {
      const ledgerId = CONFIG.INVENTORY_LEDGER_SHEET_ID;
      const [inventoryRows, transactionRows, routineRows] = await Promise.all([
        getValues(ledgerId, CONFIG.INVENTORY_SHEET_NAME),
        getValues(ledgerId, CONFIG.TRANSACTIONS_SHEET_NAME),
        getValues(ledgerId, CONFIG.ROUTINE_SHEET_NAME)
      ]);

      state.inventory = parseInventory(inventoryRows);
      state.transactions = parseTransactions(transactionRows);
      state.routines = typeof FM_ROUTINE_ALERTS !== "undefined"
        ? FM_ROUTINE_ALERTS.parseRoutineRows(routineRows)
        : parseRoutinesFallback(routineRows);
      loadLocalAnnotations();
      state.sourceMode = "sheets";
      state.dataReady = true;
      state.assets = uniqueAssets([
        ...state.inventory.map(x => x.asset),
        ...state.transactions.map(x => x.asset)
      ]);

      buildInventoryForm(state.assets);
      renderEverything();
      hideConnection();
      const info = shiftInfo();
      const txCount = shiftRows().length;
      setSyncStatus(`Live Sheets data • ${txCount.toLocaleString()} transactions in this shift • synced ${new Date().toLocaleTimeString()}`);
    } catch (error) {
      console.error(error);
      setSyncStatus(error.message || "Could not load the Assets Inventory Ledger.", true);
      if (!state.dataReady) showConnection();
    }
  }

  async function loadWorkbookFromUpload(file) {
    if (!window.XLSX) {
      showToast("Spreadsheet reader is still loading.");
      return;
    }
    try {
      setAuthStatus("");
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const readSheet = name => {
        const ws = workbook.Sheets[name];
        return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true }) : [];
      };
      const inventoryRows = readSheet(CONFIG.INVENTORY_SHEET_NAME);
      const transactionRows = readSheet(CONFIG.TRANSACTIONS_SHEET_NAME);
      const routineRows = readSheet(CONFIG.ROUTINE_SHEET_NAME);
      if (!inventoryRows.length || !transactionRows.length) {
        throw new Error("The ledger needs Asset Inventory and Asset Transactions sheets.");
      }
      state.accessToken = null;
      state.idTokenPayload = null;
      state.inventory = parseInventory(inventoryRows);
      state.transactions = parseTransactions(transactionRows);
      state.routines = typeof FM_ROUTINE_ALERTS !== "undefined"
        ? FM_ROUTINE_ALERTS.parseRoutineRows(routineRows)
        : parseRoutinesFallback(routineRows);
      loadLocalAnnotations();
      state.assets = uniqueAssets([
        ...state.inventory.map(x => x.asset),
        ...state.transactions.map(x => x.asset)
      ]);
      state.sourceMode = "upload";
      state.dataReady = true;
      buildInventoryForm(state.assets);
      renderEverything();
      hideConnection();
      setSyncStatus(`Ledger loaded from ${file.name} • ${shiftRows().length.toLocaleString()} transactions in this shift`);
      showToast("Ledger loaded");
    } catch (error) {
      console.error(error);
      showToast(error.message || "Could not read the ledger.");
    }
  }

  function handleContextChange() {
    state.scheduleDrafts = new Map();
    state.notableDrafts = new Map();
    updateShiftWindow();
    renderEverything();
  }

  function renderEverything() {
    if (!state.dataReady) return;
    applyInventoryForShift();
    renderInventorySummaryPreview();
    renderClientMovementPreview();
    renderDiscardedAssetsTodayPreview();
    renderScheduledMovements();
    renderNotableTransactions();
    renderSlackSummary();
  }

  function shiftInfo() {
    const reportDate = $("reportDate").value;
    const shift = $("shiftSelect").value;
    if (shift === "night") {
      const startDate = previousDateKey(reportDate);
      return {
        key: "night",
        label: "Night shift",
        startDate,
        endDate: reportDate,
        startTime: "19:00",
        endTime: "07:00",
        windowText: `${formatDate(startDate)}, 19:00–${formatDate(reportDate)}, 07:00`
      };
    }
    return {
      key: "day",
      label: "Day shift",
      startDate: reportDate,
      endDate: reportDate,
      startTime: "07:00",
      endTime: "19:00",
      windowText: `${formatDate(reportDate)}, 07:00–${formatDate(reportDate)}, 19:00`
    };
  }

  function updateShiftWindow() {
    const info = shiftInfo();
    $("shiftWindow").innerHTML = `<strong>${escapeHtml(info.label)}</strong><span>${escapeHtml(info.windowText)}</span><small>Report date is the shift end date.</small>`;
    $("rShift").textContent = info.label;
    $("rShiftWindow").textContent = info.windowText;
  }

  function isInSelectedShift(timestamp) {
    const parsed = parseTimestamp(timestamp);
    if (!parsed) return false;
    const p = londonParts(parsed);
    if (!p) return false;
    const info = shiftInfo();
    const mins = p.hour * 60 + p.minute;
    if (info.key === "day") return p.date === info.endDate && mins >= 420 && mins < 1140;
    return (p.date === info.startDate && mins >= 1140) || (p.date === info.endDate && mins < 420);
  }

  function shiftRows() {
    return state.transactions.filter(r => isInSelectedShift(r.timestamp));
  }

  function scheduleDateKey() {
    return shiftInfo().startDate;
  }

  function scheduledRoutineRows() {
    const date = new Date(`${scheduleDateKey()}T12:00:00`);
    const routines = state.routines.filter(r => r.active);
    const due = typeof FM_ROUTINE_ALERTS !== "undefined"
      ? routines.filter(r => FM_ROUTINE_ALERTS.routineIsDueToday(r, date))
      : routines.filter(r => routineDueFallback(r, date));
    return due.sort((a, b) => String(a.plannedTime || "").localeCompare(String(b.plannedTime || "")) || a.client.localeCompare(b.client) || a.asset.localeCompare(b.asset));
  }

  function getScheduleRows() {
    return scheduledRoutineRows().map(routine => {
      const key = scheduleKey(routine.id);
      const saved = trackingForSchedule(key);
      const draft = state.scheduleDrafts.get(key) || {};
      return {
        key,
        routine,
        status: draft.status ?? saved?.status ?? "Pending",
        actualQty: draft.actualQty ?? (saved?.actualQty ?? ""),
        note: draft.note ?? (saved?.note ?? "")
      };
    });
  }

  function scheduleKey(routineId) {
    return `${$("reportDate").value}|${$("shiftSelect").value}|${routineId}`;
  }

  function trackingForSchedule(key) {
    const [reportDate, shift, routineId] = key.split("|");
    return state.scheduleTracking.find(r => r.reportDate === reportDate && r.shift === shift && r.routineId === routineId) || null;
  }

  function renderScheduledMovements() {
    const body = $("routineBody");
    const rows = getScheduleRows();
    $("routineCount").textContent = rows.length.toLocaleString();
    if (!state.dataReady) {
      body.innerHTML = `<div class="empty-panel">Connect or upload the ledger to load scheduled movements.</div>`;
      return;
    }
    if (!rows.length) {
      body.innerHTML = `<div class="empty-panel"><strong>No scheduled routine movements for this shift.</strong><span>Check Routine Outbounds in the main dashboard to add or activate one.</span></div>`;
      return;
    }
    body.innerHTML = rows.map(item => {
      const r = item.routine;
      const statusClass = `status-${item.status.toLowerCase().replace(/\s+/g, "-")}`;
      const actual = item.actualQty === "" ? "" : Number(item.actualQty);
      return `<article class="routine-card ${statusClass}" data-routine-card="${escapeAttr(item.key)}">
        <div class="routine-card-top">
          <div>
            <span class="mini-label">${escapeHtml(r.plannedTime || "No planned time")} · ${escapeHtml(r.direction || "Movement")}</span>
            <h3>${escapeHtml(r.client)}</h3>
            <p>${escapeHtml(r.destination || "Warehouse / destination not specified")}</p>
          </div>
          <div class="routine-planned"><span>Planned</span><strong>${formatQty(r.quantity)}</strong><small>${escapeHtml(r.asset)}</small></div>
        </div>
        <div class="routine-card-fields">
          <label>Status<select data-routine-field="status" data-routine-key="${escapeAttr(item.key)}">
            ${["Pending","Completed","Partial","Not completed"].map(s => `<option ${item.status === s ? "selected" : ""}>${s}</option>`).join("")}
          </select></label>
          <label>Actual qty<input type="number" min="0" step="1" inputmode="numeric" value="${actual === "" ? "" : escapeAttr(actual)}" data-routine-field="actualQty" data-routine-key="${escapeAttr(item.key)}" placeholder="0"></label>
          <label>Notes<textarea rows="2" data-routine-field="note" data-routine-key="${escapeAttr(item.key)}" placeholder="What happened?">${escapeHtml(item.note)}</textarea></label>
          <button type="button" class="complete-routine" data-complete-routine="${escapeAttr(item.key)}">✓ Mark completed</button>
        </div>
      </article>`;
    }).join("");
  }

  function onScheduleEdit(event) {
    const input = event.target.closest("[data-routine-field]");
    if (!input) return;
    const key = input.dataset.routineKey;
    const field = input.dataset.routineField;
    const draft = state.scheduleDrafts.get(key) || {};
    draft[field] = input.value;
    if (field === "status" && input.value === "Completed") {
      const row = getScheduleRows().find(item => item.key === key);
      if (row && !input.closest(".routine-card")?.querySelector('[data-routine-field="actualQty"]')?.value) draft.actualQty = String(row.routine.quantity);
    }
    state.scheduleDrafts.set(key, draft);
    renderSlackSummary();
    queueSaveAnnotations();
  }

  function renderNotableTransactions() {
    const body = $("notableBody");
    const rows = shiftRows().slice().sort((a, b) => compareTimestamp(a.timestamp, b.timestamp));
    state.auditVisibleRows = rows;
    $("transactionCount").textContent = rows.length.toLocaleString();
    if (!state.dataReady) {
      body.innerHTML = `<div class="empty-panel">Connect or upload the ledger to load transactions for this shift.</div>`;
      return;
    }
    if (!rows.length) {
      body.innerHTML = `<div class="empty-panel">No transactions were recorded during this shift.</div>`;
      return;
    }
    body.innerHTML = rows.map((item, index) => {
      const key = transactionKey(item);
      const saved = notableForTransaction(key);
      const draft = state.notableDrafts.get(key) || {};
      const notable = Boolean(draft.notable ?? saved?.notable);
      const note = draft.note ?? (saved?.note ?? "");
      return `<div class="transaction-row ${notable ? "is-notable" : ""}" data-transaction-row="${escapeAttr(key)}">
        <label class="notable-check"><input type="checkbox" data-notable-field="notable" data-transaction-key="${escapeAttr(key)}" ${notable ? "checked" : ""}><span>Notable</span></label>
        <div class="transaction-time">${escapeHtml(formatTime(item.timestamp))}</div>
        <div class="transaction-main"><strong>${escapeHtml(item.client || "Unknown client")}</strong><span>${escapeHtml(item.asset || "Unknown asset")}</span></div>
        <div><span class="movement-pill ${movementClass(item.movement)}">${escapeHtml(item.movement || "OTHER")}</span></div>
        <div class="transaction-qty">${formatQty(item.quantity)}</div>
        <input class="transaction-note" type="text" value="${escapeAttr(note)}" data-notable-field="note" data-transaction-key="${escapeAttr(key)}" placeholder="Reason / context (optional)">
        <button type="button" class="secondary small-button" data-view-transaction="${index}">View</button>
      </div>`;
    }).join("");

    body.querySelectorAll("[data-view-transaction]").forEach(button => {
      button.addEventListener("click", () => openTransactionModal(Number(button.dataset.viewTransaction)));
    });
  }

  function onNotableEdit(event) {
    const input = event.target.closest("[data-notable-field]");
    if (!input) return;
    const key = input.dataset.transactionKey;
    const draft = state.notableDrafts.get(key) || {};
    draft[input.dataset.notableField] = input.type === "checkbox" ? input.checked : input.value;
    state.notableDrafts.set(key, draft);
    const row = input.closest(".transaction-row");
    if (row && input.dataset.notableField === "notable") row.classList.toggle("is-notable", input.checked);
    renderSlackSummary();
    queueSaveAnnotations();
  }

  function loadLocalAnnotations() {
    try {
      const data = JSON.parse(localStorage.getItem(LOCAL_ANNOTATIONS_KEY) || "{}");
      state.scheduleTracking = Array.isArray(data.schedule) ? data.schedule : [];
      state.notableTracking = Array.isArray(data.notable) ? data.notable : [];
    } catch (_) {
      state.scheduleTracking = [];
      state.notableTracking = [];
    }
    state.scheduleDrafts.clear();
    state.notableDrafts.clear();
  }

  function persistLocalAnnotations() {
    try {
      localStorage.setItem(LOCAL_ANNOTATIONS_KEY, JSON.stringify({
        version: 3,
        schedule: state.scheduleTracking,
        notable: state.notableTracking
      }));
    } catch (error) {
      console.warn("Could not save local report annotations:", error);
    }
  }

  function queueSaveAnnotations() {
    clearTimeout(saveDebounce);
    $("annotationStatus").textContent = "Saving on this device…";
    saveDebounce = setTimeout(() => saveAnnotations(false, false), 350);
  }

  async function saveAnnotations(showToastAfter = false, rerender = true) {
    const hasScheduleChanges = state.scheduleDrafts.size > 0;
    const hasNotableChanges = state.notableDrafts.size > 0;
    if (!hasScheduleChanges && !hasNotableChanges) {
      $("annotationStatus").textContent = "Saved on this device";
      if (showToastAfter) showToast("No local changes to save");
      return;
    }

    try {
      if (hasScheduleChanges) {
        const existing = state.scheduleTracking.slice();
        for (const [key, draft] of state.scheduleDrafts.entries()) {
          const row = getScheduleRows().find(item => item.key === key);
          if (!row) continue;
          const idx = existing.findIndex(x => scheduleKeyForTracking(x) === key);
          const data = {
            reportDate: $("reportDate").value,
            shift: $("shiftSelect").value,
            routineId: row.routine.id,
            client: row.routine.client,
            asset: row.routine.asset,
            plannedQty: row.routine.quantity,
            actualQty: draft.actualQty ?? row.actualQty ?? "",
            status: draft.status ?? row.status ?? "Pending",
            destination: row.routine.destination || "",
            plannedTime: row.routine.plannedTime || "",
            note: draft.note ?? row.note ?? "",
            updatedAt: new Date().toISOString(),
            updatedBy: state.idTokenPayload?.email || "Google user"
          };
          if (idx >= 0) existing[idx] = data;
          else existing.push(data);
        }
        state.scheduleTracking = existing;
      }

      if (hasNotableChanges) {
        const existing = state.notableTracking.slice();
        for (const [key, draft] of state.notableDrafts.entries()) {
          const tx = shiftRows().find(item => transactionKey(item) === key);
          if (!tx) continue;
          const idx = existing.findIndex(x => x.transactionKey === key && x.reportDate === $("reportDate").value && x.shift === $("shiftSelect").value);
          const data = {
            reportDate: $("reportDate").value,
            shift: $("shiftSelect").value,
            transactionRow: tx.rowNumber,
            transactionKey: key,
            notable: Boolean(draft.notable),
            note: String(draft.note || "").trim(),
            updatedAt: new Date().toISOString(),
            updatedBy: state.idTokenPayload?.email || "Google user"
          };
          if (idx >= 0) existing[idx] = data;
          else existing.push(data);
        }
        state.notableTracking = existing;
      }

      persistLocalAnnotations();
      state.scheduleDrafts.clear();
      state.notableDrafts.clear();
      if (rerender) {
        renderScheduledMovements();
        renderNotableTransactions();
        renderSlackSummary();
      }
      $("annotationStatus").textContent = `Saved on this device · ${new Date().toLocaleTimeString()}`;
      if (showToastAfter) showToast("Report updates saved on this device");
    } catch (error) {
      console.error(error);
      $("annotationStatus").textContent = "Could not save local report updates";
      if (showToastAfter) showToast("Could not save report updates");
    }
  }

  function scheduleKeyForTracking(item) {
    return `${item.reportDate}|${item.shift}|${item.routineId}`;
  }

  function notableForTransaction(key) {
    return state.notableTracking.find(item => item.reportDate === $("reportDate").value && item.shift === $("shiftSelect").value && item.transactionKey === key) || null;
  }

  function transactionKey(item) {
    return `TX-${item.rowNumber}-${hashString([item.timestamp, item.client, item.movement, item.asset, item.quantity, item.user].join("|"))}`;
  }

  function applyInventoryForShift() {
    const selectedRows = shiftRows();
    const { received, sent, discarded } = movementData(selectedRows);
    const info = shiftInfo();
    const endMinutes = info.endTime === "07:00" ? 420 : 1140;
    state.assets.forEach((asset, i) => {
      const current = Number(state.inventory.find(x => sameName(x.asset, asset))?.balance ?? 0);
      let afterShiftNet = 0;
      state.transactions.forEach(r => {
        if (!sameName(r.asset, asset)) return;
        const d = parseTimestamp(r.timestamp);
        if (!d) return;
        const p = londonParts(d);
        if (!p) return;
        const after = p.date > info.endDate || (p.date === info.endDate && p.hour * 60 + p.minute >= endMinutes);
        if (!after) return;
        if (["RECEIVED", "RETURNED", "COLLECTED"].includes(r.movement)) afterShiftNet += r.quantity;
        else if (["SENT", "DISCARDED", "DISCARD"].includes(r.movement)) afterShiftNet -= r.quantity;
      });
      const closing = current - afterShiftNet;
      const rec = received.filter(r => sameName(r.asset, asset)).reduce((sum, r) => sum + r.quantity, 0);
      const out = sent.filter(r => sameName(r.asset, asset)).reduce((sum, r) => sum + r.quantity, 0);
      const disc = discarded.filter(r => sameName(r.asset, asset)).reduce((sum, r) => sum + r.quantity, 0);
      const opening = closing - rec + out + disc;
      setInputValue(`open_${i}`, Math.max(0, opening));
      setInputValue(`received_${i}`, rec);
      setInputValue(`dispatch_${i}`, out);
      setInputValue(`close_${i}`, Math.max(0, closing));
      setInputValue(`discard_${i}`, disc);
    });
  }

  function buildInventoryForm(assetsList) {
    const assets = assetsList.length ? assetsList : ["Black Pallets", "Wooden Sleeves", "Wood Pallets", "THG Pallets", "Magnum Lids", "Red Sleeves", "Magnum", "Car Bags"];
    state.assets = assets;
    $("inventoryForm").innerHTML = assets.map((asset, i) => `
      <div class="inventory-item">
        <div class="inventory-item-name"><strong>${escapeHtml(asset)}</strong><span>Shift snapshot</span></div>
        <div class="inv-grid">
          <label><small>Opening</small><input type="number" min="0" id="open_${i}" readonly></label>
          <label><small>Received</small><input type="number" min="0" id="received_${i}" readonly></label>
          <label><small>Sent</small><input type="number" min="0" id="dispatch_${i}" readonly></label>
          <label><small>Closing</small><input type="number" min="0" id="close_${i}" readonly></label>
        </div>
      </div>`).join("");
  }

  function renderInventorySummaryPreview() {
    const box = $("inventorySummaryPreview");
    const data = inventorySnapshotData();
    if (!state.dataReady) {
      box.innerHTML = emptyPanel("Connect or upload the ledger to load asset inventory.");
      return;
    }
    if (!data.length) {
      box.innerHTML = emptyPanel("No inventory or movement data was found for this shift.");
      return;
    }
    box.innerHTML = data.map((x, i) => {
      const idx = state.assets.findIndex(a => sameName(a, x.asset));
      const opening = Number($(`open_${idx}`)?.value || 0);
      const closing = Number($(`close_${idx}`)?.value || 0);
      return `<div class="inventory-summary-card">
        <div class="inventory-summary-head"><strong>${escapeHtml(x.asset)}</strong><span class="shift-difference ${x.difference >= 0 ? "positive" : "negative"}">${x.difference >= 0 ? "+" : ""}${formatQty(x.difference)}</span></div>
        <div class="inventory-summary-values">
          <div><span>Opening</span><b>${formatQty(opening)}</b></div>
          <div class="movement-value in"><span>In</span><b>${formatQty(x.received)}</b></div>
          <div class="movement-value out"><span>Out</span><b>${formatQty(x.sent)}</b></div>
          <div><span>Closing</span><b>${formatQty(closing)}</b></div>
        </div>
      </div>`;
    }).join("");
  }

  function inventorySnapshotData() {
    const rows = shiftRows();
    const { received, sent } = movementData(rows);
    const rec = aggregateByAsset(received);
    const out = aggregateByAsset(sent);
    const canonical = new Map();
    const add = value => {
      const label = String(value || "").trim();
      if (!label) return;
      const key = normalize(label);
      if (!canonical.has(key)) canonical.set(key, label);
    };
    state.assets.forEach(add);
    rec.keys().forEach(add);
    out.keys().forEach(add);
    return [...canonical.entries()].map(([key, asset]) => ({
      asset,
      received: rec.get(key) || 0,
      sent: out.get(key) || 0,
      difference: (rec.get(key) || 0) - (out.get(key) || 0)
    })).sort((a, b) => a.asset.localeCompare(b.asset));
  }

  function renderClientMovementPreview() {
    const box = $("clientMovementPreview");
    if (!state.dataReady) {
      box.innerHTML = emptyPanel("Connect or upload the ledger to load shift movements.");
      return;
    }
    const data = combinedClientData(shiftRows());
    if (!data.length) {
      box.innerHTML = emptyPanel("No IN or OUT transactions were recorded during this shift.");
      return;
    }
    box.innerHTML = data.map(c => `<div class="client-preview-card">
      <div class="client-preview-head"><div><span class="client-label">Client</span><strong>${escapeHtml(c.client)}</strong></div></div>
      <div class="client-movement-head"><span>Asset</span><span>In</span><span>Out</span><span>Today</span><span>All-time</span></div>
      <div class="client-movement-list">${c.types.map(t => `<div class="client-movement-row"><span>${escapeHtml(t.asset)}</span><b class="movement-in">${formatQty(t.received)}</b><b class="movement-out">${formatQty(t.sent)}</b><b class="with-client ${t.withClient >= 0 ? "positive" : "negative"}">${t.withClient >= 0 ? "+" : ""}${formatQty(t.withClient)}</b><b class="with-client ${t.allTimeWithClient >= 0 ? "positive" : "negative"}">${t.allTimeWithClient >= 0 ? "+" : ""}${formatQty(t.allTimeWithClient)}</b></div>`).join("")}</div>
    </div>`).join("");
  }

  function renderDiscardedAssetsTodayPreview() {
    const box = $("discardedAssetsPreview");
    if (!state.dataReady) {
      box.innerHTML = emptyPanel("Connect or upload the ledger to load discarded assets.");
      return;
    }
    const data = assetTotals(shiftRows().filter(r => ["DISCARDED", "DISCARD"].includes(r.movement)));
    if (!data.length) {
      box.innerHTML = `<div class="empty-panel">No assets were discarded during this shift.</div>`;
      return;
    }
    const reportDate = $("reportDate").value;
    box.innerHTML = `<div class="discarded-today-head"><span>Asset</span><span>Count</span><span>Reason</span></div>${data.map(x => {
      const key = discardedReasonKey(reportDate, x.asset);
      return `<div class="discarded-today-row"><strong>${escapeHtml(x.asset)}</strong><b>${formatQty(x.quantity)}</b><input type="text" data-discard-reason data-key="${escapeAttr(key)}" value="${escapeAttr(discardedReasons[key] || "")}" placeholder="Enter reason"></div>`;
    }).join("")}`;
  }

  function renderReportInventorySummary() {
    const box = $("rInventorySummaryCards");
    box.innerHTML = inventorySnapshotData().map(x => {
      const idx = state.assets.findIndex(a => sameName(a, x.asset));
      const opening = Number($(`open_${idx}`)?.value || 0);
      const closing = Number($(`close_${idx}`)?.value || 0);
      return `<div class="inventory-summary-card"><div class="inventory-summary-head"><strong>${escapeHtml(x.asset)}</strong><span class="shift-difference ${x.difference >= 0 ? "positive" : "negative"}">${x.difference >= 0 ? "+" : ""}${formatQty(x.difference)}</span></div><div class="inventory-summary-values"><div><span>Opening</span><b>${formatQty(opening)}</b></div><div class="movement-value in"><span>In</span><b>${formatQty(x.received)}</b></div><div class="movement-value out"><span>Out</span><b>${formatQty(x.sent)}</b></div><div><span>Closing</span><b>${formatQty(closing)}</b></div></div></div>`;
    }).join("");
  }

  function renderReportClientMovement(rows) {
    const box = $("rClientMovement");
    const data = combinedClientData(rows);
    if (!data.length) {
      box.innerHTML = `<div class="empty-report">No IN or OUT transactions recorded for this shift.</div>`;
      return;
    }
    box.innerHTML = data.map(c => `<div class="client-report-card"><div class="client-report-head"><div><span>Client</span><h3>${escapeHtml(c.client)}</h3></div></div><div class="client-movement-head"><span>Asset</span><span>In</span><span>Out</span><span>Today</span><span>All-time</span></div><div class="client-movement-list">${c.types.map(t => `<div class="client-movement-row"><span>${escapeHtml(t.asset)}</span><b class="movement-in">${formatQty(t.received)}</b><b class="movement-out">${formatQty(t.sent)}</b><b class="with-client ${t.withClient >= 0 ? "positive" : "negative"}">${t.withClient >= 0 ? "+" : ""}${formatQty(t.withClient)}</b><b class="with-client ${t.allTimeWithClient >= 0 ? "positive" : "negative"}">${t.allTimeWithClient >= 0 ? "+" : ""}${formatQty(t.allTimeWithClient)}</b></div>`).join("")}</div></div>`).join("");
  }

  function renderReportScheduledMovements() {
    const body = $("rRoutineRows");
    const rows = getScheduleRows();
    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="6">No scheduled routine movements for this shift.</td></tr>`;
      return;
    }
    body.innerHTML = rows.map(item => {
      const actual = item.actualQty === "" ? "—" : formatQty(item.actualQty);
      const note = item.note || item.routine.notes || "";
      return `<tr><td>${escapeHtml(item.routine.client)}</td><td>${escapeHtml(item.routine.asset)}</td><td>${formatQty(item.routine.quantity)}</td><td>${actual}</td><td>${escapeHtml(item.status)}</td><td>${escapeHtml(note || "—")}</td></tr>`;
    }).join("");
  }

  function renderReportDiscardedAssets() {
    const box = $("rDiscardedAssetsToday");
    const data = assetTotals(shiftRows().filter(r => ["DISCARDED", "DISCARD"].includes(r.movement)));
    if (!data.length) {
      box.innerHTML = `<div class="empty-report">No assets were discarded during this shift.</div>`;
      return;
    }
    const date = $("reportDate").value;
    box.innerHTML = `<div class="discarded-report-head"><span>Asset</span><span>Count</span><span>Reason</span></div>${data.map(x => {
      const key = discardedReasonKey(date, x.asset);
      return `<div class="discarded-report-row"><strong>${escapeHtml(x.asset)}</strong><b>${formatQty(x.quantity)}</b><span>${escapeHtml(discardedReasons[key] || "Reason not entered")}</span></div>`;
    }).join("")}`;
  }

  async function generateReport() {
    if (!state.dataReady) {
      showToast("Load the ledger first");
      return;
    }
    try {
      applyInventoryForShift();
      renderEverything();
      updateReportDom();
      await saveAnnotations(false, false);
      await createReportImage();
    } catch (error) {
      console.error("Report generation failed:", error);
      showToast(error?.message || "Could not generate the report");
    }
  }

  function updateReportDom() {
    const info = shiftInfo();
    $("rDate").textContent = formatDate($("reportDate").value);
    $("rShift").textContent = info.label;
    $("rShiftWindow").textContent = info.windowText;
    $("rPrepared").textContent = textValue("preparedBy");
    const status = $("overallStatus").value;
    $("rStatus").textContent = status;
    $("rStatus").className = `status-pill ${status === "Critical" ? "critical" : status === "Attention Required" ? "attention" : ""}`;
    renderReportInventorySummary();
    renderReportClientMovement(shiftRows());
    renderReportScheduledMovements();
    renderReportDiscardedAssets();
    $("rReceivedNotes").textContent = textValue("receivedNotes");
    $("rSentNotes").textContent = textValue("sentNotes");
    fillIssueTable();
    $("rTomorrow").textContent = textValue("tomorrowPlan");
    $("rSignoff").textContent = formatDateTime($("signoffDate").value);
    $("rEodStatus").textContent = $("eodStatus").value;
  }

  async function createReportImage() {
    const report = $("report");
    const button = $("generateBtn");
    button.disabled = true;
    button.textContent = "Generating…";
    try {
      if (typeof window.html2canvas !== "function") {
        throw new Error("Report image tool is still loading. Please try again.");
      }

      const canvas = await Promise.race([
        window.html2canvas(report, {
          scale: 1.75,
          backgroundColor: "#fff",
          useCORS: true,
          allowTaint: false,
          logging: false,
          imageTimeout: 10000,
          removeContainer: true
        }),
        new Promise((_, reject) => setTimeout(() => reject(new Error("Report image generation timed out. Please try again.")), 30000))
      ]);

      const blob = await new Promise((resolve, reject) => {
        let settled = false;
        const finish = value => { if (!settled) { settled = true; resolve(value); } };
        const fail = error => { if (!settled) { settled = true; reject(error); } };
        try {
          canvas.toBlob(blob => blob ? finish(blob) : fail(new Error("The PNG could not be created.")), "image/png");
          setTimeout(() => fail(new Error("PNG export timed out. Please try again.")), 15000);
        } catch (error) {
          fail(error);
        }
      });

      if (state.reportBlobUrl) URL.revokeObjectURL(state.reportBlobUrl);
      state.reportBlobUrl = URL.createObjectURL(blob);
      $("reportPreview").src = state.reportBlobUrl;
      $("reportPreviewWrap").classList.remove("hidden");
      $("downloadPngBtn").onclick = () => downloadBlob(blob, reportFileName(), "image/png");
      downloadBlob(blob, reportFileName(), "image/png");
      renderSlackSummary();
      $("outputCard").classList.remove("hidden");
      showToast("PNG report generated and downloaded");
    } finally {
      button.disabled = false;
      button.textContent = "Generate report + Slack summary";
    }
  }

  function reportFileName() {
    const date = $("reportDate").value || "report";
    const shift = $("shiftSelect").value === "night" ? "Night" : "Day";
    return `EOD_Asset_Report_${date.replaceAll("-", "")}_${shift}.png`;
  }

  function renderSlackSummary() {
    const area = $("slackSummary");
    if (!area) return;
    area.value = buildSlackSummary();
    $("summaryMeta").textContent = `${shiftRows().length.toLocaleString()} shift transactions • ${getScheduleRows().length.toLocaleString()} scheduled movements • ${selectedNotables().length.toLocaleString()} notable`;
  }

  function buildSlackSummary() {
    const info = shiftInfo();
    const lines = [];
    lines.push("[:package:](https://a.slack-edge.com/production-standard-emoji-assets/16.0/apple-medium/1f4e6@2x.png) ASSET COMPLIANCE — END OF SHIFT REPORT");
    lines.push(`Shift: ${info.label} • ${info.windowText}`);
    lines.push(`Report date: ${formatDate($("reportDate").value)}`);
    lines.push("");

    lines.push("📦 ASSET INVENTORY");
    inventorySnapshotData().forEach(item => {
      const idx = state.assets.findIndex(a => sameName(a, item.asset));
      const opening = Number($(`open_${idx}`)?.value || 0);
      const closing = Number($(`close_${idx}`)?.value || 0);
      lines.push(`• ${item.asset}: ${formatQty(closing)} closing (Opening ${formatQty(opening)} | In ${formatQty(item.received)} | Out ${formatQty(item.sent)})`);
    });
    if (!inventorySnapshotData().length) lines.push("• No inventory movement data loaded.");
    lines.push("");

    lines.push("👥 CLIENT MOVEMENTS");
    const clientData = combinedClientData(shiftRows());
    if (clientData.length) {
      clientData.forEach(client => {
        client.types.forEach(t => {
          const pieces = [];
          if (t.sent) pieces.push(`sent ${formatQty(t.sent)}`);
          if (t.received) pieces.push(`received ${formatQty(t.received)}`);
          pieces.push(`with client ${t.allTimeWithClient >= 0 ? "+" : ""}${formatQty(t.allTimeWithClient)}`);
          lines.push(`• ${client.client} — ${t.asset}: ${pieces.join(" | ")}`);
        });
      });
    } else lines.push("• No client movements recorded.");
    lines.push("");

    lines.push("📅 SCHEDULED MOVEMENTS");
    const schedule = getScheduleRows();
    if (schedule.length) {
      schedule.forEach(item => {
        const actual = item.actualQty === "" ? "—" : formatQty(item.actualQty);
        const status = statusEmoji(item.status);
        const destination = item.routine.destination ? ` → ${item.routine.destination}` : "";
        const note = item.note ? ` | Note: ${item.note}` : "";
        lines.push(`• ${status} ${item.status.toUpperCase()} — ${item.routine.client} — ${item.routine.asset} — Planned ${formatQty(item.routine.quantity)} | Actual ${actual}${destination}${note}`);
      });
    } else lines.push("• No scheduled movements for this shift.");
    lines.push("");

    lines.push(`⭐ NOTABLE TRANSACTIONS (${selectedNotables().length} selected from ${shiftRows().length})`);
    const notable = selectedNotables();
    if (notable.length) {
      notable.forEach(({ item, note }) => {
        const context = note ? ` — ${note}` : (item.comment ? ` — ${item.comment}` : "");
        lines.push(`• ${formatTime(item.timestamp)} — ${item.client || "Unknown client"} — ${item.movement || "OTHER"} — ${formatQty(item.quantity)} ${item.asset}${context}`);
      });
    } else lines.push("• None selected yet.");
    lines.push("");

    lines.push("🗑️ DISCARDED ASSETS");
    const discarded = assetTotals(shiftRows().filter(r => ["DISCARDED", "DISCARD"].includes(r.movement)));
    if (discarded.length) {
      const date = $("reportDate").value;
      discarded.forEach(item => {
        const reason = discardedReasons[discardedReasonKey(date, item.asset)] || "Reason not entered";
        lines.push(`• ${item.asset}: ${formatQty(item.quantity)} — ${reason}`);
      });
    } else lines.push("• None recorded.");
    lines.push("");

    lines.push("⚠️ OUTSTANDING ISSUES / RISKS");
    const issues = collectIssueRows();
    if (issues.length) issues.forEach(row => lines.push(`• [${row[0] || "Normal"}] ${row[1] || "Issue"} | ${row[2] || ""} | Impact: ${row[3] || ""} | Action: ${row[4] || ""} | Due: ${row[5] || ""} | Status: ${row[6] || ""}`));
    else lines.push("• None recorded.");
    lines.push("");

    lines.push("📌 PLAN FOR TOMORROW");
    lines.push(textValue("tomorrowPlan", "• No plan entered."));
    lines.push("");

    lines.push(`Overall status: ${$("overallStatus").value}`);
    lines.push(`EOD status: ${$("eodStatus").value}`);
    lines.push(`Sign-off: ${formatDateTime($("signoffDate").value)}`);
    return lines.join("\n");
  }

  function selectedNotables() {
    const rows = shiftRows();
    return rows.map(item => {
      const key = transactionKey(item);
      const saved = notableForTransaction(key);
      const draft = state.notableDrafts.get(key) || {};
      const isNotable = Boolean(draft.notable ?? saved?.notable);
      return { item, note: String(draft.note ?? saved?.note ?? "").trim(), isNotable };
    }).filter(x => x.isNotable);
  }

  function openTransactionModal(index) {
    const item = state.auditVisibleRows[index];
    if (!item) return;
    $("modalClient").textContent = item.client || "—";
    $("modalMovement").textContent = item.movement || "—";
    $("modalAsset").textContent = item.asset || "—";
    $("modalQuantity").textContent = formatQty(item.quantity);
    $("modalTimestamp").textContent = `${formatDateTimeFromTimestamp(item.timestamp)} (${formatTime(item.timestamp)})`;
    $("modalUser").textContent = item.user || "—";
    $("modalComment").textContent = item.comment || "No comment recorded.";
    const photoWrap = $("modalPhotoWrap");
    const photo = $("modalPhoto");
    const link = $("modalPhotoLink");
    if (item.image) {
      photoWrap.classList.remove("hidden");
      photo.removeAttribute("src");
      link.href = item.image;
      if (window.FM_MEDIA?.loadDriveImage && state.accessToken) {
        FM_MEDIA.loadDriveImage(photo, item.image, state.accessToken).catch(() => { photo.src = item.image; });
      } else photo.src = item.image;
    } else {
      photoWrap.classList.add("hidden");
      photo.removeAttribute("src");
      link.removeAttribute("href");
    }
    $("transactionModal").classList.remove("hidden");
  }

  function closeTransactionModal() {
    $("transactionModal").classList.add("hidden");
  }

  function fillIssueTable() {
    const rows = collectIssueRows();
    $("rIssueRows").innerHTML = rows.length
      ? rows.map(r => `<tr>${r.map(v => `<td>${escapeHtml(v || "—")}</td>`).join("")}</tr>`).join("")
      : `<tr><td colspan="7">No entries recorded</td></tr>`;
  }

  function collectIssueRows() {
    return [...$("issueRows").children].map(row => [...row.querySelectorAll("input,select")].map(x => x.value.trim())).filter(row => row.some(Boolean));
  }

  function addIssueRow(values = []) {
    const row = document.createElement("div");
    row.className = "entry-row issue";
    row.innerHTML = `<select><option>Low</option><option>Medium</option><option>High</option><option>Critical</option></select><input placeholder="Issue"><input placeholder="Asset / Client"><input placeholder="Impact"><input placeholder="Action"><input type="date"><select><option>Open</option><option>In Progress</option><option>Resolved</option></select><button type="button" class="remove-row" aria-label="Remove issue">×</button>`;
    row.querySelector("button")?.addEventListener("click", () => { row.remove(); renderSlackSummary(); });
    const controls = [...row.querySelectorAll("input,select")];
    values.forEach((value, index) => { if (controls[index]) controls[index].value = value; });
    $("issueRows").appendChild(row);
    renderSlackSummary();
  }

  function clearForm() {
    if (!confirm("Clear the manual report fields? Live Sheets data and saved schedule/notable history will remain.")) return;
    document.querySelectorAll(".form-shell input,.form-shell textarea").forEach(el => {
      if (el.id === "reportDate" || el.id === "signoffDate") return;
      if (el.type === "number") el.value = "";
      else if (el.tagName === "SELECT") el.selectedIndex = 0;
      else el.value = "";
    });
    $("issueRows").innerHTML = "";
    state.scheduleDrafts.clear();
    state.notableDrafts.clear();
    updateShiftWindow();
    renderEverything();
    showToast("Manual report fields cleared");
  }

  async function copySummary() {
    const text = $("slackSummary").value;
    try {
      await navigator.clipboard.writeText(text);
    } catch (_) {
      $("slackSummary").focus();
      $("slackSummary").select();
      document.execCommand("copy");
    }
    showToast("Slack summary copied");
  }

  function downloadSummary() {
    const blob = new Blob([$("slackSummary").value], { type: "text/plain;charset=utf-8" });
    downloadBlob(blob, reportFileName().replace(/\.png$/i, ".txt"), "text/plain");
  }

  function downloadBlob(blob, filename, type) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.type = type;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function signOut() {
    window.FM_AUTH_CACHE?.clear?.();
    localStorage.removeItem(SESSION_KEY);
    state.accessToken = null;
    state.idTokenPayload = null;
    state.dataReady = false;
    state.sourceMode = "none";
    showConnection();
    setAuthStatus("");
  }

  function showConnection() {
    $("connectionCard").classList.remove("compact-connected");
    $("appShell").classList.add("hidden");
    $("signOutBtn").classList.add("hidden");
    $("googleConnectBtn")?.classList.remove("hidden");
    $("connectedUserRow")?.classList.add("hidden");
    $("connectedUser").textContent = "Not connected";
  }

  function hideConnection() {
    $("connectionCard").classList.add("compact-connected");
    $("appShell").classList.remove("hidden");
    $("signOutBtn").classList.remove("hidden");
    $("googleConnectBtn")?.classList.add("hidden");
    $("connectedUserRow")?.classList.remove("hidden");
    $("connectedUser").textContent = state.idTokenPayload?.email || window.FM_AUTH_CACHE?.read?.()?.email || "Google account";
  }

  function setUserProfile(profile) {
    $("connectedUser").textContent = profile?.email || "Google account";
  }

  function saveSession() {
    if (!state.idTokenPayload) return;
    try {
      localStorage.setItem(SESSION_KEY, JSON.stringify({
        name: state.idTokenPayload.name || "Google user",
        email: state.idTokenPayload.email || "",
        picture: state.idTokenPayload.picture || "",
        sub: state.idTokenPayload.sub || ""
      }));
    } catch (_) {}
  }

  function readSavedSession() {
    try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch (_) { return null; }
  }

  function setAuthStatus(message, isError = false) {
    $("authStatus").textContent = message || "";
    $("authStatus").className = `auth-status ${isError ? "error" : ""}`;
  }

  function setSyncStatus(message, isError = false) {
    $("syncStatus").textContent = message;
    $("syncStatus").className = `sync-status ${isError ? "error" : ""}`;
  }


  function authHeaders() { return { Authorization: `Bearer ${state.accessToken}` }; }

  async function sheetsGet(path) { return fetchJson(SHEETS_API + path, { headers: authHeaders() }); }

  async function fetchJson(url, options = {}) {
    const response = await fetch(url, options);
    if (response.status === 401 && !options.__retried) {
      try {
        await acquireAccessToken("none");
        return fetchJson(url, { ...options, __retried: true, headers: { ...(options.headers || {}), ...authHeaders() } });
      } catch (_) {
        throw new Error("Google access expired. Sign in again.");
      }
    }
    const raw = await response.text();
    let data = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch (_) {}
    if (!response.ok) throw new Error(data?.error?.message || `Request failed (${response.status})`);
    return data;
  }

  async function getValues(spreadsheetId, sheetName) {
    const range = `${quoteSheetName(sheetName)}!A:AE`;
    const data = await sheetsGet(`/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`);
    return data.values || [];
  }


  function parseInventory(rows) {
    if (!rows?.length) return [];
    const h = rows[0].map(normalizeHeader);
    const assetIndex = findHeader(h, ["asset"]);
    const balanceIndex = findHeader(h, ["balance", "quantity", "qty", "count"]);
    if (assetIndex < 0 || balanceIndex < 0) throw new Error('"Asset Inventory" needs Asset and Balance columns.');
    return rows.slice(1).map(row => ({ asset: valueAt(row, assetIndex), balance: Number(String(valueAt(row, balanceIndex)).replace(/,/g, "")) || 0 })).filter(x => x.asset);
  }

  function parseTransactions(rows) {
    if (!rows?.length) return [];
    const headerRowIndex = rows.findIndex(row => row.some(value => normalizeHeader(value) === "timestamp"));
    if (headerRowIndex < 0) throw new Error('Could not find Timestamp in "Asset Transactions".');
    const h = rows[headerRowIndex].map(normalizeHeader);
    const i = {
      timestamp: findHeader(h, ["timestamp", "date", "datetime"]), client: findHeader(h, ["client", "client name"]), movement: findHeader(h, ["movement", "type", "direction"]),
      asset: findHeader(h, ["asset", "asset name", "item"]), quantity: findHeader(h, ["quantity", "qty"]), user: findHeader(h, ["user", "entered by", "email"]),
      comment: findHeader(h, ["comment", "comments", "notes"]), image: findHeader(h, ["image link", "image", "photo", "picture", "photo link", "attachment", "drive link"])
    };
    if ([i.timestamp, i.client, i.movement, i.asset, i.quantity].some(v => v < 0)) throw new Error('"Asset Transactions" needs Timestamp, Client, Movement, Asset and Quantity columns.');
    return rows.slice(headerRowIndex + 1).map((row, offset) => ({
      rowNumber: headerRowIndex + offset + 2,
      timestamp: valueAt(row, i.timestamp), client: valueAt(row, i.client), movement: valueAt(row, i.movement).toUpperCase(), asset: valueAt(row, i.asset),
      quantity: numberAt(row, i.quantity), user: valueAt(row, i.user), comment: valueAt(row, i.comment), image: valueAt(row, i.image)
    })).filter(x => x.asset || x.client);
  }

  function parseRoutinesFallback(rows) {
    if (!rows?.length) return [];
    const h = rows[0].map(normalizeHeader);
    const i = { id: findHeader(h, ["routine id", "id"]), active: findHeader(h, ["active", "enabled"]), frequency: findHeader(h, ["frequency", "schedule"]), weekday: findHeader(h, ["weekday", "day"]), direction: findHeader(h, ["direction", "movement"]), client: findHeader(h, ["client", "client name"]), asset: findHeader(h, ["asset", "asset name", "item"]), quantity: findHeader(h, ["quantity", "qty"]), destination: findHeader(h, ["destination", "location", "site"]), plannedTime: findHeader(h, ["planned time", "time"]), notes: findHeader(h, ["notes", "comment"]) };
    return rows.slice(1).map((row, n) => ({
      rowNumber: n + 2, id: valueAt(row, i.id) || `legacy-${n + 2}`, active: isTruthy(valueAt(row, i.active)), frequency: valueAt(row, i.frequency), weekday: valueAt(row, i.weekday), direction: normalizeDirection(valueAt(row, i.direction)), client: valueAt(row, i.client), asset: valueAt(row, i.asset), quantity: numberAt(row, i.quantity), destination: valueAt(row, i.destination), plannedTime: valueAt(row, i.plannedTime), notes: valueAt(row, i.notes)
    })).filter(x => x.client && x.asset && x.quantity > 0 && x.direction);
  }

  function routineDueFallback(routine, date) {
    if (!routine.active) return false;
    const frequency = normalize(routine.frequency);
    if (frequency === "daily" || frequency === "every day" || frequency === "everyday") return true;
    if (frequency === "weekly" || frequency === "every week") return normalize(routine.weekday) === ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"][date.getDay()];
    return false;
  }

  function movementData(rows) {
    return {
      received: rows.filter(r => r.movement === "RECEIVED"),
      sent: rows.filter(r => r.movement === "SENT"),
      discarded: rows.filter(r => ["DISCARDED", "DISCARD"].includes(r.movement))
    };
  }

  function aggregateByAsset(rows) {
    const map = new Map();
    rows.forEach(r => {
      const key = normalize(r.asset);
      map.set(key, (map.get(key) || 0) + r.quantity);
    });
    return map;
  }

  function aggregateByClient(rows) {
    const clients = new Map();
    rows.forEach(r => {
      const client = r.client || "Unknown Client";
      const asset = r.asset || "Unknown Asset";
      if (!clients.has(client)) clients.set(client, new Map());
      const map = clients.get(client);
      map.set(asset, (map.get(asset) || 0) + r.quantity);
    });
    return clients;
  }

  function combinedClientData(rows) {
    const { received, sent } = movementData(rows);
    const receivedMap = aggregateByClient(received);
    const sentMap = aggregateByClient(sent);
    const allTime = allTimeClientBalances();
    const names = new Set([...receivedMap.keys(), ...sentMap.keys()]);
    return [...names].sort((a, b) => a.localeCompare(b)).map(client => {
      const r = receivedMap.get(client) || new Map();
      const s = sentMap.get(client) || new Map();
      const t = allTime.get(client) || new Map();
      const assets = new Set([...r.keys(), ...s.keys()]);
      return {
        client,
        types: [...assets].sort((a, b) => a.localeCompare(b)).map(asset => ({
          asset,
          received: r.get(asset) || 0,
          sent: s.get(asset) || 0,
          withClient: (s.get(asset) || 0) - (r.get(asset) || 0),
          allTimeWithClient: t.get(asset) || 0
        }))
      };
    });
  }

  function allTimeClientBalances() {
    const balances = new Map();
    state.transactions.forEach(r => {
      const client = r.client || "Unknown Client";
      const asset = r.asset || "Unknown Asset";
      if (!client || !asset || /hsc london\s*\(self\)/i.test(client)) return;
      if (!["SENT", "RECEIVED", "RETURNED", "COLLECTED"].includes(r.movement)) return;
      if (!balances.has(client)) balances.set(client, new Map());
      const map = balances.get(client);
      const current = map.get(asset) || 0;
      map.set(asset, r.movement === "SENT" ? current + r.quantity : current - r.quantity);
    });
    return balances;
  }

  function assetTotals(rows) {
    const map = new Map();
    rows.forEach(r => map.set(r.asset, (map.get(r.asset) || 0) + r.quantity));
    return [...map.entries()].map(([asset, quantity]) => ({ asset, quantity })).sort((a, b) => a.asset.localeCompare(b.asset));
  }

  function discardedReasonKey(date, asset) { return `${date}::${asset}`; }

  let discardedReasons = {};
  function loadDiscardReasons() {
    try { discardedReasons = JSON.parse(localStorage.getItem(DISCARD_STORAGE_KEY) || "{}"); } catch (_) { discardedReasons = {}; }
  }
  function saveDiscardReason(key, value) {
    discardedReasons[key] = value;
    try { localStorage.setItem(DISCARD_STORAGE_KEY, JSON.stringify(discardedReasons)); } catch (_) {}
    renderSlackSummary();
  }

  function fillIssueRowsFromLegacyHook() {
    // Kept as a named compatibility helper for old inline references.
    return collectIssueRows();
  }
  window.addIssueRow = addIssueRow;

  function formatImportedTime(value) { return formatTime(value); }
  function formatDateTime(v) {
    if (!v) return "—";
    const [date, time] = String(v).split("T");
    return `${formatDate(date)} • ${time || ""}`.trim();
  }
  function formatDateTimeFromTimestamp(v) {
    const d = parseTimestamp(v);
    if (!d) return String(v || "—");
    const p = londonParts(d);
    return p ? `${formatDate(p.date)} • ${String(p.hour).padStart(2, "0")}:${String(p.minute).padStart(2, "0")}:${String(p.second).padStart(2, "0")}` : String(v || "—");
  }
  function formatDate(v) {
    if (!v) return "—";
    const parts = String(v).split("-");
    return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : String(v);
  }
  function textValue(id, fallback = "—") { return ($(id)?.value || "").trim() || fallback; }
  function formatQty(v) { return numberAtValue(v).toLocaleString("en-GB"); }
  function numberAtValue(v) { const n = Number(String(v ?? "").replace(/,/g, "").trim()); return Number.isFinite(n) ? n : 0; }
  function numberAt(row, index) { return index >= 0 ? numberAtValue(row?.[index]) : 0; }
  function valueAt(row, index) { return index >= 0 ? String(row?.[index] ?? "").trim() : ""; }
  function normalizeHeader(v) { return String(v ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, ""); }
  function normalize(v) { return String(v ?? "").trim().toLowerCase().replace(/\s+/g, " "); }
  function findHeader(headers, aliases) { const wanted = aliases.map(normalizeHeader); return headers.findIndex(h => wanted.includes(h)); }
  function normalizeDirection(v) { const n = normalize(v).toUpperCase(); return n === "RECEIVED" || n === "INBOUND" || n === "IN" ? "Inbound" : n === "SENT" || n === "OUTBOUND" || n === "OUT" ? "Outbound" : ""; }
  function isTruthy(v) { return [true, 1, "true", "1", "yes", "y", "on", "active", "enabled"].includes(typeof v === "string" ? normalize(v) : v); }
  function uniqueAssets(values) { const map = new Map(); values.filter(Boolean).forEach(v => { const key = normalize(v); if (!map.has(key)) map.set(key, String(v).trim()); }); return [...map.values()]; }
  function sameName(a, b) { return normalize(a) === normalize(b); }
  function setInputValue(id, value) { if ($(id)) $(id).value = String(value); }

  function parseTimestamp(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number" && value > 0 && value < 100000) return new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    const raw = String(value ?? "").trim();
    if (!raw) return null;
    const uk = raw.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (uk) {
      const [, dd, mm, yyyy, hh = "0", min = "0", sec = "0"] = uk;
      const d = new Date(`${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}T${hh.padStart(2, "0")}:${min}:${sec}`);
      return Number.isNaN(d.getTime()) ? null : d;
    }
    const isoish = raw.includes(" ") && /^\d{4}-\d{2}-\d{2}\s/.test(raw) ? raw.replace(" ", "T") : raw;
    const d = new Date(isoish);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function londonParts(value) {
    const d = parseTimestamp(value);
    if (!d) return null;
    const parts = londonFormatter.formatToParts(d);
    const get = type => parts.find(p => p.type === type)?.value || "";
    return { date: `${get("year")}-${get("month")}-${get("day")}`, hour: Number(get("hour")), minute: Number(get("minute")), second: Number(get("second")) };
  }

  function formatTime(value) {
    const p = londonParts(value);
    return p ? `${String(p.hour).padStart(2, "0")}:${String(p.minute).padStart(2, "0")}` : String(value || "—");
  }

  function localDateKey(value) { return londonParts(value)?.date || ""; }

  function previousDateKey(date) {
    const d = new Date(`${date}T12:00:00Z`);
    d.setUTCDate(d.getUTCDate() - 1);
    return d.toISOString().slice(0, 10);
  }

  function compareTimestamp(a, b) {
    const aa = parseTimestamp(a)?.getTime() || 0;
    const bb = parseTimestamp(b)?.getTime() || 0;
    return aa - bb;
  }

  function movementClass(movement) {
    const value = String(movement || "").toUpperCase();
    if (value === "RECEIVED") return "received";
    if (value === "SENT") return "sent";
    if (["DISCARD", "DISCARDED"].includes(value)) return "discarded";
    return "other";
  }

  function statusEmoji(status) {
    if (status === "Completed") return "✅";
    if (status === "Partial") return "⚠️";
    if (status === "Not completed") return "⛔";
    return "⏳";
  }

  function hashString(value) {
    let hash = 2166136261;
    for (let i = 0; i < value.length; i++) {
      hash ^= value.charCodeAt(i);
      hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
    }
    return (hash >>> 0).toString(16).toUpperCase();
  }

  function quoteSheetName(name) { return `'${String(name).replace(/'/g, "''")}'`; }
  function escapeHtml(value) { return String(value ?? "").replace(/[&<>'"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c])); }
  function escapeAttr(value) { return escapeHtml(value).replace(/`/g, "&#96;"); }
  function showToast(message) { const el = $("toast"); el.textContent = message; el.classList.add("show"); setTimeout(() => el.classList.remove("show"), 2600); }

  // Public helpers kept for compatibility with older templates.
  window.formatImportedTime = formatImportedTime;
  window.fillIssueRowsFromLegacyHook = fillIssueRowsFromLegacyHook;
})();
