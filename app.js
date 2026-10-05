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
    reportPdfBlobUrl: "",
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
      // The Slack summary is intentionally editable. Do not regenerate it
      // while the user is typing, or their manual edits would be lost.
      if (event.target.id === "slackSummary") return;
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

    if (shift === "24hr") {
      const endDate = nextDateKey(reportDate);
      return {
        key: "24hr",
        label: "24-hour period",
        reportType: "End of Day report",
        coverageText: "Day + Night shifts",
        startDate: reportDate,
        endDate,
        startTime: "07:00",
        endTime: "07:00",
        windowText: `${formatDate(reportDate)}, 07:00–${formatDate(endDate)}, 07:00`
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

  function isEndOfDayReport() {
    return $("shiftSelect")?.value === "24hr";
  }

  function reportTypeLabel() {
    return isEndOfDayReport() ? "End of Day report" : "End of Shift report";
  }

  function coverageLabel() {
    return isEndOfDayReport() ? "Day + Night shifts" : shiftInfo().label;
  }

  function updateShiftWindow() {
    const info = shiftInfo();
    const eod = isEndOfDayReport();
    const dateNote = eod
      ? "End of Day report covering the full Day + Night shifts. Report date is the 07:00 start date."
      : info.key === "night"
        ? "Report date is the night shift end date."
        : "Report date is the shift date.";
    $("shiftWindow").innerHTML = `<strong>${escapeHtml(eod ? "End of Day • Day + Night shifts" : info.label)}</strong><span>${escapeHtml(info.windowText)}</span><small>${escapeHtml(dateNote)}</small>`;
    $("rShiftWindow").textContent = info.windowText;
    $("pageTitle") && ($("pageTitle").textContent = eod ? "Asset Compliance — End of Day Report" : "Asset Compliance — End of Shift Report");
    $("rCoverage") && ($("rCoverage").textContent = coverageLabel());
    $("formPageTitle") && ($("formPageTitle").textContent = eod ? "End-of-Day Report" : "End-of-Shift Report");
    $("formPageSubtitle") && ($("formPageSubtitle").textContent = eod
      ? "Live from the Assets Inventory Ledger · covers both the Day and Night shifts from 07:00 to 07:00."
      : "Live from the Assets Inventory Ledger · built for quick shift handover.");
    $("reportSetupTitle") && ($("reportSetupTitle").textContent = eod ? "End of Day details" : "Shift details");
    $("inventorySectionTitle") && ($("inventorySectionTitle").textContent = eod ? "Daily inventory snapshot" : "Shift inventory snapshot");
    $("reviewSectionKicker") && ($("reviewSectionKicker").textContent = eod ? "Day + night review" : "Shift review");
    $("signoffSectionTitle") && ($("signoffSectionTitle").textContent = eod ? "End-of-Day sign-off" : "End-of-shift sign-off");
    $("reportPreview") && ($("reportPreview").alt = eod ? "Generated end-of-day report preview covering day and night shifts" : "Generated end-of-shift report preview");
  }

  function isInSelectedShift(timestamp) {
    const parsed = parseTimestamp(timestamp);
    if (!parsed) return false;
    const p = londonParts(parsed);
    if (!p) return false;
    const info = shiftInfo();
    const mins = p.hour * 60 + p.minute;
    if (info.key === "day") return p.date === info.endDate && mins >= 420 && mins < 1140;
    if (info.key === "24hr") {
      return (p.date === info.startDate && mins >= 420) ||
             (p.date === info.endDate && mins < 420);
    }
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
      box.innerHTML = `<div class="empty-report">${isEndOfDayReport() ? "No IN or OUT transactions recorded for this 24-hour Day + Night period." : "No IN or OUT transactions recorded for this shift."}</div>`;
      return;
    }
    box.innerHTML = data.map((c, clientIndex) => `<div class="client-report-card" data-client-key="${escapeAttr(String(clientIndex) + "::" + hashString(c.client))}"><div class="client-report-head"><div><span>Client</span><h3>${escapeHtml(c.client)}</h3><a class="report-client-link" href="#" data-report-client-link="${escapeAttr(String(clientIndex) + "::" + hashString(c.client))}">See transactions →</a></div></div><div class="client-movement-head"><span>Asset</span><span>In</span><span>Out</span><span>Today</span><span>All-time</span></div><div class="client-movement-list">${c.types.map(t => `<div class="client-movement-row"><span>${escapeHtml(t.asset)}</span><b class="movement-in">${formatQty(t.received)}</b><b class="movement-out">${formatQty(t.sent)}</b><b class="with-client ${t.withClient >= 0 ? "positive" : "negative"}">${t.withClient >= 0 ? "+" : ""}${formatQty(t.withClient)}</b><b class="with-client ${t.allTimeWithClient >= 0 ? "positive" : "negative"}">${t.allTimeWithClient >= 0 ? "+" : ""}${formatQty(t.allTimeWithClient)}</b></div>`).join("")}</div></div>`).join("");
  }

  function renderReportScheduledMovements() {
    const body = $("rRoutineRows");
    const rows = getScheduleRows();
    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="6">${isEndOfDayReport() ? "No scheduled routine movements recorded across the Day + Night period." : "No scheduled routine movements for this shift."}</td></tr>`;
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
      box.innerHTML = `<div class="empty-report">${isEndOfDayReport() ? "No assets were discarded during the Day + Night period." : "No assets were discarded during this shift."}</div>`;
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
      // Only refresh the lightweight values that changed. Rebuilding the full
      // notable-transactions DOM at generation time can be very expensive.
      renderInventorySummaryPreview();
      renderSlackSummary();
      await saveAnnotations(false, false);
      await createReportPdf();
    } catch (error) {
      console.error("Report generation failed:", error);
      showToast(error?.message || "Could not generate the report");
    }
  }

  function updateReportDom() {
    const info = shiftInfo();
    const eod = isEndOfDayReport();
    $("rDate").textContent = formatDate($("reportDate").value);
    $("rShiftWindow").textContent = info.windowText;
    $("rPrepared").textContent = textValue("preparedBy");
    $("rReportType").textContent = reportTypeLabel();
    $("rCoverage").textContent = coverageLabel();
    $("reportKicker").textContent = eod ? "ASSET COMPLIANCE • END OF DAY" : "ASSET COMPLIANCE • END OF SHIFT";
    $("reportTitle").textContent = eod ? "End-of-Day Asset Management & Compliance Report" : "End-of-Shift Asset Management & Compliance Report";
    $("rInventoryHeading").textContent = eod ? "Daily Asset Inventory" : "Shift Inventory Snapshot";
    $("rInventorySubtitle").textContent = eod ? "Full 24-hour movement and closing stock position across Day + Night shifts" : "Movement and stock position by asset type";
    $("rClientHeading").textContent = eod ? "Client Movements • Day + Night" : "Client Movement";
    $("rClientSubtitle").textContent = eod ? "Sent, received and balances across the full Day + Night reporting period" : "Sent, received, today's net and all-time assets held with each client";
    $("rScheduledSubtitle").textContent = eod ? "Routine expectations versus actual completion across Day + Night shifts" : "Routine expectations versus actual completion";
    $("rSignoffLabel").textContent = eod ? "End-of-Day Sign-Off" : "End-of-Shift Sign-Off";
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

  async function createReportPdf() {
    const button = $("generateBtn");
    button.disabled = true;
    button.textContent = "Generating PDF…";
    try {
      const jsPdfCtor = window.jspdf?.jsPDF || window.jsPDF;
      if (typeof jsPdfCtor !== "function") {
        throw new Error("PDF tool is still loading. Please try again.");
      }

      // PDF pages are drawn directly with jsPDF instead of rasterising the
      // large hidden report DOM. This keeps generation fast and memory-light.
      const pdf = new jsPdfCtor({ orientation: "portrait", unit: "pt", format: "a4", compress: true });
      const clients = combinedClientData(shiftRows()).map((entry, index) => ({
        client: entry.client,
        key: String(index) + "::" + hashString(entry.client)
      }));

      drawPdfInventoryPage(pdf);
      const clientPageLinks = drawPdfClientMovementPages(pdf, clients);
      pdf.addPage("a4", "portrait");
      drawPdfDetailPage(pdf);

      // Ledger pages start after the three opening report sections. Keep the
      // first link target for each client for the card on Client Movements.
      const pageByClientKey = new Map();
      for (const client of clients) {
        const clientRows = getClientLedgerRows(client.client);
        await prepareLedgerPhotos(clientRows);
        clientRows.forEach(row => { row._prepared = true; });
        const ledgerPages = await drawClientLedgerPages(pdf, client.client, clientRows);
        pageByClientKey.set(client.key, ledgerPages[0]);
      }

      // Internal links are added after all ledger page numbers are known.
      clientPageLinks.forEach(item => {
        const targetPage = pageByClientKey.get(item.key);
        if (!targetPage) return;
        pdf.setPage(item.sourcePage);
        pdf.link(item.x, item.y, item.w, item.h, { pageNumber: targetPage, top: 0, left: 0 });
      });

      // Link back from every ledger page to the Client Movements page rather
      // than the inventory page, which is the page containing the client cards.
      pdf.getNumberOfPages(); // forces the page count to be final before output
      // drawClientLedgerPage already adds a back link; use page 2 as the target.
      // The native ledger renderer below is deliberately kept independent of
      // the report DOM, so there is no extra browser rendering cost here.

      // Finalise and validate the PDF before touching the lightweight preview.
      // The previous implementation could fail while converting the preview
      // data URL and prevent the PDF from ever reaching the download step.
      if (state.reportPdfBlobUrl) URL.revokeObjectURL(state.reportPdfBlobUrl);
      const pdfBlob = pdf.output("blob");
      if (!(pdfBlob instanceof Blob) || pdfBlob.size < 1000) {
        throw new Error("The generated PDF is empty or invalid.");
      }
      state.reportPdfBlobUrl = URL.createObjectURL(pdfBlob);
      const filename = reportFileName();
      $("downloadPdfBtn").onclick = () => downloadBlob(pdfBlob, filename, "application/pdf");

      // Keep the preview path completely independent from PDF creation. A
      // preview failure must never stop a valid PDF from being downloaded.
      try {
        const previewSvg = buildPdfInventoryPreviewSvg();
        const previewBlob = new Blob([previewSvg], { type: "image/svg+xml;charset=utf-8" });
        if (state.reportBlobUrl) URL.revokeObjectURL(state.reportBlobUrl);
        state.reportBlobUrl = URL.createObjectURL(previewBlob);
        $("reportPreview").src = state.reportBlobUrl;
        $("reportPreviewWrap").classList.remove("hidden");
      } catch (previewError) {
        console.warn("PDF preview could not be rendered; PDF download remains available.", previewError);
        $("reportPreviewWrap").classList.add("hidden");
      }

      // Keep the automatic download for desktop browsers, but never make it a
      // prerequisite for generation. The visible "Download PDF again" button
      // remains wired to the same validated Blob for browsers that block
      // downloads after an asynchronous operation.
      try {
        downloadBlob(pdfBlob, filename, "application/pdf");
      } catch (downloadError) {
        console.warn("Automatic PDF download was blocked; use Download PDF again.", downloadError);
      }
      renderSlackSummary();
      $("outputCard").classList.remove("hidden");
      showToast(`PDF report generated • ${clients.length} client ledger${clients.length === 1 ? "" : "s"}`);
    } finally {
      button.disabled = false;
      button.textContent = "Generate report + Slack summary";
    }
  }

  function pdfSafeText(value) {
    return String(value ?? "").replaceAll("•", "|").replaceAll("–", "-").replaceAll("—", "-").replaceAll("→", "->").replaceAll("←", "<-");
  }

  function drawPdfPageChrome(pdf, pageTitle, subtitle, pageNumberLabel = "") {
    const pageSize = pdf.internal.pageSize;
    const w = pageSize.getWidth();
    const h = pageSize.getHeight();
    // jsPDF 2.5.x exposes width/height here, but does not expose a
    // pageSize.getOrientation() function. Derive orientation from the
    // dimensions so PDF generation works reliably across jsPDF builds.
    const margin = w > h ? 24 : 30;
    let y = 24;
    pdf.setTextColor(118, 131, 141);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(7);
    pdf.text("ASSET COMPLIANCE", margin, y);
    if (pageNumberLabel) {
      pdf.setFont("helvetica", "normal");
      pdf.text(pdfSafeText(pageNumberLabel), w - margin, y, { align: "right" });
    }

    y += 19;
    pdf.setTextColor(23, 32, 42);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(16.5);
    const titleLines = pdf.splitTextToSize(pdfSafeText(pageTitle), w - margin * 2 - 8).slice(0, 2);
    pdf.text(titleLines, margin, y);
    y += titleLines.length * 17;

    if (subtitle) {
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(8.4);
      pdf.setTextColor(92, 105, 115);
      const subtitleLines = pdf.splitTextToSize(pdfSafeText(subtitle), w - margin * 2 - 8).slice(0, 2);
      pdf.text(subtitleLines, margin, y + 1);
      y += subtitleLines.length * 10.5;
    }

    pdf.setDrawColor(23, 32, 42);
    pdf.setLineWidth(1.25);
    pdf.line(margin, y + 5, w - margin, y + 5);
    return y + 20;
  }

  function drawPdfStatCard(pdf, x, y, w, h, label, value, accent = [23, 32, 42]) {
    pdf.setFillColor(247, 249, 250);
    pdf.setDrawColor(218, 225, 230);
    pdf.roundedRect(x, y, w, h, 5, 5, "FD");
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(Math.min(7, Math.max(5.8, h * 0.17)));
    pdf.setTextColor(102, 115, 125);
    pdf.text(String(label).toUpperCase(), x + 9, y + 11);
    pdf.setFontSize(Math.min(12.5, Math.max(9, h * 0.38)));
    pdf.setTextColor(...accent);
    pdf.text(String(value), x + 9, y + h - 9);
  }

  function drawPdfSectionHeading(pdf, title, subtitle, x, y, w) {
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(11.5);
    pdf.setTextColor(23, 32, 42);
    pdf.text(pdfSafeText(title).toUpperCase(), x, y);
    pdf.setDrawColor(23, 32, 42);
    pdf.setLineWidth(0.9);
    pdf.line(x, y + 4, x + Math.min(w, Math.max(85, pdf.getTextWidth(title) + 22)), y + 4);
    if (subtitle) {
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(6.7);
      pdf.setTextColor(116, 128, 138);
      pdf.text(pdfSafeText(subtitle), x, y + 14);
      return y + 28;
    }
    return y + 15;
  }

  function drawPdfInventoryPage(pdf) {
    const info = shiftInfo();
    const w = pdf.internal.pageSize.getWidth();
    const h = pdf.internal.pageSize.getHeight();
    const margin = 30;
    let y = drawPdfPageChrome(
      pdf,
      isEndOfDayReport() ? "END-OF-DAY ASSET MANAGEMENT REPORT" : "END-OF-SHIFT ASSET MANAGEMENT REPORT",
      `${formatDate($("reportDate").value)}  |  ${info.windowText}`,
      "Page 1 - Inventory snapshot"
    );

    const metaW = w - margin * 2;
    pdf.setFillColor(244, 247, 249);
    pdf.setDrawColor(218, 225, 230);
    pdf.roundedRect(margin, y, metaW, 38, 6, 6, "FD");
    const meta = [
      ["REPORT TYPE", reportTypeLabel()],
      ["COVERAGE", coverageLabel()],
      ["PREPARED BY", textValue("preparedBy")],
      ["STATUS", $("overallStatus").value]
    ];
    meta.forEach((item, i) => {
      const cellW = (metaW - 24) / 4;
      const x = margin + 12 + i * cellW;
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(5.8);
      pdf.setTextColor(111, 123, 132);
      pdf.text(item[0], x, y + 11);
      pdf.setFontSize(7.8);
      pdf.setTextColor(23, 32, 42);
      pdf.text(pdf.splitTextToSize(item[1] || "-", cellW - 12).slice(0, 1), x, y + 25);
    });
    y += 49;
    // Give the Inventory Snapshot heading a little more breathing room below
    // the report metadata block.
    y += 9;

    y = drawPdfSectionHeading(pdf, "Inventory Snapshot", "Opening, received, sent and closing position by asset type", margin, y, metaW);
    const data = inventorySnapshotData();
    if (!data.length) {
      pdf.setFillColor(250, 252, 253);
      pdf.setDrawColor(210, 219, 225);
      pdf.roundedRect(margin, y, metaW, 42, 6, 6, "FD");
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(8);
      pdf.setTextColor(107, 120, 129);
      pdf.text("No inventory or movement data was found for this report.", margin + 12, y + 25);
      return;
    }

    const gap = 8;
    const cardW = (metaW - gap * 2) / 3;
    // Give the card header and the lower metric row more breathing room.
    const cardH = 92;
    const cols = 3;
    data.forEach((item, index) => {
      const row = Math.floor(index / cols);
      const col = index % cols;
      const x = margin + col * (cardW + gap);
      const cy = y + row * (cardH + gap);
      const idx = state.assets.findIndex(a => sameName(a, item.asset));
      const opening = Number($(`open_${idx}`)?.value || 0);
      const closing = Number($(`close_${idx}`)?.value || 0);
      const net = item.received - item.sent;

      pdf.setFillColor(255, 255, 255);
      pdf.setDrawColor(210, 218, 224);
      pdf.roundedRect(x, cy, cardW, cardH, 6, 6, "FD");
      pdf.setFillColor(...(net >= 0 ? [232, 245, 236] : [249, 235, 235]));
      pdf.roundedRect(x + cardW - 43, cy + 7, 35, 16, 4, 4, "F");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(7.4);
      pdf.setTextColor(23, 32, 42);
      pdf.text(pdf.splitTextToSize(item.asset, cardW - 54).slice(0, 1), x + 9, cy + 15);
      pdf.setFontSize(7);
      pdf.setTextColor(...(net >= 0 ? [38, 118, 71] : [161, 62, 62]));
      pdf.text(`${net >= 0 ? "+" : ""}${formatQty(net)}`, x + cardW - 25.5, cy + 19, { align: "center" });

      const cells = [["Opening", opening], ["Received", item.received], ["Sent", item.sent], ["Closing", closing]];
      const cellW = (cardW - 18 - 6) / 2;
      const cellH = 20;
      const cellTop = cy + 40;
      const cellGap = 5;
      cells.forEach((cell, ci) => {
        const cx = x + 9 + (ci % 2) * (cellW + 6);
        const cY = cellTop + Math.floor(ci / 2) * (cellH + cellGap);
        pdf.setFillColor(248, 250, 251);
        pdf.setDrawColor(231, 235, 238);
        pdf.roundedRect(cx, cY, cellW, cellH, 3, 3, "FD");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(5.2);
        pdf.setTextColor(112, 124, 133);
        pdf.text(cell[0].toUpperCase(), cx + 4, cY + 7);
        pdf.setFontSize(8);
        pdf.setTextColor(23, 32, 42);
        pdf.text(formatQty(cell[1]), cx + cellW - 4, cY + 15, { align: "right" });
      });
    });

    const rows = shiftRows();
    const lastCardBottom = y + Math.ceil(data.length / cols) * (cardH + gap) - gap;
    const statsY = Math.min(lastCardBottom + 8, h - 56);
    const statsW = (metaW - 8 * 3) / 4;
    const stats = [
      ["Transactions", rows.length],
      ["Clients moved", combinedClientData(rows).length],
      ["Asset types", data.length],
      ["Scheduled", getScheduleRows().length]
    ];
    stats.forEach((stat, i) => drawPdfStatCard(pdf, margin + i * (statsW + 8), statsY, statsW, 34, stat[0], stat[1]));

    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(6.5);
    pdf.setTextColor(115, 127, 136);
    pdf.text(`Sign-off: ${formatDateTime($("signoffDate").value)}  |  EOD status: ${$("eodStatus").value}`, margin, h - 14);
  }

  function drawPdfClientMovementPages(pdf, clients) {
    const links = [];
    const info = shiftInfo();
    const data = combinedClientData(shiftRows());
    const w = pdf.internal.pageSize.getWidth();
    const h = pdf.internal.pageSize.getHeight();
    const margin = 30;
    const gap = 10;
    const cardW = (w - margin * 2 - gap) / 2;
    let pageIndex = 2;
    let y = 0;

    pdf.addPage("a4", "portrait");
    y = drawPdfPageChrome(pdf, "CLIENT MOVEMENTS", `${formatDate($("reportDate").value)}  |  ${info.windowText}`, "Page 2 - Client movements");
    const statsW = (w - margin * 2 - 8 * 2) / 3;
    const movementRows = shiftRows();
    [
      ["Clients", clients.length],
      ["Transactions", movementRows.length],
      ["Asset types", new Set(movementRows.map(r => normalize(r.asset)).filter(Boolean)).size]
    ].forEach((s, i) => drawPdfStatCard(pdf, margin + i * (statsW + 8), y, statsW, 31, s[0], s[1]));
    y += 40;

    const clientTxCounts = new Map();
    movementRows.forEach(r => {
      if (r.client) clientTxCounts.set(r.client, (clientTxCounts.get(r.client) || 0) + 1);
    });

    let slot = 0;
    let rowBottom = y;
    const linksForPage = [];
    const drawClientCard = (client, x, cy, cardData) => {
      const headerH = 36;
      const rowH = 17;
      const cardH = 50 + Math.max(18, cardData.types.length * rowH + 12);
      pdf.setFillColor(255, 255, 255);
      pdf.setDrawColor(210, 218, 224);
      pdf.roundedRect(x, cy, cardW, cardH, 6, 6, "FD");
      pdf.setFillColor(23, 32, 42);
      pdf.roundedRect(x, cy, cardW, headerH, 6, 6, "F");
      pdf.rect(x, cy + headerH - 6, cardW, 6, "F");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(8.2);
      pdf.setTextColor(255, 255, 255);
      pdf.text(pdf.splitTextToSize(client.client, cardW - 122).slice(0, 1), x + 9, cy + 13);
      pdf.setFontSize(6.2);
      pdf.setTextColor(205, 215, 221);
      pdf.text("SEE TRANSACTIONS ->", x + cardW - 9, cy + 13, { align: "right" });
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(5.7);
      pdf.text(`${cardData.types.length} asset type${cardData.types.length === 1 ? "" : "s"}`, x + 9, cy + 25);
      pdf.text(`${clientTxCounts.get(client.client) || 0} transaction${(clientTxCounts.get(client.client) || 0) === 1 ? "" : "s"}`, x + cardW - 9, cy + 25, { align: "right" });

      const cols = ["Asset", "IN", "OUT", "DATE", "ALL-TIME"];
      const colX = [x + 9, x + cardW - 152, x + cardW - 116, x + cardW - 77, x + cardW - 9];
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(5.6);
      pdf.setTextColor(111, 123, 132);
      cols.forEach((label, i) => pdf.text(label, colX[i], cy + headerH + 11, { align: i === 0 ? "left" : "right" }));
      cardData.types.forEach((t, i) => {
        const ry = cy + headerH + 22 + i * rowH;
        if (i % 2 === 0) {
          pdf.setFillColor(248, 250, 251);
          pdf.rect(x + 6, ry - 8, cardW - 12, rowH, "F");
        }
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(6.5);
        pdf.setTextColor(35, 45, 53);
        pdf.text(pdf.splitTextToSize(t.asset, 86).slice(0, 1)[0], colX[0], ry);
        pdf.text(formatQty(t.received), colX[1], ry, { align: "right" });
        pdf.text(formatQty(t.sent), colX[2], ry, { align: "right" });
        pdf.setTextColor(t.withClient < 0 ? 161 : 38, t.withClient < 0 ? 62 : 118, t.withClient < 0 ? 62 : 71);
        pdf.text(signedQty(t.withClient), colX[3], ry, { align: "right" });
        pdf.setTextColor(t.allTimeWithClient < 0 ? 161 : 38, t.allTimeWithClient < 0 ? 62 : 118, t.allTimeWithClient < 0 ? 62 : 71);
        pdf.text(signedQty(t.allTimeWithClient), colX[4], ry, { align: "right" });
      });
      links.push({ key: client.key, sourcePage: pageIndex, x: x + cardW - 112, y: cy + 4, w: 104, h: 18 });
      linksForPage.push(links[links.length - 1]);
      return cardH;
    };

    for (const client of clients) {
      const cardData = data.find(item => sameName(item.client, client.client));
      if (!cardData) continue;
      const estimatedH = 50 + Math.max(18, cardData.types.length * 17 + 12);
      if (slot === 0 && y + estimatedH > h - 54) {
        pageIndex += 1;
        pdf.addPage("a4", "portrait");
        y = drawPdfPageChrome(pdf, "CLIENT MOVEMENTS - CONTINUED", `${formatDate($("reportDate").value)}  |  ${info.windowText}`, `Page ${pageIndex} - Client movements`);
        slot = 0;
        rowBottom = y;
      }
      if (slot === 1 && y + estimatedH > h - 54) {
        pageIndex += 1;
        pdf.addPage("a4", "portrait");
        y = drawPdfPageChrome(pdf, "CLIENT MOVEMENTS - CONTINUED", `${formatDate($("reportDate").value)}  |  ${info.windowText}`, `Page ${pageIndex} - Client movements`);
        slot = 0;
        rowBottom = y;
      }
      const x = margin + slot * (cardW + gap);
      const rowY = y;
      const actualH = drawClientCard(client, x, rowY, cardData);
      rowBottom = Math.max(rowBottom, rowY + actualH);
      if (slot === 0) {
        slot = 1;
      } else {
        slot = 0;
        y = rowBottom + gap;
        rowBottom = y;
      }
    }

    return links;
  }

  function drawPdfDetailPage(pdf) {
    const w = pdf.internal.pageSize.getWidth();
    const h = pdf.internal.pageSize.getHeight();
    const margin = 30;
    const gap = 12;
    const contentW = w - margin * 2;
    let y = drawPdfPageChrome(
      pdf,
      "REPORT DETAIL",
      `${formatDate($("reportDate").value)}  |  Schedule, exceptions, risks and handover`,
      `Page ${pdf.getNumberOfPages()} - Report detail`
    );

    // Give the third page a deliberate editorial structure rather than
    // stacking small sections tightly. Each block gets its own visual card.
    y += 8;
    y = drawPdfSectionHeading(pdf, "Scheduled Movements", "Routine expectations versus actual completion", margin, y, contentW);
    const routineRows = getScheduleRows();
    const scheduleHeaders = ["Client", "Asset", "Planned", "Actual", "Status", "Notes"];
    const scheduleWidths = [88, 90, 54, 54, 76, contentW - (88 + 90 + 54 + 54 + 76)];
    y = drawPdfDataTable(
      pdf,
      scheduleHeaders,
      routineRows.length ? routineRows.map(item => [
        item.routine.client,
        item.routine.asset,
        formatQty(item.routine.quantity),
        item.actualQty === "" ? "-" : formatQty(item.actualQty),
        item.status,
        item.note || item.routine.notes || "-"
      ]) : [["-", "-", "-", "-", "No scheduled movements", "-"],],
      margin,
      y,
      scheduleWidths,
      { rowHeight: 26, headerHeight: 25, fontSize: 7 }
    );

    // Middle row: exceptions and the forward plan. The fixed card region
    // keeps these sections visually balanced even when one has little data.
    y += 18;
    const halfW = (contentW - gap) / 2;
    const leftX = margin;
    const rightX = margin + halfW + gap;
    const middleTop = y;
    const cardPadding = 10;
    const middleCardH = 128;

    pdf.setFillColor(247, 249, 250);
    pdf.setDrawColor(218, 225, 230);
    pdf.roundedRect(leftX, middleTop, halfW, middleCardH, 7, 7, "FD");
    pdf.roundedRect(rightX, middleTop, halfW, middleCardH, 7, 7, "FD");
    drawPdfSectionHeading(pdf, "Discarded Assets", "Asset type, quantity and reason", leftX + cardPadding, middleTop + 18, halfW - cardPadding * 2);
    drawPdfSectionHeading(pdf, "Plan for Tomorrow", "Priority actions and handover notes", rightX + cardPadding, middleTop + 18, halfW - cardPadding * 2);

    const discarded = assetTotals(shiftRows().filter(r => ["DISCARDED", "DISCARD"].includes(r.movement)));
    const discardRows = discarded.length
      ? discarded.map(item => [
          item.asset,
          formatQty(item.quantity),
          discardedReasons[discardedReasonKey($("reportDate").value, item.asset)] || "Reason not entered"
        ])
      : [["-", "-", "No assets were discarded"]];
    drawPdfDataTable(
      pdf,
      ["Asset", "Qty", "Reason"],
      discardRows,
      leftX + cardPadding,
      middleTop + 46,
      [Math.min(112, halfW * .32), 42, halfW - cardPadding * 2 - Math.min(112, halfW * .32) - 42],
      { rowHeight: 24, headerHeight: 23, fontSize: 6.8 }
    );

    const planText = textValue("tomorrowPlan", "No plan entered.");
    const planLines = pdf.splitTextToSize(planText, halfW - cardPadding * 2).slice(0, 6);
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(7.2);
    pdf.setTextColor(43, 54, 62);
    pdf.text(planLines.length ? planLines : ["No plan entered."], rightX + cardPadding, middleTop + 62);

    // Handover notes make better use of remaining page area and keep the
    // report useful even when there are no discarded assets or issues.
    y = middleTop + middleCardH + 18;
    const noteW = (contentW - gap) / 2;
    const noteH = 84;
    [
      ["Received notes", textValue("receivedNotes", "No received notes entered.")],
      ["Sent notes", textValue("sentNotes", "No sent notes entered.")]
    ].forEach((item, i) => {
      const x = margin + i * (noteW + gap);
      pdf.setFillColor(247, 249, 250);
      pdf.setDrawColor(218, 225, 230);
      pdf.roundedRect(x, y, noteW, noteH, 7, 7, "FD");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(7.5);
      pdf.setTextColor(23, 32, 42);
      pdf.text(item[0].toUpperCase(), x + 10, y + 16);
      pdf.setDrawColor(188, 198, 205);
      pdf.setLineWidth(0.55);
      pdf.line(x + 10, y + 21, x + 78, y + 21);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(6.9);
      pdf.setTextColor(75, 88, 98);
      const lines = pdf.splitTextToSize(item[1], noteW - 20).slice(0, 5);
      pdf.text(lines.length ? lines : ["-"], x + 10, y + 34);
    });

    y += noteH + 18;
    y = drawPdfSectionHeading(pdf, "Outstanding Issues / Risks", "Open exceptions and agreed actions", margin, y, contentW);
    const issues = collectIssueRows();
    const issueRows = issues.length ? issues.map(row => row.map(value => value || "-")) : [["-", "No entries recorded", "-", "-", "-", "-", "-"]];
    const issueWidths = [42, 82, 64, 82, 94, 52, contentW - 416];
    y = drawPdfDataTable(
      pdf,
      ["Priority", "Issue", "Asset / Client", "Impact", "Action", "Due Date", "Status"],
      issueRows,
      margin,
      y,
      issueWidths,
      { rowHeight: 24, headerHeight: 25, fontSize: 6.3 }
    );

    // Sign-off sits in its own band instead of being forced against the last
    // table. This makes the page feel complete while still respecting the
    // bottom safe area.
    y += 18;
    const footerH = 50;
    // Anchor sign-off toward the lower safe area when there is spare room,
    // while still moving it down if unusually long content needs the space.
    const footerY = Math.min(Math.max(y, h - footerH - 42), h - footerH - 12);
    pdf.setFillColor(23, 32, 42);
    pdf.roundedRect(margin, footerY, contentW, footerH, 7, 7, "F");
    const signoff = formatDateTime($("signoffDate").value);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(6.4);
    pdf.setTextColor(185, 198, 208);
    pdf.text(isEndOfDayReport() ? "END-OF-DAY SIGN-OFF" : "END-OF-SHIFT SIGN-OFF", margin + 10, footerY + 14);
    pdf.setFontSize(8.4);
    pdf.setTextColor(255, 255, 255);
    pdf.text(signoff, margin + 10, footerY + 31);
    pdf.text(`Overall EOD status: ${pdfSafeText($("eodStatus").value)}`, margin + contentW - 10, footerY + 23, { align: "right" });
  }

  function drawPdfDataTable(pdf, headers, rows, x, y, widths, options = {}) {
    const baseRowHeight = options.rowHeight || 21;
    const headerHeight = options.headerHeight || 22;
    const fontSize = options.fontSize || 6.6;
    let cursorY = y;
    let cursorX = x;

    // Draw each header independently and explicitly reset the fill/text colour.
    // This avoids the intermittent white-on-white header issue seen in PDF viewers.
    headers.forEach((header, i) => {
      pdf.setFillColor(23, 32, 42);
      pdf.setDrawColor(23, 32, 42);
      pdf.rect(cursorX, cursorY, widths[i], headerHeight, "FD");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(Math.min(6.9, Math.max(5.8, fontSize + 0.2)));
      pdf.setTextColor(255, 255, 255);
      const label = String(header ?? "").toUpperCase();
      const lines = pdf.splitTextToSize(label, Math.max(10, widths[i] - 8)).slice(0, 2);
      const lineH = 7.2;
      const startY = cursorY + Math.max(6, (headerHeight - lines.length * lineH) / 2 + 5.2);
      lines.forEach((line, lineIndex) => {
        pdf.text(line, cursorX + widths[i] / 2, startY + lineIndex * lineH, { align: "center" });
      });
      cursorX += widths[i];
    });
    cursorY += headerHeight;

    rows.forEach((row, ri) => {
      const lineH = fontSize + 1.35;
      const neededLines = Math.max(1, ...row.map((value, i) => pdf.splitTextToSize(String(value ?? ""), Math.max(14, widths[i] - 8)).slice(0, 3).length));
      const rowHeight = Math.max(baseRowHeight, neededLines * lineH + 7);
      cursorX = x;
      row.forEach((value, i) => {
        pdf.setFillColor(ri % 2 ? 250 : 255, ri % 2 ? 252 : 255, ri % 2 ? 253 : 255);
        pdf.setDrawColor(203, 210, 216);
        pdf.rect(cursorX, cursorY, widths[i], rowHeight, "FD");
        const align = [2,3].includes(i) ? "right" : "left";
        drawCellText(pdf, value, cursorX, cursorY, widths[i], rowHeight, { fontSize, align, color: [35,45,53] });
        cursorX += widths[i];
      });
      cursorY += rowHeight;
    });
    return cursorY;
  }

  function buildPdfInventoryPreviewSvg() {
    const width = 820;
    const height = 1160;
    const esc = value => String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    const info = shiftInfo();
    const data = inventorySnapshotData();
    const margin = 34;
    const gap = 12;
    const cardW = (width - margin * 2 - gap * 2) / 3;
    const cardH = 100;
    const cards = data.map((item, index) => {
      const idx = state.assets.findIndex(a => sameName(a, item.asset));
      const opening = Number($(`open_${idx}`)?.value || 0);
      const closing = Number($(`close_${idx}`)?.value || 0);
      const net = item.received - item.sent;
      const x = margin + (index % 3) * (cardW + gap);
      const y = 305 + Math.floor(index / 3) * (cardH + gap);
      return `<g><rect x="${x}" y="${y}" width="${cardW}" height="${cardH}" rx="10" fill="#ffffff" stroke="#d2dae0"/><text x="${x+12}" y="${y+21}" class="name">${esc(item.asset)}</text><rect x="${x+cardW-58}" y="${y+9}" width="46" height="22" rx="6" fill="${net>=0?'#e8f5ec':'#f9ebeb'}"/><text x="${x+cardW-35}" y="${y+24}" text-anchor="middle" class="diff ${net<0?'neg':'pos'}">${net>=0?'+':''}${esc(formatQty(net))}</text><g class="label"><text x="${x+12}" y="${y+51}">OPENING</text><text x="${x+cardW/2+2}" y="${y+51}">RECEIVED</text><text x="${x+12}" y="${y+82}">SENT</text><text x="${x+cardW/2+2}" y="${y+82}">CLOSING</text></g><g class="value"><text x="${x+cardW/2-5}" y="${y+64}" text-anchor="end">${esc(formatQty(opening))}</text><text x="${x+cardW-12}" y="${y+64}" text-anchor="end">${esc(formatQty(item.received))}</text><text x="${x+cardW/2-5}" y="${y+95}" text-anchor="end">${esc(formatQty(item.sent))}</text><text x="${x+cardW-12}" y="${y+95}" text-anchor="end">${esc(formatQty(closing))}</text></g></g>`;
    }).join("");
    const rows = shiftRows();
    const stats = [rows.length, combinedClientData(rows).length, data.length, getScheduleRows().length];
    return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#ffffff"/><style>.k{font:700 14px Arial;fill:#6f7b84;letter-spacing:2px}.t{font:700 30px Arial;fill:#17202a}.s{font:14px Arial;fill:#5f6d77}.name{font:700 13px Arial;fill:#17202a}.label{font:700 8px Arial;fill:#76838d}.value{font:700 13px Arial;fill:#17202a}.diff{font:700 10px Arial}.pos{fill:#267647}.neg{fill:#a13e3e}</style><text x="${margin}" y="36" class="k">ASSET COMPLIANCE</text><text x="${margin}" y="75" class="t">${esc(isEndOfDayReport()?'END-OF-DAY ASSET MANAGEMENT REPORT':'END-OF-SHIFT ASSET MANAGEMENT REPORT')}</text><text x="${margin}" y="101" class="s">${esc(formatDate($("reportDate").value))}  |  ${esc(info.windowText)}</text><line x1="${margin}" y1="116" x2="${width-margin}" y2="116" stroke="#17202a" stroke-width="2"/><rect x="${margin}" y="137" width="${width-margin*2}" height="48" rx="9" fill="#f4f7f9" stroke="#dae1e6"/><text x="${margin+14}" y="157" class="label">REPORT TYPE</text><text x="${margin+14}" y="175" class="s">${esc(reportTypeLabel())}</text><text x="${margin+230}" y="157" class="label">COVERAGE</text><text x="${margin+230}" y="175" class="s">${esc(coverageLabel())}</text><text x="${margin+445}" y="157" class="label">PREPARED BY</text><text x="${margin+445}" y="175" class="s">${esc(textValue('preparedBy'))}</text><text x="${margin}" y="225" class="t" style="font-size:20px">INVENTORY SNAPSHOT</text><text x="${margin}" y="247" class="s">Opening, received, sent and closing position by asset type</text>${cards}<g>${stats.map((value,i)=>{const x=margin+i*((width-margin*2-24)/4+8);const w=(width-margin*2-24)/4;return `<rect x="${x}" y="${height-100}" width="${w}" height="58" rx="9" fill="#f7f9fa" stroke="#dae1e6"/><text x="${x+12}" y="${height-81}" class="label">${['TRANSACTIONS','CLIENTS MOVED','ASSET TYPES','SCHEDULED'][i]}</text><text x="${x+12}" y="${height-57}" class="t" style="font-size:19px">${value}</text>`}).join('')}</g></svg>`)}`;
  }

  async function drawClientLedgerPages(pdf, clientName, preparedRows = null) {
    const rows = preparedRows || getClientLedgerRows(clientName);
    const chunks = paginateLedgerRows(pdf, rows);
    const pages = [];
    for (let i = 0; i < chunks.length; i++) {
      const pageNumber = pdf.getNumberOfPages() + 1;
      pdf.addPage("a4", "landscape");
      pages.push(pageNumber);
      drawClientLedgerPage(pdf, clientName, rows, chunks[i], i + 1, chunks.length);
    }
    return pages;
  }

  function getClientLedgerRows(clientName) {
    const selected = shiftRows()
      .filter(r => r.client && sameName(r.client, clientName) && ["SENT", "RECEIVED"].includes(r.movement) && r.asset && r.quantity > 0)
      .sort((a, b) => compareTimestamp(a.timestamp, b.timestamp) || (Number(a.rowNumber) || 0) - (Number(b.rowNumber) || 0));
    if (!selected.length) return [];

    const dailyTotals = new Map();
    state.transactions.forEach(r => {
      if (!r.client || !sameName(r.client, clientName) || !r.asset || !["SENT", "RECEIVED"].includes(r.movement)) return;
      const date = localDateKey(r.timestamp);
      if (!date) return;
      const key = `${date}::${normalize(r.asset)}`;
      dailyTotals.set(key, (dailyTotals.get(key) || 0) + (r.movement === "SENT" ? r.quantity : -r.quantity));
    });

    const running = new Map();
    const allTransactions = state.transactions
      .filter(r => r.client && sameName(r.client, clientName) && r.asset && ["SENT", "RECEIVED"].includes(r.movement))
      .sort((a, b) => compareTimestamp(a.timestamp, b.timestamp) || (Number(a.rowNumber) || 0) - (Number(b.rowNumber) || 0));
    const runningByRow = new Map();
    allTransactions.forEach(r => {
      const key = normalize(r.asset);
      const current = running.get(key) || 0;
      const next = current + (r.movement === "SENT" ? r.quantity : -r.quantity);
      running.set(key, next);
      runningByRow.set(transactionKey(r), next);
    });

    return selected.map(r => {
      const date = localDateKey(r.timestamp);
      const dateDifference = dailyTotals.get(`${date}::${normalize(r.asset)}`) || 0;
      return {
        ...r,
        date,
        time: formatTimeWithSeconds(r.timestamp),
        sent: r.movement === "SENT" ? r.quantity : 0,
        received: r.movement === "RECEIVED" ? r.quantity : 0,
        differenceForDate: dateDifference,
        allTimeDifference: runningByRow.get(transactionKey(r)) ?? 0
      };
    });
  }

  function formatTimeWithSeconds(value) {
    const p = londonParts(value);
    return p ? `${String(p.hour).padStart(2, "0")}:${String(p.minute).padStart(2, "0")}:${String(p.second).padStart(2, "0")}` : String(value || "—");
  }

  function ledgerRowHeight(row) {
    const commentLength = String(row?.comment || "").length;
    if (commentLength > 90) return 50;
    if (row?.image) return 48;
    return 42;
  }

  function paginateLedgerRows(pdf, rows) {
    const pageHeight = 595.28;
    const top = 24;
    const titleBlock = 55;
    const summaryBlock = 34;
    const tableHeaderHeight = 32;
    const totalsBlock = 38;
    const footerSpace = 18;
    const usable = pageHeight - top - titleBlock - summaryBlock - tableHeaderHeight - totalsBlock - footerSpace - 12;

    const groups = [...groupLedgerRowsByTimestamp(rows).entries()].map(([key, groupRows]) => ({
      key,
      rows: groupRows,
      height: groupRows.reduce((sum, row) => sum + ledgerRowHeight(row), 0)
    }));
    const chunks = [];
    let chunk = [];
    let used = 0;
    groups.forEach(group => {
      if (chunk.length && used + group.height > usable) {
        chunks.push(chunk);
        chunk = [];
        used = 0;
      }
      if (group.height <= usable) {
        group.rows.forEach(row => chunk.push({ row, h: ledgerRowHeight(row) }));
        used += group.height;
        return;
      }
      // Only split a delivery when one timestamped group is larger than an
      // entire page. This preserves merged timestamp fields in normal cases.
      group.rows.forEach(row => {
        const h = ledgerRowHeight(row);
        if (chunk.length && used + h > usable) {
          chunks.push(chunk);
          chunk = [];
          used = 0;
        }
        chunk.push({ row, h });
        used += h;
      });
    });
    if (chunk.length) chunks.push(chunk);
    return chunks.length ? chunks : [[]];
  }

  function drawClientLedgerPage(pdf, clientName, allRows, chunk, chunkIndex, chunkCount) {
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = 24;
    const cols = [48, 52, 46, 103, 52, 64, 68, 74, 108, 52, 76];
    const headers = ["Client", "Date", "Time", "Asset Type", "Qty Sent", "Qty Received", "Date Diff", "All-Time Diff", "Comment", "Photo", "Link"];
    const tableW = cols.reduce((a, b) => a + b, 0);
    const tableX = Math.max(margin, (pageW - tableW) / 2);
    let y = 22;

    // Compact, balanced title block.
    pdf.setTextColor(23, 32, 42);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(15.5);
    pdf.text("CLIENT TRANSACTION LEDGER", tableX, y + 1);
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(8.8);
    pdf.setTextColor(106, 118, 128);
    pdf.text(pdfSafeText(`Client: ${clientName}`), tableX, y + 17);
    const info = shiftInfo();
    pdf.text(pdfSafeText(`Report: ${formatDate($("reportDate").value)} | ${info.windowText}`), tableX, y + 29);
    if (chunkCount > 1) {
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(7.2);
      pdf.setTextColor(106, 118, 128);
      pdf.text(`Ledger page ${chunkIndex} of ${chunkCount}`, tableX, y + 41);
    }

    const backText = "< Back to Client Movements";
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(7.4);
    pdf.setTextColor(45, 95, 139);
    const backW = pdf.getTextWidth(backText) + 8;
    const backX = pageW - tableX - backW;
    pdf.textWithLink(backText, backX, y + 1, { pageNumber: 2, top: 0, left: 0 });
    y += 49;

    const ledgerRows = allRows;
    const sentLines = ledgerRows.filter(r => r.sent > 0).length;
    const receivedLines = ledgerRows.filter(r => r.received > 0).length;
    const assetTypes = new Set(ledgerRows.map(r => normalize(r.asset))).size;
    const statGap = 7;
    const statW = (tableW - statGap * 3) / 4;
    [["Transactions", ledgerRows.length], ["Asset Types", assetTypes], ["Sent Lines", sentLines], ["Received Lines", receivedLines]]
      .forEach((item, i) => drawPdfStatCard(pdf, tableX + i * (statW + statGap), y, statW, 28, item[0], item[1]));
    y += 37;

    const headerHeight = 32;
    let x = tableX;
    headers.forEach((header, i) => {
      // Reset BOTH fill and text colour for every cell. Older jsPDF builds can
      // retain a previous cell's graphics state, producing white-on-white headers.
      pdf.setFillColor(23, 32, 42);
      pdf.setDrawColor(23, 32, 42);
      pdf.rect(x, y, cols[i], headerHeight, "FD");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(6.15);
      pdf.setTextColor(255, 255, 255);
      const lines = pdf.splitTextToSize(header.toUpperCase(), cols[i] - 7).slice(0, 2);
      const lineH = 6.7;
      const startY = y + (headerHeight - lines.length * lineH) / 2 + 5.2;
      lines.forEach((line, idx) => {
        pdf.text(line, x + cols[i] / 2, startY + idx * lineH, { align: "center" });
      });
      x += cols[i];
    });

    const pageGroups = [];
    chunk.forEach(item => {
      const groupKey = ledgerTransactionGroupKey(item.row);
      let group = pageGroups.find(g => g.key === groupKey);
      if (!group) { group = { key: groupKey, rows: [] }; pageGroups.push(group); }
      group.rows.push(item);
    });

    let rowY = y + headerHeight;
    const groupLookup = new Map();
    pageGroups.forEach(group => {
      const groupHeight = group.rows.reduce((sum, item) => sum + item.h, 0);
      groupLookup.set(group.key, { top: rowY, height: groupHeight });
      rowY += groupHeight;
    });

    // Draw asset-level fields first.
    rowY = y + headerHeight;
    chunk.forEach((item, idx) => {
      const row = item.row;
      const h = item.h;
      x = tableX;
      const cells = [
        "", "", "", row.asset,
        row.sent ? formatQty(row.sent) : "",
        row.received ? formatQty(row.received) : "",
        signedQty(row.differenceForDate),
        signedQty(row.allTimeDifference),
        row.comment || "", "", ""
      ];
      for (let i = 0; i < cols.length; i++) {
        if ([0, 1, 2, 9, 10].includes(i)) { x += cols[i]; continue; }
        pdf.setFillColor(idx % 2 ? 250 : 255, idx % 2 ? 252 : 255, idx % 2 ? 253 : 255);
        pdf.setDrawColor(203, 210, 216);
        pdf.rect(x, rowY, cols[i], h, "FD");
        const align = [4, 5, 6, 7].includes(i) ? "right" : "left";
        const value = cells[i];
        const num = i === 6 ? row.differenceForDate : row.allTimeDifference;
        const color = [6, 7].includes(i) ? (Number(num) < 0 ? [161, 62, 62] : [38, 118, 71]) : [23, 32, 42];
        drawCellText(pdf, value, x, rowY, cols[i], h, { fontSize: 7.1, align, color });
        x += cols[i];
      }
      rowY += h;
    });

    // Merge client/date/time/photo/link cells over each exact timestamp group.
    rowY = y + headerHeight;
    pageGroups.forEach(group => {
      const first = group.rows[0].row;
      const h = groupLookup.get(group.key).height;
      const mergedFields = [
        { index: 0, text: clientName, bold: true, align: "left" },
        { index: 1, text: formatDate(first.date), bold: false, align: "center" },
        { index: 2, text: first.time, bold: false, align: "center" }
      ];
      mergedFields.forEach(field => {
        const fx = tableX + cols.slice(0, field.index).reduce((a, b) => a + b, 0);
        pdf.setFillColor(255, 255, 255);
        pdf.setDrawColor(203, 210, 216);
        pdf.rect(fx, rowY, cols[field.index], h, "FD");
        drawCellText(pdf, field.text, fx, rowY, cols[field.index], h, { fontSize: 7.1, align: field.align, bold: field.bold, color: [23, 32, 42] });
      });

      const photoRow = group.rows.find(item => item.row._pdfPhoto || item.row.image);
      const photo = photoRow?.row?._pdfPhoto || null;
      const pX = tableX + cols.slice(0, 9).reduce((a, b) => a + b, 0);
      pdf.setFillColor(250, 252, 253);
      pdf.setDrawColor(203, 210, 216);
      pdf.rect(pX, rowY, cols[9], h, "FD");
      if (photo) {
        try {
          const ph = Math.min(h - 5, 42);
          const ratio = photo.width && photo.height ? photo.width / photo.height : 1;
          const pw = Math.min(cols[9] - 6, Math.max(12, ph * ratio));
          const actualH = Math.min(ph, pw / ratio);
          pdf.addImage(photo.dataUrl, "JPEG", pX + (cols[9] - pw) / 2, rowY + (h - actualH) / 2, pw, actualH, undefined, "FAST");
        } catch (_) {
          drawCellText(pdf, "-", pX, rowY, cols[9], h, { fontSize: 7.1, align: "center", color: [135, 147, 156] });
        }
      } else {
        drawCellText(pdf, "-", pX, rowY, cols[9], h, { fontSize: 7.1, align: "center", color: [135, 147, 156] });
      }

      const linkX = pX + cols[9];
      pdf.setFillColor(255, 255, 255);
      pdf.rect(linkX, rowY, cols[10], h, "FD");
      const photoUrl = photoRow?.row?.image || first.image || "";
      if (photoUrl) {
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(6.8);
        pdf.setTextColor(45, 95, 139);
        pdf.textWithLink("Open photo", linkX + 4, rowY + h / 2 + 2, { url: photoUrl });
      } else {
        drawCellText(pdf, "-", linkX, rowY, cols[10], h, { fontSize: 7.1, align: "center", color: [135, 147, 156] });
      }
      rowY += h;
    });

    const totals = allRows.reduce((acc, r) => {
      acc.sent += r.sent; acc.received += r.received; acc.diff += r.sent - r.received; return acc;
    }, { sent: 0, received: 0, diff: 0 });
    const totalY = Math.min(rowY + 8, pageH - 62);
    pdf.setFillColor(241, 244, 245);
    pdf.setDrawColor(220, 226, 230);
    pdf.roundedRect(tableX, totalY, tableW, 33, 5, 5, "FD");
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(7.2);
    pdf.setTextColor(83, 95, 103);
    pdf.text("PERIOD TOTALS", tableX + 8, totalY + 13);
    pdf.text(`Sent: ${formatQty(totals.sent)}`, tableX + 125, totalY + 13);
    pdf.text(`Received: ${formatQty(totals.received)}`, tableX + 222, totalY + 13);
    pdf.text(`Difference: ${signedQty(totals.diff)}`, tableX + 335, totalY + 13);
    pdf.text(`Transactions: ${allRows.length}`, tableX + tableW - 92, totalY + 13);
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(5.8);
    pdf.setTextColor(125, 137, 145);
    pdf.setTextColor(125, 137, 145);
    pdf.text("All-time difference = running sent-minus-received balance by asset after each transaction.", tableX, totalY + 45);
  }

  function groupLedgerRowsByTimestamp(rows) {
    const groups = new Map();
    rows.forEach(row => {
      const key = ledgerTransactionGroupKey(row);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(row);
    });
    return groups;
  }

  function ledgerTransactionGroupKey(row) {
    const d = parseTimestamp(row.timestamp);
    return `${d ? d.getTime() : String(row.timestamp)}::${row.client || ""}`;
  }

  async function prepareLedgerPhotos(rows) {
    const groups = new Map();
    rows.forEach(r => {
      const key = ledgerTransactionGroupKey(r);
      if (!groups.has(key)) groups.set(key, r.image || "");
    });
    for (const [key, url] of groups.entries()) {
      if (!url) continue;
      try {
        const photo = await loadLedgerPhoto(url);
        if (photo) rows.filter(r => ledgerTransactionGroupKey(r) === key).forEach(r => { r._pdfPhoto = photo; });
      } catch (_) {}
    }
    return rows;
  }

  async function loadLedgerPhoto(url) {
    if (!url) return null;
    if (window.FM_MEDIA?.loadDriveImage && state.accessToken) {
      const img = new Image();
      await window.FM_MEDIA.loadDriveImage(img, url, state.accessToken);
      await waitForImage(img, 12000);
      return imageElementToPdfData(img);
    }
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.src = url;
    try {
      await waitForImage(img, 8000);
      return imageElementToPdfData(img);
    } catch (_) {
      return null;
    }
  }

  function waitForImage(img, timeout) {
    if (img.complete && img.naturalWidth) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error("Image load timeout")), timeout);
      img.onload = () => { clearTimeout(timer); resolve(); };
      img.onerror = () => { clearTimeout(timer); reject(new Error("Image load failed")); };
    });
  }

  function imageElementToPdfData(img) {
    if (!img?.naturalWidth) return null;
    const maxW = 260, maxH = 180;
    const scale = Math.min(1, maxW / img.naturalWidth, maxH / img.naturalHeight);
    const w = Math.max(1, Math.round(img.naturalWidth * scale));
    const h = Math.max(1, Math.round(img.naturalHeight * scale));
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) return null;
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0, w, h);
    return { dataUrl: canvas.toDataURL("image/jpeg", 0.78), width: w, height: h };
  }

  function drawWrappedPdfText(pdf, text, x, y, width, fontSize, fontFace = "helvetica", fontStyle = "normal", color = null) {
    pdf.setFont(fontFace, fontStyle);
    pdf.setFontSize(fontSize);
    if (color) pdf.setTextColor(...color);
    const lines = pdf.splitTextToSize(String(text || ""), width);
    pdf.text(lines.slice(0, 3), x, y, { baseline: "top" });
  }

  function drawCellText(pdf, text, x, y, w, h, options = {}) {
    const fontSize = options.fontSize || 7.5;
    const align = options.align || "left";
    pdf.setFont("helvetica", options.bold ? "bold" : "normal");
    pdf.setFontSize(fontSize);
    const color = options.color || [23,32,42];
    pdf.setTextColor(...color);
    const content = String(text ?? "");
    const maxWidth = w - 8;
    const lines = pdf.splitTextToSize(content, maxWidth).slice(0, 3);
    const lineH = fontSize + 1.5;
    const startY = y + Math.max(5, (h - lines.length * lineH) / 2 + fontSize * 0.8);
    lines.forEach((line, index) => {
      const lineY = startY + index * lineH;
      let tx = x + 4;
      if (align === "center") tx = x + w / 2;
      if (align === "right") tx = x + w - 4;
      pdf.text(line, tx, lineY, { align });
    });
  }

  function signedQty(v) {
    const n = numberAtValue(v);
    return `${n >= 0 ? "+" : ""}${formatQty(n)}`;
  }

  function reportFileName() {
    const date = $("reportDate").value || "report";
    const shiftValue = $("shiftSelect").value;
    const shift = shiftValue === "night" ? "Night" : shiftValue === "24hr" ? "EOD_DayNight" : "Day";
    return `Asset_${shiftValue === "24hr" ? "EOD" : "Shift"}_Report_${date.replaceAll("-", "")}_${shift}.pdf`;
  }

  function renderSlackSummary() {
    const area = $("slackSummary");
    if (!area) return;
    area.value = buildSlackSummary();
    $("summaryMeta").textContent = `${shiftRows().length.toLocaleString()} ${isEndOfDayReport() ? "24-hour" : "shift"} transactions • ${getScheduleRows().length.toLocaleString()} scheduled movements • ${selectedNotables().length.toLocaleString()} notable`;
  }

  function buildSlackSummary() {
    const info = shiftInfo();
    const lines = [];
    if (isEndOfDayReport()) {
      lines.push("📦 ASSET COMPLIANCE — END OF DAY REPORT");
      lines.push(`Period: 24-hour period • ${info.windowText}`);
      lines.push("Coverage: Day + Night shifts");
    } else {
      lines.push("📦 ASSET COMPLIANCE — END OF SHIFT REPORT");
      lines.push(`Shift: ${info.label} • ${info.windowText}`);
    }
    lines.push(`Report date: ${formatDate($("reportDate").value)}`);
    lines.push("");

    lines.push(isEndOfDayReport() ? "📦 ASSET INVENTORY — FULL DAY" : "📦 ASSET INVENTORY");
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
    } else lines.push(isEndOfDayReport() ? "• No scheduled movements recorded for the Day + Night period." : "• No scheduled movements for this shift.");
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
    downloadBlob(blob, reportFileName().replace(/\.pdf$/i, ".txt"), "text/plain");
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

  function nextDateKey(date) {
    const d = new Date(`${date}T12:00:00Z`);
    d.setUTCDate(d.getUTCDate() + 1);
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
