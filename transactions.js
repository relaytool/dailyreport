(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const SESSION_KEY = "fmAssetSession";
  const TOKEN_CACHE_EMAIL_KEY = "fmAssetAccessToken";
  const state = {
    accessToken: null,
    email: "",
    transactions: [],
    groups: [],
    selectedGroupId: null,
    selectedGroup: null,
    selectedHtml: "",
    selectedPhotos: [],
    receiptGroup: null,
    viewMode: "current",
    search: ""
  };

  let tokenPromise = null;
  let refreshTimer = null;

  const londonDateFormatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/London", year: "numeric", month: "2-digit", day: "2-digit"
  });
  const londonPartsFormatter = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23"
  });
  const longDateFormatter = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London", day: "2-digit", month: "short", year: "numeric"
  });

  const $ = id => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    bindEvents();
    state.viewMode = "current";
    updatePeriodLabels();
    startClockRefresh();
    initialiseGoogle();
  }

  function bindEvents() {
    $("connectBtn")?.addEventListener("click", () => requestSheetAccess(false));
    $("refreshBtn")?.addEventListener("click", () => loadSheetsData());
    $("signOutBtn")?.addEventListener("click", signOut);
    $("viewMode")?.addEventListener("change", () => {
      state.viewMode = $("viewMode").value;
      updatePeriodLabels();
      renderGroups();
    });
    $("searchInput")?.addEventListener("input", () => {
      state.search = $("searchInput").value.trim().toLowerCase();
      renderGroups();
    });
    $("shareMessageBtn")?.addEventListener("click", shareMessage);
    $("copyMessageBtn")?.addEventListener("click", copyMessage);
    $("printReceiptBtn")?.addEventListener("click", openReceiptEditor);
    $("closeReceiptBtn")?.addEventListener("click", closeReceiptEditor);
    $("cancelReceiptBtn")?.addEventListener("click", closeReceiptEditor);
    $("printReceiptConfirmBtn")?.addEventListener("click", printReceipt);
    ["receiptClientName", "receiptClientAddress", "receiptSender", "receiptReceiver"].forEach(id => {
      $(id)?.addEventListener("input", renderReceiptPreview);
    });
    $("receiptModal")?.addEventListener("click", event => {
      if (event.target === $("receiptModal")) closeReceiptEditor();
    });
    document.addEventListener("keydown", event => {
      if (event.key === "Escape" && !$("receiptModal")?.classList.contains("hidden")) closeReceiptEditor();
    });
    $("resetMessageBtn")?.addEventListener("click", () => {
      if (state.selectedGroup) buildMessageForGroup(state.selectedGroup, true);
    });
  }

  function startClockRefresh() {
    clearInterval(refreshTimer);
    refreshTimer = setInterval(() => {
      if (state.viewMode === "current") {
        const before = periodWindow("current");
        updatePeriodLabels();
        const after = periodWindow("current");
        if (before.key !== after.key) renderGroups();
      } else {
        updatePeriodLabels();
      }
    }, 30 * 1000);
  }

  function initialiseGoogle() {
    const cached = window.FM_AUTH_CACHE?.read?.();
    const saved = readSession();
    if (cached?.token) {
      state.accessToken = cached.token;
      state.email = cached.email || saved?.email || "";
      setConnectedUi(true);
      loadSheetsData();
      return;
    }
    setSyncStatus("Sign in to load live transaction data.");
    setConnectedUi(false);
  }

  function requestSheetAccess(silent = false) {
    return acquireAccessToken(silent ? "none" : "select_account")
      .then(() => loadSheetsData())
      .catch(error => {
        console.warn("Google authorization failed:", error);
        setConnectedUi(false);
        setSyncStatus(error.message || "Google sign-in failed.", true);
        if (!silent) toast("Google sign-in failed. Check the Client ID and allowed origin.");
        throw error;
      });
  }

  function acquireAccessToken(prompt = "none", email) {
    const saved = readSession();
    const cached = window.FM_AUTH_CACHE?.read?.(email || saved?.email);
    if (cached?.token) {
      state.accessToken = cached.token;
      state.email = cached.email || email || saved?.email || "";
      return Promise.resolve(cached.token);
    }
    if (tokenPromise) return tokenPromise;
    if (!window.google?.accounts?.oauth2?.initTokenClient) {
      return Promise.reject(new Error("Google sign-in is still loading. Please try again."));
    }

    tokenPromise = new Promise((resolve, reject) => {
      let settled = false;
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        tokenPromise = null;
        fn(value);
      };
      const timeout = setTimeout(() => finish(reject, new Error("Google authorization timed out.")), prompt === "none" ? 6000 : 20000);

      try {
        const client = google.accounts.oauth2.initTokenClient({
          client_id: CONFIG.GOOGLE_CLIENT_ID,
          scope: "https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/drive.readonly",
          callback: response => {
            clearTimeout(timeout);
            if (response.error) {
              finish(reject, new Error(response.error_description || response.error));
              return;
            }
            state.accessToken = response.access_token;
            const savedSession = readSession();
            state.email = email || savedSession?.email || "";
            window.FM_AUTH_CACHE?.write?.(response.access_token, response.expires_in, state.email);
            finish(resolve, response.access_token);
          }
        });
        const options = { prompt };
        const loginHint = email || saved?.email;
        if (loginHint) options.login_hint = loginHint;
        client.requestAccessToken(options);
      } catch (error) {
        clearTimeout(timeout);
        finish(reject, error);
      }
    });
    return tokenPromise;
  }

  async function loadSheetsData() {
    if (!state.accessToken) {
      return requestSheetAccess(true);
    }
    setSyncStatus("Syncing Asset Transactions…");
    try {
      const rows = await getValues(CONFIG.INVENTORY_LEDGER_SHEET_ID, CONFIG.TRANSACTIONS_SHEET_NAME);
      state.transactions = parseTransactions(rows);
      const window = periodWindow(state.viewMode);
      const count = state.transactions.filter(tx => isWithinWindow(tx.timestamp, window)).length;
      setSyncStatus(`Live Sheets data • ${count.toLocaleString()} transactions in this period • ${formatTime(new Date())}`);
      setConnectedUi(true);
      renderGroups();
    } catch (error) {
      console.error(error);
      setSyncStatus(error.message || "Could not load transactions.", true);
      if (/expired|401|403/i.test(String(error.message))) {
        window.FM_AUTH_CACHE?.clear?.();
        state.accessToken = null;
        setConnectedUi(false);
      }
    }
  }

  async function getValues(spreadsheetId, sheetName) {
    const range = `'${String(sheetName).replaceAll("'", "''")}'!A:AE`;
    const response = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`, {
      headers: { Authorization: `Bearer ${state.accessToken}` },
      cache: "no-store"
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data?.error?.message || `Sheets request failed (${response.status})`);
    return data.values || [];
  }

  function parseTransactions(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];
    const headerIndex = rows.findIndex(row => row.some(value => normalizeHeader(value) === "timestamp"));
    if (headerIndex < 0) throw new Error('Could not find Timestamp in "Asset Transactions".');
    const h = rows[headerIndex].map(normalizeHeader);
    const find = aliases => aliases.map(normalizeHeader).reduce((found, wanted) => found >= 0 ? found : h.indexOf(wanted), -1);
    const idx = {
      timestamp: find(["timestamp", "date", "datetime"]),
      client: find(["client", "client name"]),
      movement: find(["movement", "type", "direction"]),
      asset: find(["asset", "asset name", "item"]),
      quantity: find(["quantity", "qty"]),
      user: find(["user", "entered by", "email"]),
      comment: find(["comment", "comments", "notes"]),
      image: find(["image", "image link", "photo", "picture", "photo link", "attachment", "drive link"])
    };
    if ([idx.timestamp, idx.client, idx.movement, idx.asset, idx.quantity].some(v => v < 0)) {
      throw new Error('"Asset Transactions" needs Timestamp, Client, Movement, Asset and Quantity columns.');
    }
    return rows.slice(headerIndex + 1).map((row, offset) => {
      const movement = String(row[idx.movement] ?? "").trim().toUpperCase();
      const timestamp = row[idx.timestamp];
      return {
        rowNumber: headerIndex + offset + 2,
        timestamp,
        timestampMs: parseDate(timestamp).getTime(),
        client: valueAt(row, idx.client),
        movement,
        asset: valueAt(row, idx.asset),
        quantity: numberAt(row, idx.quantity),
        user: valueAt(row, idx.user),
        vehicleReg: valueAt(row, idx.comment),
        comment: valueAt(row, idx.comment),
        image: valueAt(row, idx.image)
      };
    }).filter(tx => tx.client && tx.asset && tx.quantity > 0 && tx.timestamp);
  }

  function buildGroups() {
    const period = periodWindow(state.viewMode);

    // A group is now defined ONLY by the exact transaction timestamp.
    // Different clients at the same timestamp are intentionally combined.
    const rows = state.transactions
      .filter(tx => isWithinWindow(tx.timestamp, period))
      .sort((a, b) => parseDate(a.timestamp) - parseDate(b.timestamp) || a.rowNumber - b.rowNumber);

    const buckets = new Map();

    for (const tx of rows) {
      const time = parseDate(tx.timestamp).getTime();
      if (!Number.isFinite(time)) continue;

      if (!buckets.has(time)) {
        buckets.set(time, {
          id: `time-${time}`,
          client: tx.client,
          clients: [],
          rows: [],
          firstTime: time,
          lastTime: time,
          timestampKey: time
        });
      }

      const bucket = buckets.get(time);
      bucket.rows.push(tx);
      bucket.lastTime = time;
      if (!bucket.clients.some(client => normalize(client) === normalize(tx.client))) {
        bucket.clients.push(tx.client);
      }
    }

    const groups = [...buckets.values()];

    groups.sort((a, b) => a.firstTime - b.firstTime);

    groups.forEach(group => {
      group.rows.sort(
        (a, b) => parseDate(a.timestamp) - parseDate(b.timestamp) || a.rowNumber - b.rowNumber
      );

      group.clients = unique(group.clients);
      group.isMixed = group.clients.length > 1;
      group.client = group.isMixed ? "Mixed delivery" : (group.clients[0] || "Unknown client");

      group.directions = unique(group.rows.map(row => normalizeMovement(row.movement)));
      group.photoLinks = unique(group.rows.map(row => row.image).filter(Boolean));
      group.vehicleRegs = unique(group.rows.map(row => row.vehicleReg).filter(Boolean));
      group.comments = group.vehicleRegs.slice();
      group.assetSummary = summarizeAssets(group.rows);
    });

    // Search an exact-time group as a whole. If one client matches,
    // keep every other client sharing that exact timestamp in the group.
    if (state.search) {
      return groups.filter(group => group.rows.some(matchesSearch));
    }

    return groups;
  }

  function matchesSearch(tx) {
    if (!state.search) return true;
    const haystack = [
      tx.client,
      tx.asset,
      tx.movement,
      tx.vehicleReg,
      tx.user
    ].join(" ").toLowerCase();
    return haystack.includes(state.search);
  }

  function renderGroups() {
    state.groups = buildGroups();
    const container = $("transactionGroups");
    $("groupCount").textContent = `${state.groups.length.toLocaleString()} group${state.groups.length === 1 ? "" : "s"}`;
    const count = state.transactions.filter(tx => isWithinWindow(tx.timestamp, periodWindow(state.viewMode))).length;
    $("transactionCount").textContent = `${count.toLocaleString()} transaction${count === 1 ? "" : "s"}`;

    if (!state.groups.length) {
      container.innerHTML = `<div class="no-results">No transactions match this period${state.search ? " and search" : ""}.</div>`;
      state.selectedGroup = null;
      state.selectedGroupId = null;
      resetComposer();
      return;
    }

    container.innerHTML = state.groups.map(group => groupCard(group)).join("");
    container.querySelectorAll("[data-group-id]").forEach(button => {
      button.addEventListener("click", () => selectGroup(button.dataset.groupId));
    });

    if (state.selectedGroupId) {
      const selected = state.groups.find(g => g.id === state.selectedGroupId);
      if (selected) {
        state.selectedGroup = selected;
        markSelectedGroup();
      } else {
        selectGroup(state.groups[0].id);
      }
    } else {
      selectGroup(state.groups[0].id);
    }
  }

  function groupCard(group) {
    const first = new Date(group.firstTime);
    const chips = group.assetSummary
      .map(item => `<span class="asset-chip">${escapeHtml(formatQty(item.quantity))} ${escapeHtml(item.asset)}</span>`)
      .join("");

    const dirs = group.directions
      .map(direction =>
        `<span class="direction-pill ${direction === "Outbound" ? "outbound" : direction === "Inbound" ? "inbound" : ""}">
          ${direction === "Outbound" ? "↗ OUTBOUND" : direction === "Inbound" ? "↙ INBOUND" : "↔ MOVEMENT"}
        </span>`
      )
      .join("");

    const photoCount = group.photoLinks.length;
    const clientSummary = group.isMixed
      ? group.clients.join(" · ")
      : (group.clients[0] || "Unknown client");

    return `<article class="transaction-group" data-card-id="${escapeAttr(group.id)}">
      <button type="button" class="group-button" data-group-id="${escapeAttr(group.id)}">
        <div class="group-main">
          <div class="group-time">${escapeHtml(formatMessageTime(first))}</div>
          <div>
            <div class="group-client">${escapeHtml(group.isMixed ? "Mixed delivery" : clientSummary)}</div>
            <span class="group-window">${escapeHtml(clientSummary)} · exact same timestamp</span>
          </div>
          <div class="asset-summary">${chips}</div>
          <div class="group-arrow">›</div>
        </div>
        <div class="group-meta">
          <div class="direction-pills">${dirs || `<span class="direction-pill">OTHER</span>`}</div>
          <span>${group.rows.length} rows${photoCount ? ` · ${photoCount} photo${photoCount === 1 ? "" : "s"}` : ""}</span>
        </div>
      </button>
    </article>`;
  }

  function selectGroup(id) {
    const group = state.groups.find(item => item.id === id);
    if (!group) return;
    state.selectedGroupId = id;
    state.selectedGroup = group;
    markSelectedGroup();
    buildMessageForGroup(group, true);
  }

  function markSelectedGroup() {
    document.querySelectorAll(".transaction-group").forEach(el => {
      el.classList.toggle("selected", el.dataset.cardId === state.selectedGroupId);
    });
    const group = state.selectedGroup;
    if (!group) return;
    const first = new Date(group.firstTime);
    const clientLabel = group.isMixed ? `Mixed delivery · ${group.clients.join(" · ")}` : (group.client || "Unknown client");
    $("selectedMeta").textContent = `${clientLabel} · ${formatMessageTime(first)} · ${group.rows.length} rows`;
  }

  async function buildMessageForGroup(group, overwrite = true) {
    if (!group) return;
    const editor = $("messageEditor");
    if (!overwrite && editor.innerText.trim()) return;
    state.selectedHtml = "";
    state.selectedPhotos = [];
    $("shareMessageBtn").disabled = true;
    $("copyMessageBtn").disabled = true;
    $("printReceiptBtn").disabled = true;
    $("resetMessageBtn").disabled = false;
    $("photoStatus").textContent = "Loading transaction photo(s)…";
    editor.innerHTML = messageHtmlWithoutPhotos(group) + `<p class="message-loading">Loading photo…</p>`;

    const photos = await loadGroupPhotos(group);
    const html = messageHtml(group, photos);
    editor.innerHTML = html;
    state.selectedHtml = html;
    state.selectedPhotos = photos;
    $("shareMessageBtn").disabled = false;
    $("photoStatus").textContent = photos.length
      ? `${photos.length} transaction photo${photos.length === 1 ? "" : "s"} included.`
      : "No transaction photo is attached to this group.";
    $("copyMessageBtn").disabled = false;
    $("printReceiptBtn").disabled = false;
  }

  function messageHtmlWithoutPhotos(group) {
    const lines = messageLines(group);
    return `<div class="message-bubble">${lines.map(line => `<span class="line${line.type ? ` ${line.type}` : ""}">${escapeHtml(line.text)}</span>`).join("")}</div>`;
  }

  function messageHtml(group, photos) {
    const photoHtml = photos.length
      ? photos.map(photo => photo.dataUrl
          ? `<img class="message-photo" src="${escapeAttr(photo.dataUrl)}" alt="Transaction photo" data-transaction-photo="true">`
          : "").join("")
      : "";
    return `<div class="message-content">${photoHtml}${messageHtmlWithoutPhotos(group)}</div>`;
  }

  function messageLines(group) {
    const lines = [];
    const time = formatMessageTime(group.firstTime);
    const vehicleReg = group.vehicleRegs?.length ? group.vehicleRegs.join(" / ") : "—";

    // Keep every client separate, while sharing one timestamp, vehicle
    // registration and photo set for the complete exact-time group.
    if (group.isMixed) {
      lines.push({ text: "📦 MIXED DELIVERY", type: "direction" });
      lines.push({ text: `🏭 Clients: ${group.clients.join(" · ")}` });
    }

    const byClient = new Map();
    group.rows.forEach(row => {
      const clientKey = normalize(row.client);
      if (!byClient.has(clientKey)) {
        byClient.set(clientKey, {
          client: row.client,
          rows: []
        });
      }
      byClient.get(clientKey).rows.push(row);
    });

    for (const clientGroup of byClient.values()) {
      if (group.isMixed || byClient.size > 1) {
        lines.push({ text: `🏭 Client: ${clientGroup.client}` });
      }

      const byMovement = new Map();
      clientGroup.rows.forEach(row => {
        const direction = normalizeMovement(row.movement);
        if (!byMovement.has(direction)) byMovement.set(direction, []);
        byMovement.get(direction).push(row);
      });

      ["Outbound", "Inbound", "Other"].forEach(direction => {
        const rows = byMovement.get(direction);
        if (!rows?.length) return;

        const sentence = summarizeAssets(rows)
          .map(item => `${formatQty(item.quantity)} ${item.asset}`)
          .join(", ");

        if (direction === "Outbound") {
          lines.push({ text: "↗ OUTBOUND", type: "direction" });
          lines.push({ text: `📦 Sent: ${sentence}` });
        } else if (direction === "Inbound") {
          lines.push({ text: "↙ INBOUND", type: "direction" });
          lines.push({ text: `📦 Received: ${sentence}` });
        } else {
          lines.push({ text: "↔ MOVEMENT", type: "direction" });
          lines.push({ text: `📦 Recorded: ${sentence}` });
        }
      });
    }

    lines.push({ text: `🕒 Time: ${time}` });
    lines.push({ text: `🚚 Vehicle Reg: ${vehicleReg}` });
    return lines;
  }

  async function loadGroupPhotos(group) {
    const links = group.photoLinks.slice(0, 8);
    const results = [];
    for (const url of links) {
      const id = driveFileIdFromLink(url);
      let image = null;
      if (id && state.accessToken) {
        image = await fetchDriveImage(id);
      }
      results.push({ url, dataUrl: image?.dataUrl || "", blob: image?.blob || null });
    }
    return results;
  }

  async function fetchDriveImage(fileId) {
    try {
      const response = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`, {
        headers: { Authorization: `Bearer ${state.accessToken}` },
        cache: "force-cache"
      });
      if (!response.ok) throw new Error(`Drive image request failed (${response.status})`);
      const blob = await response.blob();
      return { blob, dataUrl: await blobToDataUrl(blob) };
    } catch (error) {
      console.warn("Could not load Drive image as a blob:", error);
      try {
        const fallback = await fetch(`https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=w1600`, {
          headers: { Authorization: `Bearer ${state.accessToken}` },
          cache: "force-cache"
        });
        if (fallback.ok) {
          const blob = await fallback.blob();
          return { blob, dataUrl: await blobToDataUrl(blob) };
        }
      } catch (fallbackError) {
        console.warn("Drive thumbnail fallback failed:", fallbackError);
      }
      return null;
    }
  }

  function resetComposer() {
    $("selectedMeta").textContent = "Select a transaction group";
    $("messageEditor").innerHTML = `<p class="editor-placeholder">Select a transaction group to build the message.</p>`;
    $("photoStatus").textContent = "";
    $("shareMessageBtn").disabled = true;
    $("copyMessageBtn").disabled = true;
    $("printReceiptBtn").disabled = true;
    $("resetMessageBtn").disabled = true;
    $("copyStatus").textContent = "";
  }

  function openReceiptEditor() {
    const group = state.selectedGroup;
    if (!group) return;
    state.receiptGroup = group;

    const firstRow = group.rows[0] || {};
    $("receiptClientName").value = group.isMixed ? `Mixed delivery — ${group.clients.join(" · ")}` : (group.client || "");
    $("receiptClientAddress").value = "";
    $("receiptSender").value = firstRow.user || state.email || "";
    $("receiptReceiver").value = "";
    renderReceiptPreview();
    $("receiptModal").classList.remove("hidden");
    document.body.style.overflow = "hidden";
  }

  function closeReceiptEditor() {
    $("receiptModal")?.classList.add("hidden");
    document.body.style.overflow = "";
  }

  function renderReceiptPreview() {
    const group = state.receiptGroup || state.selectedGroup;
    if (!group || !$("receiptPreview")) return;
    const clientName = $("receiptClientName")?.value.trim() || group.client || "Client";
    const clientAddress = $("receiptClientAddress")?.value.trim() || "Address not provided";
    const sender = $("receiptSender")?.value.trim() || "";
    const receiver = $("receiptReceiver")?.value.trim() || "";
    $("receiptPreview").innerHTML = receiptDocumentHtml(group, { clientName, clientAddress, sender, receiver });
  }

  function receiptDocumentHtml(group, details) {
    const first = new Date(group.firstTime);
    const receiptRef = `REL-${group.rows[0]?.rowNumber || "TX"}-${londonDateKey(first).replaceAll("-", "")}`;
    const vehicleReg = group.vehicleRegs?.length ? group.vehicleRegs.join(" / ") : "—";
    const photoHtml = (state.selectedPhotos || [])
      .filter(photo => photo.dataUrl)
      .map(photo => `<img src="${escapeAttr(photo.dataUrl)}" alt="Transaction photo">`)
      .join("");

    const mixedClients = group.isMixed || group.clients.length > 1;

    const rowsByClient = new Map();
    group.rows.forEach(row => {
      const key = normalize(row.client);
      if (!rowsByClient.has(key)) {
        rowsByClient.set(key, {
          client: row.client,
          rows: []
        });
      }
      rowsByClient.get(key).rows.push(row);
    });

    const tableRows = [];
    for (const clientGroup of rowsByClient.values()) {
      const sentRows = clientGroup.rows.filter(row => normalizeMovement(row.movement) === "Outbound");
      const receivedRows = clientGroup.rows.filter(row => normalizeMovement(row.movement) === "Inbound");
      const otherRows = clientGroup.rows.filter(row => normalizeMovement(row.movement) === "Other");

      const makeRows = (rows, direction) => summarizeAssets(rows).map(item => ({
        client: clientGroup.client,
        direction,
        asset: item.asset,
        quantity: item.quantity
      }));

      tableRows.push(
        ...makeRows(sentRows, "Sent"),
        ...makeRows(receivedRows, "Received"),
        ...makeRows(otherRows, "Recorded")
      );
    }

    const tableRowsHtml = tableRows.length
      ? tableRows.map(item => `
          <tr>
            ${mixedClients ? `<td>${escapeHtml(item.client)}</td>` : ""}
            <td>${escapeHtml(item.direction)}</td>
            <td>${escapeHtml(item.asset)}</td>
            <td class="qty">${escapeHtml(formatQty(item.quantity))}</td>
          </tr>
        `).join("")
      : `<tr><td colspan="${mixedClients ? 4 : 3}">No item rows available for this transaction.</td></tr>`;

    const tableHeadHtml = mixedClients
      ? `<tr><th>Client</th><th>Movement</th><th>Item</th><th class="qty">Quantity</th></tr>`
      : `<tr><th>Movement</th><th>Item</th><th class="qty">Quantity</th></tr>`;

    const clientDisplay = mixedClients
      ? group.clients.join(" · ")
      : (group.client || "Client");

    return `<div class="receipt-preview-page">
      <div class="receipt-preview-header">
        <div>
          <div class="receipt-brand">Relay</div>
          <div class="receipt-address">3 Iron Bridge Road<br>West Drayton<br>UB11 1BF</div>
        </div>
        <div>
          <div class="receipt-doc-title">TRANSACTION RECEIPT</div>
          <div class="receipt-meta">Ref: ${escapeHtml(receiptRef)}<br>${escapeHtml(formatDateTime(first))}</div>
        </div>
      </div>

      <div class="receipt-address-row">
        <div class="receipt-block">
          <div class="receipt-block-label">From</div>
          <strong>Relay</strong>
          <div class="receipt-address">3 Iron Bridge Road\nWest Drayton\nUB11 1BF</div>
        </div>
        <div class="receipt-block">
          <div class="receipt-block-label">${mixedClients ? "Clients" : "To"}</div>
          <strong>${escapeHtml(details.clientName || clientDisplay)}</strong>
          <div class="receipt-address">${escapeHtml(details.clientAddress)}</div>
        </div>
      </div>

      <div class="receipt-details">
        <div><div class="receipt-detail-label">Transaction time</div><div class="receipt-detail-value">${escapeHtml(formatDateTime(first))}</div></div>
        <div><div class="receipt-detail-label">Vehicle Reg</div><div class="receipt-detail-value">${escapeHtml(vehicleReg)}</div></div>
        <div><div class="receipt-detail-label">Transaction rows</div><div class="receipt-detail-value">${escapeHtml(String(group.rows.length))}</div></div>
      </div>

      <div class="receipt-section-title">${mixedClients ? "Items by client" : "Items"}</div>
      <table class="receipt-table">
        <thead>${tableHeadHtml}</thead>
        <tbody>${tableRowsHtml}</tbody>
      </table>

      <div class="receipt-section-title">Transaction photo${(state.selectedPhotos || []).filter(photo => photo.dataUrl).length === 1 ? "" : "s"}</div>
      ${photoHtml ? `<div class="receipt-photo-grid">${photoHtml}</div>` : `<div class="receipt-no-photo">No transaction photo is available for this transaction.</div>`}

      <div class="receipt-signatures">
        <div class="receipt-signature">
          <div class="receipt-signature-label">Sender</div>
          <div class="receipt-signature-name">${escapeHtml(details.sender || "")}</div>
        </div>
        <div class="receipt-signature">
          <div class="receipt-signature-label">Receiver</div>
          <div class="receipt-signature-name">${escapeHtml(details.receiver || "")}</div>
        </div>
      </div>

      <div class="receipt-footer"><span>Relay transaction receipt</span><span>${escapeHtml(clientDisplay)}</span></div>
    </div>`;
  }

  function printReceipt() {
    const group = state.receiptGroup || state.selectedGroup;
    if (!group) return;
    const clientName = $("receiptClientName")?.value.trim() || group.client || "Client";
    const clientAddress = $("receiptClientAddress")?.value.trim() || "Address not provided";
    const sender = $("receiptSender")?.value.trim() || "";
    const receiver = $("receiptReceiver")?.value.trim() || "";
    const documentHtml = receiptDocumentHtml(group, { clientName, clientAddress, sender, receiver });
    const printWindow = window.open("", "_blank", "width=900,height=1200");
    if (!printWindow) {
      toast("Please allow pop-ups to print the receipt.");
      return;
    }

    const printCss = `
      @page { size: A4 portrait; margin: 0; }
      * { box-sizing: border-box; }
      html, body { margin: 0; padding: 0; background: #fff; color: #17202a; font-family: Inter, Arial, sans-serif; }
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .receipt-preview-page { width: 210mm; min-height: 297mm; margin: 0 auto; background: #fff; padding: 17mm 16mm 14mm; }
      .receipt-preview-header { display:flex; justify-content:space-between; gap:28px; padding-bottom:18px; border-bottom:2px solid #17202a; }
      .receipt-brand { font-size:25px; font-weight:900; letter-spacing:-.03em; }
      .receipt-doc-title { font-size:18px; font-weight:850; margin-top:4px; text-align:right; }
      .receipt-meta { margin-top:5px; color:#64748b; font-size:11px; text-align:right; line-height:1.5; }
      .receipt-address { white-space:pre-line; color:#475569; font-size:11px; line-height:1.45; }
      .receipt-address-row { display:grid; grid-template-columns:1fr 1fr; gap:24px; padding:22px 0; border-bottom:1px solid #dfe5ec; }
      .receipt-block-label,.receipt-detail-label { font-size:9px; font-weight:900; letter-spacing:.1em; text-transform:uppercase; color:#64748b; margin-bottom:5px; }
      .receipt-block strong { display:block; font-size:13px; margin-bottom:3px; }
      .receipt-details { display:grid; grid-template-columns:1fr 1fr 1fr; gap:14px; padding:18px 0; border-bottom:1px solid #dfe5ec; }
      .receipt-detail-value { font-size:12px; font-weight:750; }
      .receipt-table { width:100%; border-collapse:collapse; margin-top:20px; font-size:11px; }
      .receipt-table th { padding:9px 8px; border-bottom:2px solid #17202a; text-align:left; font-size:9px; letter-spacing:.08em; text-transform:uppercase; }
      .receipt-table td { padding:9px 8px; border-bottom:1px solid #e5eaf0; vertical-align:top; }
      .receipt-table .qty { text-align:right; font-weight:800; white-space:nowrap; }
      .receipt-section-title { font-size:11px; font-weight:900; letter-spacing:.08em; text-transform:uppercase; margin-top:22px; margin-bottom:8px; }
      .receipt-photo-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }
      .receipt-photo-grid img { display:block; width:100%; max-height:330px; object-fit:contain; border:1px solid #dfe5ec; background:#f8fafc; border-radius:8px; }
      .receipt-no-photo { padding:14px; border:1px dashed #cbd5e1; border-radius:8px; color:#64748b; font-size:11px; }
      .receipt-signatures { display:grid; grid-template-columns:1fr 1fr; gap:36px; margin-top:36px; padding-top:14px; }
      .receipt-signature { border-top:1px solid #17202a; padding-top:8px; min-height:58px; }
      .receipt-signature-label { font-size:9px; font-weight:900; letter-spacing:.08em; text-transform:uppercase; color:#64748b; }
      .receipt-signature-name { font-size:12px; font-weight:750; margin-top:5px; }
      .receipt-footer { margin-top:26px; padding-top:10px; border-top:1px solid #dfe5ec; color:#64748b; font-size:9px; display:flex; justify-content:space-between; gap:12px; }
      @media print { .receipt-preview-page { width:210mm; min-height:297mm; } }
    `;

    printWindow.document.open();
    printWindow.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Relay Transaction Receipt</title><style>${printCss}</style></head><body>${documentHtml}</body></html>`);
    printWindow.document.close();

    const waitForImagesThenPrint = () => {
      const images = [...printWindow.document.images];
      const waits = images.map(img => {
        if (img.complete) return Promise.resolve();
        return new Promise(resolve => {
          img.addEventListener("load", resolve, { once: true });
          img.addEventListener("error", resolve, { once: true });
        });
      });
      Promise.all(waits).then(() => {
        setTimeout(() => {
          printWindow.focus();
          printWindow.print();
        }, 150);
      });
    };
    if (printWindow.document.readyState === "complete") {
      waitForImagesThenPrint();
    } else {
      printWindow.onload = waitForImagesThenPrint;
    }
  }

  async function shareMessage() {
    const editor = $("messageEditor");
    const plain = editor.innerText.trim();
    if (!plain) return;

    const files = (state.selectedPhotos || [])
      .filter(photo => photo.blob)
      .slice(0, 8)
      .map((photo, index) => {
        const type = photo.blob.type || "image/jpeg";
        const ext = type.includes("png") ? "png" : type.includes("webp") ? "webp" : "jpg";
        return new File([photo.blob], `transaction-${index + 1}.${ext}`, { type });
      });

    try {
      if (navigator.share) {
        const payload = { text: plain };
        if (files.length) {
          if (navigator.canShare && !navigator.canShare({ files })) {
            throw new Error("This browser cannot share the transaction photo with the message.");
          }
          payload.files = files;
        }
        await navigator.share(payload);
        $("copyStatus").textContent = files.length ? "Shared with text and photo." : "Shared message text.";
        return;
      }
    } catch (error) {
      if (error?.name === "AbortError") return;
      console.warn("Native share failed:", error);
    }

    await copyMessage();
    $("copyStatus").textContent = files.length
      ? "Device sharing is unavailable here. Rich copy was used instead."
      : "Device sharing is unavailable here. Message copied.";
  }

  async function copyMessage() {
    const editor = $("messageEditor");
    const html = cleanEditorHtml(editor.innerHTML);
    const plain = editor.innerText.trim();
    if (!plain && !html) return;

    try {
      if (navigator.clipboard?.write && window.ClipboardItem) {
        const itemParts = {
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([plain], { type: "text/plain" })
        };
        await navigator.clipboard.write([new ClipboardItem(itemParts)]);
        $("copyStatus").textContent = "Copied rich message. On mobile, use Share with photo for the most reliable photo + text send.";
      } else if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(plain);
        $("copyStatus").textContent = "Copied text. Use Share with photo to include the image on iPhone/Android.";
      } else {
        fallbackCopy(editor);
        $("copyStatus").textContent = "Copied message.";
      }
    } catch (error) {
      console.warn("Rich clipboard failed:", error);
      try {
        await navigator.clipboard.writeText(plain);
        $("copyStatus").textContent = "Copied text. Use Share with photo to include the image on iPhone/Android.";
      } catch (_) {
        fallbackCopy(editor);
        $("copyStatus").textContent = "Copied message.";
      }
    }
  }

  function fallbackCopy(editor) {
    const range = document.createRange();
    range.selectNodeContents(editor);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    document.execCommand("copy");
    selection.removeAllRanges();
  }

  function cleanEditorHtml(html) {
    const wrapper = document.createElement("div");
    wrapper.innerHTML = html;
    wrapper.querySelectorAll(".message-loading,.editor-placeholder").forEach(el => el.remove());
    wrapper.querySelectorAll(".message-photo").forEach(img => {
      img.removeAttribute("contenteditable");
      img.style.maxWidth = "100%";
      img.style.height = "auto";
      img.style.display = "block";
    });
    return wrapper.innerHTML;
  }

  function setConnectedUi(connected) {
    $("connectBtn").classList.toggle("hidden", connected);
    $("signOutBtn").classList.toggle("hidden", !connected);
    $("statusBar").classList.toggle("ready", connected);
  }

  function setSyncStatus(message, error = false) {
    $("syncStatus").textContent = message;
    $("statusBar").classList.toggle("error", Boolean(error));
  }

  function updatePeriodLabels() {
    const period = periodWindow(state.viewMode);
    $("periodLabel").textContent = period.label;
    $("shiftSubtitle").textContent = `${period.label} · Europe/London`;
  }

  function periodWindow(mode) {
    const now = new Date();
    if (mode === "current") {
      const mins = londonMinuteOfDay(now);
      if (mins >= 420 && mins < 1140) {
        const date = londonDateKey(now);
        const end = `${date}T19:00:00`;
        return { key: `day:${date}`, label: `Day shift · ${formatLocalDate(date)} 07:00–19:00`, startDate: date, endDate: date, startMin: 420, endMin: 1140, type: "day" };
      }
      const endDate = mins < 420 ? londonDateKey(now) : nextDateKey(londonDateKey(now));
      const startDate = mins < 420 ? previousDateKey(londonDateKey(now)) : londonDateKey(now);
      return { key: `night:${startDate}`, label: `Night shift · ${formatLocalDate(startDate)} 19:00–${formatLocalDate(endDate)} 07:00`, startDate, endDate, startMin: 1140, endMin: 420, type: "night" };
    }
    const today = londonDateKey(now);
    if (mode === "day") return { key: `day:${today}`, label: `Day shift · ${formatLocalDate(today)} 07:00–19:00`, startDate: today, endDate: today, startMin: 420, endMin: 1140, type: "day" };
    if (mode === "night") {
      const startDate = londonMinuteOfDay(now) < 420 ? previousDateKey(today) : today;
      const endDate = nextDateKey(startDate);
      return { key: `night:${startDate}`, label: `Night shift · ${formatLocalDate(startDate)} 19:00–${formatLocalDate(endDate)} 07:00`, startDate, endDate, startMin: 1140, endMin: 420, type: "night" };
    }
    const endDate = nextDateKey(today);
    return { key: `24hr:${today}`, label: `24 hours · ${formatLocalDate(today)} 07:00–${formatLocalDate(endDate)} 07:00`, startDate: today, endDate, startMin: 420, endMin: 420, type: "24hr" };
  }

  function isWithinWindow(timestamp, period) {
    const p = londonDateTimeParts(parseDate(timestamp));
    if (!p) return false;
    const mins = p.hour * 60 + p.minute + p.second / 60;
    if (period.type === "day") return p.date === period.startDate && mins >= 420 && mins < 1140;
    if (period.type === "night") return (p.date === period.startDate && mins >= 1140) || (p.date === period.endDate && mins < 420);
    if (period.type === "24hr") return (p.date === period.startDate && mins >= 420) || (p.date === period.endDate && mins < 420);
    return false;
  }

  function parseDate(value) {
    let d;
    if (value instanceof Date) {
      d = new Date(value.getTime());
    } else if (typeof value === "number") {
      d = new Date(value);
    } else {
      const text = String(value ?? "").trim();
      d = new Date(text);
    }
    return Number.isNaN(d.getTime()) ? new Date(0) : d;
  }

  function londonDateKey(value) {
    return value instanceof Date ? londonDateFormatter.format(value) : londonDateFormatter.format(parseDate(value));
  }

  function londonDateTimeParts(value) {
    const parts = londonPartsFormatter.formatToParts(parseDate(value));
    const get = type => Number(parts.find(p => p.type === type)?.value || 0);
    return {
      date: `${get("year")}-${String(get("month")).padStart(2,"0")}-${String(get("day")).padStart(2,"0")}`,
      hour: get("hour"), minute: get("minute"), second: get("second")
    };
  }

  function londonMinuteOfDay(value) {
    const p = londonDateTimeParts(value);
    return p.hour * 60 + p.minute;
  }

  function nextDateKey(dateKey) {
    const [y,m,d] = dateKey.split("-").map(Number);
    return dateKeyFromLocalParts(new Date(Date.UTC(y,m-1,d+1)));
  }

  function previousDateKey(dateKey) {
    const [y,m,d] = dateKey.split("-").map(Number);
    return dateKeyFromLocalParts(new Date(Date.UTC(y,m-1,d-1)));
  }

  function dateKeyFromLocalParts(date) {
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth()+1).padStart(2,"0")}-${String(date.getUTCDate()).padStart(2,"0")}`;
  }

  function formatLocalDate(dateKey) {
    const date = new Date(`${dateKey}T12:00:00Z`);
    return longDateFormatter.format(date);
  }

  function formatTime(value) {
    return new Intl.DateTimeFormat("en-GB", { timeZone:"Europe/London", hour:"2-digit", minute:"2-digit", hourCycle:"h23" }).format(parseDate(value));
  }

  function formatMessageTime(value) {
    const raw = formatTime(value);
    return raw.replace(/^0(?=\d)/, "");
  }

  function formatDateTime(value) {
    return new Intl.DateTimeFormat("en-GB", { timeZone:"Europe/London", day:"2-digit", month:"short", hour:"2-digit", minute:"2-digit", hourCycle:"h23" }).format(parseDate(value));
  }

  function formatQty(value) {
    return Number(value || 0).toLocaleString("en-GB");
  }

  function summarizeAssets(rows) {
    const map = new Map();
    rows.forEach(row => {
      const key = normalize(row.asset);
      const existing = map.get(key) || { asset: row.asset, quantity: 0 };
      existing.quantity += numberAtValue(row.quantity);
      map.set(key, existing);
    });
    return [...map.values()];
  }

  function normalizeMovement(value) {
    const v = normalize(value).toUpperCase();
    if (["SENT","OUT","OUTBOUND"].includes(v)) return "Outbound";
    if (["RECEIVED","IN","INBOUND"].includes(v)) return "Inbound";
    return "Other";
  }

  function normalize(value) { return String(value ?? "").trim().toLowerCase().replace(/\s+/g," "); }
  function normalizeHeader(value) { return normalize(value).replace(/\r?\n/g," "); }
  function valueAt(row,index) { return index >= 0 ? String(row[index] ?? "").trim() : ""; }
  function numberAt(row,index) { return index >= 0 ? numberAtValue(row[index]) : 0; }
  function numberAtValue(value) { const n = Number(String(value ?? "").replace(/,/g,"").trim()); return Number.isFinite(n) ? n : 0; }
  function unique(values) { return [...new Set(values)]; }
  function driveFileIdFromLink(url) { const text = String(url || ""); const m = text.match(/\/d\/([a-zA-Z0-9_-]+)/) || text.match(/[?&]id=([a-zA-Z0-9_-]+)/); return m ? m[1] : ""; }

  function readSession() {
    try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch (_) { return null; }
  }

  function signOut() {
    window.FM_AUTH_CACHE?.clear?.();
    try { localStorage.removeItem(SESSION_KEY); } catch (_) {}
    state.accessToken = null;
    state.email = "";
    state.transactions = [];
    state.groups = [];
    setConnectedUi(false);
    setSyncStatus("Signed out. Connect Google Sheets to load live data.");
    renderGroups();
  }

  function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[ch]));
  }

  function escapeAttr(value) { return escapeHtml(value).replace(/`/g,"&#96;"); }

  function toast(message) {
    const el = $("toast");
    el.textContent = message;
    el.classList.add("show");
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => el.classList.remove("show"), 2800);
  }
})();
