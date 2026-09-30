(() => {
  "use strict";

  const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
  const SESSION_KEY = "fmAssetSession";
  const TOKEN_CACHE_EMAIL_KEY = "fmAssetAccessToken";
  const WINDOW_MS = 5 * 60 * 1000;
  const state = {
    accessToken: null,
    email: "",
    transactions: [],
    groups: [],
    selectedGroupId: null,
    selectedGroup: null,
    selectedHtml: "",
    selectedPhotos: [],
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
    $("copyMessageBtn")?.addEventListener("click", copyMessage);
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
        client: valueAt(row, idx.client),
        movement,
        asset: valueAt(row, idx.asset),
        quantity: numberAt(row, idx.quantity),
        user: valueAt(row, idx.user),
        comment: valueAt(row, idx.comment),
        image: valueAt(row, idx.image)
      };
    }).filter(tx => tx.client && tx.asset && tx.quantity > 0 && tx.timestamp);
  }

  function buildGroups() {
    const period = periodWindow(state.viewMode);
    const rows = state.transactions
      .filter(tx => isWithinWindow(tx.timestamp, period))
      .filter(tx => matchesSearch(tx))
      .sort((a, b) => parseDate(a.timestamp) - parseDate(b.timestamp));

    const byClient = new Map();
    for (const tx of rows) {
      const key = normalize(tx.client);
      if (!byClient.has(key)) byClient.set(key, []);
      byClient.get(key).push(tx);
    }

    const groups = [];
    for (const [clientKey, clientRows] of byClient.entries()) {
      let bucket = null;
      for (const tx of clientRows) {
        const time = parseDate(tx.timestamp).getTime();
        if (!bucket || time - bucket.anchorTime > WINDOW_MS) {
          bucket = {
            id: `${clientKey}-${tx.rowNumber}-${time}`,
            client: tx.client,
            rows: [],
            anchorTime: time,
            firstTime: time,
            lastTime: time
          };
          groups.push(bucket);
        }
        bucket.rows.push(tx);
        bucket.lastTime = time;
      }
    }

    groups.sort((a, b) => a.firstTime - b.firstTime);
    groups.forEach(group => {
      group.rows.sort((a, b) => parseDate(a.timestamp) - parseDate(b.timestamp) || a.rowNumber - b.rowNumber);
      group.directions = unique(group.rows.map(row => normalizeMovement(row.movement)));
      group.photoLinks = unique(group.rows.map(row => row.image).filter(Boolean));
      group.comments = unique(group.rows.map(row => row.comment).filter(Boolean));
      group.assetSummary = summarizeAssets(group.rows);
    });
    return groups;
  }

  function matchesSearch(tx) {
    if (!state.search) return true;
    const haystack = [tx.client, tx.asset, tx.movement, tx.comment, tx.user].join(" ").toLowerCase();
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
    const last = new Date(group.lastTime);
    const sameMinute = group.firstTime === group.lastTime || Math.round((group.lastTime - group.firstTime) / 1000) < 60;
    const windowText = sameMinute ? formatTime(first) : `${formatTime(first)}–${formatTime(last)}`;
    const chips = group.assetSummary.map(item => `<span class="asset-chip">${escapeHtml(formatQty(item.quantity))} ${escapeHtml(item.asset)}</span>`).join("");
    const dirs = group.directions.map(direction => `<span class="direction-pill ${direction === "Outbound" ? "outbound" : "inbound"}">${direction === "Outbound" ? "↗ OUTBOUND" : "↙ INBOUND"}</span>`).join("");
    const photoCount = group.photoLinks.length;
    return `<article class="transaction-group" data-card-id="${escapeAttr(group.id)}">
      <button type="button" class="group-button" data-group-id="${escapeAttr(group.id)}">
        <div class="group-main">
          <div class="group-time">${escapeHtml(formatTime(first))}</div>
          <div><div class="group-client">${escapeHtml(group.client)}</div><span class="group-window">5-min window · ${escapeHtml(windowText)}</span></div>
          <div class="asset-summary">${chips}</div>
          <div class="group-arrow">›</div>
        </div>
        <div class="group-meta"><div class="direction-pills">${dirs || `<span class="direction-pill">OTHER</span>`}</div><span>${group.rows.length} rows${photoCount ? ` · ${photoCount} photo${photoCount === 1 ? "" : "s"}` : ""}</span></div>
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
    $("selectedMeta").textContent = `${group.client} · ${formatTime(first)} · ${group.rows.length} rows`;
  }

  async function buildMessageForGroup(group, overwrite = true) {
    if (!group) return;
    const editor = $("messageEditor");
    if (!overwrite && editor.innerText.trim()) return;
    state.selectedHtml = "";
    state.selectedPhotos = [];
    $("copyMessageBtn").disabled = true;
    $("resetMessageBtn").disabled = false;
    $("photoStatus").textContent = "Loading transaction photo(s)…";
    editor.innerHTML = messageHtmlWithoutPhotos(group) + `<p class="message-loading">Loading photo…</p>`;

    const photos = await loadGroupPhotos(group);
    const html = messageHtml(group, photos);
    editor.innerHTML = html;
    state.selectedHtml = html;
    state.selectedPhotos = photos;
    $("photoStatus").textContent = photos.length
      ? `${photos.length} transaction photo${photos.length === 1 ? "" : "s"} included.`
      : "No transaction photo is attached to this group.";
    $("copyMessageBtn").disabled = false;
  }

  function messageHtmlWithoutPhotos(group) {
    const blocks = movementBlocks(group);
    const comment = group.comments.length ? group.comments.join(" · ") : "—";
    const time = formatDateTime(group.firstTime) + (group.lastTime !== group.firstTime ? `–${formatTime(new Date(group.lastTime))}` : "");
    return `<div><p><strong>📦 ${escapeHtml(group.client)}</strong></p>${blocks}<p>🕒 <strong>Time:</strong> ${escapeHtml(time)}</p><p>💬 <strong>Comment:</strong> ${escapeHtml(comment)}</p></div>`;
  }

  function messageHtml(group, photos) {
    const base = messageHtmlWithoutPhotos(group);
    const photoHtml = photos.length
      ? `<p><strong>📷 Photo${photos.length > 1 ? "s" : ""}</strong></p>${photos.map(photo => `<p><img src="${photo.dataUrl || escapeAttr(photo.url)}" alt="Transaction photo" data-transaction-photo="true"></p>`).join("")}`
      : "";
    return `${base}${photoHtml}`;
  }

  function movementBlocks(group) {
    const byMovement = new Map();
    group.rows.forEach(row => {
      const direction = normalizeMovement(row.movement);
      if (!byMovement.has(direction)) byMovement.set(direction, []);
      byMovement.get(direction).push(row);
    });
    const html = [];
    ["Outbound", "Inbound", "Other"].forEach(direction => {
      const rows = byMovement.get(direction);
      if (!rows?.length) return;
      const sentence = summarizeAssets(rows).map(item => `${formatQty(item.quantity)} ${item.asset}`).join(", ");
      const label = direction === "Outbound" ? "↗ OUTBOUND" : direction === "Inbound" ? "↙ INBOUND" : "↔ MOVEMENT";
      const verb = direction === "Outbound" ? "Sent" : direction === "Inbound" ? "Received" : "Recorded";
      html.push(`<p>${label}<br><strong>📦 ${verb}:</strong> ${escapeHtml(sentence)}</p>`);
    });
    return html.join("");
  }

  async function loadGroupPhotos(group) {
    const links = group.photoLinks.slice(0, 8);
    const results = [];
    for (const url of links) {
      const id = driveFileIdFromLink(url);
      let dataUrl = "";
      if (id && state.accessToken) {
        dataUrl = await fetchDriveDataUrl(id);
      }
      results.push({ url, dataUrl });
    }
    return results;
  }

  async function fetchDriveDataUrl(fileId) {
    try {
      const response = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`, {
        headers: { Authorization: `Bearer ${state.accessToken}` },
        cache: "force-cache"
      });
      if (!response.ok) throw new Error(`Drive image request failed (${response.status})`);
      return await blobToDataUrl(await response.blob());
    } catch (error) {
      console.warn("Could not load Drive image as a blob:", error);
      try {
        const fallback = await fetch(`https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=w1600`, {
          headers: { Authorization: `Bearer ${state.accessToken}` },
          cache: "force-cache"
        });
        if (fallback.ok) return await blobToDataUrl(await fallback.blob());
      } catch (fallbackError) {
        console.warn("Drive thumbnail fallback failed:", fallbackError);
      }
      return "";
    }
  }

  function resetComposer() {
    $("selectedMeta").textContent = "Select a transaction group";
    $("messageEditor").innerHTML = `<p class="editor-placeholder">Select a transaction group to build the message.</p>`;
    $("photoStatus").textContent = "";
    $("copyMessageBtn").disabled = true;
    $("resetMessageBtn").disabled = true;
    $("copyStatus").textContent = "";
  }

  async function copyMessage() {
    const editor = $("messageEditor");
    const html = cleanEditorHtml(editor.innerHTML);
    const plain = editor.innerText.trim();
    if (!plain && !html) return;

    try {
      if (navigator.clipboard?.write && window.ClipboardItem) {
        const item = new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([plain], { type: "text/plain" })
        });
        await navigator.clipboard.write([item]);
        $("copyStatus").textContent = "Copied rich message. Paste into WhatsApp Web or Slack.";
      } else if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(plain);
        $("copyStatus").textContent = "Copied text. Rich image clipboard is unavailable in this browser.";
      } else {
        fallbackCopy(editor);
        $("copyStatus").textContent = "Copied message.";
      }
    } catch (error) {
      console.warn("Rich clipboard failed:", error);
      try {
        await navigator.clipboard.writeText(plain);
        $("copyStatus").textContent = "Copied text. The rich image copy was blocked by the browser.";
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
    const d = value instanceof Date ? new Date(value) : new Date(String(value));
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
