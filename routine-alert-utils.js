(() => {
  "use strict";

  const DAY_NAMES = [
    "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
  ];

  const ROUTINE_HEADERS = [
    "Routine ID", "Active", "Frequency", "Weekday", "Direction", "Client",
    "Asset", "Quantity", "Destination", "Planned Time", "Notes", "Created By", "Created At"
  ];

  const ALERT_HEADERS = [
    "Alert Key", "Created At", "Alert Type", "Status", "Client", "Asset",
    "Required Qty", "Available Qty", "Shortfall", "Destination", "Message",
    "Attended At", "Attended By", "Resolution Comment", "Routine Key",
    "Ignored At", "Ignored By", "Ignore Comment"
  ];

  function normalize(value) {
    return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  }

  function columnIndex(headers, aliases) {
    const wanted = aliases.map(normalize);
    return headers.findIndex(h => wanted.includes(h));
  }

  function number(value) {
    const n = Number(String(value ?? "").replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : 0;
  }

  function normalizeDirection(value) {
    const v = normalize(value).toUpperCase();
    if (v === "RECEIVED" || v === "INBOUND" || v === "IN") return "Inbound";
    if (v === "SENT" || v === "OUTBOUND" || v === "OUT") return "Outbound";
    return "";
  }

  function normalizeFrequency(value) {
    const v = normalize(value);
    if (["daily", "every day", "everyday"].includes(v)) return "Daily";
    if (["weekly", "every week"].includes(v)) return "Weekly";
    return "";
  }

  function parseRoutineRows(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];

    const header = rows[0].map(normalize);
    const idx = {
      id: columnIndex(header, ["routine id", "id"]),
      active: columnIndex(header, ["active", "enabled"]),
      frequency: columnIndex(header, ["frequency", "schedule"]),
      weekday: columnIndex(header, ["weekday", "day"]),
      direction: columnIndex(header, ["direction", "movement"]),
      client: columnIndex(header, ["client", "client name"]),
      asset: columnIndex(header, ["asset", "asset name", "item"]),
      quantity: columnIndex(header, ["quantity", "qty"]),
      destination: columnIndex(header, ["destination", "location", "site"]),
      plannedTime: columnIndex(header, ["planned time", "time"]),
      notes: columnIndex(header, ["notes", "comment"]),
      createdBy: columnIndex(header, ["created by", "user"]),
      createdAt: columnIndex(header, ["created at", "timestamp"])
    };

    return rows.slice(1).map((row, rowOffset) => ({
      rowNumber: rowOffset + 2,
      id: String(idx.id >= 0 ? row[idx.id] ?? "" : `legacy-${rowOffset + 2}`).trim(),
      active: idx.active >= 0 ? isTruthy(row[idx.active]) : true,
      frequency: normalizeFrequency(idx.frequency >= 0 ? row[idx.frequency] : ""),
      weekday: normalizeWeekday(idx.weekday >= 0 ? row[idx.weekday] : ""),
      direction: normalizeDirection(idx.direction >= 0 ? row[idx.direction] : ""),
      client: String(idx.client >= 0 ? row[idx.client] ?? "" : "").trim(),
      asset: String(idx.asset >= 0 ? row[idx.asset] ?? "" : "").trim(),
      quantity: number(idx.quantity >= 0 ? row[idx.quantity] : 0),
      destination: String(idx.destination >= 0 ? row[idx.destination] ?? "" : "").trim(),
      plannedTime: String(idx.plannedTime >= 0 ? row[idx.plannedTime] ?? "" : "").trim(),
      notes: String(idx.notes >= 0 ? row[idx.notes] ?? "" : "").trim(),
      createdBy: String(idx.createdBy >= 0 ? row[idx.createdBy] ?? "" : "").trim(),
      createdAt: String(idx.createdAt >= 0 ? row[idx.createdAt] ?? "" : "").trim()
    })).filter(item => item.client && item.asset && item.quantity > 0 && item.direction);
  }

  function normalizeWeekday(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";
    const lower = raw.toLowerCase();
    const found = DAY_NAMES.find(day => day.toLowerCase() === lower || day.slice(0, 3).toLowerCase() === lower);
    if (found) return found;
    const numeric = Number(raw);
    if (Number.isInteger(numeric) && numeric >= 0 && numeric <= 6) return DAY_NAMES[numeric];
    return "";
  }

  function isTruthy(value) {
    if (value === true || value === 1) return true;
    return ["true", "1", "yes", "y", "on", "active", "enabled"].includes(normalize(value));
  }

  function routineIsDueToday(routine, date = new Date()) {
    if (!routine?.active) return false;
    if (routine.frequency === "Daily") return true;
    if (routine.frequency === "Weekly") return routine.weekday === DAY_NAMES[date.getDay()];
    return false;
  }

  function dueToday(routines, date = new Date()) {
    return routines.filter(r => routineIsDueToday(r, date));
  }

  function routineLabel(routine) {
    if (routine.frequency === "Daily") return "Every day";
    if (routine.frequency === "Weekly" && routine.weekday) return `Every ${routine.weekday}`;
    return routine.frequency || "Routine";
  }

  function routineLoadRows(routines, date = new Date()) {
    return dueToday(routines, date).map(item => ({
      ...item,
      loadType: `Routine · ${routineLabel(item)}`,
      plannedArrival: item.plannedTime || "Routine",
      pallets: /pallet/i.test(item.asset) ? String(item.quantity) : "0",
      loose: /pallet/i.test(item.asset) ? "0" : String(item.quantity),
      quantity: item.quantity,
      destination: item.destination || "",
      direction: item.direction
    }));
  }

  function outboundDemandByAsset(routines, date = new Date()) {
    const map = new Map();
    dueToday(routines, date).filter(r => r.direction === "Outbound").forEach(r => {
      const key = normalize(r.asset);
      const existing = map.get(key) || { asset: r.asset, required: 0, routines: [] };
      existing.required += r.quantity;
      existing.routines.push(r);
      map.set(key, existing);
    });
    return [...map.values()];
  }

  function transactionClientBalances(transactions) {
    const map = new Map();
    (transactions || []).forEach(item => {
      const client = String(item.client || "").trim();
      const asset = String(item.asset || "").trim();
      if (!client || !asset || /hsc london \(self\)/i.test(client)) return;
      const key = `${normalize(client)}||${normalize(asset)}`;
      if (!map.has(key)) map.set(key, { client, asset, sent: 0, received: 0 });
      const entry = map.get(key);
      const qty = number(item.quantity);
      const movement = normalize(item.movement).toUpperCase();
      if (movement === "SENT") entry.sent += qty;
      if (movement === "RECEIVED") entry.received += qty;
    });
    return [...map.values()].map(item => ({ ...item, net: item.sent - item.received }));
  }

  function buildAlertCandidates({ inventory = [], transactions = [], routines = [], clients = [], date = new Date() }) {
    const candidates = [];
    const dayKey = localDateKey(date);
    const activeDue = dueToday(routines, date);

    // Warehouse alerts: zero stock is always relevant. When there is a routine
    // outbound today, compare stock against the full routine demand for that asset.
    const demand = outboundDemandByAsset(routines, date);
    const demandMap = new Map(demand.map(item => [normalize(item.asset), item]));

    inventory.forEach(item => {
      const balance = number(item.balance);
      const required = demandMap.get(normalize(item.asset))?.required || 0;
      if (balance <= 0 || (required > 0 && balance < required)) {
        const shortfall = Math.max(0, required - balance);
        const routinesForAsset = demandMap.get(normalize(item.asset))?.routines || [];
        candidates.push({
          key: `inventory|${dayKey}|${normalize(item.asset)}`,
          type: "INVENTORY",
          client: "",
          asset: item.asset,
          requiredQty: required,
          availableQty: balance,
          shortfall,
          destination: [...new Set(routinesForAsset.map(r => r.destination).filter(Boolean))].join(", "),
          message: required > 0
            ? `${item.asset}: ${formatQty(balance)} available, ${formatQty(required)} required for today's routine outbounds${shortfall ? ` — short by ${formatQty(shortfall)}` : ""}.`
            : `${item.asset}: warehouse inventory is ${formatQty(balance)}.`,
          routineKey: routinesForAsset.map(r => r.id).filter(Boolean).join(",")
        });
      }
    });

    // Client alerts are scoped to active routines so a brand-new client is not
    // flagged for every asset in the warehouse. The current "With clients"
    // balance is sent minus received for that client/asset pair.
    const balances = new Map(transactionClientBalances(transactions).map(x => [
      `${normalize(x.client)}||${normalize(x.asset)}`, x
    ]));

    const clientPairs = new Map();
    activeDue.filter(r => r.direction === "Outbound").forEach(r => {
      const key = `${normalize(r.client)}||${normalize(r.asset)}`;
      clientPairs.set(key, r);
    });

    // Also include historically used client/asset pairs that currently have a
    // balance <= 0, keeping the alert tied to something the ledger knows about.
    transactionClientBalances(transactions).filter(x => x.net <= 0).forEach(x => {
      const key = `${normalize(x.client)}||${normalize(x.asset)}`;
      if (!clientPairs.has(key) && clients.some(c => normalize(c) === normalize(x.client))) {
        clientPairs.set(key, {
          client: x.client, asset: x.asset, direction: "Outbound", destination: ""
        });
      }
    });

    clientPairs.forEach(routine => {
      const key = `${normalize(routine.client)}||${normalize(routine.asset)}`;
      const entry = balances.get(key) || { net: 0, client: routine.client, asset: routine.asset };
      if (number(entry.net) <= 0) {
        candidates.push({
          key: `client|${dayKey}|${normalize(entry.client || routine.client)}|${normalize(entry.asset || routine.asset)}`,
          type: "CLIENT",
          client: entry.client || routine.client,
          asset: entry.asset || routine.asset,
          requiredQty: 0,
          availableQty: number(entry.net),
          shortfall: 0,
          destination: routine.destination || "",
          message: `${entry.client || routine.client} has ${formatQty(entry.net)} ${entry.asset || routine.asset} with client (sent − received).`,
          routineKey: routine.id || ""
        });
      }
    });

    return candidates;
  }

  function parseAlertRows(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];
    const header = rows[0].map(normalize);
    const idx = {
      key: columnIndex(header, ["alert key", "key"]),
      createdAt: columnIndex(header, ["created at", "timestamp"]),
      type: columnIndex(header, ["alert type", "type"]),
      status: columnIndex(header, ["status"]),
      client: columnIndex(header, ["client"]),
      asset: columnIndex(header, ["asset"]),
      requiredQty: columnIndex(header, ["required qty", "required"]),
      availableQty: columnIndex(header, ["available qty", "available"]),
      shortfall: columnIndex(header, ["shortfall", "short by"]),
      destination: columnIndex(header, ["destination", "location"]),
      message: columnIndex(header, ["message", "alert"]),
      attendedAt: columnIndex(header, ["attended at"]),
      attendedBy: columnIndex(header, ["attended by"]),
      resolutionComment: columnIndex(header, ["resolution comment", "resolution note"]),
      routineKey: columnIndex(header, ["routine key"]),
      ignoredAt: columnIndex(header, ["ignored at", "ignore at"]),
      ignoredBy: columnIndex(header, ["ignored by", "ignored user"]),
      ignoreComment: columnIndex(header, ["ignore comment", "ignore note"])
    };

    return rows.slice(1).map((row, offset) => ({
      rowNumber: offset + 2,
      key: String(idx.key >= 0 ? row[idx.key] ?? "" : "").trim(),
      createdAt: String(idx.createdAt >= 0 ? row[idx.createdAt] ?? "" : "").trim(),
      type: String(idx.type >= 0 ? row[idx.type] ?? "" : "").trim().toUpperCase(),
      status: String(idx.status >= 0 ? row[idx.status] ?? "OPEN" : "OPEN").trim().toUpperCase() || "OPEN",
      client: String(idx.client >= 0 ? row[idx.client] ?? "" : "").trim(),
      asset: String(idx.asset >= 0 ? row[idx.asset] ?? "" : "").trim(),
      requiredQty: number(idx.requiredQty >= 0 ? row[idx.requiredQty] : 0),
      availableQty: number(idx.availableQty >= 0 ? row[idx.availableQty] : 0),
      shortfall: number(idx.shortfall >= 0 ? row[idx.shortfall] : 0),
      destination: String(idx.destination >= 0 ? row[idx.destination] ?? "" : "").trim(),
      message: String(idx.message >= 0 ? row[idx.message] ?? "" : "").trim(),
      attendedAt: String(idx.attendedAt >= 0 ? row[idx.attendedAt] ?? "" : "").trim(),
      attendedBy: String(idx.attendedBy >= 0 ? row[idx.attendedBy] ?? "" : "").trim(),
      resolutionComment: String(idx.resolutionComment >= 0 ? row[idx.resolutionComment] ?? "" : "").trim(),
      routineKey: String(idx.routineKey >= 0 ? row[idx.routineKey] ?? "" : "").trim(),
      ignoredAt: String(idx.ignoredAt >= 0 ? row[idx.ignoredAt] ?? "" : "").trim(),
      ignoredBy: String(idx.ignoredBy >= 0 ? row[idx.ignoredBy] ?? "" : "").trim(),
      ignoreComment: String(idx.ignoreComment >= 0 ? row[idx.ignoreComment] ?? "" : "").trim(),
      rawRow: row.slice()
    })).map(item => {
      // Older versions sometimes stored an ignored alert as ATTENDED with
      // "Ignore" in the resolution/comment field. Migrate that legacy shape
      // into the new explicit IGNORED status without losing the original row.
      if (item.status === "ATTENDED" && /^ignore$/i.test(item.resolutionComment)) {
        return {
          ...item,
          status: "IGNORED",
          ignoredAt: item.attendedAt || "",
          ignoredBy: item.attendedBy || "",
          ignoreComment: item.resolutionComment || "",
          attendedAt: "",
          attendedBy: "",
          resolutionComment: ""
        };
      }
      return item;
    }).filter(item => item.key || item.message);
  }

  function localDateKey(date = new Date()) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  }

  function formatQty(value) {
    return Number(value || 0).toLocaleString("en-GB");
  }

  window.FM_ROUTINE_ALERTS = {
    DAY_NAMES,
    ROUTINE_HEADERS,
    ALERT_HEADERS,
    parseRoutineRows,
    routineIsDueToday,
    dueToday,
    routineLabel,
    routineLoadRows,
    outboundDemandByAsset,
    transactionClientBalances,
    buildAlertCandidates,
    parseAlertRows,
    localDateKey,
    normalize,
    normalizeDirection,
    normalizeFrequency,
    normalizeWeekday,
    formatQty
  };
})();
