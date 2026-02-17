/***********************
 * Blind Auction Party - Complete Game Server (Google Sheet DB)
 * Code.gs (V4.0 — full state machine)
 ***********************/

const DB_SPREADSHEET_ID = "1IDUnxLtOWoPOS_ya_FMU3B5yMtjqUWe-NBA4dz9mPCk";

const TABS = {
  ROOMS: "Rooms",
  PLAYERS: "Players",
  EVENTS: "Events",
  ITEMS: "Items",
  CARDS: "PlayerCards"
};

const ROOM_ACTIVE_MS = 15 * 60 * 1000;
const REVEAL_DURATION_MS = 10 * 1000;
const POSTROUND_DURATION_MS = 8 * 1000;

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = (params.action || "").trim();

  if (action) {
    try {
      const out = handleApiGet_(action, params);
      return apiOut_(out, params);
    } catch (err) {
      return apiOut_({ ok: false, error: String(err && err.message ? err.message : err) }, params);
    }
  }

  const tpl = HtmlService.createTemplateFromFile("index");
  tpl.apiBaseUrl = ScriptApp.getService().getUrl();

  const html = tpl.evaluate()
    .setTitle("Blind Auction Party")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return html;
}

function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) ? e.postData.contents : "";
    let payload = {};
    try {
      payload = body ? JSON.parse(body) : {};
    } catch (_) {
      return jsonOut_({ ok: false, error: "POST_BODY_NOT_JSON", bodyPreview: String(body).slice(0, 200) });
    }

    const action = (payload.action || "").trim();
    if (!action) return jsonOut_({ ok: false, error: "MISSING_ACTION" });

    const out = handleApiPost_(action, payload);
    return jsonOut_(out);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/***************
 * API Routing
 ***************/
function handleApiGet_(action, p) {
  switch (action) {
    case "debugInfo": return apiDebugInfo_();
    case "listLobbyRooms": return apiListLobbyRooms_(p);
    case "sync": return apiSync_(p);
    case "setup": return apiSetup_();
    case "createRoom": return apiCreateRoom_(p);
    case "joinRoom": return apiJoinRoom_(p);
    case "updateSettings": return apiUpdateSettings_(p);
    case "startRound": return apiStartRound_(p);
    case "bid": return apiBid_(p);
    case "resolve": return apiResolve_(p);
    case "playCard": return apiPlayCard_(p);
    case "resetDatabase": return apiResetDatabase_(p);
    case "nextRound": return apiNextRound_(p);
    default: return { ok: false, error: "UNKNOWN_GET_ACTION:" + action };
  }
}

function handleApiPost_(action, body) {
  switch (action) {
    case "setup": return apiSetup_();
    case "createRoom": return apiCreateRoom_(body);
    case "joinRoom": return apiJoinRoom_(body);
    case "updateSettings": return apiUpdateSettings_(body);
    case "startRound": return apiStartRound_(body);
    case "bid": return apiBid_(body);
    case "resolve": return apiResolve_(body);
    case "playCard": return apiPlayCard_(body);
    case "resetDatabase": return apiResetDatabase_(body);
    case "nextRound": return apiNextRound_(body);
    default: return { ok: false, error: "UNKNOWN_POST_ACTION:" + action };
  }
}

/***************
 * Core helpers
 ***************/
function db_() {
  const id = String(DB_SPREADSHEET_ID || "").trim();
  if (!id || id === "PASTE_YOUR_SPREADSHEET_ID_HERE") throw new Error("DB_SPREADSHEET_ID_NOT_SET");
  return SpreadsheetApp.openById(id);
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function apiOut_(obj, params) {
  const callback = String((params && params.callback) || "").trim();
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$\.]{0,63}$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(obj) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOut_(obj);
}

function clearSheetDataRows_(sh, width) {
  const last = sh.getLastRow();
  if (last <= 1) return;
  sh.getRange(2, 1, last - 1, width).clearContent();
}

function apiResetDatabase_(body) {
  const roomId = String(body.roomId || "").trim().toUpperCase();
  const hostToken = String(body.hostToken || "").trim();
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const im = idxMap_(headers);
    const row = getRow_(rooms, idx, rooms.getLastColumn());
    if (String(row[im.hostToken]) !== hostToken) return { ok: false, error: "HOST_TOKEN_MISMATCH" };

    [TABS.ROOMS, TABS.PLAYERS, TABS.EVENTS, TABS.CARDS].forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      clearSheetDataRows_(sh, sh.getLastColumn());
    });

    const items = ss.getSheetByName(TABS.ITEMS);
    if (items) {
      clearSheetDataRows_(items, items.getLastColumn());
      seedItems_(items);
    }

    return { ok: true };
  });
}

function nowMs_() { return Date.now(); }

function rid_(len) {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let s = "";
  for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

function normalizeLobbyCode_(raw) {
  let s = String(raw || "").trim();
  if (!s) return "";
  s = s.replace(/\s+/g, "").replace(/[a-z]/g, c => c.toUpperCase());
  return s.length >= 2 ? s.slice(0, 16) : "";
}

function newUniqueRoomId_(roomsSheet) {
  for (let i = 0; i < 50; i++) {
    const id = String(Math.floor(Math.random() * 10000)).padStart(4, "0");
    if (findRoomRow_(roomsSheet, id) < 0) return id;
  }
  throw new Error("ROOM_ID_POOL_EXHAUSTED");
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const a1 = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const empty = a1.every(v => String(v || "").trim() === "");
  if (empty) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function readAll_(sh) {
  const rng = sh.getDataRange();
  const values = rng.getValues();
  if (values.length <= 1) return { headers: values[0] || [], rows: [] };
  return { headers: values[0], rows: values.slice(1) };
}

function idxMap_(headers) {
  const m = {};
  headers.forEach((h, i) => m[String(h)] = i);
  return m;
}

function appendRow_(sh, row) {
  sh.appendRow(row);
}

function findRowIndex_(sh, colIndex, value) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, colIndex + 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(value)) return i + 2;
  }
  return -1;
}

function getRow_(sh, rowIndex, width) {
  return sh.getRange(rowIndex, 1, 1, width).getValues()[0];
}

function setRow_(sh, rowIndex, row) {
  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function seedItems_(itemsSheet) {
  const seed = [
    ["ITM1", "神秘古董", "看起來很值錢，但可能是垃圾", 30],
    ["ITM2", "限量公仔", "大家都說會漲價", 25],
    ["ITM3", "二手筆電", "螢幕有刮痕，但能用", 18],
    ["ITM4", "奇怪的箱子", "搖一搖會響", 22],
    ["ITM5", "無敵券", "規則外的力量感", 35],
    ["ITM6", "金元寶", "閃閃發光的黃金", 40],
    ["ITM7", "紅包", "裡面會放多少呢？", 15],
    ["ITM8", "古董花瓶", "據說是明朝的…", 28]
  ];
  itemsSheet.getRange(2, 1, seed.length, 4).setValues(seed);
}

/***************
 * Setup + Debug
 ***************/
function setup() {
  return apiSetup_();
}

function apiSetup_() {
  return withLock_(() => {
    const ss = db_();

    ensureSheet_(ss, TABS.ROOMS, [
      "roomId", "lobbyCode", "createdAt", "updatedAt", "state", "round", "maxRounds",
      "roundSeconds", "startingBudget", "hostToken", "bidDeadlineTs", "itemId",
      "revealUntilTs", "postUntilTs"
    ]);

    ensureSheet_(ss, TABS.PLAYERS, [
      "roomId", "playerId", "name", "isHost", "budget", "score", "joinedAt", "updatedAt"
    ]);

    ensureSheet_(ss, TABS.EVENTS, [
      "roomId", "eventId", "ts", "type", "playerId", "payloadJson"
    ]);

    ensureSheet_(ss, TABS.ITEMS, [
      "itemId", "name", "publicHint", "hiddenValue"
    ]);

    ensureSheet_(ss, TABS.CARDS, [
      "roomId", "playerId", "cardId", "name", "desc", "count"
    ]);

    // seed items if empty
    const items = ss.getSheetByName(TABS.ITEMS);
    if (items.getLastRow() < 2) {
      seedItems_(items);
    }

    return { ok: true };
  });
}

function apiDebugInfo_() {
  const ss = db_();
  return { ok: true, info: { name: ss.getName(), id: ss.getId(), url: ss.getUrl(), time: new Date().toISOString() } };
}

/***************
 * Game APIs
 ***************/
function apiCreateRoom_(body) {
  const lobbyCode = normalizeLobbyCode_(body.lobbyCode);
  if (!lobbyCode) return { ok: false, error: "MISSING_LOBBYCODE" };

  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const t = nowMs_();

    const roomId = newUniqueRoomId_(rooms);
    const hostToken = rid_(10);

    appendRow_(rooms, [
      roomId, lobbyCode, t, t,
      "LOBBY", 0, 8, 25, 100,
      hostToken, "", "",
      "", ""  // revealUntilTs, postUntilTs
    ]);

    addEvent_(roomId, "ROOM_CREATED", "", { lobbyCode });

    return { ok: true, roomId, hostToken };
  });
}

function apiJoinRoom_(body) {
  const roomId = String(body.roomId || "").trim();
  const name = String(body.name || "").trim().slice(0, 18);
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!name) return { ok: false, error: "MISSING_NAME" };

  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const players = ss.getSheetByName(TABS.PLAYERS);

    const roomRowIndex = findRoomRow_(rooms, roomId);
    if (roomRowIndex < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const t = nowMs_();
    const { headers, rows } = readAll_(players);
    const im = idxMap_(headers);
    let count = 0;
    rows.forEach(r => { if (String(r[im.roomId]) === roomId) count++; });

    const playerId = newUniquePlayerId_(rows, im);
    const roomRow = getRow_(rooms, roomRowIndex, rooms.getLastColumn());
    const roomHeaders = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(roomHeaders);

    const isHost = (count === 0);
    const startingBudget = Number(roomRow[rim.startingBudget] || 100);

    appendRow_(players, [
      roomId, playerId, name, isHost ? "1" : "0",
      startingBudget, 0, t, t
    ]);

    seedCards_(ss, roomId, playerId);

    roomRow[rim.updatedAt] = t;
    setRow_(rooms, roomRowIndex, roomRow);

    addEvent_(roomId, "PLAYER_JOINED", playerId, { name, isHost });

    return { ok: true, playerId };
  });
}

function apiListLobbyRooms_(p) {
  const lobbyCode = normalizeLobbyCode_(p.lobbyCode);
  if (!lobbyCode) return { ok: false, error: "MISSING_LOBBYCODE" };

  const ss = db_();
  const rooms = ss.getSheetByName(TABS.ROOMS);
  const players = ss.getSheetByName(TABS.PLAYERS);

  const rAll = readAll_(rooms);
  const rim = idxMap_(rAll.headers);

  const pAll = readAll_(players);
  const pim = idxMap_(pAll.headers);

  const t = nowMs_();
  const out = [];

  rAll.rows.forEach(r => {
    if (String(r[rim.lobbyCode]) !== lobbyCode) return;
    const updatedAt = Number(r[rim.updatedAt] || 0);
    if (t - updatedAt > ROOM_ACTIVE_MS) return;

    const roomId = String(r[rim.roomId]);
    let pc = 0;
    pAll.rows.forEach(pr => { if (String(pr[pim.roomId]) === roomId) pc++; });

    out.push({
      roomId,
      state: String(r[rim.state]),
      round: Number(r[rim.round] || 0),
      maxRounds: Number(r[rim.maxRounds] || 0),
      playerCount: pc
    });
  });

  return { ok: true, rooms: out };
}

/***************
 * Sync — with auto state transitions
 ***************/
function apiSync_(p) {
  const roomId = String(p.roomId || "").trim().toUpperCase();
  const sinceEventId = Number(p.sinceEventId || 0);
  const playerId = String(p.playerId || "").trim();
  const hostToken = String(p.hostToken || "").trim();

  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };

  // Try auto-transitions (needs lock)
  tryAutoTransition_(roomId);

  const ss = db_();
  const rooms = ss.getSheetByName(TABS.ROOMS);
  const players = ss.getSheetByName(TABS.PLAYERS);
  const events = ss.getSheetByName(TABS.EVENTS);
  const items = ss.getSheetByName(TABS.ITEMS);
  const cards = ss.getSheetByName(TABS.CARDS);

  const roomRowIndex = findRoomRow_(rooms, roomId);
  if (roomRowIndex < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

  const roomHeaders = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
  const rim = idxMap_(roomHeaders);
  const rr = getRow_(rooms, roomRowIndex, rooms.getLastColumn());

  // build room obj
  let item = null;
  const itemId = String(rr[rim.itemId] || "");
  if (itemId) {
    const it = findItem_(items, itemId);
    if (it) {
      // Only expose public info (not hiddenValue) unless in REVEAL/POSTROUND/ENDED
      const state = String(rr[rim.state]);
      if (state === "REVEAL" || state === "POSTROUND" || state === "ENDED") {
        item = it;
      } else {
        item = { itemId: it.itemId, name: it.name, publicHint: it.publicHint };
      }
    }
  }

  const room = {
    roomId: String(rr[rim.roomId]),
    lobbyCode: String(rr[rim.lobbyCode]),
    state: String(rr[rim.state]),
    round: Number(rr[rim.round] || 0),
    maxRounds: Number(rr[rim.maxRounds] || 0),
    roundSeconds: Number(rr[rim.roundSeconds] || 25),
    startingBudget: Number(rr[rim.startingBudget] || 100),
    bidDeadlineTs: Number(rr[rim.bidDeadlineTs] || 0),
    revealUntilTs: Number(rr[rim.revealUntilTs] || 0),
    postUntilTs: Number(rr[rim.postUntilTs] || 0),
    item
  };

  // players list
  const pAll = readAll_(players);
  const pim = idxMap_(pAll.headers);
  const plist = [];
  pAll.rows.forEach(pr => {
    if (String(pr[pim.roomId]) !== roomId) return;
    plist.push({
      roomId,
      playerId: String(pr[pim.playerId]),
      name: String(pr[pim.name]),
      isHost: String(pr[pim.isHost]) === "1",
      budget: Number(pr[pim.budget] || 0),
      score: Number(pr[pim.score] || 0)
    });
  });

  // my cards
  const cAll = readAll_(cards);
  const cim = idxMap_(cAll.headers);
  const myCards = [];
  cAll.rows.forEach(cr => {
    if (String(cr[cim.roomId]) !== roomId) return;
    if (String(cr[cim.playerId]) !== playerId) return;
    myCards.push({
      cardId: String(cr[cim.cardId]),
      name: String(cr[cim.name]),
      desc: String(cr[cim.desc]),
      count: Number(cr[cim.count] || 0)
    });
  });

  // Peek card check: if player used Peek this round, show hiddenValue hint
  let peekHint = null;
  if (room.state === "BIDDING" && playerId && room.item) {
    const round = room.round;
    const eAll = readAll_(events);
    const eim = idxMap_(eAll.headers);
    for (const er of eAll.rows) {
      if (String(er[eim.roomId]) !== roomId) continue;
      if (String(er[eim.type]) !== "CARD_PLAYED") continue;
      if (String(er[eim.playerId]) !== playerId) continue;
      let payload = {};
      try { payload = JSON.parse(String(er[eim.payloadJson] || "{}")); } catch (_) { }
      if (Number(payload.round) === round && payload.cardId === "C1") {
        // Peek: show ±5 range
        const fullItem = findItem_(items, itemId);
        if (fullItem) {
          const hv = fullItem.hiddenValue;
          peekHint = { min: Math.max(0, hv - 5), max: hv + 5 };
        }
        break;
      }
    }
  }

  // events incremental
  const eAll2 = readAll_(events);
  const eim2 = idxMap_(eAll2.headers);
  const evs = [];
  eAll2.rows.forEach(er => {
    if (String(er[eim2.roomId]) !== roomId) return;
    const eid = Number(er[eim2.eventId] || 0);
    if (eid <= sinceEventId) return;
    evs.push({
      roomId,
      eventId: eid,
      ts: Number(er[eim2.ts] || 0),
      type: String(er[eim2.type]),
      playerId: String(er[eim2.playerId] || ""),
      payloadJson: String(er[eim2.payloadJson] || "{}")
    });
  });

  // Find latest ROUND_RESOLVED for current or last round
  let resolveResult = null;
  for (let i = evs.length - 1; i >= 0; i--) {
    if (evs[i].type === "ROUND_RESOLVED") {
      try { resolveResult = JSON.parse(evs[i].payloadJson); } catch (_) { }
      break;
    }
  }
  // Also search in older events if not found
  if (!resolveResult && (room.state === "REVEAL" || room.state === "POSTROUND")) {
    eAll2.rows.forEach(er => {
      if (String(er[eim2.roomId]) !== roomId) return;
      if (String(er[eim2.type]) !== "ROUND_RESOLVED") return;
      try {
        const p = JSON.parse(String(er[eim2.payloadJson] || "{}"));
        if (Number(p.round) === room.round) resolveResult = p;
      } catch (_) { }
    });
  }

  // Check current player's bid for this round
  let myBid = null;
  if (room.state === "BIDDING" && playerId) {
    eAll2.rows.forEach(er => {
      if (String(er[eim2.roomId]) !== roomId) return;
      if (String(er[eim2.type]) !== "BID") return;
      if (String(er[eim2.playerId]) !== playerId) return;
      let payload = {};
      try { payload = JSON.parse(String(er[eim2.payloadJson] || "{}")); } catch (_) { }
      if (Number(payload.round) === room.round) {
        myBid = Number(payload.bid || 0);
      }
    });
  }

  const isHost = hostToken && (hostToken === String(rr[rim.hostToken]));

  return { ok: true, room, players: plist, myCards, events: evs, isHost, resolveResult, peekHint, myBid };
}

/***************
 * Auto state transitions (time-driven)
 ***************/
function tryAutoTransition_(roomId) {
  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx = findRoomRow_(rooms, roomId);
    if (idx < 0) return;

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(headers);
    const r = getRow_(rooms, idx, rooms.getLastColumn());
    const state = String(r[rim.state]);
    const now = nowMs_();

    // BIDDING → REVEAL (auto-resolve when deadline passed)
    if (state === "BIDDING") {
      const deadline = Number(r[rim.bidDeadlineTs] || 0);
      if (deadline > 0 && now >= deadline) {
        doResolve_(ss, rooms, idx, r, rim, roomId);
        return;
      }
    }

    // REVEAL → POSTROUND
    if (state === "REVEAL") {
      const revealUntil = Number(r[rim.revealUntilTs] || 0);
      if (revealUntil > 0 && now >= revealUntil) {
        const round = Number(r[rim.round] || 0);
        const maxRounds = Number(r[rim.maxRounds] || 8);
        if (round >= maxRounds) {
          r[rim.state] = "ENDED";
          r[rim.postUntilTs] = "";
          addEvent_(roomId, "GAME_ENDED", "", { round });
        } else {
          r[rim.state] = "POSTROUND";
          r[rim.postUntilTs] = now + POSTROUND_DURATION_MS;
          addEvent_(roomId, "STATE_TRANSITION", "", { from: "REVEAL", to: "POSTROUND" });
        }
        r[rim.updatedAt] = now;
        setRow_(rooms, idx, r);
        return;
      }
    }

    // POSTROUND → LOBBY
    if (state === "POSTROUND") {
      const postUntil = Number(r[rim.postUntilTs] || 0);
      if (postUntil > 0 && now >= postUntil) {
        r[rim.state] = "LOBBY";
        r[rim.itemId] = "";
        r[rim.bidDeadlineTs] = "";
        r[rim.revealUntilTs] = "";
        r[rim.postUntilTs] = "";
        r[rim.updatedAt] = now;
        setRow_(rooms, idx, r);
        addEvent_(roomId, "STATE_TRANSITION", "", { from: "POSTROUND", to: "LOBBY" });
        return;
      }
    }
  });
}

/***************
 * Core resolve logic (shared by manual and auto)
 ***************/
function doResolve_(ss, rooms, idx, r, rim, roomId) {
  const players = ss.getSheetByName(TABS.PLAYERS);
  const events = ss.getSheetByName(TABS.EVENTS);
  const items = ss.getSheetByName(TABS.ITEMS);
  const cards = ss.getSheetByName(TABS.CARDS);

  const round = Number(r[rim.round] || 0);
  const itemId = String(r[rim.itemId] || "");
  const item = itemId ? findItem_(items, itemId) : null;
  const hiddenValue = item ? Number(item.hiddenValue || 0) : 0;

  // read players
  const pAll = readAll_(players);
  const pim = idxMap_(pAll.headers);

  const pIndexById = {};
  pAll.rows.forEach((pr, i) => {
    if (String(pr[pim.roomId]) !== roomId) return;
    pIndexById[String(pr[pim.playerId])] = i + 2;
  });

  // gather latest bid per player for this round
  const eAll = readAll_(events);
  const eim = idxMap_(eAll.headers);
  const lastBid = {};
  const cardPlayed = {}; // playerId -> [cardId, ...]
  eAll.rows.forEach(er => {
    if (String(er[eim.roomId]) !== roomId) return;
    const type = String(er[eim.type]);
    const pid = String(er[eim.playerId] || "");
    let payload = {};
    try { payload = JSON.parse(String(er[eim.payloadJson] || "{}")); } catch (_) { }

    if (type === "BID" && Number(payload.round) === round) {
      lastBid[pid] = Number(payload.bid || 0);
    }
    if (type === "CARD_PLAYED" && Number(payload.round) === round) {
      if (!cardPlayed[pid]) cardPlayed[pid] = [];
      cardPlayed[pid].push(String(payload.cardId || ""));
    }
  });

  // find winner: highest bid that <= budget
  let winnerId = "";
  let winnerBid = -1;

  Object.keys(lastBid).forEach(pid => {
    const bid = Number(lastBid[pid]);
    const prowIdx = pIndexById[pid];
    if (!prowIdx) return;

    const row = getRow_(players, prowIdx, players.getLastColumn());
    const budget = Number(row[pim.budget] || 0);
    if (bid > budget) return;
    if (bid > winnerBid) { winnerBid = bid; winnerId = pid; }
  });

  // Apply Tax card: winner pays 20% more
  let actualCost = winnerBid;
  let taxApplied = false;
  if (winnerId && cardPlayed[winnerId] && cardPlayed[winnerId].includes("C2")) {
    actualCost = Math.ceil(winnerBid * 1.2);
    taxApplied = true;
  }

  // Calculate score deltas for all players
  const deltas = {};
  let winnerName = "";

  if (winnerId) {
    const prowIdx = pIndexById[winnerId];
    const row = getRow_(players, prowIdx, players.getLastColumn());
    const budget = Number(row[pim.budget] || 0);
    const score = Number(row[pim.score] || 0);
    winnerName = String(row[pim.name] || "");

    let delta = hiddenValue - actualCost;

    // Apply Shield card: if delta < 0 and winner has Shield, clamp to 0
    let shieldApplied = false;
    if (delta < 0 && cardPlayed[winnerId] && cardPlayed[winnerId].includes("C3")) {
      delta = 0;
      shieldApplied = true;
    }

    deltas[winnerId] = { delta, shieldApplied, taxApplied };

    row[pim.budget] = budget - actualCost;
    row[pim.score] = score + delta;
    row[pim.updatedAt] = nowMs_();
    setRow_(players, prowIdx, row);
  }

  const now = nowMs_();
  // Transition to REVEAL
  r[rim.state] = "REVEAL";
  r[rim.revealUntilTs] = now + REVEAL_DURATION_MS;
  r[rim.bidDeadlineTs] = "";
  r[rim.updatedAt] = now;
  setRow_(rooms, idx, r);

  addEvent_(roomId, "ROUND_RESOLVED", "", {
    round, winnerId, winnerName, winnerBid, actualCost,
    hiddenValue, deltas, taxApplied,
    itemName: item ? item.name : ""
  });
}

function apiResolve_(body) {
  const roomId = String(body.roomId || "").trim().toUpperCase();
  const hostToken = String(body.hostToken || "").trim();
  const playerId = String(body.playerId || "").trim();
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken && !playerId) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);

    const idx = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(headers);
    const r = getRow_(rooms, idx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, r, rim)) return { ok: false, error: "HOST_PERMISSION_DENIED" };
    if (String(r[rim.state]) !== "BIDDING") return { ok: false, error: "NOT_IN_BIDDING" };

    doResolve_(ss, rooms, idx, r, rim, roomId);

    return { ok: true };
  });
}

function apiUpdateSettings_(body) {
  const roomId = String(body.roomId || "").trim().toUpperCase();
  const hostToken = String(body.hostToken || "").trim();
  const playerId = String(body.playerId || "").trim();
  const startingBudget = Number(body.startingBudget || 100);
  const roundSeconds = Number(body.roundSeconds || 25);
  const maxRounds = Number(body.maxRounds || 8);

  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken && !playerId) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);

    const idx = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const im = idxMap_(headers);
    const r = getRow_(rooms, idx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, r, im)) return { ok: false, error: "HOST_PERMISSION_DENIED" };

    const t = nowMs_();
    r[im.startingBudget] = startingBudget;
    r[im.roundSeconds] = roundSeconds;
    r[im.maxRounds] = maxRounds;
    r[im.updatedAt] = t;
    setRow_(rooms, idx, r);

    addEvent_(roomId, "SETTINGS_UPDATED", "", { startingBudget, roundSeconds, maxRounds });
    return { ok: true };
  });
}

function apiStartRound_(body) {
  const roomId = String(body.roomId || "").trim().toUpperCase();
  const hostToken = String(body.hostToken || "").trim();
  const playerId = String(body.playerId || "").trim();
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken && !playerId) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const items = ss.getSheetByName(TABS.ITEMS);

    const idx = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const im = idxMap_(headers);
    const r = getRow_(rooms, idx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, r, im)) return { ok: false, error: "HOST_PERMISSION_DENIED" };

    const state = String(r[im.state]);
    if (state !== "LOBBY" && state !== "POSTROUND") return { ok: false, error: "NOT_IN_LOBBY_OR_POSTROUND" };

    const round = Number(r[im.round] || 0) + 1;
    const maxRounds = Number(r[im.maxRounds] || 8);
    if (round > maxRounds) return { ok: false, error: "MAX_ROUNDS_REACHED" };

    const it = pickRandomItem_(items);
    if (!it) return { ok: false, error: "NO_ITEMS" };

    const roundSeconds = Number(r[im.roundSeconds] || 25);
    const deadline = nowMs_() + Math.max(5, roundSeconds) * 1000;

    r[im.round] = round;
    r[im.state] = "BIDDING";
    r[im.itemId] = it.itemId;
    r[im.bidDeadlineTs] = deadline;
    r[im.revealUntilTs] = "";
    r[im.postUntilTs] = "";
    r[im.updatedAt] = nowMs_();
    setRow_(rooms, idx, r);

    addEvent_(roomId, "ROUND_STARTED", "", { round, itemId: it.itemId, itemName: it.name, publicHint: it.publicHint, deadline });
    return { ok: true };
  });
}

function apiNextRound_(body) {
  // Host manually triggers next round from POSTROUND
  return apiStartRound_(body);
}

function apiBid_(body) {
  const roomId = String(body.roomId || "").trim().toUpperCase();
  const playerId = String(body.playerId || "").trim();
  const round = Number(body.round || 0);
  const bid = Number(body.bid || 0);

  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!playerId) return { ok: false, error: "MISSING_PLAYERID" };
  if (!Number.isFinite(round) || round <= 0) return { ok: false, error: "BAD_ROUND" };
  if (!Number.isFinite(bid) || bid < 0) return { ok: false, error: "BAD_BID" };

  return withLock_(() => {
    addEvent_(roomId, "BID", playerId, { round, bid });
    touchRoomUpdatedAt_(roomId);
    return { ok: true };
  });
}

function apiPlayCard_(body) {
  const roomId = String(body.roomId || "").trim().toUpperCase();
  const playerId = String(body.playerId || "").trim();
  const round = Number(body.round || 0);
  const cardId = String(body.cardId || "").trim();

  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!playerId) return { ok: false, error: "MISSING_PLAYERID" };
  if (!cardId) return { ok: false, error: "MISSING_CARDID" };

  return withLock_(() => {
    const ss = db_();
    const cards = ss.getSheetByName(TABS.CARDS);

    const all = readAll_(cards);
    const im = idxMap_(all.headers);

    for (let i = 0; i < all.rows.length; i++) {
      const r = all.rows[i];
      if (String(r[im.roomId]) !== roomId) continue;
      if (String(r[im.playerId]) !== playerId) continue;
      if (String(r[im.cardId]) !== cardId) continue;

      const sheetRow = i + 2;
      const row = getRow_(cards, sheetRow, cards.getLastColumn());
      const cnt = Number(row[im.count] || 0);
      if (cnt <= 0) return { ok: false, error: "CARD_EMPTY" };
      row[im.count] = cnt - 1;
      setRow_(cards, sheetRow, row);

      addEvent_(roomId, "CARD_PLAYED", playerId, { round, cardId });
      touchRoomUpdatedAt_(roomId);
      return { ok: true };
    }

    return { ok: false, error: "CARD_NOT_FOUND" };
  });
}

/***************
 * Internal helpers
 ***************/

function canManageRoom_(ss, roomId, hostToken, playerId, roomRow, roomIndexMap) {
  if (hostToken && String(roomRow[roomIndexMap.hostToken]) === hostToken) return true;
  if (!playerId) return false;

  const players = ss.getSheetByName(TABS.PLAYERS);
  if (!players) return false;

  const all = readAll_(players);
  const pim = idxMap_(all.headers);
  for (const pr of all.rows) {
    if (String(pr[pim.roomId]) !== roomId) continue;
    if (String(pr[pim.playerId]) !== playerId) continue;
    return String(pr[pim.isHost]) === "1";
  }

  return false;
}

function findRoomRow_(roomsSheet, roomId) {
  return findRowIndex_(roomsSheet, 0, roomId);
}

function addEvent_(roomId, type, playerId, payloadObj) {
  const ss = db_();
  const events = ss.getSheetByName(TABS.EVENTS);
  const t = nowMs_();

  const lastRow = events.getLastRow();
  let nextId = 1;
  if (lastRow >= 2) {
    const lastVal = Number(events.getRange(lastRow, 2).getValue() || 0);
    nextId = lastVal + 1;
  }

  appendRow_(events, [
    roomId, nextId, t, type, playerId || "", JSON.stringify(payloadObj || {})
  ]);
}

function touchRoomUpdatedAt_(roomId) {
  const ss = db_();
  const rooms = ss.getSheetByName(TABS.ROOMS);
  const idx = findRoomRow_(rooms, roomId);
  if (idx < 0) return;
  const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
  const im = idxMap_(headers);
  const r = getRow_(rooms, idx, rooms.getLastColumn());
  r[im.updatedAt] = nowMs_();
  setRow_(rooms, idx, r);
}

function pickRandomItem_(itemsSheet) {
  const all = readAll_(itemsSheet);
  if (all.rows.length === 0) return null;
  const i = Math.floor(Math.random() * all.rows.length);
  const r = all.rows[i];
  const im = idxMap_(all.headers);
  return {
    itemId: String(r[im.itemId]),
    name: String(r[im.name]),
    publicHint: String(r[im.publicHint]),
    hiddenValue: Number(r[im.hiddenValue] || 0)
  };
}

function findItem_(itemsSheet, itemId) {
  const all = readAll_(itemsSheet);
  const im = idxMap_(all.headers);
  for (const r of all.rows) {
    if (String(r[im.itemId]) === itemId) {
      return {
        itemId: String(r[im.itemId]),
        name: String(r[im.name]),
        publicHint: String(r[im.publicHint]),
        hiddenValue: Number(r[im.hiddenValue] || 0)
      };
    }
  }
  return null;
}

function seedCards_(ss, roomId, playerId) {
  const cards = ss.getSheetByName(TABS.CARDS);
  const seed = [
    [roomId, playerId, "C1", "偷看 Peek", "窺探物品真實價值的區間（±5）", 1],
    [roomId, playerId, "C2", "加稅 Tax", "若你得標，多付 20%（心理戰用）", 1],
    [roomId, playerId, "C3", "護盾 Shield", "本回合若你分數為負，抵消一次", 1]
  ];
  cards.getRange(cards.getLastRow() + 1, 1, seed.length, 6).setValues(seed);
}

function newUniquePlayerId_(playerRows, playerIndexMap) {
  const used = new Set();
  playerRows.forEach(r => {
    const pid = String(r[playerIndexMap.playerId] || "").trim();
    if (pid) used.add(pid);
  });

  for (let i = 0; i < 20; i++) {
    const id = "P" + rid_(6);
    if (!used.has(id)) return id;
  }

  let fallback = "P" + rid_(6) + String(nowMs_()).slice(-4);
  while (used.has(fallback)) {
    fallback = "P" + rid_(6) + String(nowMs_()).slice(-4);
  }
  return fallback;
}

/***************
 * Manual debug helper
 ***************/
function showBindingInfo() {
  const ss = db_();
  const info = { name: ss.getName(), id: ss.getId(), url: ss.getUrl(), time: new Date().toISOString() };

  const sheetName = "__DEBUG__";
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  sh.getRange("A1").setValue("BOUND TO (DB): " + info.name);
  sh.getRange("A2").setValue(info.id);
  sh.getRange("A3").setValue(info.url);
  sh.getRange("A4").setValue(info.time);
  SpreadsheetApp.flush();

  console.log("DB_INFO=" + JSON.stringify(info));
  return info;
}
