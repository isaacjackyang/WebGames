/**
 * Blind Auction Party V3.1b
 * - Force DB Spreadsheet by ID (no more "I ran it but sheet didn't change")
 * - Lobby Code: list rooms in same "party space"
 * - QR Join link (generated on client)
 * - Adaptive polling on client
 * - Incremental event read: eventId roughly maps to row (fast range fetch)
 */

// =====================
// IMPORTANT: SET THIS
// =====================
const DB_SPREADSHEET_ID = "1IDUnxLtOWoPOS_ya_FMU3B5yMtjqUWe-NBA4dz9mPCk";

function db_() {
  if (!DB_SPREADSHEET_ID || DB_SPREADSHEET_ID.includes("PASTE_YOUR_SPREADSHEET_ID_HERE")) {
    throw new Error("DB_SPREADSHEET_ID_NOT_SET");
  }
  return SpreadsheetApp.openById(DB_SPREADSHEET_ID);
}

const SHEET_ROOMS = "Rooms";
const SHEET_PLAYERS = "Players";
const SHEET_ITEMS = "Items";
const SHEET_EVENTS = "Events";
const SHEET_PLAYER_CARDS = "PlayerCards";

const STATE_LOBBY = "LOBBY";
const STATE_BIDDING = "BIDDING";
const STATE_RESOLVING = "RESOLVING";
const STATE_ENDED = "ENDED";

const DEFAULT_BUDGET = 100;
const DEFAULT_ROUND_SECONDS = 25;
const DEFAULT_MAX_ROUNDS = 8;

const CARD_BOOST = "BOOST";
const CARD_CAP20 = "CAP20";

// Lobby listing rules
const LOBBY_ACTIVE_WINDOW_MS = 3 * 60 * 1000; // rooms active within 3 minutes are shown
const ROOM_HEARTBEAT_THROTTLE_MS = 15 * 1000; // host heartbeat write throttle

function doGet(e) {
  const action = e?.parameter?.action ? String(e.parameter.action) : "";
  if (!action) {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("Blind Auction Party V3.1b");
  }
  return handleApi_(action, e, "GET");
}

function doPost(e) {
  let body = {};
  try {
    body = e?.postData?.contents ? JSON.parse(e.postData.contents) : {};
  } catch (_) {
    return jsonOut_({ ok: false, error: "INVALID_JSON" });
  }
  const action = body?.action ? String(body.action) : "";
  if (!action) return jsonOut_({ ok: false, error: "MISSING_ACTION" });
  return handleApi_(action, { body }, "POST");
}

/**
 * Run once: create sheets / migrate headers / seed items
 */
function setup() {
  const ss = db_();

  ensureSheet_(ss, SHEET_ROOMS, [
    "roomId","hostToken","lobbyCode","state","round","itemId","bidDeadlineTs","seed","createdTs",
    "startingBudget","roundSeconds","maxRounds","capThisRound","lastActiveTs"
  ]);

  ensureSheet_(ss, SHEET_PLAYERS, [
    "roomId","playerId","name","budget","score","joinedTs","lastSeenTs"
  ]);

  ensureSheet_(ss, SHEET_ITEMS, [
    "itemId","name","publicHint","trueValue"
  ]);

  ensureSheet_(ss, SHEET_EVENTS, [
    "eventId","ts","roomId","playerId","type","payloadJson"
  ]);

  ensureSheet_(ss, SHEET_PLAYER_CARDS, [
    "roomId","playerId","cardId","count"
  ]);

  seedItemsIfEmpty_();
}

/**
 * Diagnostic helper: writes DB info into A1~A4 of the first sheet in DB spreadsheet.
 */
function showBindingInfo() {
  const ss = db_();
  const info = {
    name: ss.getName(),
    id: ss.getId(),
    url: ss.getUrl(),
    time: new Date().toISOString()
  };

  // 1) 一定要能看到：強制建立/取得 __DEBUG__ 分頁
  const sheetName = "__DEBUG__";
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  // 2) 強制寫入一串明顯文字
  sh.getRange("A1").setValue("BOUND TO (DB): " + info.name);
  sh.getRange("A2").setValue(info.id);
  sh.getRange("A3").setValue(info.url);
  sh.getRange("A4").setValue(info.time);

  // 3) 強制 flush，避免寫入延遲造成你以為沒寫
  SpreadsheetApp.flush();

  // 4) 同時寫到執行紀錄的 log
  console.log("DB_INFO=" + JSON.stringify(info));

  return info; // 讓你在「執行」視窗也可能看到回傳值
}


function debugPing() {
  const ss = db_();
  const sh = ss.getSheets()[0];
  sh.getRange("A1").setValue("PING " + new Date().toISOString());
}

// ---------------- API Router ----------------

function handleApi_(action, e, method) {
  try {
    switch (action) {
      case "createRoom": return apiCreateRoom_(e);
      case "joinRoom": return apiJoinRoom_(e);
      case "updateSettings": return apiUpdateSettings_(e);
      case "startRound": return apiStartRound_(e);
      case "bid": return apiBid_(e);
      case "playCard": return apiPlayCard_(e);
      case "resolve": return apiResolve_(e);
      case "sync": return apiSync_(e);
      case "listLobbyRooms": return apiListLobbyRooms_(e);
      default:
        return jsonOut_({ ok:false, error:"UNKNOWN_ACTION", action });
    }
  } catch (err) {
    return jsonOut_({ ok:false, error:"SERVER_ERROR", message: String(err?.message || err) });
  }
}

// ---------------- API ----------------

function apiCreateRoom_(e) {
  const body = e.body || {};
  const lobbyCode = normalizeLobbyCode_(body.lobbyCode || "");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = db_();
    const rooms = ss.getSheetByName(SHEET_ROOMS);

    const roomId = makeRoomId_();
    const hostToken = Utilities.getUuid();
    const seed = Utilities.getUuid().replace(/-/g,"").slice(0,16);
    const now = Date.now();

    rooms.appendRow([
      roomId, hostToken, lobbyCode, STATE_LOBBY, 0, "", 0, seed, now,
      DEFAULT_BUDGET, DEFAULT_ROUND_SECONDS, DEFAULT_MAX_ROUNDS, "", now
    ]);

    appendEvent_(roomId, "", "ROOM_CREATED", { roomId, lobbyCode });

    return jsonOut_({ ok:true, roomId, hostToken, lobbyCode });
  } finally {
    lock.releaseLock();
  }
}

function apiJoinRoom_(e) {
  const body = e.body || {};
  const roomId = mustStr_(body.roomId, "roomId");
  const name = sanitizeName_(mustStr_(body.name, "name"));

  const room = getRoom_(roomId);
  if (!room) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const room2 = getRoom_(roomId);
    if (!room2) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });

    const ss = db_();
    const playersSheet = ss.getSheetByName(SHEET_PLAYERS);

    const existing = findPlayerByName_(roomId, name);
    if (existing) {
      touchPlayer_(roomId, existing.playerId);
      appendEvent_(roomId, existing.playerId, "REJOIN", { name: existing.name });
      patchRoomActive_(roomId, Date.now());
      return jsonOut_({
        ok:true,
        roomId,
        playerId: existing.playerId,
        name: existing.name,
        budget: existing.budget,
        score: existing.score,
        rejoined:true
      });
    }

    const playerId = Utilities.getUuid();
    const now = Date.now();
    const startingBudget = Number(room2.startingBudget) || DEFAULT_BUDGET;

    playersSheet.appendRow([roomId, playerId, name, startingBudget, 0, now, now]);

    // starter cards
    setCardCount_(roomId, playerId, CARD_BOOST, 1);
    setCardCount_(roomId, playerId, CARD_CAP20, 1);

    appendEvent_(roomId, playerId, "JOIN", { name });
    patchRoomActive_(roomId, now);

    return jsonOut_({ ok:true, roomId, playerId, name, budget: startingBudget, score: 0 });
  } finally {
    lock.releaseLock();
  }
}

function apiUpdateSettings_(e) {
  const body = e.body || {};
  const roomId = mustStr_(body.roomId, "roomId");
  const hostToken = mustStr_(body.hostToken, "hostToken");

  const startingBudget = clampInt_(body.startingBudget, 10, 500, DEFAULT_BUDGET);
  const roundSeconds = clampInt_(body.roundSeconds, 5, 120, DEFAULT_ROUND_SECONDS);
  const maxRounds = clampInt_(body.maxRounds, 1, 50, DEFAULT_MAX_ROUNDS);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const room = getRoom_(roomId);
    if (!room) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });
    if (room.hostToken !== hostToken) return jsonOut_({ ok:false, error:"NOT_HOST" });
    if (room.state !== STATE_LOBBY) return jsonOut_({ ok:false, error:"NOT_IN_LOBBY" });

    patchRoom_(roomId, { startingBudget, roundSeconds, maxRounds });
    appendEvent_(roomId, "", "SETTINGS", { startingBudget, roundSeconds, maxRounds });
    patchRoomActive_(roomId, Date.now());

    return jsonOut_({ ok:true, startingBudget, roundSeconds, maxRounds });
  } finally {
    lock.releaseLock();
  }
}

function apiStartRound_(e) {
  const body = e.body || {};
  const roomId = mustStr_(body.roomId, "roomId");
  const hostToken = mustStr_(body.hostToken, "hostToken");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const room = getRoom_(roomId);
    if (!room) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });
    if (room.hostToken !== hostToken) return jsonOut_({ ok:false, error:"NOT_HOST" });
    if (room.state === STATE_ENDED) return jsonOut_({ ok:false, error:"GAME_ENDED" });
    if (room.state === STATE_BIDDING) return jsonOut_({ ok:false, error:"ALREADY_BIDDING" });

    const nextRound = Number(room.round) + 1;
    const maxRounds = Number(room.maxRounds) || DEFAULT_MAX_ROUNDS;
    if (nextRound > maxRounds) {
      patchRoom_(roomId, { state: STATE_ENDED });
      appendEvent_(roomId, "", "ENDED", { reason:"MAX_ROUNDS", maxRounds });
      patchRoomActive_(roomId, Date.now());
      return jsonOut_({ ok:false, error:"MAX_ROUNDS_REACHED" });
    }

    const item = pickItemByRound_(roomId, nextRound);
    if (!item) return jsonOut_({ ok:false, error:"NO_ITEMS" });

    const now = Date.now();
    const seconds = Number(room.roundSeconds) || DEFAULT_ROUND_SECONDS;
    const deadline = now + seconds * 1000;

    patchRoom_(roomId, {
      state: STATE_BIDDING,
      round: nextRound,
      itemId: item.itemId,
      bidDeadlineTs: deadline,
      capThisRound: ""
    });

    appendEvent_(roomId, "", "START_ROUND", {
      round: nextRound,
      itemId: item.itemId,
      itemName: item.name,
      publicHint: item.publicHint,
      bidDeadlineTs: deadline
    });

    patchRoomActive_(roomId, now);

    return jsonOut_({
      ok:true,
      round: nextRound,
      item: { itemId: item.itemId, name: item.name, publicHint: item.publicHint },
      bidDeadlineTs: deadline
    });
  } finally {
    lock.releaseLock();
  }
}

function apiBid_(e) {
  const body = e.body || {};
  const roomId = mustStr_(body.roomId, "roomId");
  const playerId = mustStr_(body.playerId, "playerId");
  const round = Number(body.round);
  const bid = Number(body.bid);

  if (!Number.isFinite(round) || round <= 0) return jsonOut_({ ok:false, error:"BAD_ROUND" });
  if (!Number.isFinite(bid) || bid < 0) return jsonOut_({ ok:false, error:"BAD_BID" });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const room = getRoom_(roomId);
    if (!room) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });
    if (room.state !== STATE_BIDDING) return jsonOut_({ ok:false, error:"NOT_BIDDING" });
    if (Number(room.round) !== round) return jsonOut_({ ok:false, error:"ROUND_MISMATCH", serverRound:Number(room.round) });

    const now = Date.now();
    if (now > Number(room.bidDeadlineTs)) return jsonOut_({ ok:false, error:"DEADLINE_PASSED" });

    const player = getPlayer_(roomId, playerId);
    if (!player) return jsonOut_({ ok:false, error:"PLAYER_NOT_FOUND" });

    if (bid > Number(player.budget)) return jsonOut_({ ok:false, error:"INSUFFICIENT_BUDGET", budget:Number(player.budget) });

    appendEvent_(roomId, playerId, "BID", { round, bid });
    touchPlayer_(roomId, playerId);
    return jsonOut_({ ok:true });
  } finally {
    lock.releaseLock();
  }
}

function apiPlayCard_(e) {
  const body = e.body || {};
  const roomId = mustStr_(body.roomId, "roomId");
  const playerId = mustStr_(body.playerId, "playerId");
  const round = Number(body.round);
  const cardId = mustStr_(body.cardId, "cardId");

  if (!Number.isFinite(round) || round <= 0) return jsonOut_({ ok:false, error:"BAD_ROUND" });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const room = getRoom_(roomId);
    if (!room) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });
    if (room.state !== STATE_BIDDING) return jsonOut_({ ok:false, error:"NOT_BIDDING" });
    if (Number(room.round) !== round) return jsonOut_({ ok:false, error:"ROUND_MISMATCH", serverRound:Number(room.round) });

    const now = Date.now();
    if (now > Number(room.bidDeadlineTs)) return jsonOut_({ ok:false, error:"DEADLINE_PASSED" });

    const player = getPlayer_(roomId, playerId);
    if (!player) return jsonOut_({ ok:false, error:"PLAYER_NOT_FOUND" });

    const count = getCardCount_(roomId, playerId, cardId);
    if (count <= 0) return jsonOut_({ ok:false, error:"NO_CARD" });

    if (cardId === CARD_BOOST) {
      const fee = 5;
      if (Number(player.budget) < fee) return jsonOut_({ ok:false, error:"INSUFFICIENT_BUDGET_FOR_CARD", fee });
      setPlayerFields_(roomId, playerId, { budget: Number(player.budget) - fee });
    } else if (cardId === CARD_CAP20) {
      const cap = 20;
      const cur = room.capThisRound ? Number(room.capThisRound) : 0;
      const newCap = (cur > 0) ? Math.min(cur, cap) : cap;
      patchRoom_(roomId, { capThisRound: newCap });
    } else {
      return jsonOut_({ ok:false, error:"UNKNOWN_CARD" });
    }

    setCardCount_(roomId, playerId, cardId, count - 1);
    appendEvent_(roomId, playerId, "PLAY_CARD", { round, cardId });
    patchRoomActive_(roomId, Date.now());

    return jsonOut_({ ok:true });
  } finally {
    lock.releaseLock();
  }
}

function apiResolve_(e) {
  const body = e.body || {};
  const roomId = mustStr_(body.roomId, "roomId");
  const hostToken = body.hostToken ? String(body.hostToken) : "";

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const room = getRoom_(roomId);
    if (!room) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });
    if (hostToken && room.hostToken !== hostToken) return jsonOut_({ ok:false, error:"NOT_HOST" });
    if (room.state !== STATE_BIDDING) return jsonOut_({ ok:false, error:"NOT_BIDDING" });

    const now = Date.now();
    if (now < Number(room.bidDeadlineTs)) return jsonOut_({ ok:false, error:"NOT_YET_DEADLINE", bidDeadlineTs:Number(room.bidDeadlineTs) });

    patchRoom_(roomId, { state: STATE_RESOLVING });

    const round = Number(room.round);
    const item = getItem_(room.itemId);
    if (!item) {
      patchRoom_(roomId, { state: STATE_LOBBY });
      return jsonOut_({ ok:false, error:"ITEM_NOT_FOUND" });
    }

    const cap = room.capThisRound ? Number(room.capThisRound) : 0;

    const played = getPlayedCardsForRound_(roomId, round, Number(room.bidDeadlineTs));
    const boostedPlayers = new Set(played.filter(x => x.cardId === CARD_BOOST).map(x => x.playerId));

    const bids = getBidsForRound_(roomId, round, Number(room.bidDeadlineTs));

    let winner = null;
    for (const b of bids) {
      let eff = b.bid;
      if (boostedPlayers.has(b.playerId)) eff += 10;
      if (cap > 0) eff = Math.min(eff, cap);

      const cand = { ...b, effectiveBid: eff };
      if (!winner) { winner = cand; continue; }
      if (cand.effectiveBid > winner.effectiveBid) { winner = cand; continue; }
      if (cand.effectiveBid === winner.effectiveBid) {
        if (cand.ts < winner.ts) { winner = cand; continue; }
        if (cand.ts === winner.ts && cand.eventId < winner.eventId) { winner = cand; continue; }
      }
    }

    const result = {
      round,
      itemId: item.itemId,
      itemName: item.name,
      publicHint: item.publicHint,
      trueValue: Number(item.trueValue),
      capApplied: cap || 0,
      boostsApplied: Array.from(boostedPlayers),
      winnerPlayerId: "",
      winnerName: "",
      winningBid: 0,
      winningEffectiveBid: 0,
      hadBids: bids.length > 0
    };

    if (winner) {
      const player = getPlayer_(roomId, winner.playerId);
      if (player && winner.bid <= Number(player.budget)) {
        const newBudget = Number(player.budget) - winner.bid;
        const newScore = Number(player.score) + Number(item.trueValue);
        setPlayerFields_(roomId, winner.playerId, { budget: newBudget, score: newScore });

        result.winnerPlayerId = winner.playerId;
        result.winnerName = player.name;
        result.winningBid = winner.bid;
        result.winningEffectiveBid = winner.effectiveBid;
      }
    }

    appendEvent_(roomId, "", "RESOLVE", result);

    const maxRounds = Number(room.maxRounds) || DEFAULT_MAX_ROUNDS;
    if (round >= maxRounds) {
      patchRoom_(roomId, { state: STATE_ENDED, itemId:"", bidDeadlineTs:0 });
      appendEvent_(roomId, "", "ENDED", { reason:"MAX_ROUNDS", maxRounds });
    } else {
      patchRoom_(roomId, { state: STATE_LOBBY, itemId:"", bidDeadlineTs:0 });
    }

    patchRoomActive_(roomId, Date.now());
    return jsonOut_({ ok:true, result });
  } finally {
    lock.releaseLock();
  }
}

function apiSync_(e) {
  const p = e?.parameter || {};
  const roomId = p.roomId ? String(p.roomId) : "";
  if (!roomId) return jsonOut_({ ok:false, error:"MISSING_roomId" });

  const sinceEventId = p.sinceEventId ? Number(p.sinceEventId) : 0;
  const playerId = p.playerId ? String(p.playerId) : "";
  const hostToken = p.hostToken ? String(p.hostToken) : "";

  const room = getRoom_(roomId);
  if (!room) return jsonOut_({ ok:false, error:"ROOM_NOT_FOUND" });

  // Host heartbeat (throttled): keeps lobby list fresh without every player writing
  if (hostToken && hostToken === room.hostToken) {
    const now = Date.now();
    const last = Number(room.lastActiveTs) || 0;
    if (now - last >= ROOM_HEARTBEAT_THROTTLE_MS) {
      patchRoomActive_(roomId, now);
    }
  }

  const players = listPlayers_(roomId).map(ply => ({
    playerId: ply.playerId,
    name: ply.name,
    budget: Number(ply.budget),
    score: Number(ply.score)
  }));

  let itemPublic = null;
  if (room.state === STATE_BIDDING && room.itemId) {
    const item = getItem_(room.itemId);
    if (item) itemPublic = { itemId: item.itemId, name: item.name, publicHint: item.publicHint };
  }

  const events = listEventsSinceFast_(roomId, sinceEventId);
  const myCards = playerId ? getPlayerCards_(roomId, playerId) : [];

  // throttle player lastSeen writes: only if >= 12s
  if (playerId) {
    try { touchPlayerThrottled_(roomId, playerId, 12000); } catch (_) {}
  }

  return jsonOut_({
    ok:true,
    room: {
      roomId: room.roomId,
      lobbyCode: room.lobbyCode || "",
      state: room.state,
      round: Number(room.round),
      bidDeadlineTs: Number(room.bidDeadlineTs),
      item: itemPublic,
      startingBudget: Number(room.startingBudget) || DEFAULT_BUDGET,
      roundSeconds: Number(room.roundSeconds) || DEFAULT_ROUND_SECONDS,
      maxRounds: Number(room.maxRounds) || DEFAULT_MAX_ROUNDS,
      capThisRound: room.capThisRound ? Number(room.capThisRound) : 0,
      lastActiveTs: Number(room.lastActiveTs) || 0
    },
    players,
    myCards,
    events,
    serverNow: Date.now()
  });
}

function apiListLobbyRooms_(e) {
  const p = e?.parameter || {};
  const lobbyCode = normalizeLobbyCode_(p.lobbyCode || "");
  if (!lobbyCode) return jsonOut_({ ok:false, error:"MISSING_lobbyCode" });

  const now = Date.now();
  const rooms = listRoomsByLobby_(lobbyCode, now - LOBBY_ACTIVE_WINDOW_MS);

  // player counts (single scan)
  const counts = countPlayersForRooms_(rooms.map(r => r.roomId));

  const out = rooms.map(r => ({
    roomId: r.roomId,
    lobbyCode: r.lobbyCode,
    state: r.state,
    round: Number(r.round) || 0,
    maxRounds: Number(r.maxRounds) || DEFAULT_MAX_ROUNDS,
    roundSeconds: Number(r.roundSeconds) || DEFAULT_ROUND_SECONDS,
    lastActiveTs: Number(r.lastActiveTs) || 0,
    playerCount: counts[r.roomId] || 0
  }));

  // order: most recent active first
  out.sort((a,b)=> (b.lastActiveTs - a.lastActiveTs) || (a.roomId.localeCompare(b.roomId)));

  return jsonOut_({ ok:true, rooms: out, serverNow: now });
}

// ---------------- Storage helpers ----------------

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const lastCol = Math.max(sh.getLastColumn(), headers.length);
  const row1 = sh.getLastRow() ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];
  let changed = (sh.getLastRow() === 0);

  if (!changed) {
    for (let i=0;i<headers.length;i++) {
      if (row1[i] !== headers[i]) { changed = true; break; }
    }
  }

  if (changed) {
    const existing = sh.getDataRange().getValues();
    sh.clear();

    sh.getRange(1,1,1,headers.length).setValues([headers]);

    if (existing.length > 1) {
      const oldHeaders = existing[0] || [];
      const oldMap = {};
      oldHeaders.forEach((h, i) => { oldMap[String(h)] = i; });

      const newRows = [];
      for (let r=1;r<existing.length;r++) {
        const oldRow = existing[r];
        const newRow = headers.map(h => {
          const idx = oldMap[String(h)];
          return (idx === undefined) ? "" : oldRow[idx];
        });
        newRows.push(newRow);
      }
      if (newRows.length) {
        sh.getRange(2,1,newRows.length,headers.length).setValues(newRows);
      }
    }
  }
}

function seedItemsIfEmpty_() {
  const ss = db_();
  const items = ss.getSheetByName(SHEET_ITEMS);
  if (items.getLastRow() > 1) return;

  const sample = [
    ["I001","生鏽懷錶","提示：出自某位名人後代的遺物（真假不明）", 14],
    ["I002","玻璃瓶裡的沙","提示：標籤寫著“月球”", 6],
    ["I003","折角地圖","提示：角落有手寫座標與咖啡漬", 11],
    ["I004","黑色錄音帶","提示：只有 30 秒，但有人願意付錢買沉默", 16],
    ["I005","怪異小雕像","提示：看起來像護身符，也像詛咒", 13],
    ["I006","泛黃劇照","提示：背後簽名疑似偽造", 8],
    ["I007","金屬鑰匙","提示：沒有鎖孔，但你就是想要它", 9],
    ["I008","絲質手帕","提示：淡淡香味，像某個你忘不掉的人", 12],
    ["I009","破損望遠鏡","提示：看得到遠方，看不到真相", 10],
    ["I010","薄薄一本筆記","提示：裡面有一頁被撕掉，大家都在猜那頁寫了什麼", 15],
  ];
  items.getRange(2,1,sample.length,4).setValues(sample);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function mustStr_(v, field) {
  if (v === undefined || v === null) throw new Error("MISSING_" + field);
  const s = String(v).trim();
  if (!s) throw new Error("MISSING_" + field);
  return s;
}

function sanitizeName_(name) {
  name = String(name).trim();
  if (!name) return "Player";
  if (name.length > 18) name = name.slice(0,18);
  return name.replace(/[<>]/g,"");
}

function clampInt_(v, min, max, def) {
  const n = Number(v);
  if (!Number.isFinite(n)) return def;
  return Math.max(min, Math.min(max, Math.floor(n)));
}

function normalizeLobbyCode_(v) {
  const s = String(v || "").trim().toUpperCase();
  if (!s) return "";
  const cleaned = s.replace(/[^A-Z0-9]/g, "").slice(0, 8);
  if (cleaned.length < 2) return "";
  return cleaned;
}

function makeRoomId_() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let id = "";
  for (let i=0;i<6;i++) id += chars[Math.floor(Math.random()*chars.length)];
  return id;
}

function indexMap_(headers) {
  const m = {};
  headers.forEach((h,i)=>{ m[String(h)] = i; });
  return m;
}

function getRoom_(roomId) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_ROOMS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId)) {
      return {
        row: r+1,
        roomId: String(data[r][idx.roomId]),
        hostToken: String(data[r][idx.hostToken]),
        lobbyCode: String(data[r][idx.lobbyCode] || ""),
        state: String(data[r][idx.state]),
        round: Number(data[r][idx.round]) || 0,
        itemId: String(data[r][idx.itemId] || ""),
        bidDeadlineTs: Number(data[r][idx.bidDeadlineTs]) || 0,
        seed: String(data[r][idx.seed] || ""),
        createdTs: Number(data[r][idx.createdTs]) || 0,
        startingBudget: Number(data[r][idx.startingBudget]) || DEFAULT_BUDGET,
        roundSeconds: Number(data[r][idx.roundSeconds]) || DEFAULT_ROUND_SECONDS,
        maxRounds: Number(data[r][idx.maxRounds]) || DEFAULT_MAX_ROUNDS,
        capThisRound: String(data[r][idx.capThisRound] || ""),
        lastActiveTs: Number(data[r][idx.lastActiveTs]) || 0
      };
    }
  }
  return null;
}

function patchRoom_(roomId, patch) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_ROOMS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId)) {
      const row = r+1;
      const setIf = (key, colName) => {
        if (patch[key] !== undefined) sh.getRange(row, idx[colName]+1).setValue(patch[key]);
      };
      setIf("lobbyCode","lobbyCode");
      setIf("state","state");
      setIf("round","round");
      setIf("itemId","itemId");
      setIf("bidDeadlineTs","bidDeadlineTs");
      setIf("startingBudget","startingBudget");
      setIf("roundSeconds","roundSeconds");
      setIf("maxRounds","maxRounds");
      setIf("capThisRound","capThisRound");
      setIf("lastActiveTs","lastActiveTs");
      return;
    }
  }
  throw new Error("ROOM_NOT_FOUND");
}

function patchRoomActive_(roomId, ts) {
  patchRoom_(roomId, { lastActiveTs: ts });
}

function listRoomsByLobby_(lobbyCode, minActiveTs) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_ROOMS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);
  const out = [];

  for (let r=1;r<data.length;r++) {
    const lc = String(data[r][idx.lobbyCode] || "");
    if (lc !== lobbyCode) continue;
    const lastActiveTs = Number(data[r][idx.lastActiveTs]) || 0;
    if (lastActiveTs < minActiveTs) continue;

    out.push({
      roomId: String(data[r][idx.roomId]),
      lobbyCode: lc,
      state: String(data[r][idx.state]),
      round: Number(data[r][idx.round]) || 0,
      maxRounds: Number(data[r][idx.maxRounds]) || DEFAULT_MAX_ROUNDS,
      roundSeconds: Number(data[r][idx.roundSeconds]) || DEFAULT_ROUND_SECONDS,
      lastActiveTs
    });
  }
  return out;
}

function countPlayersForRooms_(roomIds) {
  const want = new Set(roomIds.map(String));
  const counts = {};
  for (const id of want) counts[id] = 0;

  const ss = db_();
  const sh = ss.getSheetByName(SHEET_PLAYERS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    const rid = String(data[r][idx.roomId] || "");
    if (!want.has(rid)) continue;
    counts[rid] = (counts[rid] || 0) + 1;
  }
  return counts;
}

function getPlayer_(roomId, playerId) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_PLAYERS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId) &&
        String(data[r][idx.playerId]) === String(playerId)) {
      return {
        row: r+1,
        roomId: String(data[r][idx.roomId]),
        playerId: String(data[r][idx.playerId]),
        name: String(data[r][idx.name]),
        budget: Number(data[r][idx.budget]) || 0,
        score: Number(data[r][idx.score]) || 0,
        joinedTs: Number(data[r][idx.joinedTs]) || 0,
        lastSeenTs: Number(data[r][idx.lastSeenTs]) || 0
      };
    }
  }
  return null;
}

function setPlayerFields_(roomId, playerId, patch) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_PLAYERS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId) &&
        String(data[r][idx.playerId]) === String(playerId)) {
      const row = r+1;
      if (patch.budget !== undefined) sh.getRange(row, idx.budget+1).setValue(patch.budget);
      if (patch.score !== undefined) sh.getRange(row, idx.score+1).setValue(patch.score);
      if (patch.lastSeenTs !== undefined) sh.getRange(row, idx.lastSeenTs+1).setValue(patch.lastSeenTs);
      return;
    }
  }
  throw new Error("PLAYER_NOT_FOUND");
}

function touchPlayer_(roomId, playerId) {
  setPlayerFields_(roomId, playerId, { lastSeenTs: Date.now() });
}

function touchPlayerThrottled_(roomId, playerId, throttleMs) {
  const p = getPlayer_(roomId, playerId);
  if (!p) return;
  const now = Date.now();
  if (now - (Number(p.lastSeenTs)||0) >= throttleMs) {
    setPlayerFields_(roomId, playerId, { lastSeenTs: now });
  }
}

function listPlayers_(roomId) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_PLAYERS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  const out = [];
  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId)) {
      out.push({
        playerId: String(data[r][idx.playerId]),
        name: String(data[r][idx.name]),
        budget: Number(data[r][idx.budget]) || 0,
        score: Number(data[r][idx.score]) || 0
      });
    }
  }
  out.sort((a,b)=> b.score - a.score || b.budget - a.budget || a.name.localeCompare(b.name));
  return out;
}

function findPlayerByName_(roomId, name) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_PLAYERS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId) &&
        String(data[r][idx.name]) === String(name)) {
      return {
        playerId: String(data[r][idx.playerId]),
        name: String(data[r][idx.name]),
        budget: Number(data[r][idx.budget]) || 0,
        score: Number(data[r][idx.score]) || 0
      };
    }
  }
  return null;
}

function getItem_(itemId) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_ITEMS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.itemId]) === String(itemId)) {
      return {
        itemId: String(data[r][idx.itemId]),
        name: String(data[r][idx.name]),
        publicHint: String(data[r][idx.publicHint]),
        trueValue: Number(data[r][idx.trueValue]) || 0
      };
    }
  }
  return null;
}

function getAllItems_() {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_ITEMS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  const items = [];
  for (let r=1;r<data.length;r++) {
    const id = String(data[r][idx.itemId] || "");
    if (!id) continue;
    items.push({
      itemId: id,
      name: String(data[r][idx.name]),
      publicHint: String(data[r][idx.publicHint]),
      trueValue: Number(data[r][idx.trueValue]) || 0
    });
  }
  return items;
}

function pickItemByRound_(roomId, round) {
  const room = getRoom_(roomId);
  if (!room) return null;

  const items = getAllItems_();
  if (!items.length) return null;

  const seed = String(room.seed || "seed");
  const shuffled = seededShuffle_(items, seed);

  const idx = round - 1;
  const cycle = Math.floor(idx / shuffled.length);
  const offset = idx % shuffled.length;
  const shuffled2 = (cycle === 0) ? shuffled : seededShuffle_(items, seed + "_" + cycle);
  return shuffled2[offset] || null;
}

function seededShuffle_(arr, seed) {
  const a = arr.slice();
  const rng = mulberry32_(hash32_(seed));
  for (let i=a.length-1;i>0;i--) {
    const j = Math.floor(rng() * (i+1));
    const tmp = a[i]; a[i] = a[j]; a[j] = tmp;
  }
  return a;
}

function hash32_(s) {
  s = String(s);
  let h = 2166136261;
  for (let i=0;i<s.length;i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

function mulberry32_(seed) {
  let a = seed >>> 0;
  return function() {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// ---------------- Events (FAST) ----------------

function appendEvent_(roomId, playerId, type, payloadObj) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_EVENTS);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const cur = Number(props.getProperty("EVENT_SEQ") || "0");
    const next = cur + 1;
    props.setProperty("EVENT_SEQ", String(next));

    sh.appendRow([next, Date.now(), roomId, playerId, type, JSON.stringify(payloadObj || {})]);
    return next;
  } finally {
    lock.releaseLock();
  }
}

function listEventsSinceFast_(roomId, sinceEventId) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  const since = Math.max(0, Number(sinceEventId) || 0);
  const startRow = since + 2; // header at row 1, eventId 1 at row 2
  if (startRow > lastRow) return [];

  const numRows = lastRow - startRow + 1;
  const range = sh.getRange(startRow, 1, numRows, 6).getValues();

  const out = [];
  for (let i=0;i<range.length;i++) {
    const row = range[i];
    const eventId = Number(row[0]) || 0;
    if (eventId <= since) continue;
    if (String(row[2]) !== String(roomId)) continue;
    out.push({
      eventId,
      ts: Number(row[1]) || 0,
      roomId: String(row[2] || ""),
      playerId: String(row[3] || ""),
      type: String(row[4] || ""),
      payloadJson: String(row[5] || "{}")
    });
  }
  out.sort((a,b)=>a.eventId-b.eventId);
  return out;
}

function getBidsForRound_(roomId, round, deadlineTs) {
  const events = listEventsSinceFast_(roomId, 0).filter(ev => ev.type === "BID");
  const lastByPlayer = new Map();

  for (const ev of events) {
    if (ev.ts > deadlineTs) continue;
    let p = {};
    try { p = JSON.parse(ev.payloadJson || "{}"); } catch(_) { continue; }
    if (Number(p.round) !== Number(round)) continue;

    const bid = Number(p.bid);
    if (!Number.isFinite(bid) || bid < 0) continue;

    const key = String(ev.playerId);
    const prev = lastByPlayer.get(key);
    if (!prev || ev.eventId > prev.eventId) {
      lastByPlayer.set(key, { eventId: ev.eventId, ts: ev.ts, playerId: key, bid });
    }
  }
  return Array.from(lastByPlayer.values());
}

function getPlayedCardsForRound_(roomId, round, deadlineTs) {
  const events = listEventsSinceFast_(roomId, 0).filter(ev => ev.type === "PLAY_CARD");
  const out = [];
  for (const ev of events) {
    if (ev.ts > deadlineTs) continue;
    let p = {};
    try { p = JSON.parse(ev.payloadJson || "{}"); } catch(_) { continue; }
    if (Number(p.round) !== Number(round)) continue;
    if (!p.cardId) continue;
    out.push({ eventId: ev.eventId, ts: ev.ts, playerId: String(ev.playerId), cardId: String(p.cardId) });
  }
  out.sort((a,b)=>a.eventId-b.eventId);

  const seen = new Set();
  const filtered = [];
  for (const x of out) {
    const k = x.playerId + ":" + x.cardId + ":" + round;
    if (seen.has(k)) continue;
    seen.add(k);
    filtered.push(x);
  }
  return filtered;
}

// ---------------- Cards ----------------

function getCardCount_(roomId, playerId, cardId) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_PLAYER_CARDS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId) &&
        String(data[r][idx.playerId]) === String(playerId) &&
        String(data[r][idx.cardId]) === String(cardId)) {
      return Number(data[r][idx.count]) || 0;
    }
  }
  return 0;
}

function setCardCount_(roomId, playerId, cardId, count) {
  const ss = db_();
  const sh = ss.getSheetByName(SHEET_PLAYER_CARDS);
  const data = sh.getDataRange().getValues();
  const idx = indexMap_(data[0]);

  for (let r=1;r<data.length;r++) {
    if (String(data[r][idx.roomId]) === String(roomId) &&
        String(data[r][idx.playerId]) === String(playerId) &&
        String(data[r][idx.cardId]) === String(cardId)) {
      sh.getRange(r+1, idx.count+1).setValue(count);
      return;
    }
  }
  sh.appendRow([roomId, playerId, cardId, count]);
}

function getPlayerCards_(roomId, playerId) {
  const cards = [
    { cardId: CARD_BOOST, name: "強推(BOOST)", desc: "本回合有效出價 +10，但立刻扣 5 預算。" },
    { cardId: CARD_CAP20, name: "封頂(CAP20)", desc: "本回合所有人有效出價上限 = 20。" }
  ];
  return cards.map(c => ({ ...c, count: getCardCount_(roomId, playerId, c.cardId) }));
}
