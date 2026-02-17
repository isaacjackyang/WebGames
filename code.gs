/*******************************************************************************
 * Blind Auction Party — 完整遊戲伺服器
 * Code.gs  V4.1（帶註解 + 修正版）
 *
 * 技術選擇：Apps Script + Google Sheet 當 DB + 前端輪詢
 * 狀態機：LOBBY → BIDDING → REVEAL → POSTROUND → LOBBY/ENDED
 *
 * 修正內容 (V4.1)：
 *   - roomId 改為 1000-9999 避免 Google Sheets 吃掉前導零
 *   - ensureSheet_ 會比對既有 header，缺欄自動補齊
 *   - 移除未使用的 sheet_() 函式
 *   - 全面加上中英文註解
 ******************************************************************************/

/* ── 全域設定 ── */
const DB_SPREADSHEET_ID = "1IDUnxLtOWoPOS_ya_FMU3B5yMtjqUWe-NBA4dz9mPCk";

/** 工作表分頁名稱 */
const TABS = {
  ROOMS:   "Rooms",
  PLAYERS: "Players",
  EVENTS:  "Events",
  ITEMS:   "Items",
  CARDS:   "PlayerCards"
};

const ROOM_ACTIVE_MS      = 15 * 60 * 1000;  // 房間列表只顯示 15 分鐘內更新過的房間
const REVEAL_DURATION_MS  = 10 * 1000;        // REVEAL 頁面停留 10 秒
const POSTROUND_DURATION_MS = 8 * 1000;       // POSTROUND 排行榜停留 8 秒


/* ══════════════════════════════════════════════════════════════════════════════
 *  HTTP Endpoints（Apps Script Web App 的入口）
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * GET 請求處理
 * - 有 action：當 JSON API 使用
 * - 無 action：回傳前端 HTML 頁面
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = (params.action || "").trim();

  if (action) {
    try {
      return apiOut_(routeAction_(action, params), params);
    } catch (err) {
      return apiOut_({ ok: false, error: errMsg_(err) }, params);
    }
  }

  // 無 action → 回傳前端頁面
  const tpl = HtmlService.createTemplateFromFile("index");
  tpl.apiBaseUrl = ScriptApp.getService().getUrl();   // 注入 API base URL
  return tpl.evaluate()
    .setTitle("Blind Auction Party")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * POST 請求處理（payload 為 JSON body，含 action 欄位）
 */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || "";
    let body = {};
    try { body = raw ? JSON.parse(raw) : {}; } catch (_) {
      return jsonOut_({ ok: false, error: "POST_BODY_NOT_JSON" });
    }
    const action = (body.action || "").trim();
    if (!action) return jsonOut_({ ok: false, error: "MISSING_ACTION" });
    return jsonOut_(routeAction_(action, body));
  } catch (err) {
    return jsonOut_({ ok: false, error: errMsg_(err) });
  }
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  API 路由（GET / POST 共用同一張表）
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * 統一路由：根據 action 名稱分派到對應處理函式
 * GET 和 POST 都走這裡，差別只在參數來源（query vs body）
 */
function routeAction_(action, p) {
  switch (action) {
    case "debugInfo":       return apiDebugInfo_();
    case "setup":           return apiSetup_();
    case "createRoom":      return apiCreateRoom_(p);
    case "joinRoom":        return apiJoinRoom_(p);
    case "listLobbyRooms":  return apiListLobbyRooms_(p);
    case "sync":            return apiSync_(p);
    case "updateSettings":  return apiUpdateSettings_(p);
    case "startRound":      return apiStartRound_(p);
    case "bid":             return apiBid_(p);
    case "resolve":         return apiResolve_(p);
    case "playCard":        return apiPlayCard_(p);
    case "nextRound":       return apiStartRound_(p);  // nextRound 複用 startRound 邏輯
    case "restartGame":     return apiRestartGame_(p);
    case "kickPlayer":      return apiKickPlayer_(p);
    case "resetDatabase":   return apiResetDatabase_(p);
    default: return { ok: false, error: "UNKNOWN_ACTION:" + action };
  }
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  底層工具函式（DB 存取、鎖、序列化）
 * ════════════════════════════════════════════════════════════════════════════ */

/** 取得 DB Spreadsheet 物件 */
function db_() {
  const id = String(DB_SPREADSHEET_ID || "").trim();
  if (!id || id === "PASTE_YOUR_SPREADSHEET_ID_HERE") throw new Error("DB_SPREADSHEET_ID_NOT_SET");
  return SpreadsheetApp.openById(id);
}

/** 回傳 JSON ContentService */
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 回傳 API 回應（支援 JSONP callback，讓跨域靜態頁也能呼叫）
 */
function apiOut_(obj, params) {
  const cb = String((params && params.callback) || "").trim();
  if (cb && /^[A-Za-z_$][\w$.]{0,63}$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + JSON.stringify(obj) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOut_(obj);
}

/** 安全取得 error message 字串 */
function errMsg_(err) {
  return String(err && err.message ? err.message : err);
}

/** 取得當前毫秒時間戳 */
function nowMs_() { return Date.now(); }

/**
 * 產生隨機字串（字母+數字，排除易混淆字元 I/O/0/1）
 * @param {number} len 長度
 */
function rid_(len) {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let s = "";
  for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

/**
 * 正規化 lobbyCode：移除空白、英文轉大寫、至少 2 字
 * 支援中文/英文/數字混合輸入
 */
function normalizeLobbyCode_(raw) {
  let s = String(raw || "").trim();
  if (!s) return "";
  s = s.replace(/\s+/g, "").replace(/[a-z]/g, c => c.toUpperCase());
  return s.length >= 2 ? s.slice(0, 16) : "";
}

/**
 * 產生不重複的 4 位數 roomId（1000-9999）
 * 注意：不用 padStart("0") 是因為 Google Sheets 會把 "0123" 存成數字 123
 */
function newUniqueRoomId_(roomsSheet) {
  for (let i = 0; i < 50; i++) {
    const id = String(1000 + Math.floor(Math.random() * 9000)); // 1000~9999
    if (findRoomRow_(roomsSheet, id) < 0) return id;
  }
  throw new Error("ROOM_ID_POOL_EXHAUSTED");
}

/**
 * 用 Script Lock 保護寫入操作，避免同時寫入造成資料錯亂
 * @param {Function} fn 要在鎖內執行的函式
 */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/**
 * 確保工作表存在且 header 完整
 * - 若工作表不存在：建立並寫入 header
 * - 若工作表已存在但 header 不足：補齊缺少的欄位（不破壞既有資料）
 */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  // 檢查並補齊 header
  const existingWidth = sh.getLastColumn();
  if (existingWidth < headers.length) {
    // 現有 header 欄數不夠 → 補齊缺少的欄位
    const missing = headers.slice(existingWidth);
    sh.getRange(1, existingWidth + 1, 1, missing.length).setValues([missing]);
  } else if (existingWidth === 0) {
    // 完全空白的表 → 寫入完整 header
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sh;
}

/** 讀取整張表的 headers + data rows（防禦 null sheet） */
function readAll_(sh) {
  if (!sh) return { headers: [], rows: [] };
  const last = sh.getLastRow();
  if (last < 1) return { headers: [], rows: [] };
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return { headers: vals[0] || [], rows: [] };
  return { headers: vals[0], rows: vals.slice(1) };
}

/** 把 header 陣列轉成 { headerName: columnIndex } 的 map */
function idxMap_(headers) {
  const m = {};
  headers.forEach((h, i) => { if (String(h).trim()) m[String(h)] = i; });
  return m;
}

/** 新增一列資料到表尾 */
function appendRow_(sh, row) { sh.appendRow(row); }

/**
 * 在指定欄位搜尋 value，找到後回傳 sheet row index（1-based）；找不到回傳 -1
 * - 比對方式：String(cellValue) === String(value)，處理 Google Sheets 數字/字串差異
 */
function findRowIndex_(sh, colIndex, value) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, colIndex + 1, last - 1, 1).getValues();
  const target = String(value);
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === target) return i + 2;
  }
  return -1;
}

/** 讀取指定 row 的所有欄位 */
function getRow_(sh, rowIndex, width) {
  return sh.getRange(rowIndex, 1, 1, width).getValues()[0];
}

/** 覆寫指定 row 的所有欄位 */
function setRow_(sh, rowIndex, row) {
  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

/** 清除 header 以外的所有資料列（刪除整列避免 ghost rows） */
function clearSheetDataRows_(sh, _width) {
  const last = sh.getLastRow();
  if (last <= 1) return;
  sh.deleteRows(2, last - 1);
}

/** 在 Rooms 表中以 roomId 尋找列（封裝 findRowIndex_） */
function findRoomRow_(roomsSheet, roomId) {
  return findRowIndex_(roomsSheet, 0, roomId);
}

/** 寫入 seed items（100 項，每次 hiddenValue 隨機） */
function seedItems_(itemsSheet) {
  // [itemId, name, publicHint, minValue, maxValue]  — hiddenValue 在 min~max 間隨機
  const defs = [
    ["ITM001", "神秘古董",       "看起來很值錢，但可能是垃圾",   10, 60],
    ["ITM002", "限量公仔",       "大家都說會漲價",               15, 50],
    ["ITM003", "二手筆電",       "螢幕有刮痕，但能用",           10, 35],
    ["ITM004", "奇怪的箱子",     "搖一搖會響",                   5,  45],
    ["ITM005", "無敵券",         "規則外的力量感",               20, 70],
    ["ITM006", "金元寶",         "閃閃發光的黃金",               30, 80],
    ["ITM007", "紅包",           "裡面會放多少呢？",             5,  40],
    ["ITM008", "古董花瓶",       "據說是明朝的…",                15, 55],
    ["ITM009", "翡翠手鐲",       "通透碧綠，不知真假",           20, 90],
    ["ITM010", "舊照相機",       "底片還能用嗎？",               5,  25],
    ["ITM011", "名牌手錶",       "錶帶有點舊，但機芯很順",       30, 95],
    ["ITM012", "破損的吉他",     "缺一根弦",                     3,  15],
    ["ITM013", "陳年紅酒",       "1990 年份，保存狀況未知",       20, 75],
    ["ITM014", "貓咪背包",       "毛茸茸的設計，很搶眼",         8,  35],
    ["ITM015", "VR 眼鏡",        "二手的，鏡片有點花",           15, 45],
    ["ITM016", "黃金魚鉤",       "據說能釣到龍王",               20, 70],
    ["ITM017", "過期零食",       "已經過期三天了",               1,  10],
    ["ITM018", "鑽石耳環",       "小小一顆，閃閃的",             40, 100],
    ["ITM019", "手工肥皂",       "薰衣草味，很香",               3,  18],
    ["ITM020", "古地圖",         "標示著寶藏位置？",             15, 65],
    ["ITM021", "智慧音箱",       "只支援英文指令",               10, 35],
    ["ITM022", "生鏽寶劍",       "劍身刻著奇怪的文字",           12, 55],
    ["ITM023", "蘭花盆栽",       "開了三朵，很漂亮",             10, 40],
    ["ITM024", "簽名棒球",       "看不清是誰的簽名",             15, 75],
    ["ITM025", "機械鍵盤",       "Cherry 軟軸，手感一流",        20, 55],
    ["ITM026", "迷你無人機",     "電池續航 5 分鐘",              12, 45],
    ["ITM027", "神秘藥水",       "瓶身寫著『喝了變強』",         5,  30],
    ["ITM028", "油畫作品",       "署名看不清楚",                 25, 90],
    ["ITM029", "復古收音機",     "還能收到 FM 電台",             8,  30],
    ["ITM030", "水晶球",         "裡面有氣泡，閃閃發光",         15, 55],
    ["ITM031", "狗狗雨衣",       "XL 號，適合大型犬",            5,  20],
    ["ITM032", "珍珠項鍊",       "每顆大小不太一致",             20, 80],
    ["ITM033", "二手遊戲機",     "附兩個手把",                   20, 60],
    ["ITM034", "瑜伽墊",         "微微有使用痕跡",               3,  15],
    ["ITM035", "銀色懷錶",       "能走，但偶爾慢幾分鐘",         15, 55],
    ["ITM036", "神秘信封",       "封口用火漆封著",               5,  45],
    ["ITM037", "有機蜂蜜",       "農場直送，純天然",             8,  28],
    ["ITM038", "電競滑鼠",       "RGB 燈效全開",                 12, 40],
    ["ITM039", "陶瓷茶壺",       "日式風格，釉色美麗",           15, 50],
    ["ITM040", "登山背包",       "60 公升，防水材質",             18, 55],
    ["ITM041", "古銅色指南針",   "磁針還能動",                   10, 35],
    ["ITM042", "手繪撲克牌",     "每張都是獨特畫作",             12, 45],
    ["ITM043", "桌上型風扇",     "三段風速，有點吵",             3,  18],
    ["ITM044", "鍍金相框",       "適合放全家福",                 5,  25],
    ["ITM045", "真皮皮夾",       "義大利進口，有壓紋",           20, 65],
    ["ITM046", "迷你投影機",     "畫質勉強能看",                 15, 55],
    ["ITM047", "竹編籃子",       "手工製作，很精緻",             5,  22],
    ["ITM048", "大理石棋盤",     "附完整棋子",                   20, 70],
    ["ITM049", "香氛蠟燭組",     "六種味道",                     8,  30],
    ["ITM050", "折疊腳踏車",     "輪胎需要打氣",                 25, 80],
    ["ITM051", "老式打字機",     "鍵帽有些卡住",                 15, 50],
    ["ITM052", "太陽能充電器",   "陰天充電超慢",                 8,  35],
    ["ITM053", "手沖咖啡組",     "含磨豆機和濾杯",               12, 45],
    ["ITM054", "毛筆套組",       "書法愛好者必備",               5,  25],
    ["ITM055", "藍芽喇叭",       "防水，音質不錯",               15, 50],
    ["ITM056", "盆栽仙人掌",     "三年沒澆水都活著",             3,  15],
    ["ITM057", "皮革公事包",     "有商務質感",                   25, 70],
    ["ITM058", "復古桌燈",       "鎢絲燈泡，暖黃光",             10, 35],
    ["ITM059", "木雕擺件",       "手工雕刻的貓頭鷹",             15, 55],
    ["ITM060", "行李秤",         "出國旅行神器",                 3,  18],
    ["ITM061", "紫砂茶壺",       "刻有大師印章",                 40, 100],
    ["ITM062", "寵物自動餵食器", "可設定時間",                   10, 40],
    ["ITM063", "古典吊燈",       "需要重新接電線",               18, 60],
    ["ITM064", "天文望遠鏡",     "入門款，能看月球坑洞",         25, 80],
    ["ITM065", "手工皮帶",       "牛皮材質，銅扣環",             12, 45],
    ["ITM066", "老唱片",         "鄧麗君經典專輯",               15, 55],
    ["ITM067", "保溫便當盒",     "三層式，很實用",               5,  25],
    ["ITM068", "迷彩帳篷",       "雙人帳，附營釘",               20, 65],
    ["ITM069", "骨董電話",       "轉盤式，充滿年代感",           15, 50],
    ["ITM070", "水彩顏料組",     "24 色專業級",                  8,  35],
    ["ITM071", "黑膠唱片機",     "復古造型，能正常播放",         30, 90],
    ["ITM072", "編織毛毯",       "手工羊毛，冬天暖和",           10, 40],
    ["ITM073", "象棋組",         "檀木材質，質感好",             18, 60],
    ["ITM074", "螢幕掛燈",       "護眼不反光",                   10, 35],
    ["ITM075", "露營椅",         "鋁合金骨架，輕便",             8,  30],
    ["ITM076", "琥珀墜飾",       "裡面好像有蟲子",               20, 80],
    ["ITM077", "按摩槍",         "三段力道，筋膜放鬆",           15, 55],
    ["ITM078", "傳統摺扇",       "檀香木骨架",                   8,  35],
    ["ITM079", "電子書閱讀器",   "螢幕有一條淡淡的線",           18, 60],
    ["ITM080", "手工果醬",       "季節限定草莓口味",             3,  20],
    ["ITM081", "黃銅燭台",       "成對的，很有氣氛",             12, 45],
    ["ITM082", "運動水壺",       "保冷 24 小時",                 5,  25],
    ["ITM083", "老式鐘擺鐘",     "每小時會響一次",               25, 80],
    ["ITM084", "手作陶杯",       "釉色獨一無二",                 8,  30],
    ["ITM085", "防水相機",       "適合水下攝影",                 20, 70],
    ["ITM086", "薄荷精油",       "提神醒腦好物",                 3,  18],
    ["ITM087", "紀念幣套組",     "十二生肖完整版",               30, 90],
    ["ITM088", "羽毛筆",         "附墨水瓶",                     5,  25],
    ["ITM089", "空拍機零件",     "不確定能不能裝上",             2,  15],
    ["ITM090", "日式風鈴",       "玻璃材質，聲音清脆",           5,  20],
    ["ITM091", "皮革日記本",     "附鎖扣，很有質感",             10, 35],
    ["ITM092", "迷你魚缸",       "含過濾器和LED燈",              10, 40],
    ["ITM093", "古董懷爐",       "銅製的，可以用現代油",         12, 50],
    ["ITM094", "摺紙藝術品",     "1000 隻紙鶴串",                8,  35],
    ["ITM095", "手工吉他撥片",   "玳瑁材質，音色溫暖",           5,  25],
    ["ITM096", "景泰藍花瓶",     "掐絲琺瑯工藝",                35, 100],
    ["ITM097", "貓眼石戒指",     "有神秘光芒在流動",             25, 85],
    ["ITM098", "老式地球儀",     "還是蘇聯時代的版本",           15, 55],
    ["ITM099", "機械式鋼筆",     "14K 金筆尖，書寫流暢",        20, 70],
    ["ITM100", "龍銀(舊台幣)",   "日治時期流通的銀幣",           40, 100]
  ];
  // 每次隨機生成 hiddenValue
  const seed = defs.map(d => [d[0], d[1], d[2], d[3] + Math.floor(Math.random() * (d[4] - d[3] + 1))]);
  itemsSheet.getRange(2, 1, seed.length, 4).setValues(seed);
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  初始化 / 偵錯
 * ════════════════════════════════════════════════════════════════════════════ */

/** 外部可直接跑的初始化入口（也可從 Apps Script 控制台呼叫） */
function setup() { return apiSetup_(); }

/**
 * 建立所有必要的工作表並確保 header 完整
 * 若 Items 表為空則寫入 seed 資料
 */
function apiSetup_() {
  return withLock_(() => {
    const ss = db_();

    ensureSheet_(ss, TABS.ROOMS, [
      "roomId", "lobbyCode", "createdAt", "updatedAt",
      "state", "round", "maxRounds", "roundSeconds", "startingBudget",
      "hostToken", "bidDeadlineTs", "itemId",
      "revealUntilTs", "postUntilTs"
    ]);

    ensureSheet_(ss, TABS.PLAYERS, [
      "roomId", "playerId", "name", "isHost",
      "budget", "score", "joinedAt", "updatedAt"
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

    // Seed items（僅首次 — 避免遊戲進行中重新整理頁面時改變物品價值）
    const items = ss.getSheetByName(TABS.ITEMS);
    if (items.getLastRow() < 2) seedItems_(items);

    SpreadsheetApp.flush();
    return { ok: true };
  });
}

/** 回傳 DB 基本資訊，用來驗證 API 是否正常運作 */
function apiDebugInfo_() {
  const ss = db_();
  return {
    ok: true,
    info: { name: ss.getName(), id: ss.getId(), url: ss.getUrl(), time: new Date().toISOString() }
  };
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  遊戲 API：房間管理
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * 建立新房間
 * 回傳 roomId + hostToken（hostToken 用來驗證 Host 身分）
 */
function apiCreateRoom_(p) {
  const lobbyCode = normalizeLobbyCode_(p.lobbyCode);
  if (!lobbyCode) return { ok: false, error: "MISSING_LOBBYCODE" };

  return withLock_(() => {
    const ss = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const t = nowMs_();
    const roomId = newUniqueRoomId_(rooms);
    const hostToken = rid_(10);

    // 14 欄：對應 Rooms header
    appendRow_(rooms, [
      roomId, lobbyCode, t, t,
      "LOBBY", 0, 8, 25, 100,
      hostToken, "", "",
      "", ""   // revealUntilTs, postUntilTs
    ]);

    // 強制寫入：確保其他使用者的 listLobbyRooms 能立即看到新房間
    SpreadsheetApp.flush();

    addEvent_(roomId, "ROOM_CREATED", "", { lobbyCode });
    return { ok: true, roomId, hostToken };
  });
}

/**
 * 加入房間
 * - 若帶有正確的 hostToken，該玩家標記為 Host（解決 race condition）
 * - 否則若房間內無人，第一個加入者自動成為 Host
 * - 發放初始手牌
 */
function apiJoinRoom_(p) {
  const roomId   = String(p.roomId   || "").trim();
  const name     = String(p.name     || "").trim().slice(0, 18);
  const joinHostToken = String(p.hostToken || "").trim();  // 建房者帶入的 hostToken
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!name)   return { ok: false, error: "MISSING_NAME" };

  return withLock_(() => {
    const ss = db_();
    const rooms   = ss.getSheetByName(TABS.ROOMS);
    const players = ss.getSheetByName(TABS.PLAYERS);

    const roomRowIdx = findRoomRow_(rooms, roomId);
    if (roomRowIdx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const t = nowMs_();
    const { headers, rows } = readAll_(players);
    const pim = idxMap_(headers);

    // 計算房間內現有玩家數
    let count = 0;
    rows.forEach(r => { if (String(r[pim.roomId]) === roomId) count++; });

    const playerId = newUniquePlayerId_(rows, pim);
    const roomRow    = getRow_(rooms, roomRowIdx, rooms.getLastColumn());
    const roomHeaders = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(roomHeaders);

    // 判定 Host 身分：
    // 1. 若帶有正確的 hostToken → 一定是 Host（建房者）
    // 2. 否則若房間內無人 → 第一個加入者為 Host
    const realHostToken = String(roomRow[rim.hostToken] || "");
    const isHost = (joinHostToken && joinHostToken === realHostToken) || (count === 0);
    const startingBudget = Number(roomRow[rim.startingBudget] || 100);

    appendRow_(players, [
      roomId, playerId, name, isHost ? "1" : "0",
      startingBudget, 0, t, t
    ]);

    // 發放初始卡牌
    seedCards_(ss, roomId, playerId);

    // 更新房間的 updatedAt
    roomRow[rim.updatedAt] = t;
    setRow_(rooms, roomRowIdx, roomRow);

    addEvent_(roomId, "PLAYER_JOINED", playerId, { name, isHost });

    // 強制寫入：確保其他玩家的 sync 能立即看到新加入的玩家
    SpreadsheetApp.flush();

    return { ok: true, playerId };
  });
}

/**
 * Host 踢除玩家（刪除 Players + Cards 中該玩家的資料）
 */
function apiKickPlayer_(p) {
  const roomId       = String(p.roomId       || "").trim();
  const hostToken    = String(p.hostToken    || "").trim();
  const playerId     = String(p.playerId     || "").trim();
  const targetId     = String(p.targetPlayerId || "").trim();
  if (!roomId)   return { ok: false, error: "MISSING_ROOMID" };
  if (!targetId) return { ok: false, error: "MISSING_TARGET" };

  return withLock_(() => {
    const ss = db_();
    const rooms   = ss.getSheetByName(TABS.ROOMS);
    const players = ss.getSheetByName(TABS.PLAYERS);
    const cards   = ss.getSheetByName(TABS.CARDS);

    const roomRowIdx = findRoomRow_(rooms, roomId);
    if (roomRowIdx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const roomHeaders = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(roomHeaders);
    const rr  = getRow_(rooms, roomRowIdx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, rr, rim))
      return { ok: false, error: "HOST_PERMISSION_DENIED" };

    // 不能踢自己
    if (targetId === playerId) return { ok: false, error: "CANNOT_KICK_SELF" };

    // 刪除目標玩家（從下往上刪避免 row shift 問題）
    const pAll = readAll_(players);
    const pim  = idxMap_(pAll.headers);
    let found = false;
    for (let i = pAll.rows.length - 1; i >= 0; i--) {
      const pr = pAll.rows[i];
      if (String(pr[pim.roomId]) !== roomId) continue;
      if (String(pr[pim.playerId]) !== targetId) continue;
      players.deleteRow(i + 2);
      found = true;
    }
    if (!found) return { ok: false, error: "PLAYER_NOT_FOUND" };

    // 刪除該玩家的卡片
    if (cards) {
      const cAll = readAll_(cards);
      const cim  = idxMap_(cAll.headers);
      for (let i = cAll.rows.length - 1; i >= 0; i--) {
        const cr = cAll.rows[i];
        if (String(cr[cim.roomId]) !== roomId) continue;
        if (String(cr[cim.playerId]) !== targetId) continue;
        cards.deleteRow(i + 2);
      }
    }

    addEvent_(roomId, "PLAYER_KICKED", targetId, { by: playerId });
    touchRoomUpdatedAt_(roomId);
    SpreadsheetApp.flush();
    return { ok: true };
  });
}

/**
 * 列出活躍房間（15 分鐘內有更新）
 * lobbyCode 為可選參數：有填則過濾該場域，空值則列出所有房間
 */
function apiListLobbyRooms_(p) {
  const lobbyCode = normalizeLobbyCode_(p.lobbyCode);  // 可能為空字串

  const ss      = db_();
  let rooms   = ss.getSheetByName(TABS.ROOMS);
  let players = ss.getSheetByName(TABS.PLAYERS);

  // 若表不存在則自動初始化
  if (!rooms || !players) {
    apiSetup_();
    rooms   = ss.getSheetByName(TABS.ROOMS);
    players = ss.getSheetByName(TABS.PLAYERS);
  }
  if (!rooms || !players) return { ok: true, rooms: [] };

  const rAll = readAll_(rooms);
  const rim  = idxMap_(rAll.headers);
  const pAll = readAll_(players);
  const pim  = idxMap_(pAll.headers);

  const now = nowMs_();
  const out = [];

  rAll.rows.forEach(r => {
    const rowLobby = String(r[rim.lobbyCode] || "");

    // 若有指定 lobbyCode，則只顯示該場域的房間
    if (lobbyCode && rowLobby !== lobbyCode) return;

    // 過濾太久沒更新的房間
    const updatedAt = Number(r[rim.updatedAt] || 0);
    if (now - updatedAt > ROOM_ACTIVE_MS) return;

    // 跳過 roomId 為空的列（可能是 clearContent 殘留的空行）
    const rid = String(r[rim.roomId] || "").trim();
    if (!rid) return;

    // 計算玩家數
    let pc = 0;
    pAll.rows.forEach(pr => { if (String(pr[pim.roomId]) === rid) pc++; });

    out.push({
      roomId: rid,
      lobbyCode: rowLobby,
      state: String(r[rim.state]),
      round: Number(r[rim.round] || 0),
      maxRounds: Number(r[rim.maxRounds] || 0),
      playerCount: pc
    });
  });

  return { ok: true, rooms: out };
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  同步 API + 自動狀態轉換
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * 核心同步端點 — 前端每 2 秒呼叫一次
 * 功能：
 *   1. 觸發時間驅動的自動狀態轉換（BIDDING→REVEAL→POSTROUND→LOBBY/ENDED）
 *   2. 回傳房間狀態、玩家列表、手牌、事件增量、結算結果
 */
function apiSync_(p) {
  const roomId      = String(p.roomId      || "").trim().toUpperCase();
  const sinceEventId = Number(p.sinceEventId || 0);
  const playerId    = String(p.playerId    || "").trim();
  const hostToken   = String(p.hostToken   || "").trim();
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };

  // ── 1. 嘗試自動狀態轉換（需要鎖） ──
  tryAutoTransition_(roomId);

  // ── 2. 讀取所有資料（不需要鎖，純讀取） ──
  const ss     = db_();
  const rooms  = ss.getSheetByName(TABS.ROOMS);
  const roomRowIdx = findRoomRow_(rooms, roomId);
  if (roomRowIdx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

  const roomHeaders = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
  const rim = idxMap_(roomHeaders);
  const rr  = getRow_(rooms, roomRowIdx, rooms.getLastColumn());
  const state = String(rr[rim.state]);

  // ── 3. 組裝 room 物件 ──
  const items = ss.getSheetByName(TABS.ITEMS);
  let item = null;
  const itemId = String(rr[rim.itemId] || "");
  if (itemId) {
    const it = findItem_(items, itemId);
    if (it) {
      // REVEAL/POSTROUND/ENDED 時才公開 hiddenValue
      if (state === "REVEAL" || state === "POSTROUND" || state === "ENDED") {
        item = it;
      } else {
        item = { itemId: it.itemId, name: it.name, publicHint: it.publicHint };
      }
    }
  }

  const room = {
    roomId:         String(rr[rim.roomId]),
    lobbyCode:      String(rr[rim.lobbyCode]),
    state, 
    round:          Number(rr[rim.round]          || 0),
    maxRounds:      Number(rr[rim.maxRounds]      || 0),
    roundSeconds:   Number(rr[rim.roundSeconds]   || 25),
    startingBudget: Number(rr[rim.startingBudget] || 100),
    bidDeadlineTs:  Number(rr[rim.bidDeadlineTs]  || 0),
    revealUntilTs:  Number(rr[rim.revealUntilTs]  || 0),
    postUntilTs:    Number(rr[rim.postUntilTs]    || 0),
    item
  };

  // ── 4. 玩家列表 ──
  const players = ss.getSheetByName(TABS.PLAYERS);
  const pAll = readAll_(players);
  const pim  = idxMap_(pAll.headers);
  const plist = [];
  pAll.rows.forEach(pr => {
    if (String(pr[pim.roomId]) !== roomId) return;
    plist.push({
      playerId: String(pr[pim.playerId]),
      name:     String(pr[pim.name]),
      isHost:   String(pr[pim.isHost]) === "1",
      budget:   Number(pr[pim.budget]  || 0),
      score:    Number(pr[pim.score]   || 0)
    });
  });

  // ── 5. 我的手牌 ──
  const cards = ss.getSheetByName(TABS.CARDS);
  const cAll = readAll_(cards);
  const cim  = idxMap_(cAll.headers);
  const myCards = [];
  cAll.rows.forEach(cr => {
    if (String(cr[cim.roomId]) !== roomId || String(cr[cim.playerId]) !== playerId) return;
    myCards.push({
      cardId: String(cr[cim.cardId]),
      name:   String(cr[cim.name]),
      desc:   String(cr[cim.desc]),
      count:  Number(cr[cim.count] || 0)
    });
  });

  // ── 6. 事件增量 ──
  const events = ss.getSheetByName(TABS.EVENTS);
  const eAll = readAll_(events);
  const eim  = idxMap_(eAll.headers);
  const evs = [];
  eAll.rows.forEach(er => {
    if (String(er[eim.roomId]) !== roomId) return;
    const eid = Number(er[eim.eventId] || 0);
    if (eid <= sinceEventId) return;
    evs.push({
      eventId:     eid,
      ts:          Number(er[eim.ts] || 0),
      type:        String(er[eim.type]),
      playerId:    String(er[eim.playerId] || ""),
      payloadJson: String(er[eim.payloadJson] || "{}")
    });
  });

  // ── 7. Peek 卡牌效果：BIDDING 期間已使用 Peek 的玩家可看到 hiddenValue 區間 ──
  let peekHint = null;
  if (state === "BIDDING" && playerId && itemId) {
    const round = room.round;
    for (const er of eAll.rows) {
      if (String(er[eim.roomId]) !== roomId) continue;
      if (String(er[eim.type]) !== "CARD_PLAYED") continue;
      if (String(er[eim.playerId]) !== playerId) continue;
      let payload = {};
      try { payload = JSON.parse(String(er[eim.payloadJson] || "{}")); } catch (_) {}
      if (Number(payload.round) === round && payload.cardId === "C1") {
        const fullItem = findItem_(items, itemId);
        if (fullItem) {
          peekHint = { min: Math.max(0, fullItem.hiddenValue - 5), max: fullItem.hiddenValue + 5 };
        }
        break;
      }
    }
  }

  // ── 8. 目前我的出價（讓前端顯示「已出價：X」） ──
  let myBid = null;
  if (state === "BIDDING" && playerId) {
    eAll.rows.forEach(er => {
      if (String(er[eim.roomId]) !== roomId) return;
      if (String(er[eim.type]) !== "BID") return;
      if (String(er[eim.playerId]) !== playerId) return;
      let payload = {};
      try { payload = JSON.parse(String(er[eim.payloadJson] || "{}")); } catch (_) {}
      if (Number(payload.round) === room.round) myBid = Number(payload.bid || 0);
    });
  }

  // ── 9. 最新結算結果（REVEAL / POSTROUND 用） ──
  let resolveResult = null;
  if (state === "REVEAL" || state === "POSTROUND" || state === "ENDED") {
    // 從所有事件中找本回合的 ROUND_RESOLVED
    eAll.rows.forEach(er => {
      if (String(er[eim.roomId]) !== roomId) return;
      if (String(er[eim.type]) !== "ROUND_RESOLVED") return;
      try {
        const pl = JSON.parse(String(er[eim.payloadJson] || "{}"));
        if (Number(pl.round) === room.round) resolveResult = pl;
      } catch (_) {}
    });
  }

  // ── 10. Host 身分判定 ──
  const isHost = !!(hostToken && hostToken === String(rr[rim.hostToken]));

  return { ok: true, room, players: plist, myCards, events: evs, isHost, resolveResult, peekHint, myBid };
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  自動狀態轉換（時間驅動）
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * 在 sync 時呼叫，檢查是否該自動切換狀態
 * 因為 Apps Script 沒有可靠的常駐背景工作，靠玩家輪詢最穩定
 *
 * 轉換規則：
 *   BIDDING  → REVEAL    ：now >= bidDeadlineTs（自動結算）
 *   REVEAL   → POSTROUND ：now >= revealUntilTs
 *   REVEAL   → ENDED     ：now >= revealUntilTs 且 round >= maxRounds
 *   POSTROUND → LOBBY    ：now >= postUntilTs
 */
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
    const now   = nowMs_();

    // BIDDING → REVEAL（自動結算）
    if (state === "BIDDING") {
      const deadline = Number(r[rim.bidDeadlineTs] || 0);
      if (deadline > 0 && now >= deadline) {
        doResolve_(ss, rooms, idx, r, rim, roomId);
      }
      return;
    }

    // REVEAL → POSTROUND 或 ENDED
    if (state === "REVEAL") {
      const until = Number(r[rim.revealUntilTs] || 0);
      if (until > 0 && now >= until) {
        const round     = Number(r[rim.round]     || 0);
        const maxRounds = Number(r[rim.maxRounds]  || 8);
        if (round >= maxRounds) {
          r[rim.state] = "ENDED";
          addEvent_(roomId, "GAME_ENDED", "", { round });
        } else {
          r[rim.state]       = "POSTROUND";
          r[rim.postUntilTs] = now + POSTROUND_DURATION_MS;
          addEvent_(roomId, "STATE_TRANSITION", "", { from: "REVEAL", to: "POSTROUND" });
        }
        r[rim.updatedAt] = now;
        setRow_(rooms, idx, r);
      }
      return;
    }

    // POSTROUND → LOBBY
    if (state === "POSTROUND") {
      const until = Number(r[rim.postUntilTs] || 0);
      if (until > 0 && now >= until) {
        r[rim.state]          = "LOBBY";
        r[rim.itemId]         = "";
        r[rim.bidDeadlineTs]  = "";
        r[rim.revealUntilTs]  = "";
        r[rim.postUntilTs]    = "";
        r[rim.updatedAt]      = now;
        setRow_(rooms, idx, r);
        addEvent_(roomId, "STATE_TRANSITION", "", { from: "POSTROUND", to: "LOBBY" });
      }
    }
  });
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  結算邏輯（手動結算 + 自動結算 共用）
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * 執行本回合結算：找出得標者、套卡牌效果、更新分數/預算、切到 REVEAL
 * 被 apiResolve_（Host 手動）和 tryAutoTransition_（自動）呼叫
 */
function doResolve_(ss, rooms, idx, r, rim, roomId) {
  const players = ss.getSheetByName(TABS.PLAYERS);
  const events  = ss.getSheetByName(TABS.EVENTS);
  const items   = ss.getSheetByName(TABS.ITEMS);

  const round       = Number(r[rim.round]  || 0);
  const itemId      = String(r[rim.itemId] || "");
  const item        = itemId ? findItem_(items, itemId) : null;
  const hiddenValue = item ? Number(item.hiddenValue || 0) : 0;

  // 讀取玩家資料，建立 playerId → sheetRowIndex 的映射
  const pAll = readAll_(players);
  const pim  = idxMap_(pAll.headers);
  const pRowById = {};  // playerId → sheet row index (1-based)
  pAll.rows.forEach((pr, i) => {
    if (String(pr[pim.roomId]) !== roomId) return;
    pRowById[String(pr[pim.playerId])] = i + 2;
  });

  // 收集本回合的出價和卡牌使用紀錄
  const eAll = readAll_(events);
  const eim  = idxMap_(eAll.headers);
  const lastBid    = {};   // playerId → 最後一次出價金額
  const bidOrder   = {};   // playerId → 該出價在 events 中的索引（用於 tie-break）
  const cardPlayed = {};   // playerId → [cardId, ...]

  eAll.rows.forEach((er, idx) => {
    if (String(er[eim.roomId]) !== roomId) return;
    const type = String(er[eim.type]);
    const pid  = String(er[eim.playerId] || "");
    let payload = {};
    try { payload = JSON.parse(String(er[eim.payloadJson] || "{}")); } catch (_) {}

    if (type === "BID" && Number(payload.round) === round) {
      lastBid[pid]  = Number(payload.bid || 0);
      bidOrder[pid] = idx;  // 越後面 = 越晚出價
    }
    if (type === "CARD_PLAYED" && Number(payload.round) === round) {
      if (!cardPlayed[pid]) cardPlayed[pid] = [];
      cardPlayed[pid].push(String(payload.cardId || ""));
    }
  });

  // 找出得標者：出價最高且 ≤ 預算的玩家
  // Tie-break：同出價時，先出價者（bidOrder 較小）得標
  let winnerId  = "";
  let winnerBid = -1;
  let winnerOrder = Infinity;

  Object.keys(lastBid).forEach(pid => {
    const bid = lastBid[pid];
    const rowIdx = pRowById[pid];
    if (!rowIdx) return;
    const row    = getRow_(players, rowIdx, players.getLastColumn());
    const budget = Number(row[pim.budget] || 0);
    if (bid > budget) return;          // 超過預算 → 無效
    const order = bidOrder[pid] || 0;
    if (bid > winnerBid || (bid === winnerBid && order < winnerOrder)) {
      winnerBid = bid; winnerId = pid; winnerOrder = order;
    }
  });

  // 卡牌效果：Tax（加稅 +20%）
  let actualCost = winnerBid;
  let taxApplied = false;
  if (winnerId && cardPlayed[winnerId] && cardPlayed[winnerId].includes("C2")) {
    actualCost  = Math.ceil(winnerBid * 1.2);
    taxApplied = true;
  }

  // 計算分數變化
  const deltas     = {};
  let winnerName = "";

  if (winnerId) {
    const rowIdx = pRowById[winnerId];
    const row    = getRow_(players, rowIdx, players.getLastColumn());
    winnerName   = String(row[pim.name] || "");
    const budget = Number(row[pim.budget] || 0);
    const score  = Number(row[pim.score]  || 0);

    let delta = hiddenValue - actualCost;

    // 卡牌效果：Shield（護盾）— 若分數增量為負則抵消為 0
    let shieldApplied = false;
    if (delta < 0 && cardPlayed[winnerId] && cardPlayed[winnerId].includes("C3")) {
      delta = 0;
      shieldApplied = true;
    }

    deltas[winnerId] = { delta, shieldApplied, taxApplied };

    // 更新玩家資料
    row[pim.budget]    = budget - actualCost;
    row[pim.score]     = score + delta;
    row[pim.updatedAt] = nowMs_();
    setRow_(players, rowIdx, row);
  }

  // 狀態切換：BIDDING → REVEAL
  const now = nowMs_();
  r[rim.state]         = "REVEAL";
  r[rim.revealUntilTs] = now + REVEAL_DURATION_MS;
  r[rim.bidDeadlineTs] = "";
  r[rim.updatedAt]     = now;
  setRow_(rooms, idx, r);

  // 記錄結算事件（前端用來顯示結果面板）
  addEvent_(roomId, "ROUND_RESOLVED", "", {
    round, winnerId, winnerName, winnerBid, actualCost,
    hiddenValue, deltas, taxApplied,
    itemName: item ? item.name : ""
  });
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  遊戲 API：設定 / 回合 / 出價 / 卡牌
 * ════════════════════════════════════════════════════════════════════════════ */

/** Host 手動結算（按「立即結算」按鈕） */
function apiResolve_(p) {
  const roomId    = String(p.roomId    || "").trim().toUpperCase();
  const hostToken = String(p.hostToken || "").trim();
  const playerId  = String(p.playerId  || "").trim();
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken && !playerId) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss    = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx   = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(headers);
    const r   = getRow_(rooms, idx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, r, rim))
      return { ok: false, error: "HOST_PERMISSION_DENIED" };
    if (String(r[rim.state]) !== "BIDDING")
      return { ok: false, error: "NOT_IN_BIDDING" };

    doResolve_(ss, rooms, idx, r, rim, roomId);
    return { ok: true };
  });
}

/** Host 更新房間設定（回合數、秒數、起始預算） */
function apiUpdateSettings_(p) {
  const roomId        = String(p.roomId       || "").trim().toUpperCase();
  const hostToken     = String(p.hostToken    || "").trim();
  const playerId      = String(p.playerId     || "").trim();
  const startingBudget = Number(p.startingBudget || 100);
  const roundSeconds   = Number(p.roundSeconds   || 25);
  const maxRounds      = Number(p.maxRounds      || 8);
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken && !playerId) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss    = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx   = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const im = idxMap_(headers);
    const r  = getRow_(rooms, idx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, r, im))
      return { ok: false, error: "HOST_PERMISSION_DENIED" };

    r[im.startingBudget] = startingBudget;
    r[im.roundSeconds]   = roundSeconds;
    r[im.maxRounds]      = maxRounds;
    r[im.updatedAt]      = nowMs_();
    setRow_(rooms, idx, r);

    addEvent_(roomId, "SETTINGS_UPDATED", "", { startingBudget, roundSeconds, maxRounds });
    return { ok: true };
  });
}

/**
 * Host 開始新回合
 * 隨機選一個 item，設定出價截止時間，切到 BIDDING
 */
function apiStartRound_(p) {
  const roomId    = String(p.roomId    || "").trim().toUpperCase();
  const hostToken = String(p.hostToken || "").trim();
  const playerId  = String(p.playerId  || "").trim();
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken && !playerId) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss    = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const items = ss.getSheetByName(TABS.ITEMS);
    const idx   = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const im = idxMap_(headers);
    const r  = getRow_(rooms, idx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, r, im))
      return { ok: false, error: "HOST_PERMISSION_DENIED" };

    const curState = String(r[im.state]);
    if (curState !== "LOBBY" && curState !== "POSTROUND")
      return { ok: false, error: "NOT_IN_LOBBY_OR_POSTROUND" };

    const round     = Number(r[im.round] || 0) + 1;
    const maxRounds = Number(r[im.maxRounds] || 8);
    if (round > maxRounds) return { ok: false, error: "MAX_ROUNDS_REACHED" };

    const it = pickRandomItem_(items);
    if (!it) return { ok: false, error: "NO_ITEMS" };

    const secs     = Math.max(5, Number(r[im.roundSeconds] || 25));
    const deadline = nowMs_() + secs * 1000;

    r[im.round]          = round;
    r[im.state]          = "BIDDING";
    r[im.itemId]         = it.itemId;
    r[im.bidDeadlineTs]  = deadline;
    r[im.revealUntilTs]  = "";
    r[im.postUntilTs]    = "";
    r[im.updatedAt]      = nowMs_();
    setRow_(rooms, idx, r);

    addEvent_(roomId, "ROUND_STARTED", "", {
      round, itemId: it.itemId, itemName: it.name,
      publicHint: it.publicHint, deadline
    });
    return { ok: true };
  });
}

/**
 * 玩家送出出價（同回合可多次更新，結算取最後一次）
 * 會驗證房間處於 BIDDING 狀態且回合數一致
 */
function apiBid_(p) {
  const roomId   = String(p.roomId   || "").trim().toUpperCase();
  const playerId = String(p.playerId || "").trim();
  const round    = Number(p.round || 0);
  const bid      = Number(p.bid   || 0);
  if (!roomId)   return { ok: false, error: "MISSING_ROOMID" };
  if (!playerId) return { ok: false, error: "MISSING_PLAYERID" };
  if (!Number.isFinite(round) || round <= 0) return { ok: false, error: "BAD_ROUND" };
  if (!Number.isFinite(bid)   || bid < 0)    return { ok: false, error: "BAD_BID" };

  return withLock_(() => {
    // 驗證房間狀態
    const ss    = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx   = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };
    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(headers);
    const r   = getRow_(rooms, idx, rooms.getLastColumn());
    if (String(r[rim.state]) !== "BIDDING")
      return { ok: false, error: "NOT_IN_BIDDING" };
    if (Number(r[rim.round] || 0) !== round)
      return { ok: false, error: "ROUND_MISMATCH" };

    addEvent_(roomId, "BID", playerId, { round, bid });
    touchRoomUpdatedAt_(roomId);
    return { ok: true };
  });
}

/** 玩家使用卡牌（扣除卡牌數量 + 記錄事件） */
function apiPlayCard_(p) {
  const roomId   = String(p.roomId   || "").trim().toUpperCase();
  const playerId = String(p.playerId || "").trim();
  const round    = Number(p.round    || 0);
  const cardId   = String(p.cardId   || "").trim();
  if (!roomId)   return { ok: false, error: "MISSING_ROOMID" };
  if (!playerId) return { ok: false, error: "MISSING_PLAYERID" };
  if (!cardId)   return { ok: false, error: "MISSING_CARDID" };

  return withLock_(() => {
    const ss    = db_();
    const cards = ss.getSheetByName(TABS.CARDS);
    const all   = readAll_(cards);
    const im    = idxMap_(all.headers);

    for (let i = 0; i < all.rows.length; i++) {
      const r = all.rows[i];
      if (String(r[im.roomId])   !== roomId)   continue;
      if (String(r[im.playerId]) !== playerId) continue;
      if (String(r[im.cardId])   !== cardId)   continue;

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

/**
 * Host 重新開始遊戲（不離開房間）
 * 將房間重設為 LOBBY 狀態，重設所有玩家 budget/score，重新發牌
 */
function apiRestartGame_(p) {
  const roomId    = String(p.roomId    || "").trim().toUpperCase();
  const hostToken = String(p.hostToken || "").trim();
  const playerId  = String(p.playerId  || "").trim();
  if (!roomId) return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken && !playerId) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss    = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx   = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const roomHeaders = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(roomHeaders);
    const r   = getRow_(rooms, idx, rooms.getLastColumn());

    if (!canManageRoom_(ss, roomId, hostToken, playerId, r, rim))
      return { ok: false, error: "HOST_PERMISSION_DENIED" };

    const t = nowMs_();
    const startingBudget = Number(r[rim.startingBudget] || 100);

    // 重設房間狀態
    r[rim.state]          = "LOBBY";
    r[rim.round]          = 0;
    r[rim.itemId]         = "";
    r[rim.bidDeadlineTs]  = "";
    r[rim.revealUntilTs]  = "";
    r[rim.postUntilTs]    = "";
    r[rim.updatedAt]      = t;
    setRow_(rooms, idx, r);

    // 清除本房間所有舊卡牌（避免 restart 多次後 PlayerCards 表膨脹）
    const cards = ss.getSheetByName(TABS.CARDS);
    if (cards) {
      const cAll = readAll_(cards);
      const cim  = idxMap_(cAll.headers);
      for (let i = cAll.rows.length - 1; i >= 0; i--) {
        if (String(cAll.rows[i][cim.roomId]) === roomId) cards.deleteRow(i + 2);
      }
    }

    // 重設所有玩家的 budget 和 score，並重新發牌
    const players = ss.getSheetByName(TABS.PLAYERS);
    const pAll = readAll_(players);
    const pim  = idxMap_(pAll.headers);
    pAll.rows.forEach((pr, i) => {
      if (String(pr[pim.roomId]) !== roomId) return;
      const pRowIdx = i + 2;
      const pRow = getRow_(players, pRowIdx, players.getLastColumn());
      pRow[pim.budget]    = startingBudget;
      pRow[pim.score]     = 0;
      pRow[pim.updatedAt] = t;
      setRow_(players, pRowIdx, pRow);

      // 重新發牌（舊卡已在上面清除）
      seedCards_(ss, roomId, String(pr[pim.playerId]));
    });

    addEvent_(roomId, "GAME_RESTARTED", "", { startingBudget });
    SpreadsheetApp.flush();
    return { ok: true };
  });
}

/**
 * Host 重置資料庫（清空全部房間 / 玩家 / 事件 / 卡牌，重新 seed items）
 */
function apiResetDatabase_(p) {
  const roomId    = String(p.roomId    || "").trim().toUpperCase();
  const hostToken = String(p.hostToken || "").trim();
  if (!roomId)    return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken) return { ok: false, error: "MISSING_HOST_CREDENTIAL" };

  return withLock_(() => {
    const ss    = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx   = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };

    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const im  = idxMap_(headers);
    const row = getRow_(rooms, idx, rooms.getLastColumn());
    if (String(row[im.hostToken]) !== hostToken)
      return { ok: false, error: "HOST_TOKEN_MISMATCH" };

    // 只有名稱為 admin 的房主才能重置資料庫
    const players = ss.getSheetByName(TABS.PLAYERS);
    const pAll = readAll_(players);
    const pim  = idxMap_(pAll.headers);
    let hostName = "";
    pAll.rows.forEach(pr => {
      if (String(pr[pim.roomId]) !== roomId) return;
      if (String(pr[pim.isHost]) === "1") hostName = String(pr[pim.name] || "");
    });
    if (hostName.toLowerCase() !== "admin")
      return { ok: false, error: "只有名稱為 admin 的房主才能重置資料庫" };

    // 清空所有資料（保留 header）
    [TABS.ROOMS, TABS.PLAYERS, TABS.EVENTS, TABS.CARDS].forEach(name => {
      const sh = ss.getSheetByName(name);
      if (sh) clearSheetDataRows_(sh, sh.getLastColumn());
    });

    // 重新 seed items
    const items = ss.getSheetByName(TABS.ITEMS);
    if (items) {
      clearSheetDataRows_(items, items.getLastColumn());
      seedItems_(items);
    }

    return { ok: true };
  });
}


/* ══════════════════════════════════════════════════════════════════════════════
 *  內部輔助函式
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * 檢查呼叫者是否有 Host 權限
 * 優先比對 hostToken，其次比對 playerId 的 isHost 欄位
 */
function canManageRoom_(ss, roomId, hostToken, playerId, roomRow, rim) {
  if (hostToken && String(roomRow[rim.hostToken]) === hostToken) return true;
  if (!playerId) return false;

  const players = ss.getSheetByName(TABS.PLAYERS);
  if (!players) return false;

  const all = readAll_(players);
  const pim = idxMap_(all.headers);
  for (const pr of all.rows) {
    if (String(pr[pim.roomId])   !== roomId)   continue;
    if (String(pr[pim.playerId]) !== playerId) continue;
    return String(pr[pim.isHost]) === "1";
  }
  return false;
}

/**
 * 新增一筆事件到 Events 表（自動遞增 eventId）
 * 掃描整欄找最大 ID，避免 deleteRows 或多房間造成 ID 衝突
 */
function addEvent_(roomId, type, playerId, payloadObj) {
  const ss     = db_();
  const events = ss.getSheetByName(TABS.EVENTS);
  const t      = nowMs_();

  let maxId = 0;
  const lastRow = events.getLastRow();
  if (lastRow >= 2) {
    const col = events.getRange(2, 2, lastRow - 1, 1).getValues();
    col.forEach(r => { const v = Number(r[0] || 0); if (v > maxId) maxId = v; });
  }

  appendRow_(events, [
    roomId, maxId + 1, t, type, playerId || "", JSON.stringify(payloadObj || {})
  ]);
}

/** 更新房間的 updatedAt 時間戳（觸碰一下讓列表不會過期） */
function touchRoomUpdatedAt_(roomId) {
  const ss    = db_();
  const rooms = ss.getSheetByName(TABS.ROOMS);
  const idx   = findRoomRow_(rooms, roomId);
  if (idx < 0) return;
  const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
  const im = idxMap_(headers);
  const r  = getRow_(rooms, idx, rooms.getLastColumn());
  r[im.updatedAt] = nowMs_();
  setRow_(rooms, idx, r);
}

/** 從 Items 表隨機選一個物品 */
function pickRandomItem_(itemsSheet) {
  const all = readAll_(itemsSheet);
  if (!all.rows.length) return null;
  const r  = all.rows[Math.floor(Math.random() * all.rows.length)];
  const im = idxMap_(all.headers);
  return {
    itemId:      String(r[im.itemId]),
    name:        String(r[im.name]),
    publicHint:  String(r[im.publicHint]),
    hiddenValue: Number(r[im.hiddenValue] || 0)
  };
}

/** 在 Items 表中以 itemId 尋找特定物品 */
function findItem_(itemsSheet, itemId) {
  const all = readAll_(itemsSheet);
  const im  = idxMap_(all.headers);
  for (const r of all.rows) {
    if (String(r[im.itemId]) === itemId) {
      return {
        itemId:      String(r[im.itemId]),
        name:        String(r[im.name]),
        publicHint:  String(r[im.publicHint]),
        hiddenValue: Number(r[im.hiddenValue] || 0)
      };
    }
  }
  return null;
}

/**
 * 為新玩家發放初始卡牌
 * V1 三張：Peek（偷看）、Tax（加稅）、Shield（護盾）
 */
function seedCards_(ss, roomId, playerId) {
  const cards = ss.getSheetByName(TABS.CARDS);
  const seed = [
    [roomId, playerId, "C1", "偷看 Peek",    "窺探物品真實價值的區間（±5）", 1],
    [roomId, playerId, "C2", "加稅 Tax",     "若你得標，多付 20%（心理戰用）", 1],
    [roomId, playerId, "C3", "護盾 Shield",  "本回合若你分數為負，抵消一次",   1]
  ];
  cards.getRange(cards.getLastRow() + 1, 1, seed.length, 6).setValues(seed);
}

/**
 * 產生不重複的玩家 ID（P + 6 位隨機字元）
 */
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
  // 極端情況 fallback：加入時間戳片段
  return "P" + rid_(6) + String(nowMs_()).slice(-4);
}
