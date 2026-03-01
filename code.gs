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
    case "leaveRoom":       return apiLeaveRoom_(p);
    case "shuffleItems":    return apiShuffleItems_(p);
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

/** 取得所有物品定義 [itemId, name, publicHint, minValue, maxValue] — 300 項 */
function getItemDefs_() {
  return [
    ["ITM001","神秘古董 Antique","看起來很值錢 Looks valuable",10,60],
    ["ITM002","限量公仔 Figure","大家都說會漲價 May appreciate",15,50],
    ["ITM003","二手筆電 Laptop","螢幕有刮痕 Scratched screen",10,35],
    ["ITM004","奇怪的箱子 Mystery Box","搖一搖會響 Rattles when shaken",5,45],
    ["ITM005","無敵券 Power Ticket","規則外的力量 Rule-breaking",20,70],
    ["ITM006","金元寶 Gold Ingot","閃閃發光 Shiny gold",30,80],
    ["ITM007","紅包 Red Envelope","裡面放多少呢 How much inside?",5,40],
    ["ITM008","古董花瓶 Antique Vase","據說明朝的 Ming dynasty?",15,55],
    ["ITM009","翡翠手鐲 Jade Bracelet","通透碧綠 Translucent green",20,90],
    ["ITM010","舊照相機 Old Camera","底片還能用嗎 Film still works?",5,25],
    ["ITM011","名牌手錶 Watch","機芯很順 Smooth movement",30,95],
    ["ITM012","破損的吉他 Guitar","缺一根弦 Missing a string",3,15],
    ["ITM013","陳年紅酒 Wine","1990年份 Vintage 1990",20,75],
    ["ITM014","貓咪背包 Cat Bag","毛茸茸 Fluffy design",8,35],
    ["ITM015","VR眼鏡 VR Headset","鏡片有點花 Slightly worn",15,45],
    ["ITM016","黃金魚鉤 Gold Hook","能釣到龍王 Catch a dragon?",20,70],
    ["ITM017","過期零食 Expired Snack","過期三天 3 days expired",1,10],
    ["ITM018","鑽石耳環 Diamond Earring","小小一顆 Tiny but shiny",40,100],
    ["ITM019","手工肥皂 Handmade Soap","薰衣草味 Lavender scent",3,18],
    ["ITM020","古地圖 Old Map","標示寶藏位置 Treasure map?",15,65],
    ["ITM021","智慧音箱 Smart Speaker","只支援英文 English only",10,35],
    ["ITM022","生鏽寶劍 Rusty Sword","刻著奇怪文字 Strange runes",12,55],
    ["ITM023","蘭花盆栽 Orchid","開了三朵 Three blooms",10,40],
    ["ITM024","簽名棒球 Signed Ball","看不清簽名 Unreadable sig",15,75],
    ["ITM025","機械鍵盤 Mech Keyboard","Cherry軟軸 Cherry switches",20,55],
    ["ITM026","迷你無人機 Mini Drone","續航5分鐘 5min battery",12,45],
    ["ITM027","神秘藥水 Potion","喝了變強 Drink to power up",5,30],
    ["ITM028","油畫作品 Oil Painting","署名看不清 Unsigned?",25,90],
    ["ITM029","復古收音機 Retro Radio","還能收FM Can still tune FM",8,30],
    ["ITM030","水晶球 Crystal Ball","閃閃發光 Sparkling bubbles",15,55],
    ["ITM031","狗狗雨衣 Dog Raincoat","XL號 For large dogs",5,20],
    ["ITM032","珍珠項鍊 Pearl Necklace","大小不一 Uneven pearls",20,80],
    ["ITM033","二手遊戲機 Console","附兩個手把 2 controllers",20,60],
    ["ITM034","瑜伽墊 Yoga Mat","微有痕跡 Slightly used",3,15],
    ["ITM035","銀色懷錶 Pocket Watch","偶爾慢幾分 Sometimes slow",15,55],
    ["ITM036","神秘信封 Mystery Letter","火漆封口 Wax sealed",5,45],
    ["ITM037","有機蜂蜜 Organic Honey","農場直送 Farm fresh",8,28],
    ["ITM038","電競滑鼠 Gaming Mouse","RGB全開 Full RGB",12,40],
    ["ITM039","陶瓷茶壺 Tea Pot","日式風格 Japanese style",15,50],
    ["ITM040","登山背包 Hiking Pack","60L防水 60L waterproof",18,55],
    ["ITM041","古銅指南針 Compass","磁針還能動 Needle works",10,35],
    ["ITM042","手繪撲克牌 Art Cards","每張都獨特 Each one unique",12,45],
    ["ITM043","桌上風扇 Desk Fan","三段風速 3 speeds",3,18],
    ["ITM044","鍍金相框 Gold Frame","適合全家福 For family photo",5,25],
    ["ITM045","真皮皮夾 Leather Wallet","義大利進口 Italian leather",20,65],
    ["ITM046","迷你投影機 Projector","畫質勉強 Barely watchable",15,55],
    ["ITM047","竹編籃子 Bamboo Basket","手工製作 Handcrafted",5,22],
    ["ITM048","大理石棋盤 Marble Chess","附完整棋子 Full set",20,70],
    ["ITM049","香氛蠟燭組 Candle Set","六種味道 6 scents",8,30],
    ["ITM050","折疊腳踏車 Folding Bike","輪胎需打氣 Needs air",25,80],
    ["ITM051","老式打字機 Typewriter","鍵帽有些卡 Some keys stuck",15,50],
    ["ITM052","太陽能充電器 Solar Charger","陰天很慢 Slow on cloudy days",8,35],
    ["ITM053","手沖咖啡組 Pour Over Set","含磨豆機 Includes grinder",12,45],
    ["ITM054","毛筆套組 Brush Set","書法必備 Calligraphy kit",5,25],
    ["ITM055","藍芽喇叭 BT Speaker","防水音質好 Waterproof",15,50],
    ["ITM056","仙人掌 Cactus","三年沒澆水 3yr no water",3,15],
    ["ITM057","皮革公事包 Briefcase","商務質感 Business style",25,70],
    ["ITM058","復古桌燈 Retro Lamp","鎢絲燈泡 Tungsten bulb",10,35],
    ["ITM059","木雕擺件 Wood Carving","手工貓頭鷹 Carved owl",15,55],
    ["ITM060","行李秤 Luggage Scale","旅行神器 Travel essential",3,18],
    ["ITM061","紫砂茶壺 Yixing Teapot","大師印章 Master stamp",40,100],
    ["ITM062","自動餵食器 Auto Feeder","可設定時間 Programmable",10,40],
    ["ITM063","古典吊燈 Chandelier","需重新接線 Needs rewiring",18,60],
    ["ITM064","天文望遠鏡 Telescope","能看月坑 See moon craters",25,80],
    ["ITM065","手工皮帶 Leather Belt","牛皮銅扣 Brass buckle",12,45],
    ["ITM066","老唱片 Vinyl Record","經典專輯 Classic album",15,55],
    ["ITM067","保溫便當盒 Lunch Box","三層式 3-tier insulated",5,25],
    ["ITM068","迷彩帳篷 Camo Tent","雙人帳 2-person",20,65],
    ["ITM069","骨董電話 Rotary Phone","轉盤式 Vintage rotary",15,50],
    ["ITM070","水彩顏料 Watercolors","24色專業 24 pro colors",8,35],
    ["ITM071","黑膠唱片機 Turntable","能正常播放 Fully working",30,90],
    ["ITM072","編織毛毯 Wool Blanket","手工羊毛 Hand-knit wool",10,40],
    ["ITM073","象棋組 Chess Set","檀木材質 Sandalwood",18,60],
    ["ITM074","螢幕掛燈 Monitor Light","護眼不反光 Anti-glare",10,35],
    ["ITM075","露營椅 Camp Chair","鋁合金輕便 Aluminum frame",8,30],
    ["ITM076","琥珀墜飾 Amber Pendant","裡面有蟲 Bug inside",20,80],
    ["ITM077","按摩槍 Massage Gun","三段力道 3 intensity",15,55],
    ["ITM078","傳統摺扇 Folding Fan","檀香木 Sandalwood frame",8,35],
    ["ITM079","電子書閱讀器 E-Reader","螢幕有線 Screen line",18,60],
    ["ITM080","手工果醬 Handmade Jam","草莓限定 Strawberry ltd",3,20],
    ["ITM081","黃銅燭台 Brass Candle","成對有氣氛 Pair, romantic",12,45],
    ["ITM082","運動水壺 Sports Bottle","保冷24小時 24hr cold",5,25],
    ["ITM083","老鐘擺鐘 Pendulum Clock","每小時響 Chimes hourly",25,80],
    ["ITM084","手作陶杯 Ceramic Cup","釉色獨一 Unique glaze",8,30],
    ["ITM085","防水相機 Dive Camera","水下攝影 Underwater",20,70],
    ["ITM086","薄荷精油 Mint Oil","提神醒腦 Refreshing",3,18],
    ["ITM087","紀念幣套組 Coin Set","十二生肖 12 zodiac",30,90],
    ["ITM088","羽毛筆 Quill Pen","附墨水瓶 Includes ink",5,25],
    ["ITM089","空拍機零件 Drone Parts","不確定能裝 May not fit",2,15],
    ["ITM090","日式風鈴 Wind Chime","玻璃清脆 Glass, crisp",5,20],
    ["ITM091","皮革日記本 Leather Diary","附鎖扣 With lock",10,35],
    ["ITM092","迷你魚缸 Mini Aquarium","含過濾LED Filter+LED",10,40],
    ["ITM093","古董懷爐 Antique Warmer","銅製 Copper, works",12,50],
    ["ITM094","摺紙藝術 Origami Art","1000紙鶴 1000 cranes",8,35],
    ["ITM095","吉他撥片 Guitar Pick","玳瑁材質 Tortoiseshell",5,25],
    ["ITM096","景泰藍花瓶 Cloisonne","掐絲琺瑯 Enamel craft",35,100],
    ["ITM097","貓眼石戒指 Cat-eye Ring","神秘光芒 Mysterious glow",25,85],
    ["ITM098","老地球儀 Globe","蘇聯版本 Soviet era",15,55],
    ["ITM099","機械鋼筆 Fountain Pen","14K金筆尖 14K gold nib",20,70],
    ["ITM100","龍銀 Silver Coin","日治銀幣 Colonial era",40,100],
    ["ITM101","迷你烤箱 Mini Oven","能烤吐司 Toasts bread",8,30],
    ["ITM102","絲綢圍巾 Silk Scarf","真絲手工 Handmade silk",15,55],
    ["ITM103","古銅望遠鏡 Brass Spyglass","海盜風 Pirate style",20,65],
    ["ITM104","竹製滑板 Bamboo Board","環保設計 Eco-friendly",12,45],
    ["ITM105","手工巧克力 Artisan Choco","限量口味 Limited flavor",5,25],
    ["ITM106","老舊地毯 Vintage Rug","波斯風格 Persian style",25,85],
    ["ITM107","LED植物燈 Grow Light","全光譜 Full spectrum",8,35],
    ["ITM108","骨瓷茶杯 Bone China Cup","英式下午茶 English tea",12,45],
    ["ITM109","木質音樂盒 Music Box","旋轉芭蕾 Ballet dancer",10,40],
    ["ITM110","露營燈 Lantern","太陽能充電 Solar powered",6,25],
    ["ITM111","手工皮鞋 Leather Shoes","義大利製 Italian made",30,90],
    ["ITM112","古董鐘 Desk Clock","需上發條 Needs winding",15,50],
    ["ITM113","望遠鏡 Binoculars","10x50倍率 10x50 zoom",20,65],
    ["ITM114","茶葉禮盒 Tea Gift Box","高山烏龍 Oolong tea",10,40],
    ["ITM115","機械錶 Mechanical Watch","自動上鏈 Auto-wind",35,95],
    ["ITM116","陶笛 Ocarina","12孔 12-hole ceramic",5,22],
    ["ITM117","皮革手環 Leather Cuff","編織設計 Braided design",8,30],
    ["ITM118","古董鏡 Antique Mirror","木框雕花 Carved frame",18,60],
    ["ITM119","電動牙刷 E-Toothbrush","充電式 Rechargeable",6,25],
    ["ITM120","手工蠟封 Wax Seal Kit","含印章 With stamp",5,22],
    ["ITM121","黑曜石 Obsidian","火山玻璃 Volcanic glass",10,40],
    ["ITM122","復古打火機 Retro Lighter","黃銅機身 Brass body",8,35],
    ["ITM123","砧板 Cutting Board","相思木 Acacia wood",5,20],
    ["ITM124","乾燥花束 Dried Flowers","永生花 Everlasting",8,30],
    ["ITM125","銅製書擋 Bookends","馬頭造型 Horse head",12,45],
    ["ITM126","迷你吸塵器 Mini Vacuum","桌面用 Desktop size",5,22],
    ["ITM127","日式便當 Bento Box","漆器風格 Lacquerware",10,35],
    ["ITM128","手搖磨豆機 Hand Grinder","陶瓷磨芯 Ceramic burr",8,30],
    ["ITM129","摺疊傘 Folding Umbrella","抗UV UV-proof",3,18],
    ["ITM130","木製拼圖 Wood Puzzle","1000片 1000 pieces",8,30],
    ["ITM131","香水 Perfume","花香調 Floral notes",15,55],
    ["ITM132","陶製花瓶 Ceramic Vase","手拉坯 Hand-thrown",10,40],
    ["ITM133","行動電源 Power Bank","20000mAh capacity",8,30],
    ["ITM134","棉麻圍裙 Apron","廚師風格 Chef style",5,20],
    ["ITM135","銅鈴 Brass Bell","清脆響亮 Clear ring",3,15],
    ["ITM136","手工紙 Handmade Paper","花瓣嵌入 Petal embedded",5,22],
    ["ITM137","竹蒸籠 Bamboo Steamer","三層 3-tier",8,28],
    ["ITM138","懷舊糖果 Retro Candy","鐵罐裝 Tin can",3,15],
    ["ITM139","羊毛氈 Wool Felt","手作材料 Craft supply",5,20],
    ["ITM140","擴香瓶 Diffuser","藤條式 Reed type",8,30],
    ["ITM141","金箔畫 Gold Leaf Art","手工貼金 Hand gilded",25,80],
    ["ITM142","竹木筷組 Chopsticks Set","雷射雕刻 Laser engraved",5,22],
    ["ITM143","手提燈 Hand Lantern","復古造型 Vintage style",10,35],
    ["ITM144","貝殼項鍊 Shell Necklace","海邊拾取 Beach found",5,25],
    ["ITM145","銀髮簪 Silver Hairpin","古典美 Classical beauty",15,55],
    ["ITM146","迷你盆景 Mini Bonsai","五年樹齡 5-year-old",10,40],
    ["ITM147","書法卷軸 Calligraphy Scroll","大師真跡? Master's work?",20,75],
    ["ITM148","老照片 Old Photo","黑白年代 B&W era",5,25],
    ["ITM149","手工編織袋 Woven Bag","天然材料 Natural fiber",8,30],
    ["ITM150","銅製天秤 Brass Scale","科學風格 Scientific look",15,50],
    ["ITM151","玻璃筆 Glass Pen","彩虹色 Rainbow color",10,40],
    ["ITM152","手工餅乾 Cookies","禮盒裝 Gift box",3,18],
    ["ITM153","古典書籤 Bookmark","金屬雕花 Metal carved",3,15],
    ["ITM154","胡桃木盒 Walnut Box","珠寶收納 Jewelry box",12,45],
    ["ITM155","老式眼鏡 Vintage Glasses","圓框金屬 Round metal",10,35],
    ["ITM156","手工冰淇淋 Artisan Gelato","義式風味 Italian style",3,18],
    ["ITM157","銅製門環 Brass Knocker","獅頭造型 Lion head",10,40],
    ["ITM158","迷你地球儀 Mini Globe","桌上型 Desktop size",8,30],
    ["ITM159","手繪明信片 Postcards","一組十張 Set of 10",3,15],
    ["ITM160","陶塑擺飾 Clay Figure","動物造型 Animal shape",8,30],
    ["ITM161","紫水晶 Amethyst","天然原石 Natural raw",15,55],
    ["ITM162","老軍帽 Military Cap","二戰風格 WWII style",12,45],
    ["ITM163","手工鈎針 Crochet Set","附毛線 With yarn",5,22],
    ["ITM164","銅版畫 Etching Print","限量簽名 Signed limited",20,65],
    ["ITM165","竹蜻蜓 Bamboo Copter","手工製 Handcrafted",2,12],
    ["ITM166","錫杯 Pewter Cup","中世紀風 Medieval style",10,35],
    ["ITM167","手工領帶 Handmade Tie","真絲 Pure silk",12,45],
    ["ITM168","木雕面具 Wood Mask","部落風格 Tribal style",15,55],
    ["ITM169","手工香包 Sachet","中藥材料 Herbal scent",3,15],
    ["ITM170","老式秤錘 Old Weight","黃銅製 Brass made",8,30],
    ["ITM171","珊瑚飾品 Coral Jewelry","粉紅色 Pink coral",20,70],
    ["ITM172","棉花糖機 Cotton Candy","迷你型 Mini machine",10,40],
    ["ITM173","手工燈罩 Lampshade","彩繪玻璃 Stained glass",15,50],
    ["ITM174","木製陀螺 Spinning Top","手工車削 Hand-turned",3,15],
    ["ITM175","象牙白棋 Ivory-look Chess","仿象牙 Faux ivory",20,65],
    ["ITM176","手搖風琴 Hurdy-Gurdy","需要調音 Needs tuning",30,90],
    ["ITM177","印度薰香 Indian Incense","檀香味 Sandalwood",3,15],
    ["ITM178","手工橡皮章 Rubber Stamp","可客製 Customizable",5,22],
    ["ITM179","老鑰匙 Antique Key","不知開什麼 Unknown lock",5,25],
    ["ITM180","琉璃珠 Glass Bead","手工吹製 Hand-blown",8,35],
    ["ITM181","手工麵條 Handmade Noodle","日曬乾燥 Sun-dried",3,15],
    ["ITM182","老式計算機 Calculator","機械式 Mechanical",10,40],
    ["ITM183","手工石鹼 Stone Soap","火山泥 Volcanic mud",3,18],
    ["ITM184","銀質書籤 Silver Bookmark","花紋精緻 Fine pattern",8,30],
    ["ITM185","竹搖椅 Bamboo Rocker","手編座面 Hand-woven seat",20,65],
    ["ITM186","瑪瑙戒指 Agate Ring","天然紋路 Natural grain",12,45],
    ["ITM187","手沖壺 Gooseneck Kettle","細嘴 Pour-over style",10,35],
    ["ITM188","老舊郵票 Vintage Stamps","一頁20枚 Sheet of 20",8,35],
    ["ITM189","珐瑯別針 Enamel Pin","限量設計 Limited edition",3,18],
    ["ITM190","手工燒酒 Craft Spirits","小批量 Small batch",10,40],
    ["ITM191","木刻版畫 Woodblock Print","浮世繪風 Ukiyo-e style",15,55],
    ["ITM192","老玩具車 Toy Car","鐵皮製 Tin plate",8,30],
    ["ITM193","手工涼鞋 Sandals","皮革手縫 Leather handmade",12,40],
    ["ITM194","迷你望遠鏡 Pocket Scope","摺疊式 Foldable",5,22],
    ["ITM195","手工牛軋糖 Nougat","花生口味 Peanut flavor",3,15],
    ["ITM196","青銅壺 Bronze Pot","綠鏽古味 Green patina",20,65],
    ["ITM197","手繪陶盤 Painted Plate","花鳥圖案 Bird & flower",10,40],
    ["ITM198","老式鬧鐘 Alarm Clock","雙鈴 Twin bell",5,22],
    ["ITM199","麻繩吊籃 Macrame Hanger","波西米亞 Boho style",5,22],
    ["ITM200","手工護唇膏 Lip Balm","蜂蠟配方 Beeswax formula",2,12],
    ["ITM201","銀質袖扣 Silver Cufflinks","紳士風 Gentleman style",15,50],
    ["ITM202","老算盤 Abacus","花梨木 Rosewood frame",10,40],
    ["ITM203","手工貝雕 Shell Carving","精細工藝 Fine craft",15,55],
    ["ITM204","復古鉛筆盒 Pencil Case","鐵製 Tin case",3,15],
    ["ITM205","手工乳酪 Artisan Cheese","熟成三月 3-month aged",5,25],
    ["ITM206","老式煤油燈 Kerosene Lamp","附燈芯 With wick",10,35],
    ["ITM207","手作拇指琴 Kalimba","17鍵 17-key",8,30],
    ["ITM208","玉石擺件 Jade Ornament","和田玉? Hetian jade?",25,85],
    ["ITM209","手工紮染 Tie-dye Cloth","天然染料 Natural dye",5,22],
    ["ITM210","老漫畫 Vintage Comic","初版? First edition?",10,45],
    ["ITM211","銅鏡 Bronze Mirror","仿漢代 Han dynasty style",15,50],
    ["ITM212","手工年糕 Rice Cake","傳統配方 Traditional",3,15],
    ["ITM213","復古胸針 Vintage Brooch","寶石鑲嵌 Gemstone inlay",12,45],
    ["ITM214","竹笛 Bamboo Flute","手工調音 Hand-tuned",5,25],
    ["ITM215","手工繡花 Embroidery","蘇繡風格 Suzhou style",15,55],
    ["ITM216","老相框 Vintage Frame","銅製雕花 Brass carved",8,30],
    ["ITM217","手工豆腐乳 Fermented Tofu","傳統釀造 Traditional",2,12],
    ["ITM218","紅木書架 Rosewood Shelf","迷你桌上型 Desktop mini",20,65],
    ["ITM219","老式電扇 Vintage Fan","鐵製 Iron blade",12,40],
    ["ITM220","手工毛線帽 Knit Hat","冬季限定 Winter special",5,20],
    ["ITM221","玻璃花瓶 Glass Vase","穆拉諾 Murano style",15,55],
    ["ITM222","老火車模型 Train Model","HO比例 HO scale",20,65],
    ["ITM223","手工辣醬 Hot Sauce","魔鬼椒 Ghost pepper",3,18],
    ["ITM224","檜木名片盒 Cypress Case","芳香 Aromatic",8,30],
    ["ITM225","手工刺繡包 Embroidery Bag","民族風 Ethnic style",10,40],
    ["ITM226","老銅鎖 Brass Lock","附鑰匙 With key",8,30],
    ["ITM227","紫檀木梳 Sandalwood Comb","防靜電 Anti-static",10,35],
    ["ITM228","手工酸菜 Sauerkraut","自然發酵 Natural ferment",2,12],
    ["ITM229","錫製酒壺 Pewter Flask","隨身攜帶 Pocket size",10,35],
    ["ITM230","古典吉他弦 Guitar Strings","尼龍 Nylon set",3,15],
    ["ITM231","手工鑰匙圈 Keychain","皮革編織 Leather braided",3,15],
    ["ITM232","月光石 Moonstone","藍色光暈 Blue sheen",12,45],
    ["ITM233","竹製筷架 Chopstick Rest","一組六個 Set of 6",3,15],
    ["ITM234","手工肉乾 Jerky","黑胡椒味 Black pepper",3,18],
    ["ITM235","銅質名片夾 Card Holder","商務風 Business style",5,22],
    ["ITM236","老式顯微鏡 Microscope","教學用 Educational",20,65],
    ["ITM237","手工果乾 Dried Fruit","無添加 No additives",3,15],
    ["ITM238","海星標本 Starfish","乾燥處理 Preserved",3,15],
    ["ITM239","手工花冠 Flower Crown","永生花 Preserved flowers",5,22],
    ["ITM240","磁鐵玩具 Magnet Toy","創意造型 Creative shapes",5,20],
    ["ITM241","手工蛋捲 Egg Roll","古早味 Traditional taste",3,15],
    ["ITM242","老式放大鏡 Magnifier","黃銅柄 Brass handle",5,22],
    ["ITM243","手工木湯匙 Wood Spoon","橄欖木 Olive wood",3,15],
    ["ITM244","復古手帕 Vintage Hanky","蕾絲邊 Lace trimmed",3,18],
    ["ITM245","手工陶鈴 Clay Bell","清脆聲 Clear tone",3,15],
    ["ITM246","古董火柴盒 Match Box","收藏用 Collectible",2,12],
    ["ITM247","手工魚丸 Fish Ball","新鮮現做 Freshly made",2,12],
    ["ITM248","老式筆筒 Pen Holder","竹製 Bamboo carved",5,22],
    ["ITM249","手工米酒 Rice Wine","傳統釀造 Traditional brew",5,25],
    ["ITM250","皮影戲偶 Shadow Puppet","牛皮 Cowhide craft",15,50],
    ["ITM251","鍛鐵燭台 Iron Candle","工業風 Industrial style",10,35],
    ["ITM252","手工鳳梨酥 Pineapple Cake","土鳳梨 Native pineapple",3,18],
    ["ITM253","古董拐杖 Walking Cane","藤製龍頭 Dragon handle",20,65],
    ["ITM254","玻璃彈珠 Marbles","一袋50顆 Bag of 50",2,12],
    ["ITM255","手工洗髮餅 Shampoo Bar","草本配方 Herbal formula",3,15],
    ["ITM256","老式墨水瓶 Ink Bottle","寶藍色 Royal blue",5,22],
    ["ITM257","手工扇墜 Fan Charm","玉石 Jade pendant",8,30],
    ["ITM258","黃楊木雕 Boxwood Carving","微型佛像 Mini Buddha",20,65],
    ["ITM259","手工花生糖 Peanut Candy","芝麻口味 Sesame flavor",2,12],
    ["ITM260","復古郵筒模型 Mailbox Model","鐵皮 Tin plate",5,22],
    ["ITM261","手工陶壺 Clay Pot","柴燒 Wood-fired",15,50],
    ["ITM262","老式溫度計 Thermometer","水銀 Mercury glass",5,22],
    ["ITM263","手工鞋墊 Insole","竹炭 Bamboo charcoal",2,12],
    ["ITM264","龍涎香 Ambergris","真偽未知 Authenticity unknown",30,100],
    ["ITM265","手工刮痧板 Gua Sha","牛角製 Buffalo horn",5,22],
    ["ITM266","老式煙斗 Vintage Pipe","石楠木 Briar wood",15,50],
    ["ITM267","手工棉被 Cotton Quilt","手工彈棉 Hand-carded",20,60],
    ["ITM268","銅製獎杯 Brass Trophy","無銘文 Unmarked",10,35],
    ["ITM269","手工太妃糖 Toffee","海鹽焦糖 Salted caramel",2,12],
    ["ITM270","老式收銀機 Cash Register","裝飾用 Decorative",25,80],
    ["ITM271","手工蒲扇 Palm Fan","夏日必備 Summer essential",2,10],
    ["ITM272","玉簪 Jade Hairpin","白玉 White jade",15,55],
    ["ITM273","手工雪花酥 Nougat Crisp","杏仁口味 Almond flavor",3,15],
    ["ITM274","舊版桌遊 Board Game","絕版 Out of print",10,45],
    ["ITM275","手工掛毯 Tapestry","幾何圖案 Geometric pattern",15,50],
    ["ITM276","鐵壺 Iron Kettle","南部鐵器 Nambu ironware",25,80],
    ["ITM277","手工咖哩粉 Curry Powder","自製配方 Homemade blend",3,15],
    ["ITM278","銅製羅盤 Brass Compass","風水用 Feng shui tool",10,40],
    ["ITM279","手工蜜餞 Preserved Fruit","梅子口味 Plum flavor",2,12],
    ["ITM280","老式手搖鑽 Hand Drill","木工用 Woodworking",8,30],
    ["ITM281","手工玻璃杯 Blown Glass","彩色 Colorful",8,30],
    ["ITM282","古董撲克牌 Antique Cards","完整一副 Full deck",5,25],
    ["ITM283","手工竹蓆 Bamboo Mat","清涼 Cool in summer",5,20],
    ["ITM284","老式電表 Voltmeter","銅製指針 Copper needle",10,35],
    ["ITM285","手工魚酥 Fish Crisp","傳統零食 Traditional snack",2,12],
    ["ITM286","銅製號角 Brass Horn","裝飾用 Decorative",15,50],
    ["ITM287","手工花枕 Flower Pillow","刺繡 Embroidered",8,30],
    ["ITM288","老式縫紉機 Sewing Machine","腳踏式 Pedal type",25,80],
    ["ITM289","手工牛角梳 Horn Comb","水牛角 Buffalo horn",5,22],
    ["ITM290","田黃石 Tianhuang Stone","壽山石? Shoushan stone?",35,100],
    ["ITM291","手工竹杯 Bamboo Cup","天然漆 Natural lacquer",3,15],
    ["ITM292","老式油燈 Oil Lamp","黃銅 Brass body",8,30],
    ["ITM293","手工肉鬆 Pork Floss","手工炒製 Hand-fried",3,15],
    ["ITM294","琉璃擺件 Glass Ornament","多色 Multicolored",15,50],
    ["ITM295","手工綠豆糕 Mung Cake","傳統糕點 Traditional",2,12],
    ["ITM296","古董門把 Antique Handle","鑄鐵花紋 Cast iron",8,30],
    ["ITM297","手工藕粉 Lotus Powder","杭州出產 From Hangzhou",3,15],
    ["ITM298","銅製望遠鏡 Brass Scope","三段伸縮 3-section",15,50],
    ["ITM299","手工芋圓 Taro Balls","Q彈 Chewy texture",2,12],
    ["ITM300","傳家寶玉 Heirloom Jade","家族傳承 Family legacy",50,100]
  ];
}

/**
 * 寫入初始物品資料（300 項）到 Items 工作表
 * 每次呼叫時 hiddenValue 會在 minValue ~ maxValue 之間重新隨機
 *
 * Sheet 欄位：itemId | name | publicHint | hiddenValue | minValue | maxValue
 * - minValue / maxValue 會一併寫入 Sheet，這樣之後 apiShuffleItems_ 可以
 *   直接從 Sheet 讀取範圍來重新隨機化，不需依賴 code 中的定義
 * - 使用者也可以在 Google Sheet 直接修改物品的名稱、提示、min/max
 */
function seedItems_(itemsSheet) {
  const defs = getItemDefs_();
  // 將 [id, name, hint, min, max] 轉換為 [id, name, hint, randomizedValue, min, max]
  const seed = defs.map(d => {
    const hv = d[3] + Math.floor(Math.random() * (d[4] - d[3] + 1));
    return [d[0], d[1], d[2], hv, d[3], d[4]];
  });
  // 批次寫入所有列（比逐列寫入快 100 倍以上）
  itemsSheet.getRange(2, 1, seed.length, 6).setValues(seed);
}


/**
 * Host 重新隨機化所有物品的 hiddenValue（真實價值）
 *
 * volatility 參數控制隨機範圍有多大：
 *   "low"  (0.3) — 金額集中在 minValue~maxValue 的中位數附近，波動小（適合新手場）
 *   "mid"  (0.6) — 適度變化，預設值
 *   "high" (1.0) — 使用完整 min~max 範圍，波動最大（刺激場）
 *
 * 計算方式：以 (min+max)/2 為中心點，向兩側展開 halfRange × volMult 的範圍
 * 例如 min=10, max=60, volMult=0.6:
 *   mid=35, halfRange=25×0.6=15 → 隨機範圍 [20, 50]
 *
 * 注意：min/max 是從 Sheet 的 minValue/maxValue 欄位讀取的，
 *       可以在 Google Sheet 直接修改而不需要改 code
 */
function apiShuffleItems_(p) {
  const roomId    = String(p.roomId   || "").trim();
  const hostToken = String(p.hostToken|| "").trim();
  const vol       = String(p.volatility || "mid").trim();
  if (!roomId)    return { ok: false, error: "MISSING_ROOMID" };
  if (!hostToken) return { ok: false, error: "HOST_ONLY" };

  // 波動倍率對照表：low=30%, mid=60%, high=100% 的 min~max 範圍
  const volMult = { low: 0.3, mid: 0.6, high: 1.0 }[vol] || 0.6;

  return withLock_(() => {
    const ss    = db_();
    const rooms = ss.getSheetByName(TABS.ROOMS);
    const idx   = findRoomRow_(rooms, roomId);
    if (idx < 0) return { ok: false, error: "ROOM_NOT_FOUND" };
    const headers = rooms.getRange(1, 1, 1, rooms.getLastColumn()).getValues()[0];
    const rim = idxMap_(headers);
    const r   = getRow_(rooms, idx, rooms.getLastColumn());
    if (String(r[rim.hostToken]) !== hostToken) return { ok: false, error: "NOT_HOST" };

    // 從 Items Sheet 讀取所有物品並重新隨機化 hiddenValue
    const items = ss.getSheetByName(TABS.ITEMS);
    if (!items) return { ok: false, error: "ITEMS_NOT_FOUND" };
    const iAll = readAll_(items);
    if (!iAll.rows.length) return { ok: false, error: "NO_ITEMS" };
    const iim = idxMap_(iAll.headers);

    // 逐項從 Sheet 的 minValue / maxValue 計算新的 hiddenValue
    iAll.rows.forEach((row, i) => {
      const mn = Number(row[iim.minValue] || 1);   // 從 Sheet 讀取最小值
      const mx = Number(row[iim.maxValue] || mn);  // 從 Sheet 讀取最大值
      if (mx <= mn) return;                        // 無效範圍 → 跳過

      // 以中位數為中心，根據 volatility 計算實際隨機範圍
      const mid = (mn + mx) / 2;
      const halfRange = ((mx - mn) / 2) * volMult;
      const lo = Math.max(1, Math.round(mid - halfRange));  // 確保不低於 1
      const hi = Math.round(mid + halfRange);
      const newVal = lo + Math.floor(Math.random() * (hi - lo + 1));

      // 更新該行的 hiddenValue 欄位
      const pRow = getRow_(items, i + 2, items.getLastColumn());
      pRow[iim.hiddenValue] = newVal;
      setRow_(items, i + 2, pRow);
    });

    SpreadsheetApp.flush();  // 強制寫入 Sheet
    addEvent_(roomId, "ITEMS_SHUFFLED", "", { volatility: vol });
    return { ok: true, volatility: vol };
  });
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
      "itemId", "name", "publicHint", "hiddenValue", "minValue", "maxValue"
    ]);

    ensureSheet_(ss, TABS.CARDS, [
      "roomId", "playerId", "cardId", "name", "desc", "count"
    ]);

    // Seed items — 每次 Setup 都重新寫入（清除舊資料 + 寫入最新 300 項含 minValue/maxValue）
    const items = ss.getSheetByName(TABS.ITEMS);
    // 用 clearContent 而非 deleteRows — deleteRows 在只有 header 時會報錯
    // 「你無法刪除所有非凍結的列」
    const lastItemRow = items.getLastRow();
    if (lastItemRow > 1) {
      items.getRange(2, 1, lastItemRow - 1, items.getLastColumn()).clearContent();
    }
    seedItems_(items);

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
 * 玩家離開房間（刪除 Players + Cards 中該玩家的資料）
 * 若房間內無剩餘玩家，自動清除房間及相關資料
 */
function apiLeaveRoom_(p) {
  const roomId   = String(p.roomId   || "").trim();
  const playerId = String(p.playerId || "").trim();
  if (!roomId)   return { ok: false, error: "MISSING_ROOMID" };
  if (!playerId) return { ok: false, error: "MISSING_PLAYERID" };

  return withLock_(() => {
    const ss      = db_();
    const rooms   = ss.getSheetByName(TABS.ROOMS);
    const players = ss.getSheetByName(TABS.PLAYERS);
    const cards   = ss.getSheetByName(TABS.CARDS);
    const events  = ss.getSheetByName(TABS.EVENTS);

    // 刪除該玩家（從下往上刪避免 row shift 問題）
    const pAll = readAll_(players);
    const pim  = idxMap_(pAll.headers);
    for (let i = pAll.rows.length - 1; i >= 0; i--) {
      const pr = pAll.rows[i];
      if (String(pr[pim.roomId]) !== roomId) continue;
      if (String(pr[pim.playerId]) !== playerId) continue;
      players.deleteRow(i + 2);
    }

    // 刪除該玩家的卡片
    if (cards) {
      const cAll = readAll_(cards);
      const cim  = idxMap_(cAll.headers);
      for (let i = cAll.rows.length - 1; i >= 0; i--) {
        const cr = cAll.rows[i];
        if (String(cr[cim.roomId]) !== roomId) continue;
        if (String(cr[cim.playerId]) !== playerId) continue;
        cards.deleteRow(i + 2);
      }
    }

    // 重新計算房間內剩餘玩家數
    const pAllAfter = readAll_(players);
    const pimAfter  = idxMap_(pAllAfter.headers);
    let remaining = 0;
    pAllAfter.rows.forEach(pr => {
      if (String(pr[pimAfter.roomId]) === roomId) remaining++;
    });

    // 若無剩餘玩家 → 清除整個房間資料
    if (remaining === 0) {
      // 刪除房間列
      const roomRowIdx = findRoomRow_(rooms, roomId);
      if (roomRowIdx > 0) rooms.deleteRow(roomRowIdx);

      // 刪除該房間所有事件（從下往上刪）
      if (events) {
        const eAll = readAll_(events);
        const eim  = idxMap_(eAll.headers);
        for (let i = eAll.rows.length - 1; i >= 0; i--) {
          if (String(eAll.rows[i][eim.roomId]) === roomId) events.deleteRow(i + 2);
        }
      }

      // 刪除該房間所有卡片（可能還有非該玩家的殘留）
      if (cards) {
        const cAll2 = readAll_(cards);
        const cim2  = idxMap_(cAll2.headers);
        for (let i = cAll2.rows.length - 1; i >= 0; i--) {
          if (String(cAll2.rows[i][cim2.roomId]) === roomId) cards.deleteRow(i + 2);
        }
      }
    }

    SpreadsheetApp.flush();
    return { ok: true, roomCleared: remaining === 0 };
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
 * 被 apiResolve_（Host 手動）、tryAutoTransition_（自動計時）、apiBid_（全員出價完畢）呼叫
 *
 * 流程概述：
 *   1. 讀取本回合物品的 hiddenValue（真實價值）
 *   2. 掃描 Events 收集本回合所有 BID 和 CARD_PLAYED 紀錄
 *   3. 找出出價最高且 ≤ 預算的玩家為得標者（Tie-break：先出價者優先）
 *   4. 套用卡牌效果：Tax（加稅 +20%）、Shield（負分抵消為 0）
 *   5. 計算分數 = hiddenValue − actualCost，更新得標者的 budget / score
 *   6. 將房間狀態切換為 REVEAL（停留 REVEAL_DURATION_MS 毫秒）
 *   7. 寫入 ROUND_RESOLVED 事件，供前端渲染結果面板
 *
 * @param {Spreadsheet} ss    - 資料庫 Spreadsheet 物件
 * @param {Sheet}       rooms - Rooms 工作表
 * @param {number}      idx   - 房間在 Rooms 表中的 row index (1-based)
 * @param {Array}       r     - 房間 row 陣列
 * @param {Object}      rim   - Rooms header 的 { 欄名: index } map
 * @param {string}      roomId
 */
function doResolve_(ss, rooms, idx, r, rim, roomId) {
  const players = ss.getSheetByName(TABS.PLAYERS);
  const events  = ss.getSheetByName(TABS.EVENTS);
  const items   = ss.getSheetByName(TABS.ITEMS);

  // ── Step 1: 取得本回合物品及其真實價值 ──
  const round       = Number(r[rim.round]  || 0);
  const itemId      = String(r[rim.itemId] || "");
  const item        = itemId ? findItem_(items, itemId) : null;
  const hiddenValue = item ? Number(item.hiddenValue || 0) : 0;

  // ── Step 2a: 讀取房間內所有玩家，建立 playerId → Sheet 列號的查找表 ──
  const pAll = readAll_(players);
  const pim  = idxMap_(pAll.headers);
  const pRowById = {};  // playerId → sheet row index (1-based)
  pAll.rows.forEach((pr, i) => {
    if (String(pr[pim.roomId]) !== roomId) return;
    pRowById[String(pr[pim.playerId])] = i + 2; // +2: 第1列是 header，forEach 從 0 開始
  });

  // ── Step 2b: 掃描事件表，收集本回合的出價 & 使用卡牌紀錄 ──
  const eAll = readAll_(events);
  const eim  = idxMap_(eAll.headers);
  const lastBid    = {};   // playerId → 最後一次出價金額（同回合可多次出價，取最後一次）
  const bidOrder   = {};   // playerId → 該出價在 events 中的索引（越小 = 越早出價，用於 tie-break）
  const cardPlayed = {};   // playerId → [cardId, ...]（本回合使用的卡牌清單）

  eAll.rows.forEach((er, idx) => {
    if (String(er[eim.roomId]) !== roomId) return;
    const type = String(er[eim.type]);
    const pid  = String(er[eim.playerId] || "");
    let payload = {};
    try { payload = JSON.parse(String(er[eim.payloadJson] || "{}")); } catch (_) {}

    // 記錄本回合的出價（每次更新覆蓋上一次，最終 lastBid[pid] 為最後出價）
    if (type === "BID" && Number(payload.round) === round) {
      lastBid[pid]  = Number(payload.bid || 0);
      bidOrder[pid] = idx;  // 越後面 = 越晚出價（同金額時，越早出價者優先得標）
    }
    // 記錄本回合的卡牌使用
    if (type === "CARD_PLAYED" && Number(payload.round) === round) {
      if (!cardPlayed[pid]) cardPlayed[pid] = [];
      cardPlayed[pid].push(String(payload.cardId || ""));
    }
  });

  // ── Step 3: 找出得標者 ──
  // 規則：出價最高的玩家得標，但出價不能超過自己的預算
  // Tie-break：同出價金額時，先出價者（bidOrder 數字較小）優先
  let winnerId  = "";
  let winnerBid = -1;
  let winnerOrder = Infinity;

  Object.keys(lastBid).forEach(pid => {
    const bid = lastBid[pid];
    const rowIdx = pRowById[pid];
    if (!rowIdx) return;                  // 該玩家已不在房間中
    const row    = getRow_(players, rowIdx, players.getLastColumn());
    const budget = Number(row[pim.budget] || 0);
    if (bid > budget) return;             // 超過預算 → 出價無效
    const order = bidOrder[pid] || 0;
    // 出價更高，或同出價但更早出價 → 更新得標候選人
    if (bid > winnerBid || (bid === winnerBid && order < winnerOrder)) {
      winnerBid = bid; winnerId = pid; winnerOrder = order;
    }
  });

  // ── Step 4a: 卡牌效果 — Tax（加稅）──
  // C2 = Tax 卡：得標者的實際支付金額 = 出價 × 1.2（向上取整）
  let actualCost = winnerBid;
  let taxApplied = false;
  if (winnerId && cardPlayed[winnerId] && cardPlayed[winnerId].includes("C2")) {
    actualCost  = Math.ceil(winnerBid * 1.2);
    taxApplied = true;
  }

  // ── Step 5: 計算分數並更新玩家資料 ──
  const deltas     = {};   // playerId → { delta, shieldApplied, taxApplied }（前端用）
  let winnerName = "";

  if (winnerId) {
    const rowIdx = pRowById[winnerId];
    const row    = getRow_(players, rowIdx, players.getLastColumn());
    winnerName   = String(row[pim.name] || "");
    const budget = Number(row[pim.budget] || 0);
    const score  = Number(row[pim.score]  || 0);

    // 分數 = 物品真實價值 − 實際支付金額
    // 正值 = 賺到（買到便宜貨），負值 = 虧損（買貴了）
    let delta = hiddenValue - actualCost;

    // ── Step 4b: 卡牌效果 — Shield（護盾）──
    // C3 = Shield 卡：若分數增量為負（買貴了），則將虧損抵消為 0
    let shieldApplied = false;
    if (delta < 0 && cardPlayed[winnerId] && cardPlayed[winnerId].includes("C3")) {
      delta = 0;
      shieldApplied = true;
    }

    deltas[winnerId] = { delta, shieldApplied, taxApplied };

    // 更新得標者：預算扣除實際支付，分數加上 delta
    row[pim.budget]    = budget - actualCost;
    row[pim.score]     = score + delta;
    row[pim.updatedAt] = nowMs_();
    setRow_(players, rowIdx, row);
  }

  // ── Step 6: 切換房間狀態 BIDDING → REVEAL ──
  // REVEAL 狀態持續 REVEAL_DURATION_MS 毫秒，時間到後由 tryAutoTransition_ 切到 POSTROUND
  const now = nowMs_();
  r[rim.state]         = "REVEAL";
  r[rim.revealUntilTs] = now + REVEAL_DURATION_MS;
  r[rim.bidDeadlineTs] = "";               // 清除出價截止時間（已結算）
  r[rim.updatedAt]     = now;
  setRow_(rooms, idx, r);

  // ── Step 7: 記錄 ROUND_RESOLVED 事件 ──
  // 前端的 syncOnce 會讀取此事件來渲染結果面板（得標者、出價、分數變動等）
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

    // 同步更新所有玩家的預算為新的 startingBudget
    const players = ss.getSheetByName(TABS.PLAYERS);
    const pAll = readAll_(players);
    const pim  = idxMap_(pAll.headers);
    const t = nowMs_();
    pAll.rows.forEach((pr, i) => {
      if (String(pr[pim.roomId]) !== roomId) return;
      const pRowIdx = i + 2;
      const pRow = getRow_(players, pRowIdx, players.getLastColumn());
      pRow[pim.budget]    = startingBudget;
      pRow[pim.updatedAt] = t;
      setRow_(players, pRowIdx, pRow);
    });

    addEvent_(roomId, "SETTINGS_UPDATED", "", { startingBudget, roundSeconds, maxRounds });
    SpreadsheetApp.flush();
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


    // 驗證預算 & 收集活躍玩家（合併迴圈以提升效率）
    // Step A: 收集房間內所有 budget > 0 的玩家 ID，同時驗證當前玩家是否有資格
    const players = ss.getSheetByName(TABS.PLAYERS);
    const pAll = readAll_(players);
    const pim  = idxMap_(pAll.headers);
    const activePlayers = [];
    let myBudget = 0;
    let foundMe = false;

    for (const pr of pAll.rows) {
      if (String(pr[pim.roomId]) !== roomId) continue;
      
      const pid = String(pr[pim.playerId]);
      const budget = Number(pr[pim.budget] || 0);

      if (pid === playerId) {
        foundMe = true;
        myBudget = budget;
      }

      if (budget > 0) {
        activePlayers.push(pid);
      }
    }

    if (!foundMe) return { ok: false, error: "PLAYER_NOT_FOUND" };
    if (myBudget <= 0) return { ok: false, error: "ZERO_BUDGET" };


    // Step B: 掃描本回合所有 BID 事件，建立已出價玩家的 Set
    const events = ss.getSheetByName(TABS.EVENTS);
    const eAll = readAll_(events);
    const eim  = idxMap_(eAll.headers);
    const biddedSet = new Set();  // 本回合已出價的 playerId 集合

    // 預先產生 round 的搜尋字串，避免對舊事件進行 JSON.parse
    // JSON 中 round 通常為 "round":N 或 "round": N
    const roundStr = `"round":${round}`; 
    
    eAll.rows.forEach(ev => {
      if (String(ev[eim.roomId]) !== roomId) return;
      if (String(ev[eim.type]) !== "BID") return;
      
      const payload = String(ev[eim.payloadJson] || "{}");
      // 優化：先檢查字串是否包含 round，減少 JSON.parse 次數
      if (!payload.includes(roundStr) && !payload.includes(`"round": ${round}`)) return;

      try {
        const pl = JSON.parse(payload);

        if (Number(pl.round) === round) biddedSet.add(String(ev[eim.playerId]));
      } catch (_) {}
    });

    // Step C: 檢查是否全員都已出價
    const allBidded = activePlayers.length > 0 && activePlayers.every(pid => biddedSet.has(pid));
    if (allBidded) {
      // 全員出價完畢 → 不等倒數計時，立即呼叫 doResolve_ 結算本回合
      // 先重新讀取 room row（因為可能有其他併發請求修改）
      const rr = getRow_(rooms, idx, rooms.getLastColumn());
      if (String(rr[rim.state]) === "BIDDING") {
        doResolve_(ss, rooms, idx, rr, rim, roomId);
      }
    }

    return { ok: true, allBidded };
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
