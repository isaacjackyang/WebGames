Blind Auction Party（盲拍派對）完整化規格書（務實版）
0. 核心原則（別再做「會動的 demo」）
你現在的技術選擇是：Apps Script + Google Sheet + 輪詢。

要做成「完整遊戲」，最務實的策略是：
先把玩法閉環做到乾淨（回合 → 出價 → 結算 → 結果面板 → 下一回合）
再加心理戰與卡牌（但要用事件驅動，不要前端亂改數字）
最後才談 QR、同網路偵測、效能極限（這些是 UX 不是核心玩法）
這份文件就是照這三層走。
1. 完整遊戲玩法（V1 完整版）
1.1 遊戲資源
每位玩家擁有：

Budget（預算）：用來出價
Score（分數）：最後排名依據
Cards（卡牌手牌）：改變規則或干擾對手（心理戰的「工具箱」）
每回合抽出一個 Item：

name（公開）
publicHint（公開提示）
hiddenValue（真實價值：結算後才揭露或只揭露差值）
1.2 回合流程（強制閉環）
回合狀態固定四段：

LOBBY（準備）
玩家加入/離開
Host 調整設定（回合數、秒數、起始預算、卡牌配置）
Host 按「開始回合」
BIDDING（出價中）（倒數 N 秒）
玩家提交出價（可多次更新，採最後一次）
玩家可使用卡牌（每張牌需定義允許的時機，例如：僅出價中/僅結算前）
REVEAL（揭露結算）
倒數結束或 Host 按「立即結算」
系統結算：得標者、扣預算、算分
顯示本回合結果面板（至少 8–10 秒，讓派對有節奏）
POSTROUND（回合後）
顯示排行榜、每人 budget/score 變化
Host 按「下一回合」或系統自動進 LOBBY（依設定）
注意：現在你的系統缺 REVEAL/POSTROUND 這種儀式感。沒有它就不像遊戲，只像表單。
2. UI 全面設計（讓它像「遊戲」）
2.1 大廳 UI（Lobby）
目標：3 秒內讓所有人找到房間並進去。

輸入：lobbyCode、name
按鈕：
測試API：顯示後端是否回 JSON（防止權限/部署問題）
初始化DB：建立工作表與 seed
開新房間：建立房間 + 加入
刷新房間列表：列出活躍房間
加入：用 roomId 直接加入
房間列表：顯示 roomId、state、玩家數、回合進度、加入按鈕
加入後：自動跳房間 UI
驗收條件：

任何錯誤都必須出現在畫面（不能只在 console）
房間列表 1 次 refresh 能看到同 lobbyCode 的活躍房
2.2 房間 UI（Room）— 四段狀態切換
用「狀態頁」做 UI，而不是堆一堆按鈕：

A) LOBBY（準備頁）
顯示：玩家列表、budget/score、Host 標記
Host 設定區：
回合數 maxRounds
每回合秒數 roundSeconds
起始預算 startingBudget
卡牌模式（簡單：開/關；進階：卡牌池）
按鈕：
Host：開始回合
任何人：離開房間
驗收：

Host 改設定後，所有玩家 2 秒內同步看到。
B) BIDDING（出價頁）
顯示：item name、publicHint、倒數計時（大字體）
出價輸入與送出
卡牌區（手牌顯示 + 可用按鈕）
顯示「目前你最後一次出價」（避免玩家不確定有沒有送出）
驗收：

倒數到 0 必須自動進 REVEAL（不靠 Host 手動）。
C) REVEAL（揭露結果頁）
顯示：得標者、得標價、hiddenValue（或至少顯示本回合得分 delta）
顯示：每人本回合分數變化（+/-）
10 秒倒數後自動進 POSTROUND
驗收：

所有玩家看到相同結果，且 UI 有「節奏感」。
D) POSTROUND（排行榜頁）
顯示：排行榜（score 排序），每人 budget/score
Host：下一回合（或自動進 LOBBY）
若回合已滿：顯示「遊戲結束」+ 重新開始（Host）
驗收：

最後一回合結束會停在「遊戲結束」狀態，不會亂跳。
3. 後端狀態機與規則（要硬，不然一定爆）
3.1 Room state machine（不可模糊）
Rooms 表新增欄位（或你已有的擴充）：

state: LOBBY | BIDDING | REVEAL | POSTROUND | ENDED
round: 0..maxRounds
itemId
bidDeadlineTs
revealUntilTs（揭露頁倒數結束時間）
postUntilTs（排行榜頁倒數結束時間）
轉移規則：

LOBBY → BIDDING：Host startRound
BIDDING → REVEAL：now >= bidDeadlineTs（任何玩家 sync 時觸發也可以）或 Host forceResolve
REVEAL → POSTROUND：now >= revealUntilTs
POSTROUND → LOBBY：now >= postUntilTs 且 round < maxRounds
POSTROUND → ENDED：round == maxRounds
務實做法：把「時間到了就轉狀態」寫在 sync 裡。

因為 Apps Script 沒有可靠常駐背景工作，靠玩家輪詢最穩。
3.2 出價規則（最小完整集）
每位玩家對每回合最多只取「最後一次出價」
若出價 > budget：該出價無效（或視為 all-in，二選一要固定）
平手：用「最早提交的最高價」或「隨機」— 但要寫死規則
3.3 結算公式（保持簡單）
得標者：

budget -= winnerBid
score += (hiddenValue - winnerBid)
非得標者：

score += 0（先保持簡單）
4. 卡牌系統（心理戰落地的最務實版本）
4.1 卡牌要用事件驅動（不要前端改數字）
卡牌使用 → Events 記一筆 CARD_PLAYED

結算時後端讀事件，套用效果。
4.2 建議 V1 卡牌（好做、好玩、可驗收）
每張牌要定義：

可用時機：BIDDING / REVEAL 前 / POSTROUND
作用對象：自己 / 指定玩家 / 全體
效果：怎麼影響結算
V1 三張（務實能做完）：

Peek（偷看）
時機：BIDDING
效果：只讓使用者看到 hiddenValue 的區間（例如 ±5），而不是精確值
UI：只顯示給該玩家（前端判斷 playerId）
Tax（加稅）
時機：BIDDING
效果：如果你得標，你多付 +20%（是自殘牌，心理戰用來假裝自己會出高價、或逼別人錯判）
實作：resolve 時檢查該玩家是否打了 Tax，winnerBid 變成 ceil(bid*1.2)
Shield（護盾）
時機：REVEAL 前
效果：本回合若你分數為負，抵消一次（delta 最多補到 0）
實作：resolve 後計算 delta，若 delta<0 且有 Shield → delta=0 並扣牌
這三張能做出「看起來像完整遊戲」的心理博弈，而且後端改動集中在 resolve。
5. 資料表（Sheet）與事件設計（完整化後）
5.1 Sheet Tabs
Rooms
Players
Events
Items
PlayerCards
5.2 Events 類型（建議固定枚舉）
ROOM_CREATED
PLAYER_JOINED
SETTINGS_UPDATED
ROUND_STARTED
BID
CARD_PLAYED
ROUND_RESOLVED
STATE_TRANSITION（可選：每次狀態切換都記）
GAME_ENDED（可選）
5.3 sync 回傳應該包含（完整遊戲必備）
room（含 state、deadline、revealUntil、postUntil、item public info）
players（含 score、budget）
myCards（含 count）
events（增量）
6. 程式架構與開發順序（兩天能做完的那種務實）
Stage 1：把「回合閉環」做完（先別想卡牌）
要改：

Code.gs
startRound：寫入 itemId、bidDeadlineTs、state=BIDDING、round+1
sync：若 now >= bidDeadlineTs 且 state=BIDDING → 自動 resolve → state=REVEAL → 設 revealUntilTs
sync：若 now >= revealUntilTs 且 state=REVEAL → state=POSTROUND → 設 postUntilTs
sync：若 now >= postUntilTs 且 state=POSTROUND → 進 LOBBY 或 ENDED
index.html
加 Host「開始回合」按鈕（LOBBY 可見）
做倒數計時顯示（bidDeadlineTs）
做 REVEAL/POSTROUND 畫面（至少兩個 panel）
結算結果用漂亮面板呈現（不要只印 log）
驗收：

Host 按一次開始回合 → 全員進 BIDDING
倒數到 0 → 全員自動進 REVEAL → 10 秒後進 POSTROUND → 再回 LOBBY
不需要 Host 手動按結算也能跑完（Host 可以有「立即結算」）
Stage 2：加卡牌（只加 1–3 張，但要完整）
要改：

Code.gs
playCard：扣手牌 count + 記事件 CARD_PLAYED
resolve：讀 CARD_PLAYED 事件套效果（先做 Peek/Tax/Shield）
index.html
手牌 UI（顯示 count）
只有合法時機顯示可用按鈕
Peek 的資訊只對該玩家顯示
驗收：

使用卡牌後 count 減 1
Peek 玩家看到不同資訊
Tax/Shield 會影響本回合 delta
Stage 3：讓它「派對可用」的 UX（不是花拳繡腿）
要改：

QR 進房：URL 加 ?roomId=...（已可做）
房間列表顯示「可加入/正在出價中/正在揭露」
音效或簡單震動（可選，純前端）
顯示「你已出價：X」（避免重複按）
結果面板要夠大，大家圍著看才像派對
驗收：

10–15 人連線時不會因為輪詢堆疊卡死（single-flight 保持）
7. 你現在最容易犯的錯（我直接罵醒你）
你之前一直想加「同網路偵測玩家」「local wifi 玩家顯示」——那是典型的把 UX 當核心玩法。

玩法沒閉環、狀態機沒定義、結算沒結果面板，你就算偵測到隔壁路由器裡有誰，也只是在更華麗地展示一個不完整 demo。
先把「回合節奏」做出來。派對遊戲的靈魂是節奏，不是網路黑魔法。
8. 最小可交付版本（MVP）定義
以下全部做到，這就叫「完整遊戲（V1）」：
視覺風格要是喜氣的紅色，華人過年氣氛，增加一些紅包 元寶 等元素

4 段狀態：LOBBY/BIDDING/REVEAL/POSTROUND（必須）
倒數與自動狀態切換（必須）
結果面板與排行榜（必須）
Host 設定（至少回合數、秒數、起始預算）
至少 1 張可用卡牌（可選但強烈建議）
