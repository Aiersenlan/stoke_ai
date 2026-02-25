# 股票大戶動向（買賣超）終端機 - 系統開發需求規格書 (PRD)

## 系統概述
本專案包含一個 Python 爬蟲與資料分析腳本 (`analyze.py`)，以及一個輕量級的 Flask 網頁應用程式供資料視覺化 (`app.py` 及 `templates/index.html`)。目標是從台灣證券交易所與櫃買中心獲取每日的三大法人買賣超資料與收盤行情，進行資料整理、計算估價後，產出經過排版與顏色標註的 Excel 報表，並透過網頁提供專業看盤終端機（Trading Terminal）風格的互動式介面。

## 1. 後端資料爬取與分析 (`analyze.py`)
### 1.1 資料源網址規則
- **上市買賣超 (TWSE T86)**: `https://www.twse.com.tw/rwd/zh/fund/T86?response=json&selectType=ALL&date={date}`
- **上櫃買賣超 (TPEX)**: `https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&se=EW&t=D&d={tpex_date}` （注意：tpex_date 為民國年格式，如 115/02/24）
- **上市收盤價 (TWSE MI_INDEX)**: `https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?response=json&type=ALL&date={date}`
- **上櫃收盤價 (TPEX)**: `https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&d={tpex_date}`

### 1.2 資料過濾與欄位處理
1. 取得每日交易資料後，**排除 ETF 與權證** (邏輯：股號長度大於 4，或是首字不為數字者，皆跳過)。
2. 計算屬性：
   - `price`: 該日收盤價。
   - `vwap` (均價): 成交金額(元) / 成交股數 (若缺金額或股數則 fallback 為收盤價)。
   - `foreign_shares`: 外資買賣超股數。
   - `it_shares`: 投信買賣超股數。
   - `foreign_val`: 外資買賣超估計金額 = `foreign_shares * vwap`。
   - `it_val`: 投信買賣超估計金額 = `it_shares * vwap`。

### 1.3 產出 Excel 報表 (`market_analysis_YYYYMMDD.xlsx`)
1. 區分 `上市` 與 `上櫃` 兩個工作表 (Sheet)。
2. **版面配置**:
   - 第 1 列: 放報表日期 (字體微軟正黑，靠左對齊，僅需於 A1 放置一次)。
   - 第 2 列 (合併儲存格): 依序為 `外資買超`, `外資賣超`, `投信買超`, `投信賣超`。
   - 第 3 列 (子標題): 每個分類下並列 `證券代號`, `證券名稱`, `收盤價`, `均價`, `股數`, `估價(百萬)`。
   - 欄寬設定：代號/收盤/均價等為 10~11，名稱 12，股數 16，估價 13。
   - 證券代號需轉換為「整數」(Number) 型態以避免 Excel 跳出格式綠色警告。
3. **資料排序與切分**:
   - `foreign_shares > 0` (外買), `foreign_shares < 0` (外賣), `it_shares > 0` (投買), `it_shares < 0` (投賣)。
   - 四個分支的清單皆**依照估價金額絕對值由大到小排序**。
4. **同向與對作重點著色 (Highlight)**:
   - 只針對「名稱」欄位上背景色。
   - 依據股數絕對值大小判斷誰為主導方。
   - **同向者 (同為買或賣)**: 股數較大者塗「深紅 `#FF8080`」，較小者塗「淺紅 `#FFCCCC`」。
   - **對作者 (一買一賣)**: 股數較大者塗「深綠 `#80FF80`」，較小者塗「淺綠 `#E6FFE6`」。

## 2. Web 伺服器 (`app.py`)
- 使用 `Flask` 及 `pandas` 處理與提供後端資料。
-路由規則：
   - `/`: 渲染 `index.html`。
   - `/get_available_dates`: 掃描目錄尋找 `market_analysis_YYYYMMDD.xlsx`，以陣列回傳排序後的可用日期。
   - `/get_report/<date>`: 用 pandas (`pd.read_excel(..., header=None)`) 讀取 Excel 檔所有 Sheet，去除空行後重新包裝為包含 `main_headers`, `sub_headers`, `data` 的 JSON。
   - `/download/<date>`: 提供 Excel 檔案本機端直接下載。

## 3. 前端看盤終端機 (`templates/index.html`)
### 3.1 整體視覺風格 (Dark Mode Trading Terminal)
- HTML 掛載 `data-bs-theme="dark"`。
- 背景改為專業深色系 (`#0d0d0d`, `#161616`)，隱藏原生捲軸並實作深灰色細捲軸。
- 表格最上層區塊標題套用特別設計：
   - `買超` 區塊標題：暗紅茶色 (#5a1a1a) ＋ 亮紅底線。
   - `賣超` 區塊標題：暗青綠色 (#1a4a1a) ＋ 亮綠底線。
- 大標題顯示：「📈 當日大戶動向（買賣超）」。

### 3.2 頂部控制列
- 包含四種互動介面：
   1. **選擇日期** (動態載入)。
   2. **顯示筆數** (`前 35 名`、`全部顯示`)。
   3. **排序方式** (`依估價排名`、`依股數排名`)。
   4. **操作按鈕** (`⬇️ 下載 Excel`, `🔄 重新整理`)。

### 3.3 互動效果與動態渲染 (JavaScript)
- **同向/對作著色邏輯**: 
   切片 (`slice`) 資料前，必須先蒐集該板塊內所有存在的股票資訊，還原出與 Excel 完全相同的紅綠渲染邏輯（見 1.3）。確認此筆資料沒有因為過濾 35 筆而在畫面上找不到其對作方（若畫面另一方已被裁切使得資料內無法讀取，則此股呈現無顏色著色，達到「所見即所得」）。
- **滑鼠懸停 (Hover) 標記**: 移至名稱顯示 `🔍`。全表同樣代號的收盤/均/股數/估價格子底色瞬間同步改變為深黃高亮 (`#3d3511` / 金黃文字 `#ffe600`)。
- **一鍵瞬間移動 (Click to Scroll)**: 點擊股名時，呼叫 `scrollIntoView(center)` 平滑跳轉至該股在其他買賣區塊的位置，並加上視覺化導引（目標背景閃爍純黃色 `#ffff00`、黑字、紅邊框 `#ff0000` 長達 2 秒）。
- **回到頂端 (Back to Top)**: 下滑超過 300px 時浮現右下角 `⬆️` 按鈕。

---
**提示與驗證目標**：
請另外建立一個獨立資料夾供測試，代理 Agent 需能完全只透過此份 md 建立出 `analyze_test.py`, `app_test.py` 與 `templates/index.html` 並確保執行結果與原專案吻合。
