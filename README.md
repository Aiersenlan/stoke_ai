# Institutional Tracker (股票大戶動向追蹤系統)

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/flask-app-green.svg)](#)
[![Deployment](https://img.shields.io/badge/deployment-render-purple.svg)](#)

👉 **Live Demo**: [https://institutional-tracker.onrender.com](https://institutional-tracker.onrender.com)

## 摘要 (Abstract)

Institutional Tracker 是一款專為台灣證券市場（TWSE/TPEX）設計的機構投資人（外資、投信）資金動向追蹤系統。本專案透過自動化數據採集與清洗，建構出即時的法人買賣超矩陣，並結合專業交易終端機 (Trading Terminal) 的視覺化介面，提供具量化價值的市場籌碼洞察。

本專案由高效能 Python 爬蟲引擎與基於 Flask 的 Web 視覺化後端組成，內建抗 WAF 防護規避機制，並原生支援 GitHub Actions 與 Render 雲端環境之自動化部署。

## 核心特性 (Key Features)

- **籌碼數據清洗與加權估值 (VWAP Analysis)**
  自動過濾 ETF 與權證等非原生股票標的，基於個股成交均價 (VWAP) 計算法人真實買賣部位（估價金額），還原大戶真實動向。
  
- **抗 WAF 爬蟲引擎 (Anti-WAF Engine)**
  基於 `requests` 重構底層網路連線層，深度偽裝 HTTP Headers（含 `Referer` 與進階 `User-Agent` 指紋），具備高可用性，精準繞過台灣櫃買中心 (TPEX) 的 403 阻擋機制。
  
- **交易終端機視覺化 (Trading Terminal UI)**
  採用 Dark Mode 介面設計，實現「法人同向/對作」資金矩陣之色彩標註機制。內建動態游標連動高亮 (Hover Sync) 與跨板塊平滑導航 (Click-to-Scroll) 功能，達到所見即所得之分析體驗。
  
- **無人值守自動化 (Automated Pipeline)**
  前端具備「懶人全自動抓取 (Auto-Fetch)」邏輯，自動推算開盤日補齊回看數據。後台內建 GitHub Actions 腳本 (`daily_analysis.yml`)，支援每日盤後排程執行。
  
- **零配置雲端部署 (Zero-Config Deployment)**
  提供 `render.yaml`，配合 `gunicorn` 與 `requirements.txt`，支援一鍵式 Render 容器化部署與擴展。

## 系統架構 (Architecture)

- **後端與數據處理**: Python 3.10+, Pandas, Requests, Openpyxl
- **Web 伺服器**: Flask, Gunicorn
- **前端介面**: HTML5, Vanilla JS (ES11+), Bootstrap 5

## 快速啟動 (Quick Start)

### 1. 環境安裝
請準備 Python 3.10 或以上環境，並依賴以下指令安裝必要套件：
```bash
pip install -r requirements.txt
```

### 2. 數據獲取 (CLI 模式)
預設擷取最近一營業日之數據。亦可透過參數指定歷史日期（格式：`YYYYMMDD`）：
```bash
python analyze.py 20260224
```
執行完畢後，系統將於工作目錄產出 `market_analysis_YYYYMMDD.xlsx` 報表。

### 3. 啟動 Web 伺服器
```bash
python app.py
```
啟動後使用瀏覽器訪問 `http://127.0.0.1:5000` 即可進入視覺化交易終端。

## 開發與貢獻 (Development & Agents)
針對 AI 代碼代理人 (AI Coding Agents) 或二次開發者，核心商業邏輯與規避策略之還原規格，請參閱 [Agent Recovery Specification](agent_recover.md)。

## 授權 (License)
本專案為開源軟體，僅供學術與技術交流使用，投資風險請自行評估。
