# Rate
自動抓取台銀網站的匯率
# 🏦 Taiwan Bank Exchange Rate Crawler & Reporter
### 財務自動化：台銀歷史匯率自動爬取與 Excel 報表生成工具

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg) ![Pandas](https://img.shields.io/badge/Library-Pandas-orange.svg) ![Openpyxl](https://img.shields.io/badge/Library-Openpyxl-green.svg)

## 📖 專案簡介
本專案專為財務會計人員設計，旨在解決每月手動查詢並錄入多國匯率的繁瑣工作。透過 Python 自動化爬取 **臺灣銀行 (Bank of Taiwan)** 公開的歷史匯率資訊，並自動進行數據清洗、計算平均值，最後產出符合財務審核格式的 Excel 報表。

## ✨ 核心功能
* **多幣別自動爬取：** 一次性獲取 USD, JPY, EUR 等 19 種主要貨幣的歷史數據。
* **數據異常處理：** 內建自動重試機制與請求延遲，確保爬蟲穩定性並友善對待伺服器。
* **自動化財務計算：** * 自動計算全月平均匯率。
    * 自動計算各幣別買入/賣出的中價平均。
* **專業格式報表：** 使用 `openpyxl` 進行 Excel 底層操作，包含自動合併儲存格、寫入 Excel 計算公式。

## 🛠️ 技術棧
* **語言：** Python 3
* **數據處理：** `Pandas`, `Numpy`
* **網路請求：** `Requests`
* **Excel 操作：** `Openpyxl`

## 🚀 如何使用
1. **複製專案：**
   ```bash
   git clone [https://github.com/Eddddddy437/Rate.git](https://github.com/Eddddddy437/Rate.git)
