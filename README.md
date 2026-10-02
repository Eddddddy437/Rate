# 💱 Rate｜台銀匯率抓取小幫手

### 財務自動化：一鍵抓取臺灣銀行 19 種外幣牌告匯率，自動產出月報 Excel

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Playwright](https://img.shields.io/badge/Browser-Playwright-2EAD33.svg)
![Pandas](https://img.shields.io/badge/Data-Pandas-150458.svg)
![Openpyxl](https://img.shields.io/badge/Excel-Openpyxl-217346.svg)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-yellow.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6.svg)

---

## 📖 專案簡介

每月整理合併報表時，會計人員需要到臺灣銀行網站逐一查詢多國匯率、複製貼上再計算平均，既耗時又容易出錯。

**匯率抓取小幫手** 把這件事變成「選年月 → 按開始 → 等完成」三個步驟：程式會自動開啟瀏覽器，到 [臺灣銀行歷史牌告匯率](https://rate.bot.com.tw/xrt) 抓取指定月份 19 種外幣的每日匯率，整理成格式固定的 Excel，並算好月平均與買賣中價。

> 💡 打包後的 exe 版內建瀏覽器，同事電腦**不需要安裝 Python** 就能使用。

---

## ✨ 核心功能

| 功能 | 說明 |
| --- | --- |
| 🖱️ **圖形化操作介面** | Tkinter 視窗，下拉選單選年月、檔名自動帶入（例：`exrate_AUG.xlsx`），執行紀錄即時顯示 |
| 🌐 **瀏覽器自動化抓取** | 以 Playwright（headless Chromium）載入台銀頁面，先開首頁暖機再逐一抓取 19 種幣別 |
| 🔁 **穩定性機制** | 每個幣別失敗自動重試 2 次（間隔 3 秒）、幣別間延遲 1 秒，對伺服器友善 |
| ✅ **完整性檢查** | 自動比對是否抓滿 19 幣別 × 4 欄 = 76 欄，缺少的幣別或欄位會在紀錄中列出 |
| 🧮 **自動財務計算** | 月平均數、各幣別買入／賣出平均（中價）、最後一天買入／賣出平均（寫入 Excel 公式） |
| 📊 **報表格式處理** | openpyxl 合併買賣儲存格、寫入 `=AVERAGE()` 公式、數字保留小數 4 位 |
| 🧵 **不卡畫面** | 抓取在背景執行緒進行，視窗不會凍結；同名檔案會先詢問是否覆蓋 |
| 📦 **可打包成 exe** | PyInstaller（onedir）＋隨附 `ms-playwright` 瀏覽器，程式會自動找到隨附的執行檔 |

**抓取幣別（19 種）**

`USD` `JPY` `EUR` `GBP` `AUD` `CAD` `CNY` `HKD` `SGD` `CHF` `SEK` `ZAR` `NZD` `THB` `PHP` `IDR` `KRW` `VND` `MYR`

每個幣別抓取：**現金買入、現金賣出、即期買入、即期賣出**

---

## 🛠️ 技術棧

| 類別 | 使用工具 |
| --- | --- |
| 語言 | Python 3 |
| 網頁抓取 | Playwright（Chromium, headless） |
| 資料處理 | pandas（`read_html`）、lxml |
| Excel 輸出 | openpyxl |
| 圖形介面 | tkinter、threading、queue |
| 打包發佈 | PyInstaller（onedir，spec 檔管理） |

---

## 🚀 快速開始

### 方法一：直接使用 exe（一般使用者）

1. 取得打包好的 `匯率抓取小幫手` 資料夾，**整包複製**，不要只複製 exe。
2. 雙擊 `匯率抓取小幫手.exe`。
3. 選擇年份與月份 → 確認檔名與存放位置 → 按 **🚀 開始抓取**。
4. 看到「✅ 抓取完成！」後，按 **📁 開啟檔案位置** 即可打開 Excel 所在資料夾。

資料夾結構：

```
匯率抓取小幫手\
├── 匯率抓取小幫手.exe      ← 雙擊執行
├── ms-playwright\          ← 內建瀏覽器，請勿刪除或改名
└── _internal\              ← 程式執行所需檔案（PyInstaller 6 以上產生），請勿刪除
```

### 方法二：從原始碼執行（開發者）

**1. 複製專案**

```bash
git clone https://github.com/Eddddddy437/Rate.git
cd Rate
```

**2. 建立虛擬環境並安裝套件**

```bash
python -m venv exrate_env

# cmd
exrate_env\Scripts\activate
# PowerShell（若被執行原則擋下，請改用 cmd）
exrate_env\Scripts\Activate.ps1

pip install playwright pandas openpyxl lxml
playwright install chromium
```

> ⚠️ `lxml` 一定要安裝在虛擬環境內，否則 `pandas.read_html` 會無法解析網頁表格。

**3. 執行**

```bash
python 匯率抓取小幫手.py
```

---

## 📊 輸出結果

產出一個 Excel 工作表：

- **A 欄**：掛牌日期，每列一個掛牌日
- **B 欄起**：依幣別排列，每幣別 4 欄（例：`USD_現金買入`、`USD_現金賣出`、`USD_即期買入`、`USD_即期賣出`），共 76 欄
- **最下方 3 列彙總**：

| 彙總列 | 算法 |
| --- | --- |
| `N月平均數` | 每一欄取整月平均 |
| `各幣別買入/賣出平均` | 先算每日（買入＋賣出）÷ 2，再取整月平均；現金、即期分開計算，買賣兩格合併 |
| `各幣別最後一天買入/賣出平均` | 以 Excel 公式 `=AVERAGE()` 取最後一個掛牌日的買賣平均，買賣兩格合併 |

> 台銀網頁日期為**由新到舊**排列，因此 Excel 第 2 列即為該月最後一個掛牌日。台銀未報價的欄位（例如部分幣別無現金匯率）會保留空白。

---

## ⚙️ 自訂設定

所有參數都在 `匯率抓取小幫手.py` 內：

| 想調整 | 位置 |
| --- | --- |
| 增減幣別 | `CURRENCIES` 清單（代碼須與台銀網址一致） |
| 存檔資料夾 | `build_output_folder()`，預設為公司公用磁碟路徑，**其他環境使用前請改成自己的資料夾** |
| 台銀網址 | `run_fetch()` 中的 `https://rate.bot.com.tw/xrt/quote/ltm/{cur}` |
| 重試次數 | `fetch_html_with_retries(..., max_retries=2)` |
| 頁面逾時 | `timeout=60000`（毫秒） |
| 幣別間隔 | `time.sleep(1.0)` |
| 小數位數 | `round(4)` 與 `number_format = '0.####'` |

---

## 📦 打包成 exe

```bash
# 在 exrate_env 中執行
pip install pyinstaller
pyinstaller 匯率抓取小幫手.spec
```

1. 依 `匯率抓取小幫手.spec` 打包（onedir、無主控台視窗，已設定 `collect_all` 收齊 pandas、openpyxl、playwright、lxml）。
2. **手動把 `ms-playwright` 資料夾複製到 `dist\匯率抓取小幫手\` 內、與 exe 同層**。PyInstaller 不會自動帶上瀏覽器，預設下載位置通常在 `C:\Users\<使用者>\AppData\Local\ms-playwright`。
3. 在沒有安裝 Python 的電腦實測一次，執行紀錄應出現「使用隨附瀏覽器」與「✅ 幣別與欄位完整」。

> 請一律用 spec 檔打包。若直接對 `.py` 打包，PyInstaller 會重新產生 spec 並覆蓋原有設定，exe 可能在其他電腦上無法執行。

---

## ❓ 常見問題

| 訊息 | 原因與處理 |
| --- | --- |
| 無法建立或存取輸出資料夾 | 存檔路徑不存在或無權限，請確認磁碟連線或修改 `build_output_folder()` |
| ⚠️ XXX 該月份無資料，略過 | 台銀此頁面只提供近期資料（約三個月），太久以前的月份無法抓取 |
| 第 1 次抓取 XXX 失敗，3 秒後重試 | 網站回應較慢，程式會自動重試，不需處理 |
| ❌ 所有幣別都沒抓到資料 | 網路斷線或台銀網站改版；若網站改版，需調整 `run_fetch()` 中的欄位重排邏輯 |
| 瀏覽器無法啟動 | `ms-playwright` 資料夾遺失或未與 exe 同層，請重新複製整包工具 |
| ⚠️ lxml 模組載入診斷：失敗 | 打包時缺少 lxml，請確認在虛擬環境內安裝後重新打包 |
| 存檔時出錯（Permission denied） | 同名 Excel 正被開啟，請關閉後重跑或更換檔名 |

---

## 📝 版本紀錄

- **v2**：改用 Playwright 瀏覽器自動化抓取，新增 Tkinter 圖形介面、背景執行緒、完整性檢查、覆蓋確認，並支援 PyInstaller 打包與隨附瀏覽器
- **v1**：以 Requests 抓取台銀歷史匯率，產出含月平均的 Excel 報表

---

## ⚖️ 聲明

匯率資料來源為[臺灣銀行](https://rate.bot.com.tw/xrt)公開網頁，本工具僅供內部作業自動化使用，實際匯率請以臺灣銀行公告為準。
