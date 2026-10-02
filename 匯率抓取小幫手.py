import os
import sys
import time
import threading
import queue
import subprocess
import glob
from io import StringIO
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ============================================================
# 路徑設定：跟毛利分析、預提工具一樣的寫法，方便打包成 exe
# ============================================================
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 指定 Playwright 要去哪裡找瀏覽器本體（Chromium）
# 打包成 exe 後，瀏覽器資料夾會跟 exe 放在同一層、取名叫 ms-playwright
_bundled_browser_path = os.path.join(BASE_DIR, "ms-playwright")
if os.path.isdir(_bundled_browser_path):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = _bundled_browser_path


def find_bundled_chromium_executable():
    """
    直接搜尋隨附的 ms-playwright 資料夾裡，實際的瀏覽器執行檔完整路徑。
    打包成 exe 後，Playwright 內部對「去哪裡找瀏覽器」的判斷邏輯不完全可靠，
    只設定 PLAYWRIGHT_BROWSERS_PATH 環境變數不夠保險，
    所以改成直接把執行檔的確切路徑告訴 Playwright，不讓它自己猜。
    找不到就回傳 None（例如在開發環境直接跑 .py 時，讓 Playwright 用預設方式找即可）。
    """
    if not os.path.isdir(_bundled_browser_path):
        return None

    # chrome-headless-shell.exe：Playwright 在 headless=True 時預設使用的執行檔
    pattern_headless_shell = os.path.join(
        _bundled_browser_path, "chromium_headless_shell-*",
        "chrome-headless-shell-win64", "chrome-headless-shell.exe"
    )
    matches = glob.glob(pattern_headless_shell)
    if matches:
        return matches[0]

    # 保險：如果找不到 headless-shell 版本，退而求其次找一般的 chrome.exe
    pattern_regular = os.path.join(
        _bundled_browser_path, "chromium-*", "chrome-win64", "chrome.exe"
    )
    matches = glob.glob(pattern_regular)
    if matches:
        return matches[0]

    return None


# ============================================================
# 幣別清單、月份對照
# ============================================================
CURRENCIES = ['USD', 'JPY', 'EUR', 'GBP', 'AUD', 'CAD', 'CNY', 'HKD',
              'SGD', 'CHF', 'SEK', 'ZAR', 'NZD', 'THB', 'PHP', 'IDR',
              'KRW', 'VND', 'MYR']

# 月份英文縮寫對照表，用來組成 exrate_AUG.xlsx 這種檔名
MONTH_ABBR = {
    "01": "JAN", "02": "FEB", "03": "MAR", "04": "APR",
    "05": "MAY", "06": "JUN", "07": "JUL", "08": "AUG",
    "09": "SEP", "10": "OCT", "11": "NOV", "12": "DEC",
}


# 依照選擇的年度，組出公用資料夾的目標路徑
def build_output_folder(year: str) -> str:
    return os.path.join(
        r"Y:\Account_Dept\管理報表",
        f"{year}損益",
        f"{year} 合併財務報表",
        f"{year}匯率統計",
    )


# ============================================================
# 核心抓取邏輯
# ============================================================
def run_fetch(month: str, output_path: str, log_func):
    from playwright.sync_api import sync_playwright

    all_data = {}

    def fetch_html(page, cur):
        url = f'https://rate.bot.com.tw/xrt/quote/ltm/{cur}'
        page.goto(url, wait_until="networkidle", timeout=60000)
        page.wait_for_selector("table", timeout=60000)
        return page.content()

    def fetch_html_with_retries(page, cur, max_retries=2):
        last_err = None
        for attempt in range(1, max_retries + 2):
            try:
                return fetch_html(page, cur)
            except Exception as e:
                last_err = e
                if attempt <= max_retries:
                    log_func(f"   ↳ 第 {attempt} 次抓取 {cur} 失敗（{e}），3 秒後重試...")
                    time.sleep(3.0)
        raise last_err

        # 診斷用：直接測試 lxml 能不能正常 import，並印出最原始、最詳細的錯誤
    try:
        import lxml.etree  # noqa: F401
        log_func("✅ lxml 模組載入診斷：正常")
    except Exception as e:
        log_func(f"⚠️ lxml 模組載入診斷：失敗！詳細原因 -> {type(e).__name__}: {e}")

    log_func("🚀 開始啟動瀏覽器，準備抓取資料...")

    with sync_playwright() as p:
        # 優先使用隨附在 ms-playwright 資料夾裡的瀏覽器執行檔完整路徑，
        # 不依賴 Playwright 自己打包後的內部路徑判斷（那個判斷在打包環境下不可靠）
        bundled_exe = find_bundled_chromium_executable()
        if bundled_exe:
            log_func(f"🔧 使用隨附瀏覽器：{bundled_exe}")
            browser = p.chromium.launch(headless=True, executable_path=bundled_exe)
        else:
            browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # 暖機：第一次連線台銀網站時，防機器人驗證/建立連線可能比較慢
        try:
            log_func("🔥 暖機中，先連線一次台銀首頁...")
            page.goto("https://rate.bot.com.tw/xrt", wait_until="networkidle", timeout=60000)
            time.sleep(3.0)
        except Exception as e:
            log_func(f"⚠️ 暖機時發生小狀況（不影響後續抓取）：{e}")

        for cur in CURRENCIES:
            log_func(f"抓取 {cur} ...")
            try:
                html = fetch_html_with_retries(page, cur, max_retries=2)

                bot_data = pd.read_html(StringIO(html))
                df = bot_data[0]

                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = df.columns.get_level_values(1)

                # 新版 pandas 型態檢查較嚴格，先轉成通用型態避免下面交換欄位時報錯
                df = df.astype(object)

                df = df.iloc[:, [0, 2, 1, 3, 4, 5]]
                df.columns = ['掛牌日期', '幣別', '現金買入', '現金賣出', '即期買入', '即期賣出']
                df.loc[:, ['幣別', '現金買入']] = df[['現金買入', '幣別']].values
                df = df.replace('-----------------------', pd.NA)

                num_cols = ['現金買入', '現金賣出', '即期買入', '即期賣出']
                df[num_cols] = df[num_cols].apply(pd.to_numeric, errors='coerce')

                df = df.set_index('掛牌日期')
                df = df[df.index.astype(str).str.startswith(month.replace('-', '/'))]

                if df.empty:
                    log_func(f"⚠️ {cur} 該月份無資料，略過")
                    continue

                df = df.drop(columns=['幣別'])
                df.columns = [f'{cur}_{col}' for col in df.columns]
                all_data[cur] = df
                time.sleep(1.0)

            except Exception as e:
                log_func(f"⚠️ {cur} 無法抓取: {e}")

        browser.close()

    if not all_data:
        log_func("❌ 所有幣別都沒抓到資料，請確認來源網址是否又改版")
        return False, None

    final_df = pd.concat(all_data.values(), axis=1)

    expected = [f'{c}_{t}' for c in CURRENCIES
                for t in ['現金買入', '現金賣出', '即期買入', '即期賣出']]
    got = list(final_df.columns)
    missing_cur = [c for c in CURRENCIES if not any(col.startswith(f'{c}_') for col in got)]
    missing_col = [c for c in expected if c not in got]

    log_func(f"共抓到 {len(got)} 欄（應為 {len(expected)} 欄）")
    if missing_cur:
        log_func(f"⚠️ 完全沒抓到的幣別: {missing_cur}")
    if missing_col:
        log_func(f"⚠️ 缺少的欄位: {missing_col}")
    if not missing_col:
        log_func("✅ 幣別與欄位完整，跟原版結構一致")

    month_label = f'{int(month[5:7])}月平均數'
    avg_row = final_df.mean(numeric_only=True).to_frame().T
    avg_row.index = [month_label]

    rows = {}
    for cur in CURRENCIES:
        try:
            cash_cols = [f'{cur}_現金買入', f'{cur}_現金賣出']
            spot_cols = [f'{cur}_即期買入', f'{cur}_即期賣出']
            rows[f'{cur}_現金買入'] = final_df[cash_cols].mean(axis=1, numeric_only=True).mean()
            rows[f'{cur}_現金賣出'] = rows[f'{cur}_現金買入']
            rows[f'{cur}_即期買入'] = final_df[spot_cols].mean(axis=1, numeric_only=True).mean()
            rows[f'{cur}_即期賣出'] = rows[f'{cur}_即期買入']
        except KeyError:
            continue

    avg_by_type_row = pd.DataFrame(rows, index=['各幣別買入/賣出平均'])
    final_df = pd.concat([final_df, avg_row, avg_by_type_row])
    final_df = final_df.round(4)

    final_df.to_excel(output_path, engine='openpyxl')

    wb = load_workbook(output_path)
    ws = wb.active

    target_row = ws.max_row
    for cur in CURRENCIES:
        try:
            cash_buy_col = list(final_df.columns).index(f'{cur}_現金買入') + 2
            cash_sell_col = list(final_df.columns).index(f'{cur}_現金賣出') + 2
            spot_buy_col = list(final_df.columns).index(f'{cur}_即期買入') + 2
            spot_sell_col = list(final_df.columns).index(f'{cur}_即期賣出') + 2

            ws.merge_cells(start_row=target_row, end_row=target_row,
                           start_column=cash_buy_col, end_column=cash_sell_col)
            ws.merge_cells(start_row=target_row, end_row=target_row,
                           start_column=spot_buy_col, end_column=spot_sell_col)
        except ValueError:
            continue

    ws.append(['各幣別最後一天買入/賣出平均'] + [None] * (len(final_df.columns) - 1))
    target_row = ws.max_row

    for cur in CURRENCIES:
        try:
            cash_buy_col = list(final_df.columns).index(f'{cur}_現金買入') + 2
            cash_sell_col = list(final_df.columns).index(f'{cur}_現金賣出') + 2
            spot_buy_col = list(final_df.columns).index(f'{cur}_即期買入') + 2
            spot_sell_col = list(final_df.columns).index(f'{cur}_即期賣出') + 2

            ws.cell(row=target_row, column=cash_buy_col,
                    value=f'=AVERAGE({get_column_letter(cash_buy_col)}2:{get_column_letter(cash_sell_col)}2)')
            ws.cell(row=target_row, column=spot_buy_col,
                    value=f'=AVERAGE({get_column_letter(spot_buy_col)}2:{get_column_letter(spot_sell_col)}2)')

            ws.merge_cells(start_row=target_row, end_row=target_row,
                           start_column=cash_buy_col, end_column=cash_sell_col)
            ws.merge_cells(start_row=target_row, end_row=target_row,
                           start_column=spot_buy_col, end_column=spot_sell_col)
        except ValueError:
            continue

    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            if cell.value is not None:
                cell.number_format = '0.####'

    wb.save(output_path)

    log_func(f"✅ {int(month[5:7])} 月所有幣別匯率已存成功！")
    log_func(f"📁 檔案完整路徑：{output_path}")

    return True, output_path


# ============================================================
# 🎨 GUI 介面
# ============================================================
class ExRateApp:
    BRAND_BLUE = "#3B7DD8"
    BRAND_LIGHT_BLUE = "#EAF2FD"
    BRAND_GREEN = "#4CAF50"
    BRAND_GRAY = "#6B7280"

    def __init__(self, root):
        self.root = root
        self.root.title("💱 Wellell 匯率抓取小幫手")
        self.root.geometry("580x620")
        self.root.configure(bg="white")
        self.root.resizable(False, False)

        self.log_queue = queue.Queue()
        self.last_output_path = None
        self.is_running = False

        self._build_style()
        self._build_header()
        self._build_form()
        self._build_log_area()
        self._build_status_bar()

        self.root.after(150, self._poll_log_queue)

    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Brand.TButton",
                         background=self.BRAND_BLUE,
                         foreground="white",
                         font=("Microsoft JhengHei UI", 11, "bold"),
                         padding=8,
                         borderwidth=0)
        style.map("Brand.TButton",
                  background=[("active", "#2F63AD"), ("disabled", "#B7C6DE")])

        style.configure("Secondary.TButton",
                         background="#F1F5FB",
                         foreground=self.BRAND_BLUE,
                         font=("Microsoft JhengHei UI", 10),
                         padding=6,
                         borderwidth=1)
        style.map("Secondary.TButton",
                  background=[("active", "#DCE8FA"), ("disabled", "#F5F5F5")])

        style.configure("TCombobox", padding=4)

    def _build_header(self):
        header = tk.Frame(self.root, bg=self.BRAND_BLUE, height=70)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        title = tk.Label(header, text="💱 Wellell 匯率抓取小幫手",
                          bg=self.BRAND_BLUE, fg="white",
                          font=("Microsoft JhengHei UI", 16, "bold"))
        title.pack(side="left", padx=20, pady=15)

        subtitle = tk.Label(header, text="自動抓取台灣銀行 19 種幣別匯率",
                             bg=self.BRAND_BLUE, fg="#DCEBFF",
                             font=("Microsoft JhengHei UI", 9))
        subtitle.pack(side="left", padx=0, pady=(28, 0))

    def _build_form(self):
        form = tk.Frame(self.root, bg="white", padx=24, pady=18)
        form.pack(fill="x")

        tk.Label(form, text="📅 要抓取的年月", bg="white",
                  font=("Microsoft JhengHei UI", 11, "bold"),
                  fg=self.BRAND_GRAY).grid(row=0, column=0, sticky="w", pady=(0, 4))

        date_frame = tk.Frame(form, bg="white")
        date_frame.grid(row=1, column=0, sticky="w", pady=(0, 16))

        now = datetime.now()
        year_options = [str(y) for y in range(now.year - 5, now.year + 2)]
        month_options = [f"{m:02d}" for m in range(1, 13)]

        self.year_var = tk.StringVar(value=str(now.year))
        self.month_var = tk.StringVar(value=f"{now.month:02d}")

        year_box = ttk.Combobox(date_frame, textvariable=self.year_var,
                                 values=year_options, width=8, state="normal")
        year_box.grid(row=0, column=0, padx=(0, 4))
        tk.Label(date_frame, text="年（可直接輸入）", bg="white",
                  font=("Microsoft JhengHei UI", 8), fg="#9CA3AF").grid(row=0, column=1, padx=(0, 12))

        month_box = ttk.Combobox(date_frame, textvariable=self.month_var,
                                  values=month_options, width=6, state="readonly")
        month_box.grid(row=0, column=2, padx=(0, 4))
        tk.Label(date_frame, text="月", bg="white").grid(row=0, column=3)

        year_box.bind("<<ComboboxSelected>>", self._update_default_filename)
        year_box.bind("<KeyRelease>", self._update_default_filename)
        month_box.bind("<<ComboboxSelected>>", self._update_default_filename)

        tk.Label(form, text="📄 輸出檔名", bg="white",
                  font=("Microsoft JhengHei UI", 11, "bold"),
                  fg=self.BRAND_GRAY).grid(row=2, column=0, sticky="w", pady=(0, 4))

        self.filename_var = tk.StringVar()

        filename_entry = tk.Entry(form, textvariable=self.filename_var,
                                   font=("Microsoft JhengHei UI", 10),
                                   relief="solid", bd=1, width=48)
        filename_entry.grid(row=3, column=0, sticky="w", pady=(0, 4), ipady=4)

        self.folder_preview_var = tk.StringVar()
        tk.Label(form, textvariable=self.folder_preview_var, bg="white",
                  font=("Microsoft JhengHei UI", 8), fg="#9CA3AF", justify="left"
                  ).grid(row=4, column=0, sticky="w", pady=(0, 16))

        self._update_default_filename()

        btn_frame = tk.Frame(form, bg="white")
        btn_frame.grid(row=5, column=0, sticky="w")

        self.run_btn = ttk.Button(btn_frame, text="🚀 開始抓取",
                                   style="Brand.TButton", command=self._on_run_clicked)
        self.run_btn.grid(row=0, column=0, padx=(0, 10))

        self.open_folder_btn = ttk.Button(btn_frame, text="📁 開啟檔案位置",
                                           style="Secondary.TButton",
                                           command=self._on_open_folder_clicked,
                                           state="disabled")
        self.open_folder_btn.grid(row=0, column=1)

    def _update_default_filename(self, event=None):
        year = self.year_var.get().strip()
        month = self.month_var.get().strip()
        abbr = MONTH_ABBR.get(month, month)
        self.filename_var.set(f"exrate_{abbr}.xlsx")

        if year.isdigit() and len(year) == 4:
            folder = build_output_folder(year)
            self.folder_preview_var.set(f"將存到：{folder}")
        else:
            self.folder_preview_var.set("將存到：（請輸入正確的 4 位數年份）")

    def _build_log_area(self):
        log_frame = tk.Frame(self.root, bg="white", padx=24)
        log_frame.pack(fill="both", expand=True)

        tk.Label(log_frame, text="📋 執行紀錄", bg="white",
                  font=("Microsoft JhengHei UI", 10, "bold"),
                  fg=self.BRAND_GRAY).pack(anchor="w")

        text_frame = tk.Frame(log_frame, bg=self.BRAND_LIGHT_BLUE, bd=1, relief="solid")
        text_frame.pack(fill="both", expand=True, pady=(4, 12))

        self.log_text = tk.Text(text_frame, height=10, bg=self.BRAND_LIGHT_BLUE,
                                 fg="#374151", font=("Consolas", 9),
                                 relief="flat", wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(text_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        scrollbar.pack(side="right", fill="y")

    def _build_status_bar(self):
        self.status_var = tk.StringVar(value="🟢 準備就緒")
        status_bar = tk.Label(self.root, textvariable=self.status_var,
                               bg="#F3F4F6", fg=self.BRAND_GRAY, anchor="w",
                               font=("Microsoft JhengHei UI", 9), padx=12, pady=6)
        status_bar.pack(fill="x", side="bottom")

    def _append_log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log_from_thread(self, msg: str):
        self.log_queue.put(msg)

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log_queue)

    def _on_run_clicked(self):
        if self.is_running:
            return

        year_input = self.year_var.get().strip()
        if not year_input.isdigit() or len(year_input) != 4:
            messagebox.showwarning("提醒", "年份請輸入 4 位數字，例如 2026")
            return
        year_num = int(year_input)
        if year_num < 2000 or year_num > 2100:
            messagebox.showwarning("提醒", "年份數字看起來不太合理，請確認輸入是否正確")
            return

        month_str = f"{year_input}-{self.month_var.get()}"
        filename = self.filename_var.get().strip()

        if not filename:
            messagebox.showwarning("提醒", "請輸入輸出檔名")
            return
        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"
            self.filename_var.set(filename)

        output_folder = build_output_folder(year_input)
        try:
            os.makedirs(output_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror(
                "錯誤",
                f"無法建立或存取輸出資料夾：\n{output_folder}\n\n錯誤訊息：{e}\n\n"
                "請確認公用磁碟機（Y:）是否已連線，或該路徑是否有存取權限。"
            )
            return

        output_path = os.path.join(output_folder, filename)

        if os.path.exists(output_path):
            confirm = messagebox.askyesno(
                "檔案已存在",
                f"這個檔案已經存在：\n{output_path}\n\n"
                "繼續執行會覆蓋掉原本的檔案內容，確定要覆蓋嗎？"
            )
            if not confirm:
                self.status_var.set("🟡 已取消，未覆蓋原有檔案")
                return

        self.is_running = True
        self.run_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.status_var.set(f"🔵 正在抓取 {month_str} 的匯率資料，請稍候...")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        thread = threading.Thread(
            target=self._run_fetch_worker,
            args=(month_str, output_path),
            daemon=True
        )
        thread.start()

    def _run_fetch_worker(self, month_str, output_path):
        try:
            success, final_path = run_fetch(month_str, output_path, self._log_from_thread)
        except Exception as e:
            self._log_from_thread(f"❌ 發生未預期的錯誤：{e}")
            success, final_path = False, None

        self.root.after(0, self._on_fetch_done, success, final_path)

    def _on_fetch_done(self, success, final_path):
        self.is_running = False
        self.run_btn.configure(state="normal")

        if success:
            self.last_output_path = final_path
            self.open_folder_btn.configure(state="normal")
            self.status_var.set("✅ 抓取完成！")
            messagebox.showinfo("完成", f"匯率資料已成功抓取並存檔：\n{final_path}")
        else:
            self.status_var.set("❌ 抓取失敗，請查看執行紀錄")
            messagebox.showerror("失敗", "抓取失敗，請查看執行紀錄區塊的詳細訊息。")

    def _on_open_folder_clicked(self):
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            messagebox.showwarning("提醒", "找不到檔案，請先執行一次抓取。")
            return

        folder = os.path.dirname(self.last_output_path)
        try:
            if sys.platform == "win32":
                subprocess.run(["explorer", "/select,", os.path.normpath(self.last_output_path)])
            elif sys.platform == "darwin":
                subprocess.run(["open", folder])
            else:
                subprocess.run(["xdg-open", folder])
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟資料夾：{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExRateApp(root)
    root.mainloop()