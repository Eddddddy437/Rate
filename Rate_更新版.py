import pandas as pd
import numpy as np
import requests
import time
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --- 設定抓取參數 ---
# 將幣別與月份設為可調整變數
currencies = ['USD', 'JPY', 'EUR', 'GBP', 'AUD', 'CAD', 'CNY', 'HKD',
              'SGD', 'CHF', 'SEK', 'ZAR', 'NZD', 'THB', 'PHP', 'IDR', 
              'KRW', 'VND', 'MYR']
target_month = '2026-03'
output_file = f'exrate_{target_month}.xlsx'
all_data = {}

# === 建立 Session (模擬標準瀏覽器行為) ===
session = requests.Session()
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

print(f"🚀 開始抓取 {target_month} 匯率資料...")

for cur in currencies:
    url = f'https://rate.bot.com.tw/xrt/quote/{target_month}/{cur}'
    try:
        response = session.get(url, headers=headers, timeout=15)
        
        # 簡易自動重試機制 (處理 503 Service Unavailable)
        if response.status_code != 200:
            print(f"⏳ {cur} 伺服器忙碌，3秒後重試...")
            time.sleep(3)
            response = session.get(url, headers=headers, timeout=15)
        
        response.raise_for_status()
        
        # 解析 HTML 表格
        bot_data = pd.read_html(response.text)
        df = bot_data[0]

        # 處理 MultiIndex 欄位標題
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(1)

        # 篩選並重新命名欄位
        df = df.iloc[:, [0, 2, 1, 3, 4, 5]]
        df.columns = ['掛牌日期', '幣別', '現金買入', '現金賣出', '即期買入', '即期賣出']

        # 資料清洗與校正
        df.loc[:, ['幣別', '現金買入']] = df[['現金買入', '幣別']].values
        df = df.replace('-----------------------', pd.NA)
        df = df.apply(pd.to_numeric, errors='ignore')
        
        df = df.set_index('掛牌日期')
        df = df.drop(columns=['幣別'])
        df.columns = [f'{cur}_{col}' for col in df.columns]

        all_data[cur] = df
        print(f"✅ {cur} 抓取成功")
        
        # 禮貌性延遲，避免對銀行伺服器造成負擔
        time.sleep(1.2)

    except Exception as e:
        print(f"❌ {cur} 無法抓取，錯誤原因: {e}")

# === 數據整合與統計計算 ===
if all_data:
    final_df = pd.concat(all_data.values(), axis=1)

    # 1. 計算整月平均
    avg_row = final_df.mean(numeric_only=True).to_frame().T
    avg_row.index = [f'{target_month}月平均數']

    # 2. 計算各幣別買賣中間價平均
    rows = {}
    for cur in currencies:
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

    # 存檔至 Excel
    final_df.to_excel(output_file, engine='openpyxl')

    # === Excel 格式美化 (合併儲存格) ===
    wb = load_workbook(output_file)
    ws = wb.active
    target_row = ws.max_row
    
    # 這裡保留你原本合併「各幣別買入/賣出平均」的邏輯
    # ... (省略重複的 openpyxl 合併程式碼) ...

    wb.save(output_file)
    print(f"🎉 任務完成！檔案已存至: {output_file}")
else:
    print("⚠️ 未能成功抓取任何資料。")
