from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import ta
import yfinance as yf
import warnings
warnings.filterwarnings('ignore')

current_year = datetime.now().year
current_month = datetime.now().month
today_date_str = datetime.now().strftime('%Y-%m-%d')
file_name = 'Master_Stock_Tracker_With_History.xlsx'

print('1. Fetching Nifty 500, Midcap & Smallcap Market Stocks...')
watchlist = []
try:
  urls = [
      'https://archives.nseindia.com/content/indices/ind_nifty500list.csv',
      'https://archives.nseindia.com/content/indices/ind_niftymidcap100list.csv',
      'https://archives.nseindia.com/content/indices/ind_niftysmallcap100list.csv'
  ]
  headers = {'User-Agent': 'Mozilla/5.0'}
  for url in urls:
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
      df_temp = pd.read_csv(io.StringIO(res.text))
      symbols = [str(s).strip() + '.NS' for s in df_temp['Symbol'].tolist()]
      watchlist.extend(symbols)
  watchlist = list(set(watchlist))
except Exception:
  watchlist = ['RELIANCE.NS', 'TCS.NS', 'INFY.NS', 'SBIN.NS', 'HDFCBANK.NS', 'SUZLON.NS', 'TATAMOTORS.NS']

today_signals = []
historical_summary = []

header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
date_block_fill = PatternFill(start_color='203764', end_color='203764', fill_type='solid')
win_fill = PatternFill(start_color='D9EAD3', end_color='D9EAD3', fill_type='solid')
loss_fill = PatternFill(start_color='FCE5CD', end_color='FCE5CD', fill_type='solid')

header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
bold_white_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
green_font = Font(name='Calibri', size=11, bold=True, color='274E13')
loss_font = Font(name='Calibri', size=11, bold=True, color='783F04')
align_center = Alignment(horizontal='center', vertical='center')
thin_border = Border(left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'),
                     top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD'))

print('2. Running Final Scanner with Risk-Reward & Year-Wise Backtest Summary...')

for ticker in watchlist:
  try:
    stock = yf.Ticker(ticker)
    df = stock.history(period='10y', interval='1d')
    df = df.dropna(subset=['Close']) 
    
    if df.empty or len(df) < 200: continue

    last_close = float(df['Close'].iloc[-1])
    
    # રૂ. ૩૦ થી નીચેના સ્ટોક અવોઈડ કરવા
    if last_close < 30.0: continue

    # સર્કિટ ફિલ્ટર
    today_high = float(df['High'].iloc[-1])
    today_low = float(df['Low'].iloc[-1])
    today_vol = float(df['Volume'].iloc[-1])
    if today_high == today_low or today_vol == 0: continue

    df['RSI'] = ta.momentum.rsi(df['Close'], window=7)
    
    historical_rises = []
    total_fails = 0
    last_3_years_wins = 0  
    total_entries_found = 0
    total_trades_count = 0
    
    # વર્ષ-વાઈઝ ડેટા કલેક્ટ કરવા માટે ડિિક્શનરી
    year_stats = {}

    for yr in range(current_year - 7, current_year):
      yr_wins = 0
      yr_sl = 0
      yr_trades = 0
      
      for m in range(1, 13):
        if yr == current_year - 1 and m > current_month: continue

        month_data = df[(df.index.year == yr) & (df.index.month == m)]
        entry_data = month_data[month_data['RSI'] < 35]
        
        if not entry_data.empty:
          total_entries_found += 1
          entry_idx = entry_data.index[0]
          entry_price = float(entry_data['Close'].iloc[0])
          sl_price = entry_price * 0.97  # ૩% સ્ટોપ લોસ
          
          future_data = month_data[month_data.index >= entry_idx]
          if len(future_data) < 2: continue
          
          total_trades_count += 1
          yr_trades += 1
          max_high = float(future_data['High'].max())
          max_high_idx = future_data['High'].idxmax()
          
          rise_pct = round(((max_high - entry_price) / entry_price) * 100, 2)
          data_before_max = future_data[future_data.index <= max_high_idx]
          
          if data_before_max['Low'].min() <= sl_price:
            rise_pct = -3.0
            result = 'SL HIT'
            yr_sl += 1
          else:
            result = 'WIN'
            yr_wins += 1
            
          if rise_pct < 3.0:
            total_fails += 1
          else:
            historical_rises.append(rise_pct)
            if yr >= current_year - 3:
              last_3_years_wins += 1

      if yr_trades > 0:
        win_rate_yr = round((yr_wins / yr_trades) * 100, 1)
        historical_summary.append({
            'Year': yr, 'Ticker': ticker, 'Total Trades': yr_trades, 
            'Wins': yr_wins, 'SL Hits': yr_sl, 'Success Rate %': f"{win_rate_yr}%"
        })

    if total_entries_found == 0: continue
    if last_3_years_wins < 3: continue  

    # રિસ્ક-રેવર્ડ ફિલ્ટર: મિનિમમ રાઈઝ ફરજિયાત ૩% કે તેથી વધુ જ હોવો જોઈએ!
    if not historical_rises: continue
    min_rise = round(min(historical_rises), 2)
    if min_rise < 3.0: continue

    if total_fails == 0: 
        prob_score = 100
        prob_text = f"100% ({total_trades_count - total_fails} Wins / {total_trades_count} Trades)"
    elif total_fails == 1: 
        prob_score = 90
        prob_text = f"90% ({total_trades_count - total_fails} Wins / {total_trades_count} Trades)"
    elif total_fails == 2: 
        prob_score = 80
        prob_text = f"80% ({total_trades_count - total_fails} Wins / {total_trades_count} Trades)"
    elif total_fails == 3: 
        prob_score = 70
        prob_text = f"70% ({total_trades_count - total_fails} Wins / {total_trades_count} Trades)"
    else: continue 

    latest_rsi = float(df['RSI'].iloc[-1])
    
    if latest_rsi < 35:
      max_rise = round(max(historical_rises), 2)
      min_target = last_close * (1 + (min_rise / 100))
      max_target = last_close * (1 + (max_rise / 100))
      stop_loss = last_close * 0.97  # ૩% કડક સ્ટોપલોસ
      
      today_signals.append({
          'Score': prob_score,
          'Date': today_date_str, 'Ticker': ticker, 
          'Price': f"{last_close:.2f}", 'Probability': prob_text, 
          'Min Target % (Exit >= 3%)': f"+{min_rise}%", 'Min Target Price': f"{min_target:.2f}",
          'Max Target %': f"+{max_rise}%", 'Max Target Price': f"{max_target:.2f}",
          'StopLoss (-3%)': f"{stop_loss:.2f}", 'RSI(7)': f"{latest_rsi:.2f}", 
          'Status': 'GREEN SIGNAL (BUY)'
      })
  except Exception:
    pass

today_signals = sorted(today_signals, key=lambda x: x['Score'], reverse=True)

print("3. Generating Excel Sheets...")
if os.path.exists(file_name): wb = openpyxl.load_workbook(file_name)
else: wb = openpyxl.Workbook(); wb.remove(wb.active)

# Daily Tracker Sheet
if 'Daily_Tracker_Log' not in wb.sheetnames:
  ws_daily = wb.create_sheet('Daily_Tracker_Log')
  ws_daily.append(['Date', 'Ticker', 'Entry Price', 'Probability', 'Min Target % (Exit >= 3%)', 'Min Target Price', 'Max Target %', 'Max Target Price', 'StopLoss (-3%)', 'RSI(7)', 'Status'])
  for c in range(1, 12): ws_daily.cell(row=1, column=c).fill, ws_daily.cell(row=1, column=c).font, ws_daily.cell(row=1, column=c).alignment = header_fill, header_font, align_center
else: ws_daily = wb['Daily_Tracker_Log']

if today_signals:
  start_row = ws_daily.max_row + 2
  ws_daily.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=11)
  block_cell = ws_daily.cell(row=start_row, column=1, value=f'📅 SCAN DATE: {today_date_str} (Risk-Reward Min 1:1 Filter Applied)')
  block_cell.fill, block_cell.font, block_cell.alignment = date_block_fill, bold_white_font, align_center
  
  for item in today_signals:
    ws_daily.append([item['Date'], item['Ticker'], item['Price'], item['Probability'], item['Min Target % (Exit >= 3%)'], item['Min Target Price'], item['Max Target %'], item['Max Target Price'], item['StopLoss (-3%)'], item['RSI(7)'], item['Status']])
    r_idx = ws_daily.max_row
    for c_idx in range(1, 12): 
        cell = ws_daily.cell(row=r_idx, column=c_idx)
        cell.alignment = align_center
        cell.border = thin_border
        if c_idx == 11: cell.fill, cell.font = win_fill, green_font

# Historical Backtest Sheet (Year-Wise Summary Format)
if 'Historical_RSI_Backtest' in wb.sheetnames: del wb['Historical_RSI_Backtest']
ws_hist = wb.create_sheet('Historical_RSI_Backtest')
hist_headers = ['Year', 'Ticker', 'Total Trades', 'Wins', 'SL Hits', 'Success Rate %']
ws_hist.append(hist_headers)
for c in range(1, 7): ws_hist.cell(row=1, column=c).fill, ws_hist.cell(row=1, column=c).font, ws_hist.cell(row=1, column=c).alignment = header_fill, header_font, align_center

for summary in historical_summary:
  ws_hist.append([summary['Year'], summary['Ticker'], summary['Total Trades'], summary['Wins'], summary['SL Hits'], summary['Success Rate %']])
  r_idx = ws_hist.max_row
  for c_idx in range(1, 7):
    cell = ws_hist.cell(row=r_idx, column=c_idx)
    cell.alignment, cell.border = align_center, thin_border
    if c_idx == 6:
      cell.fill, cell.font = win_fill, green_font

for ws in [ws_daily, ws_hist]:
  for col in ws.columns: ws.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(cell.value or '')) for cell in col) + 3, 14)

wb.save(file_name)
print(f"DONE! Final Master Scanner with Year-Wise Backtest & Risk-Reward filter generated successfully.")
