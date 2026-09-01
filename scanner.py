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

current_month = datetime.now().month
current_year = datetime.now().year
today_date_str = datetime.now().strftime('%Y-%m-%d')
file_name = 'Master_Stock_Tracker_With_History.xlsx'

print('1. Fetching Nifty 500 Market Stocks...')
try:
  url = 'https://archives.nseindia.com/content/indices/ind_nifty500list.csv'
  headers = {'User-Agent': 'Mozilla/5.0'}
  res = requests.get(url, headers=headers)
  df_nifty = pd.read_csv(io.StringIO(res.text))
  watchlist = [str(symbol).strip() + '.NS' for symbol in df_nifty['Symbol'].tolist()]
except Exception:
  watchlist = ['RELIANCE.NS', 'TCS.NS', 'INFY.NS', 'SBIN.NS', 'HDFCBANK.NS', 'SUZLON.NS', 'TATAMOTORS.NS']

today_signals = []
historical_trades = []

# સ્ટાઇલિંગ સેટિંગ્સ
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

print('2. Running Strict Scan (Last 3 Years PERFECT Entry & Win Required)...')

for ticker in watchlist:
  try:
    stock = yf.Ticker(ticker)
    df = stock.history(period='10y', interval='1d')
    df = df.dropna(subset=['Close']) 
    
    if df.empty or len(df) < 200: continue

    avg_volume = df['Volume'].tail(20).mean()
    last_close = float(df['Close'].iloc[-1])
    if avg_volume * last_close < 10000000: continue

    df['RSI'] = ta.momentum.rsi(df['Close'], window=7)
    
    historical_rises = []
    total_fails = 0
    last_3_years_wins = 0  # છેલ્લા ૩ વર્ષમાં કેટલી વાર પાસ થયો તેનું ટ્રેકિંગ

    # પાછલા 7 વર્ષનું લૂપ (2019 થી 2025)
    for yr in range(current_year - 7, current_year):
      month_data = df[(df.index.year == yr) & (df.index.month == current_month)]
      entry_data = month_data[month_data['RSI'] < 35]
      
      if not entry_data.empty:
        entry_idx = entry_data.index[0]
        entry_price = float(entry_data['Close'].iloc[0])
        sl_price = entry_price * 0.95
        
        future_data = month_data[month_data.index >= entry_idx]
        
        max_high = float(future_data['High'].max())
        max_high_idx = future_data['High'].idxmax()
        
        rise_pct = round(((max_high - entry_price) / entry_price) * 100, 2)
        
        data_before_max = future_data[future_data.index <= max_high_idx]
        if data_before_max['Low'].min() <= sl_price:
          rise_pct = -5.0
          result = 'SL HIT'
        else:
          result = 'WIN'
          
        if rise_pct < 3.0:
          total_fails += 1
        else:
          historical_rises.append(rise_pct)
          # જો આ વર્ષ છેલ્લા 3 વર્ષમાં આવતું હોય (2025, 2024, 2023) તો જ Win ગણતરીમાં લેવો
          if yr >= current_year - 3:
            last_3_years_wins += 1
            
        historical_trades.append({
            'Year': yr, 'Ticker': ticker, 'Entry Date (RSI<35)': entry_idx.strftime('%Y-%m-%d'),
            'Entry Price': f"{entry_price:.2f}", 
            'Max High Date': max_high_idx.strftime('%Y-%m-%d'),
            'Max High Price': f"{max_high:.2f}", 
            'Max Rise %': f"{rise_pct}%", 
            'Result': result
        })
      else:
        # જો એન્ટ્રી જ ના મળી હોય તો કંઈ નહિ, બસ ખાલી એટલું ધ્યાન રાખો કે છેલ્લા ૩ વર્ષમાં એકાદ વર્ષ ખાલી નથી ગયું ને!
        pass

    # --- તમારી અસલી શરત અહી કામ કરશે ---
    # જો છેલ્લા 3 વર્ષમાં (2023, 2024, 2025) લગાતાર 3 એન્ટ્રી અને 3 ટાર્ગેટ હિટ ના થયા હોય તો સ્ટોક રીજેક્ટ!
    if last_3_years_wins < 3: 
        continue 
        
    # ટકાવારી લોજિક (બાકીના ૪ વર્ષોની ભૂલો પરથી)
    if total_fails == 0:
        prob_text = "100%"
    elif total_fails == 1:
        prob_text = "90%"
    elif total_fails == 2:
        prob_text = "80%"
    elif total_fails == 3:
        prob_text = "70%"
    else:
        continue # 3 થી વધુ વાર ફેલ થયો હોય તો કચરો કાઢો

    latest_rsi = float(df['RSI'].iloc[-1])
    
    if latest_rsi < 35 and historical_rises:
      min_rise = round(min(historical_rises), 2)
      target = last_close * (1 + (min_rise / 100))
      stop_loss = last_close * 0.95
      
      today_signals.append({
          'Date': today_date_str, 'Ticker': ticker, 
          'Price': f"{last_close:.2f}",
          'Probability': prob_text, 
          'Target %': f"+{min_rise}%", 
          'Target Price': f"{target:.2f}",
          'StopLoss': f"{stop_loss:.2f}", 
          'RSI(7)': f"{latest_rsi:.2f}", 
          'Status': 'GREEN SIGNAL (BUY)'
      })
  except Exception:
    pass

print("3. Generating Excel Sheets...")
if os.path.exists(file_name): wb = openpyxl.load_workbook(file_name)
else: wb = openpyxl.Workbook(); wb.remove(wb.active)

# Daily Tracker Sheet
if 'Daily_Tracker_Log' not in wb.sheetnames:
  ws_daily = wb.create_sheet('Daily_Tracker_Log')
  ws_daily.append(['Date', 'Ticker', 'Entry Price', 'Probability', 'Target %', 'Target Price', 'StopLoss', 'RSI(7)', 'Status'])
  for c in range(1, 10): ws_daily.cell(row=1, column=c).fill, ws_daily.cell(row=1, column=c).font, ws_daily.cell(row=1, column=c).alignment = header_fill, header_font, align_center
else: ws_daily = wb['Daily_Tracker_Log']

if today_signals:
  start_row = ws_daily.max_row + 2
  ws_daily.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=9)
  block_cell = ws_daily.cell(row=start_row, column=1, value=f'📅 SCAN DATE: {today_date_str} (RSI < 35 Triggered)')
  block_cell.fill, block_cell.font, block_cell.alignment = date_block_fill, bold_white_font, align_center
  
  for item in today_signals:
    ws_daily.append([item['Date'], item['Ticker'], item['Price'], item['Probability'], item['Target %'], item['Target Price'], item['StopLoss'], item['RSI(7)'], item['Status']])
    
    r_idx = ws_daily.max_row
    for c_idx in range(1, 10): 
        cell = ws_daily.cell(row=r_idx, column=c_idx)
        cell.alignment = align_center
        cell.border = thin_border
        if c_idx == 9:
            cell.fill = win_fill
            cell.font = green_font

# Historical Backtest Sheet
if 'Historical_RSI_Backtest' in wb.sheetnames: del wb['Historical_RSI_Backtest']
ws_hist = wb.create_sheet('Historical_RSI_Backtest')
hist_headers = ['Year', 'Ticker', 'Entry Date (RSI<35)', 'Entry Price', 'Max High Date', 'Max High Price', 'Max Rise %', 'Result']
ws_hist.append(hist_headers)
for c in range(1, 9): ws_hist.cell(row=1, column=c).fill, ws_hist.cell(row=1, column=c).font, ws_hist.cell(row=1, column=c).alignment = header_fill, header_font, align_center

for trade in historical_trades:
  ws_hist.append([trade['Year'], trade['Ticker'], trade['Entry Date (RSI<35)'], trade['Entry Price'], trade['Max High Date'], trade['Max High Price'], trade['Max Rise %'], trade['Result']])
  r_idx = ws_hist.max_row
  for c_idx in range(1, 9):
    cell = ws_hist.cell(row=r_idx, column=c_idx)
    cell.alignment, cell.border = align_center, thin_border
    if c_idx == 8:
      if trade['Result'] == 'WIN': cell.fill, cell.font = win_fill, green_font
      else: cell.fill, cell.font = loss_fill, loss_font

for ws in [ws_daily, ws_hist]:
  for col in ws.columns: ws.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(cell.value or '')) for cell in col) + 3, 14)

wb.save(file_name)
print(f"DONE! Strict 3-Year consecutive logic applied.")