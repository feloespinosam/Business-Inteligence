import yfinance as yf

tickers = ['CGAU', 'GOOGL', 'GOOG', 'ABEV', 'SAN', 'BBVA', 'GMED', 'MWA', 'OHI', 'TSM']

all_income    = set()
all_balance   = set()
all_cashflow  = set()

for ticker_symbol in tickers:
    try:
        ticker = yf.Ticker(ticker_symbol)
        
        if not ticker.income_stmt.empty:
            all_income.update(ticker.income_stmt.index.tolist())
        
        if not ticker.balance_sheet.empty:
            all_balance.update(ticker.balance_sheet.index.tolist())
        
        if not ticker.cash_flow.empty:
            all_cashflow.update(ticker.cash_flow.index.tolist())
            
        print(f'{ticker_symbol} OK')
    except Exception as e:
        print(f'{ticker_symbol} error: {e}')

print('\n─── INCOME STATEMENT ───')
for row in sorted(all_income):
    print(f"  '{row}',")

print('\n─── BALANCE SHEET ───')
for row in sorted(all_balance):
    print(f"  '{row}',")

print('\n─── CASH FLOW ───')
for row in sorted(all_cashflow):
    print(f"  '{row}',")