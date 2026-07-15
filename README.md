Business Intelligence - Financial Data Analysis & Visualization
https://img.shields.io/badge/Python-100%2525-blue
https://img.shields.io/badge/Streamlit-FF4B4B?logo=streamlit&logoColor=white
https://img.shields.io/badge/yFinance-Data-green

This repository contains a set of tools for financial data analysis, visualization, and extraction using Python, yFinance, Streamlit, and Plotly.

📁 Project Structure
text
Business-Inteligence/
├── Financial Dashboard/
│   └── financial_dashboard.py      # Interactive dashboard for financial statement visualization
├── Financial Data Fetcher/
│   ├── Financial Analysis Visa.py  # Detailed financial ratio analysis and Excel export
│   └── Financial Statements Accounts.py  # Exploration of financial statement accounts
└── README.md
🚀 Modules
1. Financial Dashboard (financial_dashboard.py)
Interactive dashboard built with Streamlit that allows you to visualize the financial statements of companies listed on the NYSE.

Key Features:

Selection of popular NYSE stocks (JPM, V, WMT, DIS, KO, MCD, PFE, GS, BRK-B, CAT)

Visualization of financial statements (Annual or Quarterly)

Key metrics: Revenue, Net Income, Total Assets, Total Debt

Interactive charts with Plotly for revenue, net income, and operating cash flow

Detailed tables of the three main financial statements

Technologies: Streamlit, yFinance, Pandas, Plotly

2. Financial Data Fetcher
A set of scripts for extracting and analyzing financial data.

a) Financial Analysis Visa.py
Script that downloads financial data for a given ticker (default: TSLA) and generates a comprehensive financial ratio analysis exported to a professionally formatted Excel file.

Calculated Ratios:

Liquidity: Current Ratio, Quick Ratio, Cash Ratio

Profitability: Gross Margin, Operating Margin, Net Profit Margin, EBITDA Margin, ROE, ROA, ROIC

Leverage: Debt to Equity, Total Debt to Assets, Interest Coverage, Financial Leverage

Efficiency: Asset Turnover, Inventory Turnover

Valuation: EPS (Trailing and Forward), Revenue per Share

Growth: Revenue Growth, Earnings Growth

Cash Flow: Operating Cash Flow, Free Cash Flow

Excel Format: Data is exported with accounting formats, percentages, and professional styling (fonts, borders, alignment).

Technologies: yFinance, Pandas, OpenPyXL

b) Financial Statements Accounts.py
Exploration script that iterates over a list of tickers (example: ['CGAU', 'GOOGL', 'GOOG', 'ABEV', 'SAN', 'BBVA', 'GMED', 'MWA', 'OHI', 'TSM']) and extracts all available accounts from the three financial statements:

Income Statement

Balance Sheet

Cash Flow

Useful for identifying which metrics are available for each company and standardizing analysis.

Technologies: yFinance

🔧 Prerequisites
Python 3.7 or higher

Internet connection (to download data from Yahoo Finance)

📦 Installation
Clone the repository:

bash
git clone https://github.com/feloespinosam/Business-Inteligence.git
cd Business-Inteligence
Install the required dependencies:

bash
pip install streamlit yfinance pandas plotly openpyxl
▶️ Usage
Dashboard (Streamlit)
bash
cd "Financial Dashboard"
streamlit run financial_dashboard.py
Ratio Analysis (Visa/TSLA)
bash
cd "Financial Data Fetcher"
python "Financial Analysis Visa.py"
This script will generate an Excel file with the financial ratios.

Account Exploration
bash
cd "Financial Data Fetcher"
python "Financial Statements Accounts.py"
Prints all available accounts for the configured tickers to the console.

📝 Notes
Data is fetched in real-time from Yahoo Finance using the yfinance library.

The dashboard uses caching (@st.cache_data) to optimize performance.

The generated Excel files include accounting formats and professional styles for easy interpretation.

👤 Author
Felo Espinosa M.
GitHub: @feloespinosam

📄 License
This project is open-source and available for educational and financial analysis purposes.
