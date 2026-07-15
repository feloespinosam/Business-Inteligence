# Business Intelligence - Financial Data Analysis & Visualization

![Python](https://img.shields.io/badge/Python-100%25-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?logo=streamlit&logoColor=white)
![yFinance](https://img.shields.io/badge/yFinance-Data-green)

A collection of Python tools for **financial data extraction, analysis, and visualization** using **yFinance**, **Streamlit**, **Plotly**, and **Pandas**.

---

## 📁 Project Structure

```text
Business-Inteligence/
│
├── Financial Dashboard/
│   └── financial_dashboard.py
│
├── Financial Data Fetcher/
│   ├── Financial Analysis Visa.py
│   └── Financial Statements Accounts.py
│
└── README.md
```

---

# 🚀 Modules

## 📊 Financial Dashboard (`financial_dashboard.py`)

An interactive dashboard built with **Streamlit** for visualizing the financial statements of companies listed on the NYSE.

### Features

- Select from popular NYSE companies:
  - JPM
  - V
  - WMT
  - DIS
  - KO
  - MCD
  - PFE
  - GS
  - BRK-B
  - CAT

- View:
  - Annual Financial Statements
  - Quarterly Financial Statements

- Key financial metrics:
  - Revenue
  - Net Income
  - Total Assets
  - Total Debt

- Interactive Plotly charts:
  - Revenue
  - Net Income
  - Operating Cash Flow

- Complete financial statements:
  - Income Statement
  - Balance Sheet
  - Cash Flow Statement

### Technologies

- Streamlit
- yFinance
- Pandas
- Plotly

---

## 📈 Financial Data Fetcher

A collection of scripts for downloading and analyzing financial information.

---

### 1. Financial Analysis Visa.py

Downloads financial data for a selected ticker (default: **TSLA**) and generates a professionally formatted Excel workbook containing financial ratio analysis.

### Financial Ratios

#### Liquidity

- Current Ratio
- Quick Ratio
- Cash Ratio

#### Profitability

- Gross Margin
- Operating Margin
- Net Profit Margin
- EBITDA Margin
- Return on Equity (ROE)
- Return on Assets (ROA)
- Return on Invested Capital (ROIC)

#### Leverage

- Debt to Equity
- Total Debt to Assets
- Interest Coverage
- Financial Leverage

#### Efficiency

- Asset Turnover
- Inventory Turnover

#### Valuation

- Trailing EPS
- Forward EPS
- Revenue per Share

#### Growth

- Revenue Growth
- Earnings Growth

#### Cash Flow

- Operating Cash Flow
- Free Cash Flow

### Excel Output

The generated workbook includes:

- Professional formatting
- Accounting number formats
- Percentage formatting
- Borders and alignment
- Styled headers

### Technologies

- yFinance
- Pandas
- OpenPyXL

---

### 2. Financial Statements Accounts.py

Utility script that explores all available financial statement accounts for multiple companies.

Example tickers:

```python
[
    "CGAU",
    "GOOGL",
    "GOOG",
    "ABEV",
    "SAN",
    "BBVA",
    "GMED",
    "MWA",
    "OHI",
    "TSM"
]
```

Extracts every available account from:

- Income Statement
- Balance Sheet
- Cash Flow Statement

Useful for:

- Understanding available financial metrics
- Standardizing company analysis
- Exploring Yahoo Finance financial datasets

### Technologies

- yFinance

---

# 🔧 Prerequisites

- Python 3.7+
- Internet connection
- Yahoo Finance access through yFinance

---

# 📦 Installation

Clone the repository:

```bash
git clone https://github.com/feloespinosam/Business-Inteligence.git
cd Business-Inteligence
```

Install dependencies:

```bash
pip install streamlit yfinance pandas plotly openpyxl
```

---

# ▶️ Usage

## Launch the Dashboard

```bash
cd "Financial Dashboard"
streamlit run financial_dashboard.py
```

---

## Run Financial Ratio Analysis

```bash
cd "Financial Data Fetcher"
python "Financial Analysis Visa.py"
```

This script generates an Excel workbook containing the calculated financial ratios.

---

## Explore Financial Statement Accounts

```bash
cd "Financial Data Fetcher"
python "Financial Statements Accounts.py"
```

Prints all available financial statement accounts for the configured tickers.

---

# 📝 Notes

- Financial data is downloaded in real time from **Yahoo Finance**.
- The dashboard uses `@st.cache_data` to improve performance.
- Excel exports include professional accounting formatting.
- Compatible with any ticker supported by **yFinance**.

---

# 🛠 Built With

- Python
- Streamlit
- Plotly
- Pandas
- yFinance
- OpenPyXL

---

# 👤 Author

**Felo Espinosa M.**

GitHub: **[@feloespinosam](https://github.com/feloespinosam)**

---

# 📄 License

This project is open-source and intended for **educational purposes** and **financial analysis**.
