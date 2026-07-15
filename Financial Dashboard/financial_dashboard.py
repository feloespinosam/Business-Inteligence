import streamlit as st
import yfinance as yf
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="NYSE Financial Insights",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for a premium, dark-mode-optimized look
st.markdown("""
<style>
    /* Main container styling */
    .main {
        background-color: #0e1117;
    }
    
    /* Header styling */
    .stHeading h1 {
        color: #f0f2f6;
        font-family: 'Inter', sans-serif;
        font-weight: 700;
        letter-spacing: -1px;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-image: linear-gradient(#1e1e1e, #0e1117);
        border-right: 1px solid #333;
    }
    
    /* Metric container styling */
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
        color: #00d4ff;
    }
    
    /* Custom tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
        background-color: transparent;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px;
        color: #808495;
        font-weight: 600;
        font-size: 1rem;
    }

    .stTabs [aria-selected="true"] {
        color: #00d4ff !important;
        border-bottom-color: #00d4ff !important;
    }
    
    /* Dataframe styling */
    [data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# Predefined popular NYSE stock list
NYSE_STOCKS = {
    "JPM": "JPMorgan Chase & Co.",
    "V": "Visa Inc.",
    "WMT": "Walmart Inc.",
    "DIS": "The Walt Disney Company",
    "KO": "The Coca-Cola Company",
    "MCD": "McDonald's Corporation",
    "PFE": "Pfizer Inc.",
    "GS": "The Goldman Sachs Group",
    "BRK-B": "Berkshire Hathaway Inc.",
    "CAT": "Caterpillar Inc."
}

@st.cache_data(ttl=3600)
def fetch_financial_data(ticker_symbol, period_type):
    """
    Fetches financial statements for a given ticker and period type.
    """
    ticker = yf.Ticker(ticker_symbol)
    
    if period_type == "Annual":
        income_stmt = ticker.financials
        balance_sheet = ticker.balance_sheet
        cash_flow = ticker.cashflow
    else:
        income_stmt = ticker.quarterly_financials
        balance_sheet = ticker.quarterly_balance_sheet
        cash_flow = ticker.quarterly_cashflow
        
    # Slice for last 2 periods
    income_stmt = income_stmt.iloc[:, :2] if not income_stmt.empty else income_stmt
    balance_sheet = balance_sheet.iloc[:, :2] if not balance_sheet.empty else balance_sheet
    cash_flow = cash_flow.iloc[:, :2] if not cash_flow.empty else cash_flow
    
    # Format headers to readable dates
    def format_headers(df):
        if not df.empty:
            df.columns = [col.strftime('%Y-%m-%d') if isinstance(col, datetime) else col for col in df.columns]
        return df

    return format_headers(income_stmt), format_headers(balance_sheet), format_headers(cash_flow)

# Sidebar UI
st.sidebar.title("💎 Finance Visualizer")
st.sidebar.divider()

selected_symbol = st.sidebar.selectbox(
    "Select NYSE Stock",
    options=list(NYSE_STOCKS.keys()),
    format_func=lambda x: f"{x} - {NYSE_STOCKS[x]}"
)

selected_period = st.sidebar.radio(
    "Reporting Period",
    options=["Annual", "Quarterly"],
    index=0,
    horizontal=True
)

st.sidebar.info(f"Displaying {selected_period} data for the last 2 periods of availability.")

# Main content
st.title(f"{NYSE_STOCKS[selected_symbol]} ({selected_symbol})")

try:
    with st.spinner("Fetching financial data..."):
        income_df, balance_df, cash_df = fetch_financial_data(selected_symbol, selected_period)

    if income_df.empty or balance_df.empty or cash_df.empty:
        st.warning("Could not retrieve all financial statements for this ticker. Yahoo Finance may have limited data for this period.")
    else:
        # Key Metrics Row
        cols = st.columns(4)
        
        # Extract some key metrics for the summary
        try:
            rev_label = "Total Revenue" if "Total Revenue" in income_df.index else income_df.index[0]
            net_inc_label = "Net Income" if "Net Income" in income_df.index else income_df.index[income_df.index.str.contains("Net Income", case=False)][0]
            
            latest_rev = income_df.loc[rev_label].iloc[0]
            prev_rev = income_df.loc[rev_label].iloc[1]
            rev_delta = ((latest_rev - prev_rev) / prev_rev) * 100
            
            latest_ni = income_df.loc[net_inc_label].iloc[0]
            prev_ni = income_df.loc[net_inc_label].iloc[1]
            ni_delta = ((latest_ni - prev_ni) / prev_ni) * 100
            
            cols[0].metric("Revenue", f"${latest_rev/1e9:.2f}B", f"{rev_delta:.1f}%")
            cols[1].metric("Net Income", f"${latest_ni/1e9:.2f}B", f"{ni_delta:.1f}%")
            
            # Asset metric
            asset_label = "Total Assets" if "Total Assets" in balance_df.index else balance_df.index[balance_df.index.str.contains("Total Assets", case=False)][0]
            latest_assets = balance_df.loc[asset_label].iloc[0]
            cols[2].metric("Total Assets", f"${latest_assets/1e9:.2f}B")
            
            # Cash metric
            cash_label = "Free Cash Flow" if "Free Cash Flow" in cash_df.index else cash_df.index[cash_df.index.str.contains("Free Cash Flow", case=False)][0]
            latest_fcf = cash_df.loc[cash_label].iloc[0]
            cols[3].metric("Free Cash Flow", f"${latest_fcf/1e9:.2f}B")
        except Exception:
            pass # Gracefully skip if specific metric extraction fails
            
        st.divider()

        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Income Statement", "🏥 Balance Sheet", "💸 Cash Flow", "📈 Visual Analytics"])

        with tab1:
            st.subheader("Income Statement")
            st.dataframe(income_df.style.format(precision=0, thousands=","), use_container_width=True)

        with tab2:
            st.subheader("Balance Sheet")
            st.dataframe(balance_df.style.format(precision=0, thousands=","), use_container_width=True)

        with tab3:
            st.subheader("Cash Flow Statement")
            st.dataframe(cash_df.style.format(precision=0, thousands=","), use_container_width=True)

        with tab4:
            st.subheader("Visual Analytics")
            
            # Chart 1: Revenue & Net Income Trend
            try:
                rev_data = income_df.loc[rev_label]
                ni_data = income_df.loc[net_inc_label]
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=rev_data.index[::-1],
                    y=rev_data.values[::-1],
                    name="Total Revenue",
                    marker_color="#00d4ff"
                ))
                fig.add_trace(go.Scatter(
                    x=ni_data.index[::-1],
                    y=ni_data.values[::-1],
                    name="Net Income",
                    line=dict(color="#ff4b4b", width=4),
                    mode='lines+markers'
                ))
                
                fig.update_layout(
                    title="Revenue vs Net Income Trend",
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    yaxis_title="USD",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig, use_container_width=True)
            except:
                st.info("Chart data not available for this ticker.")

            # Chart 2: Asset Composition (Placeholder for Assets/Liabilities)
            try:
                liab_label = "Total Liabilities Net Minority Interest" if "Total Liabilities Net Minority Interest" in balance_df.index else balance_df.index[balance_df.index.str.contains("Total Liabilities", case=False)][0]
                equity_label = "Total Equity Gross Minority Interest" if "Total Equity Gross Minority Interest" in balance_df.index else balance_df.index[balance_df.index.str.contains("Total Equity", case=False)][0]
                
                latest_liab = balance_df.loc[liab_label].iloc[0]
                latest_equity = balance_df.loc[equity_label].iloc[0]
                
                labels = ['Total Liabilities', 'Total Equity']
                values = [latest_liab, latest_equity]
                
                fig2 = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.4, marker_colors=["#333", "#00d4ff"])])
                fig2.update_layout(
                    title=f"Capital Structure ({balance_df.columns[0]})",
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)"
                )
                st.plotly_chart(fig2, use_container_width=True)
            except:
                pass

except Exception as e:
    st.error(f"An error occurred: {e}")
    st.info("Try selecting a different ticker or reporting period.")

# Footer
st.divider()
st.caption("Data provided by Yahoo Finance via yfinance library. Visuals built with Streamlit and Plotly.")