import streamlit as st
import pandas as pd
import openpyxl
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go

# --- PAGE SETUP ---
st.set_page_config(
    page_title="Capital Utilisation Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS FOR EXACT MOCKUP AESTHETICS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap');
    
    /* Global Styles */
    html, body, [class*="css"], .stApp {
        font-family: 'Outfit', sans-serif !important;
        background-color: #0c0a1c !important;
        color: #f3f4f6 !important;
    }
    
    /* Sidebar Overrides */
    section[data-testid="stSidebar"] {
        background-color: #070614 !important;
        border-right: 1px solid #1b1735 !important;
    }
    
    /* Hide Streamlit Default UI Elements */
    .stDeployButton {
        display: none !important;
    }
    header[data-testid="stHeader"] {
        display: none !important;
    }
    #MainMenu {
        visibility: hidden;
    }
    footer {
        visibility: hidden;
    }
    
    /* Card Styles */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 16px;
        margin-bottom: 24px;
    }
    .card-purple-pink {
        background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%);
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 8px 16px rgba(236, 72, 153, 0.15);
        color: white;
    }
    .card-cyan-blue {
        background: linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%);
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 8px 16px rgba(6, 182, 212, 0.15);
        color: white;
    }
    .card-dark {
        background-color: #15132d;
        border: 1px solid #232049;
        border-radius: 16px;
        padding: 20px;
        position: relative;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
    }
    .card-title {
        font-size: 14px;
        font-weight: 300;
        opacity: 0.8;
    }
    .card-title-dark {
        font-size: 14px;
        font-weight: 400;
        color: #9ca3af;
    }
    .card-value {
        font-size: 26px;
        font-weight: 700;
        margin-top: 8px;
    }
    .badge {
        position: absolute;
        top: 15px;
        right: 15px;
        padding: 3px 8px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 600;
    }
    .badge-green {
        background-color: rgba(63, 185, 80, 0.15);
        color: #3fb950;
    }
    .badge-red {
        background-color: rgba(248, 81, 73, 0.15);
        color: #f85149;
    }
    
    /* Tabs Overrides */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #070614;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #1b1735;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 16px !important;
        border-radius: 8px !important;
        color: #9ca3af !important;
        background-color: transparent !important;
        border: none !important;
        font-weight: 600 !important;
    }
    .stTabs [aria-selected="true"] {
        color: white !important;
        background-color: #15132d !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.4) !important;
    }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR STRATEGY SELECTOR ---
st.sidebar.image("https://img.icons8.com/color/120/000000/dashboard.png", width=65)
st.sidebar.markdown("<h2 style='color:#ec4899; font-weight:700; font-size:22px; margin-top:10px;'>CRM Dashboard</h2>", unsafe_allow_html=True)
st.sidebar.write("---")

selected_strategy = st.sidebar.selectbox(
    "Select Strategy Model",
    [
        "Case 1: No Filter (Baseline)",
        "Case 2: Above High EMA",
        "Case 3: Within Band",
        "Case 4: Entry Below Band, Watch Cross",
        "Case 5: Within=Drop, Below=Watch Cross",
        "Case 6: No Filter, Fixed SL",
        "Case 7: No Filter, Trailing SL",
        "Case 8: Filtered Trailing SL (EMA50 Low Filter)"
    ]
)

ROOT = Path("D:/behaviour analysis")
case_num = int(selected_strategy.split(":")[0].split(" ")[1])
CONSOLIDATED_FILE = ROOT / f"12_Quarters_Reports_Case{case_num}" / f"12_Quarters_Consolidated_Case{case_num}.xlsx"
strategy_label = f"12-Quarters Walk-Forward Strategy ({selected_strategy})"
download_filename = f"12_Quarters_Consolidated_Case{case_num}.xlsx"

@st.cache_data
def load_consolidated_data(file_path):
    if not file_path.exists():
        return None, None
    
    # Load sheet names to get available quarters
    wb = openpyxl.load_workbook(file_path, read_only=True)
    sheets = [s for s in wb.sheetnames if s != "Detailed_Trades"]
    wb.close()
    
    # Read master detailed trades
    df_details = pd.read_excel(file_path, sheet_name="Detailed_Trades")
    
    # Load slots summary for each quarter
    quarters_data = {}
    for q in sheets:
        df_q = pd.read_excel(file_path, sheet_name=q, skiprows=6)
        # Drop the overall total row at the end
        if not df_q.empty and "Total" in str(df_q.iloc[-1, 0]):
            df_q = df_q.iloc[:-1]
        quarters_data[q] = df_q
        
    return quarters_data, df_details

quarters_data, df_details = load_consolidated_data(CONSOLIDATED_FILE)

if quarters_data is None:
    st.error("Error: Consolidated Excel file not found! Please run the consolidated build script first.")
    st.stop()

# --- SIDEBAR FILTERS ---
available_quarters = list(quarters_data.keys())
selected_q = st.sidebar.selectbox("Select Quarter", available_quarters)

# Get details for selected quarter
q_details = df_details[df_details["Quarter"] == selected_q].copy()

# Sector Filter
sectors = ["All"] + sorted(list(q_details["Sector"].dropna().unique()))
selected_sector = st.sidebar.selectbox("Filter by Sector", sectors)

# Symbol Search
search_symbol = st.sidebar.text_input("Search Stock Symbol", "").strip().upper()

# Apply Filters to Detailed Trades
filtered_details = q_details.copy()
if selected_sector != "All":
    filtered_details = filtered_details[filtered_details["Sector"] == selected_sector]
if search_symbol:
    filtered_details = filtered_details[filtered_details["Symbol"].str.contains(search_symbol)]

# --- MAIN PAGE HEADER ---
st.title("📊 Nifty 50 Capital Utilisation & Re-entry Dashboard")
st.markdown(f"**{strategy_label}** &middot; Quarter: `{selected_q}`")
st.write("---")

# --- CALCULATE LIVE KPIs FOR SELECTED QUARTER & FILTERS ---
slots_df = quarters_data[selected_q].copy()

# Filter Slots based on Sector and Symbol Search in their trade paths
filtered_slots = slots_df.copy()
if selected_sector != "All" or search_symbol:
    matching_trades = q_details.copy()
    if selected_sector != "All":
        matching_trades = matching_trades[matching_trades["Sector"] == selected_sector]
    if search_symbol:
        matching_trades = matching_trades[matching_trades["Symbol"].str.contains(search_symbol)]
    
    # Extract unique Slot IDs (e.g., "Slot 1", "Slot 2") involved in matching trades
    matching_slots = matching_trades["Assigned Slot"].unique()
    filtered_slots = filtered_slots[filtered_slots["Slot ID"].isin(matching_slots)]

# Calculate KPIs dynamically based on filtered slots
if filtered_slots.empty:
    fund_utilised = 0.0
    final_value = 0.0
    net_profit = 0.0
    compounded_return = 0.0
    total_reentries = 0
else:
    fund_utilised = filtered_slots["Start Capital (Rs.)"].sum()
    final_value = filtered_slots["Final Value (Rs.)"].sum()
    net_profit = filtered_slots["Net Profit (Rs.)"].sum()
    compounded_return = (final_value / fund_utilised - 1) * 100 if fund_utilised > 0 else 0.0
    total_reentries = int(filtered_slots["Number of Re-entries"].sum())

total_trades = len(filtered_details)

# --- KPI DISPLAY CARDS (CUSTOM HTML MATCHING MOCKUP) ---
profit_sign = "+" if net_profit >= 0 else ""
profit_color_class = "badge-green" if net_profit >= 0 else "badge-red"
comp_sign = "+" if compounded_return >= 0 else ""

html_kpi_block = f"""
<div class="kpi-container">
    <div class="card-purple-pink">
        <div class="card-title">Fund Utilised</div>
        <div class="card-value">₹{fund_utilised:,.2f}</div>
    </div>
    <div class="card-cyan-blue">
        <div class="card-title">Final Value</div>
        <div class="card-value">₹{final_value:,.2f}</div>
    </div>
    <div class="card-dark">
        <div class="card-title-dark">Net Profit</div>
        <div class="card-value" style="color: {'#3fb950' if net_profit >= 0 else '#f85149'};">₹{profit_sign}{net_profit:,.2f}</div>
        <span class="badge {profit_color_class}">{profit_sign}{compounded_return:.1f}%</span>
    </div>
    <div class="card-dark">
        <div class="card-title-dark">Compounded Return</div>
        <div class="card-value" style="color: {'#3fb950' if compounded_return >= 0 else '#f85149'};">{comp_sign}{compounded_return:.2f}%</div>
    </div>
    <div class="card-dark" style="display: flex; flex-direction: column; justify-content: center; gap: 8px;">
        <div style="font-size: 13px; font-weight: 400; color: #9ca3af;">
            Trades: <strong style="color: white; font-size:16px; margin-left:5px;">{total_trades}</strong>
        </div>
        <div style="font-size: 13px; font-weight: 400; color: #9ca3af;">
            Re-entries: <strong style="color: #a855f7; font-size:16px; margin-left:5px;">{total_reentries}</strong>
        </div>
    </div>
</div>
"""
st.markdown(html_kpi_block, unsafe_allow_html=True)

# --- TABS LAYOUT ---
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Slots Compounding Flow", 
    "📋 Detailed Transaction Audit Log", 
    "⚖️ Expected vs. Realised Performance",
    "💼 Sector Allocations & Returns"
])

# --- TAB 1: SLOTS FLOW ---
with tab1:
    st.subheader("Quarterly Capital Slot Summary")
    st.markdown("Each **Slot** represents a unit of capital compounding as trades exit and new ones re-enter the timeline.")
    
    if filtered_slots.empty:
        st.info("No compounded slots match the selected filters.")
    else:
        # A. Gantt Timeline Chart using Plotly
        st.write("### 📅 Capital Utilisation Timeline (Gantt Chart)")
        
        timeline_data = []
        for _, s_row in filtered_slots.iterrows():
            s_name = s_row["Slot ID"]
            s_trades = q_details[q_details["Assigned Slot"] == s_name].sort_values("Entry Date")
            for _, t in s_trades.iterrows():
                is_match = True
                if selected_sector != "All" and t["Sector"] != selected_sector:
                    is_match = False
                if search_symbol and search_symbol not in t["Symbol"]:
                    is_match = False
                    
                timeline_data.append({
                    "Slot ID": s_name,
                    "Symbol": t["Symbol"],
                    "Company Name": t["Company Name"],
                    "Sector": t["Sector"],
                    "Start": t["Entry Date"],
                    "End": t["Exit Date"],
                    "Buy Price": t["Buy Price (Rs.)"],
                    "Exit Price": t["Exit Price (Rs.)"],
                    "Return (%)": t["Realised Return (%)"],
                    "Expected (%)": t["Expected Return (%)"],
                    "Type": t["Re-entry Type"],
                    "Highlight": "Match" if is_match else "Other"
                })
        
        if timeline_data:
            df_tl = pd.DataFrame(timeline_data)
            df_tl["Label"] = df_tl["Symbol"] + " (" + df_tl["Return (%)"].map(lambda x: f"{x:+.1f}%") + ")"
            
            # Draw Gantt
            fig_tl = px.timeline(
                df_tl,
                x_start="Start",
                x_end="End",
                y="Slot ID",
                color="Return (%)",
                color_continuous_scale=["#f85149", "#f8a09d", "#d2e8d5", "#3fb950"], # Diverging color scheme
                text="Label",
                hover_data={
                    "Company Name": True,
                    "Sector": True,
                    "Buy Price": ":,.2f",
                    "Exit Price": ":,.2f",
                    "Expected (%)": ":+.2f",
                    "Return (%)": ":+.2f",
                    "Type": True,
                    "Start": "|%d-%b-%Y",
                    "End": "|%d-%b-%Y",
                    "Slot ID": False
                },
                template="plotly_dark"
            )
            fig_tl.update_layout(
                plot_bgcolor="#15132d",
                paper_bgcolor="#15132d",
                yaxis=dict(autorange="reversed", gridcolor="#232049"),
                xaxis=dict(gridcolor="#232049"),
                height=max(300, len(filtered_slots) * 22 + 100),
                coloraxis_colorbar=dict(title="Return %")
            )
            fig_tl.update_traces(textposition="inside", insidetextanchor="middle")
            st.plotly_chart(fig_tl, use_container_width=True)
        
        st.write("---")
        st.write("### 📑 Compound Slot Allocation Details")
        st.markdown("*Click on any Slot card below to expand and view its full sequential trade flow path.*")
        
        # B. Expandable List of Slot Cards
        for _, s_row in filtered_slots.iterrows():
            s_name = s_row["Slot ID"]
            start_cap = s_row["Start Capital (Rs.)"]
            final_cap = s_row["Final Value (Rs.)"]
            net_prof = s_row["Net Profit (Rs.)"]
            ret_pct = s_row["Compounded Return (%)"]
            re_entries = s_row["Number of Re-entries"]
            
            color_indicator = "🟢" if net_prof >= 0 else "🔴"
            expander_title = (
                f"{color_indicator} **{s_name}** &middot; "
                f"Utilised: **₹{start_cap:,.2f}** &rarr; Final: **₹{final_cap:,.2f}** &middot; "
                f"Profit: **₹{net_prof:+,.2f} ({ret_pct:+.2f}%)** &middot; Re-entries: **{re_entries}**"
            )
            
            with st.expander(expander_title):
                s_trades_detail = q_details[q_details["Assigned Slot"] == s_name].sort_values("Entry Date").copy()
                s_trades_display = s_trades_detail[[
                    "Symbol", "Company Name", "Sector", "Entry Date", "Exit Date",
                    "Buy Price (Rs.)", "Exit Price (Rs.)", "Expected Return (%)", "Realised Return (%)"
                ]].reset_index(drop=True)
                
                s_trades_display["Entry Date"] = pd.to_datetime(s_trades_display["Entry Date"]).dt.strftime("%Y-%m-%d")
                s_trades_display["Exit Date"] = pd.to_datetime(s_trades_display["Exit Date"]).dt.strftime("%Y-%m-%d")
                
                def color_rows(val):
                    color = '#3fb950' if val >= 0 else '#f85149'
                    return f'color: {color}; font-weight: bold;'
                
                st.dataframe(
                    s_trades_display.style.format({
                        "Buy Price (Rs.)": "₹{:,.2f}",
                        "Exit Price (Rs.)": "₹{:,.2f}",
                        "Expected Return (%)": "{:.2f}%",
                        "Realised Return (%)": "{:.2f}%"
                    }).map(color_rows, subset=["Realised Return (%)", "Expected Return (%)"]),
                    use_container_width=True
                )

# --- TAB 2: DETAILED TRADES ---
with tab2:
    st.subheader("Transaction Log Audit")
    st.write(f"Showing filtered trades for quarter {selected_q} (Matches: {len(filtered_details)} / {len(q_details)})")
    
    def color_returns(val):
        color = '#3fb950' if val >= 0 else '#f85149'
        return f'color: {color}; font-weight: bold;'
        
    format_cols = {
        "Buy Price (Rs.)": "₹{:,.2f}",
        "Exit Price (Rs.)": "₹{:,.2f}",
        "Expected Return (%)": "{:.2f}%",
        "Realised Return (%)": "{:.2f}%"
    }
    if "Slot Compounded Value" in filtered_details.columns:
        format_cols["Slot Compounded Value"] = "₹{:,.2f}"
    if "Realised Profit/Loss (Rs.)" in filtered_details.columns:
        format_cols["Realised Profit/Loss (Rs.)"] = "₹{:,.2f}"
    if "Realised Profit (Rs.)" in filtered_details.columns:
        format_cols["Realised Profit (Rs.)"] = "₹{:,.2f}"

    st.dataframe(
        filtered_details.style.format(format_cols).map(color_returns, subset=["Realised Return (%)", "Expected Return (%)"]),
        use_container_width=True
    )

# --- TAB 3: EXPECTED VS REALISED (SMOOTH AREA CHART MATCHING MOCKUP) ---
with tab3:
    st.subheader("Model Predictions vs. Actual Realised Outcomes")
    st.markdown("Compare the machine learning model's expected return with the actual realized trade performance.")
    
    q_details_sorted = q_details.sort_values("Entry Date")
    fig_comp = go.Figure()
    
    # Expected line
    fig_comp.add_trace(go.Scatter(
        x=q_details_sorted["Symbol"], 
        y=q_details_sorted["Expected Return (%)"],
        mode='lines+markers',
        name='Expected Return (%)',
        line=dict(color='#a855f7', width=2, shape='spline'), # Spline curve in pink-purple
        marker=dict(size=4)
    ))
    # Realised area chart matching the mockup's smooth glowing curves
    fig_comp.add_trace(go.Scatter(
        x=q_details_sorted["Symbol"], 
        y=q_details_sorted["Realised Return (%)"],
        mode='lines+markers',
        name='Realised Return (%)',
        line=dict(color='#00f3ff', width=3, shape='spline'), # Neon cyan line
        fill='tozeroy',
        fillcolor='rgba(0, 243, 255, 0.08)', # Translucent cyan area fill
        marker=dict(size=5, color='#00f3ff')
    ))
    
    fig_comp.update_layout(
        title="Expected vs. Realised Returns side-by-side (Spline Area Chart)",
        xaxis_title="Stock Symbol",
        yaxis_title="Return %",
        template="plotly_dark",
        plot_bgcolor="#15132d",
        paper_bgcolor="#15132d",
        xaxis=dict(tickangle=45, gridcolor="#232049"),
        yaxis=dict(gridcolor="#232049")
    )
    st.plotly_chart(fig_comp, use_container_width=True)
    
    # Summary stats
    mae = (q_details["Expected Return (%)"] - q_details["Realised Return (%)"]).abs().mean()
    win_rate = (q_details["Realised Return (%)"] > 0).mean() * 100
    st.markdown(f"""
    * **Average Absolute Deviation (MAE)**: `{mae:.2f}%` (How close the model's prediction was to actual outcomes on average)
    * **Historical Win Rate in Quarter**: `{win_rate:.1f}%` (Percent of positive realized return trades)
    """)

# --- TAB 4: SECTOR ANALYSIS (NEON DONUT AND GRADIENT BAR CHARTS) ---
with tab4:
    st.subheader("Sector Breakdown & Performance Analysis")
    
    col_l, col_r = st.columns(2)
    
    with col_l:
        # Donut chart with large center hole matching the mockup's circular charts
        sector_counts = q_details["Sector"].value_counts().reset_index()
        sector_counts.columns = ["Sector", "Count"]
        fig_sec_pie = px.pie(
            sector_counts, 
            values="Count", 
            names="Sector", 
            hole=0.62, # Large donut hole
            title="Trade Count Distribution by Sector",
            color_discrete_sequence=['#a855f7', '#00f3ff', '#3fb950', '#3b82f6', '#ec4899', '#f59e0b'],
            template="plotly_dark"
        )
        fig_sec_pie.update_layout(plot_bgcolor="#15132d", paper_bgcolor="#15132d")
        st.plotly_chart(fig_sec_pie, use_container_width=True)
        
    with col_r:
        # Horizontal Bar chart colored with gradient-like scale
        sector_perf = q_details.groupby("Sector")["Realised Return (%)"].mean().reset_index()
        sector_perf = sector_perf.sort_values("Realised Return (%)")
        fig_sec_bar = px.bar(
            sector_perf,
            x="Realised Return (%)",
            y="Sector",
            orientation="h",
            title="Average Realised Return % by Sector",
            color="Realised Return (%)",
            color_continuous_scale=["#f85149", "#00f3ff"], # Diverging color scheme
            template="plotly_dark"
        )
        fig_sec_bar.update_layout(
            plot_bgcolor="#15132d", 
            paper_bgcolor="#15132d",
            xaxis=dict(gridcolor="#232049"),
            yaxis=dict(gridcolor="#232049")
        )
        st.plotly_chart(fig_sec_bar, use_container_width=True)

# --- FOOTER & DOWNLOAD DIRECT Excel BUTTON ---
st.write("---")
with open(CONSOLIDATED_FILE, "rb") as f:
    st.download_button(
        label="📥 Download Consolidated Excel Report Workbook",
        data=f.read(),
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
