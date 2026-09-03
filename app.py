import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import re

# 頁面配置
st.set_page_config(
    page_title="MCD Forecast 管理系統 V2",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自訂 CSS 樣式 (隱藏透明頂部導覽列，一頁式緊湊排版)
st.markdown("""
<style>
    header[data-testid="stHeader"] {
        display: none !important;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        padding-bottom: 0.5rem !important;
    }
    .stApp {
        background-color: #ffffff;
    }
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1e293b;
        margin-top: 0px !important;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .metric-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 10px 14px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
        height: 100%;
    }
    .metric-title {
        font-size: 12px;
        color: #64748b;
        font-weight: 500;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 17px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 4px;
    }
    /* 💡 精準對齊的匯率雙欄容器 */
    .rate-grid {
        display: grid;
        grid-template-columns: auto 1fr;
        column-gap: 8px;
        row-gap: 2px;
        font-size: 15px;
        font-weight: 700;
        color: #0f172a;
        margin-top: 4px;
    }
    .rate-label {
        display: flex;
        justify-content: space-between;
        width: 85px; /* 鎖定文字區域寬度，實現兩端對齊 */
    }
    .current-month-badge {
        display: inline-block;
        background-color: #fef3c7;
        color: #b45309;
        font-size: 11px;
        font-weight: 700;
        padding: 1px 6px;
        border-radius: 10px;
    }
    .metric-badge {
        display: inline-block;
        background-color: #dbeafe;
        color: #1e40af;
        font-size: 11px;
        font-weight: 600;
        padding: 1px 6px;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 載入與解析數據
@st.cache_data(ttl=60)
def load_v2_data():
    file_path = None
    
    fcst_files = [
        f for f in os.listdir('.') 
        if f.startswith("FCST匯總表") and (f.endswith(".xlsx") or f.endswith(".xls"))
    ]
    
    if fcst_files:
        fcst_files.sort(key=lambda x: re.sub(r'\D', '', x), reverse=True)
        file_path = fcst_files[0]
    else:
        for fallback in ["data.xlsx", "total.xlsx", "total.txt"]:
            if os.path.exists(fallback):
                file_path = fallback
                break

    if file_path and os.path.exists(file_path):
        xls = pd.ExcelFile(file_path)
        sheet_name = '汇总' if '汇总' in xls.sheet_names else 0
        df_hz = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # 自動抓取 Excel 上方的最新匯率 (Row 0 D欄: 人民幣, Row 1 D欄: 台幣/美金)
        rmb_rate = float(df_hz.iloc[0, 3]) if pd.notna(df_hz.iloc[0, 3]) else 6.7894
        ntd_rate = float(df_hz.iloc[1, 3]) if pd.notna(df_hz.iloc[1, 3]) else 32.4650
        
        row3 = df_hz.iloc[3].tolist()
        row5 = df_hz.iloc[5].tolist()
        
        curr_year = "2025"
        time_cols = []
        cols_mapping = {}
        
        data_start_col = 7
        for idx in range(data_start_col, len(row5)):
            r3_val = str(row3[idx]).strip() if pd.notna(row3[idx]) else ""
            r5_val = str(row5[idx]).strip() if pd.notna(row5[idx]) else ""
            
            if re.match(r'^20\d{2}$', r3_val):
                curr_year = r3_val
                
            if '月' in r5_val and not any(k in r5_val for k in ['Total', 'TTL', '小計', '合計']):
                mo_num = re.sub(r'\D', '', r5_val)
                if mo_num:
                    col_key = f"{curr_year}-{int(mo_num):02d}"
                    time_cols.append(col_key)
                    cols_mapping[idx] = col_key

        df_data = df_hz.iloc[6:].copy()
        df_data[1] = df_data[1].ffill()
        
        clean_dict = {
            'Customer': df_data[1].astype(str).str.strip(),
            'HH_PN': df_data[2].astype(str).str.strip(),
            'Cust_PN': df_data[3].astype(str).str.strip(),
            'Category': df_data[4].fillna('未分類').astype(str).str.strip(),
            'UnitPrice': pd.to_numeric(df_data[5].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0),
            'Currency': df_data[6].astype(str).str.strip().str.upper()
        }
        
        for c_idx, c_name in cols_mapping.items():
            clean_dict[c_name] = pd.to_numeric(df_data[c_idx], errors='coerce').fillna(0)
            
        clean_df = pd.DataFrame(clean_dict)
        clean_df['Category'] = clean_df['Category'].replace({'nan': '未分類', 'None': '未分類', '': '未分類'})
        
        clean_df = clean_df[~clean_df['HH_PN'].str.contains(r'TOTAL|小計|合計|nan', case=False, na=False)]
        clean_df = clean_df[~clean_df['Customer'].str.contains(r'TOTAL|小計|合計|nan', case=False, na=False)]

        def get_rate_factor(curr):
            curr_str = str(curr).upper().strip()
            if any(x in curr_str for x in ['USD', 'US$', '美元']):
                return ntd_rate
            elif any(x in curr_str for x in ['RMB', 'CNY', '人民幣']):
                return ntd_rate / rmb_rate
            else:
                return 1.0
                
        clean_df['Rate_Factor'] = clean_df['Currency'].apply(get_rate_factor)
        
        for c in time_cols:
            clean_df[f'{c}_Rev'] = clean_df[c] * clean_df['UnitPrice'] * clean_df['Rate_Factor']
            
        filename_str = os.path.basename(file_path)
        match = re.search(r'(20\d{2})(\d{2})\d{2}', filename_str)
        if match:
            detected_mo = f"{match.group(1)}-{match.group(2)}"
        else:
            last_yr, last_mo = time_cols[-1].split('-')
            detected_mo = f"{int(last_yr) - 1}-{last_mo}"
            
        return clean_df, time_cols, rmb_rate, ntd_rate, detected_mo, file_path
    else:
        st.error(f"⚠️ 找不到資料檔案，請確認 Excel 檔案已上傳。")
        return pd.DataFrame(), [], 6.7894, 32.4650, "2026-08", ""

df_raw, time_cols, rmb_rate, ntd_rate, current_mo_str, current_loaded_file = load_v2_data()

if not df_raw.empty:
    # 主標題
    st.markdown('<div class="main-header">📊 MCD Forecast管理系統 V2</div>', unsafe_allow_html=True)

    # ---------------- 頂部控制與切換區 ----------------
    col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4 = st.columns([2.5, 2.5, 3, 2])
    
    with col_ctrl1:
        cust_options = ["所有客戶"] + sorted([x for x in df_raw['Customer'].unique() if x != "nan" and x != ""])
        selected_cust = st.selectbox("🏢 選擇客戶：", options=cust_options)

    with col_ctrl2:
        cat_options = sorted([x for x in df_raw['Category'].unique() if x != "nan" and x != ""])
        selected_cats = st.multiselect("🏷️ 選擇機種：", options=cat_options, default=[], placeholder="預設全部（可複選）")
        
    with col_ctrl3:
        search_pn = st.text_input("🔎 搜尋料號 (HH P/N 或 Cust P/N)：", placeholder="請輸入料號關鍵字...")
        
    with col_ctrl4:
        display_mode = st.radio(
            "🔘 顯示視角：",
            ["📦 出貨數量 (pcs)", "💰 預估營業額 (NTD)"],
            horizontal=False
        )

    is_revenue_mode = "營業額" in display_mode

    # ---------------- 資料過濾 ----------------
    df_filtered = df_raw.copy()
    if selected_cust != "所有客戶":
        df_filtered = df_filtered[df_filtered['Customer'] == selected_cust]

    if selected_cats:
        df_filtered = df_filtered[df_filtered['Category'].isin(selected_cats)]
        
    if search_pn:
        mask = (df_filtered['HH_PN'].astype(str).str.contains(search_pn, case=False, na=False)) | \
               (df_filtered['Cust_PN'].astype(str).str.contains(search_pn, case=False, na=False))
        df_filtered = df_filtered[mask]

    cols_2026 = [c for c in time_cols if '2026-' in c]
    active_cols = [f'{c}_Rev' for c in cols_2026] if is_revenue_mode else cols_2026

    # 計算統計數字
    total_pn = len(df_filtered)
    total_2026_val = df_filtered[active_cols].sum().sum()
    avg_2026_val = total_2026_val / len(cols_2026) if cols_2026 else 0
    
    cols_h1 = [f'{c}_Rev' for c in cols_2026[:6]] if is_revenue_mode else cols_2026[:6]
    cols_h2 = [f'{c}_Rev' for c in cols_2026[6:]] if is_revenue_mode else cols_2026[6:]
    val_h1 = df_filtered[cols_h1].sum().sum()
    val_h2 = df_filtered[cols_h2].sum().sum()

    st.markdown("---")

    # ---------------- 頂部 Metrics 卡片展示 ----------------
    unit_str = "NTD" if is_revenue_mode else "pcs"
    fmt_str = "${:,.0f}" if is_revenue_mode else "{:,.0f}"

    m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns([1.5, 2.2, 2.2, 2.1, 2.1])

    with m_col1:
        st.markdown(f'<div class="metric-card"><div class="metric-title">涵蓋料號數</div><div class="metric-value">{total_pn:,} 筆</div><div class="current-month-badge">📍 當前月份: {current_mo_str}</div></div>', unsafe_allow_html=True)

    with m_col2:
        st.markdown(f'<div class="metric-card"><div class="metric-title">2026 全年總{"營業額" if is_revenue_mode else "需求量"}</div><div class="metric-value">{fmt_str.format(total_2026_val)}</div><div class="metric-badge">↑ 均 {fmt_str.format(avg_2026_val)}/月</div></div>', unsafe_allow_html=True)

    avg_h1 = val_h1 / 6 if len(cols_h1) == 6 else 0
    avg_h2 = val_h2 / 6 if len(cols_h2) == 6 else 0

    with m_col3:
        st.markdown(f'<div class="metric-card"><div class="metric-title">2026 H1 (1-6月)</div><div class="metric-value">{fmt_str.format(val_h1)}</div><div class="metric-badge">↑ 均 {fmt_str.format(avg_h1)}/月</div></div>', unsafe_allow_html=True)

    with m_col4:
        st.markdown(f'<div class="metric-card"><div class="metric-title">2026 H2 (7-12月)</div><div class="metric-value">{fmt_str.format(val_h2)}</div><div class="metric-badge">↑ 均 {fmt_str.format(avg_h2)}/月</div></div>', unsafe_allow_html=True)

    # 💡 計算人民幣對台幣匯率 (RMB:NTD)
    rmb_to_ntd = ntd_rate / rmb_rate if rmb_rate > 0 else 0

    with m_col5:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-title">參考換算匯率</div>
            <div class="rate-grid">
                <div class="rate-label"><span>美</span><span>金</span><span>匯</span><span>率</span></div>
                <div>: {ntd_rate:.4f}</div>
                <div class="rate-label"><span>人</span><span>民</span><span>幣</span><span>匯</span><span>率</span></div>
                <div>: {rmb_to_ntd:.4f}</div>
            </div>
        </div>
        ''', unsafe_allow_html=True)

    # ---------------- 功能頁籤 ----------------
    tab1, tab2, tab3 = st.tabs(["📈 月度走勢分析", "📋 明細數據表", "📊 機種/客戶貢獻度分析"])

    # --- TAB 1: 月度走勢分析 ---
    with tab1:
        st.subheader(f"📈 各月份需求走勢圖 — [{display_mode}]")
        trend_cols = [f'{c}_Rev' for c in time_cols] if is_revenue_mode else time_cols
        sums = df_filtered[trend_cols].sum().reset_index()
        sums.columns = ["原始欄位", "數值"]
        sums["標準時間點"] = sums["原始欄位"].apply(lambda x: str(x).replace('_Rev', ''))
        
        seen_years = set()
        tick_text_list = []
        for col_name in sums["標準時間點"]:
            yr, mo = col_name.split('-')
            mo_num = int(mo)
            if yr not in seen_years:
                seen_years.add(yr)
                tick_text_list.append(f"{yr}年{mo_num}月")
            else:
                tick_text_list.append(f"{mo_num}月")
        
        fig_line = px.line(
            sums, x="標準時間點", y="數值", text="數值", markers=True,
            title=f"各月份 [{ '預估營業額 (NTD)' if is_revenue_mode else '出貨數量 (pcs)' }] 走勢 (📍 當前基準月：{current_mo_str})"
        )
        fig_line.update_traces(texttemplate='%{text:,.0f}', textposition="top center", line_color="#2563eb", line_width=3)
        
        if current_mo_str in sums["標準時間點"].values:
            curr_val = sums[sums["標準時間點"] == current_mo_str]["數值"].values[0]
            
            fig_line.add_shape(
                type="line",
                x0=current_mo_str,
                x1=current_mo_str,
                y0=0,
                y1=1,
                yref="paper",
                line=dict(color="#dc2626", width=2, dash="dash")
            )
            
            fig_line.add_annotation(
                x=current_mo_str,
                y=1,
                yref="paper",
                text=f"📍 當前月份 ({current_mo_str})",
                showarrow=False,
                font=dict(color="#dc2626", size=13),
                xanchor="right",
                yanchor="bottom"
            )
            
            fig_line.add_trace(go.Scatter(
                x=[current_mo_str],
                y=[curr_val],
                mode="markers+text",
                marker=dict(color="#dc2626", size=12, symbol="circle"),
                text=[f"📍 {curr_val:,.0f}"],
                textposition="bottom center",
                name="當前月份"
            ))

        fig_line.update_layout(
            height=380, 
            yaxis_title=unit_str, 
            xaxis_title="時間 (月份)", 
            showlegend=False,
            margin=dict(l=20, r=20, t=40, b=20),
            xaxis=dict(
                tickmode='array',
                tickvals=sums["標準時間點"].tolist(),
                ticktext=tick_text_list
            )
        )
        st.plotly_chart(fig_line, use_container_width=True)

    # --- TAB 2: 明細數據表 ---
    with tab2:
        st.subheader(f"📋 需求明細表 — [{display_mode}]")
        base_cols = ['Customer', 'HH_PN', 'Cust_PN', 'Category', 'UnitPrice', 'Currency']
        df_disp = df_filtered[base_cols].copy()
        
        for c in time_cols:
            if is_revenue_mode:
                df_disp[c] = df_filtered[f'{c}_Rev'].apply(lambda x: round(x, 0))
            else:
                df_disp[c] = df_filtered[c]
                
        rename_dict = {'Customer': '客戶', 'Category': '機種/分類', 'UnitPrice': '單價', 'Currency': '幣別'}
        df_disp = df_disp.rename(columns=rename_dict)
        st.dataframe(df_disp, use_container_width=True, height=450)

    # --- TAB 3: 機種與客戶雙維度貢獻度分析 ---
    with tab3:
        st.subheader(f"📊 機種 (Category) 與客戶雙維度分析 — [{display_mode}]")
        
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("#### 🏷️ 1. 各機種 (Category) 出貨/營收分布")
            cat_group = df_filtered.groupby("Category")[active_cols].sum()
            cat_group["Total"] = cat_group.sum(axis=1)
            cat_sorted = cat_group.sort_values(by="Total", ascending=False).head(10).reset_index()
            
            fig_cat = px.bar(
                cat_sorted, x="Category", y="Total", text="Total",
                color="Category", title=f"TOP 10 機種 [{ '營業額 (NTD)' if is_revenue_mode else '出貨量 (pcs)' }]"
            )
            fig_cat.update_traces(texttemplate='%{text:,.0f}', textposition="outside")
            fig_cat.update_layout(height=380, yaxis_title=unit_str, xaxis_title="機種", showlegend=False)
            st.plotly_chart(fig_cat, use_container_width=True)

        with c2:
            st.markdown("#### 🏢 2. 前 10 大客戶營收/出貨堆疊分析")
            cust_group = df_filtered.groupby("Customer")[active_cols].sum()
            cust_group["2026_Total"] = cust_group.sum(axis=1)
            top10_cust = cust_group.sort_values(by="2026_Total", ascending=False).head(10).reset_index()
            
            melt_cols = [c for c in top10_cust.columns if c != "Customer" and c != "2026_Total"]
            top10_melt = top10_cust.melt(id_vars=["Customer"], value_vars=melt_cols, var_name="月份", value_name="數值")
            top10_melt["月份標籤"] = top10_melt["月份"].apply(lambda x: str(x).replace('_Rev', ''))
            
            fig_bar = px.bar(
                top10_melt, x="月份標籤", y="數值", color="Customer",
                barmode="stack", title=f"TOP 10 客戶月份分布"
            )
            fig_bar.update_layout(height=380, yaxis_title=unit_str, xaxis_title="月份")
            st.plotly_chart(fig_bar, use_container_width=True)