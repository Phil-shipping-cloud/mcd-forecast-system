import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os

# 頁面配置
st.set_page_config(
    page_title="MCD Forecast 管理系統 V2",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自訂 CSS 樣式 (隱藏透明頂點工具列，解放頂部空白)
st.markdown("""
<style>
    /* 隱藏 Streamlit 原生頂部透明導覽列，徹底解決空白過大與文字裁切問題 */
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
        font-size: 18px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 4px;
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
    .current-month-badge {
        display: inline-block;
        background-color: #fef3c7;
        color: #b45309;
        font-size: 11px;
        font-weight: 700;
        padding: 1px 6px;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 載入與解析 Excel 資料
@st.cache_data(ttl=60)
def load_v2_data():
    file_name = "FCST匯總表_20260429.xlsx"
    file_path = file_name
    if not os.path.exists(file_path):
        file_path = os.path.join(os.path.dirname(__file__), file_name)
    
    if os.path.exists(file_path):
        xls = pd.ExcelFile(file_path)
        sheet_name = '汇总' if '汇总' in xls.sheet_names else 0
        df_hz = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # 讀取匯率
        rmb_rate = float(df_hz.iloc[0, 3]) if pd.notna(df_hz.iloc[0, 3]) else 6.9194
        ntd_rate = float(df_hz.iloc[1, 3]) if pd.notna(df_hz.iloc[1, 3]) else 31.335
        
        # 擷取明細數據列
        df_data = df_hz.iloc[6:1533].copy()
        df_data[1] = df_data[1].ffill()
        
        cols = [
            'Index', 'Customer', 'HH_PN', 'Cust_PN', 'Category', 'UnitPrice', 'Currency',
            '2025-01', '2025-02', '2025-03', '2025-04', '2025-05', '2025-06',
            '2025-07', '2025-08', '2025-09', '2025-10', '2025-11', '2025-12', '2025_Total',
            '2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06',
            '2026-07', '2026-08', '2026-09', '2026-10', '2026-11', '2026-12', '2026_Total',
            '2027-01', '2027-02', '2027-03', '2027-04'
        ]
        df_data.columns = cols
        
        # 清理欄位
        df_data['Customer'] = df_data['Customer'].astype(str).str.strip()
        df_data['Category'] = df_data['Category'].fillna('未分類').astype(str).str.strip()
        df_data['Category'] = df_data['Category'].replace({'nan': '未分類', 'None': '未分類', '': '未分類'})
        df_data['Currency'] = df_data['Currency'].astype(str).str.strip().str.upper()
        df_data['UnitPrice'] = pd.to_numeric(df_data['UnitPrice'], errors='coerce').fillna(0)
        
        time_cols = [c for c in cols if ('2025-' in c or '2026-' in c or '2027-' in c) and '_Total' not in c]
        for c in time_cols:
            df_data[c] = pd.to_numeric(df_data[c], errors='coerce').fillna(0)
            
        def get_rate_factor(curr):
            if curr == 'USD':
                return ntd_rate
            elif curr == 'RMB':
                return ntd_rate / rmb_rate
            else:
                return 1.0
                
        df_data['Rate_Factor'] = df_data['Currency'].apply(get_rate_factor)
        
        for c in time_cols:
            df_data[f'{c}_Rev'] = df_data[c] * df_data['UnitPrice'] * df_data['Rate_Factor']
            
        return df_data, time_cols, rmb_rate, ntd_rate
    else:
        st.error(f"⚠️ 找不到檔案 `{file_name}`")
        return pd.DataFrame(), [], 6.9194, 31.335

df_raw, time_cols, rmb_rate, ntd_rate = load_v2_data()

if not df_raw.empty:
    # 💡 自動推算「當前基準月份」邏輯：取表格最後一欄 (如 2027-04) 扣除 1 年 => 2026-04
    last_time_col = time_cols[-1]
    try:
        last_yr, last_mo = last_time_col.split('-')
        current_mo_str = f"{int(last_yr) - 1}-{last_mo}"
    except Exception:
        current_mo_str = "2026-04"

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

    with m_col5:
        st.markdown(f'<div class="metric-card"><div class="metric-title">參考換算匯率</div><div class="metric-value">31.335</div><div class="metric-badge">USD={ntd_rate} | RMB={rmb_rate}</div></div>', unsafe_allow_html=True)

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