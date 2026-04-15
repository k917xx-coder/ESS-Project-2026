"""
凱綸儲能效益評估系統 - 專業整合版
品牌: 凱綸 KaiLun Energy
整合特色：
1. 完整 EMS 時序模擬 + ROI 策略矩陣
2. 多策略組合分析 (14種策略)
3. 累計現金流瀑布圖
4. 月度電費比較分析
5. 專業凱綸品牌設計
"""

import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ==================== 頁面設定 ====================
st.set_page_config(
    page_title="凱綸儲能效益評估系統",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 凱綸品牌 CSS 設計
st.markdown("""
<style>
    /* 主標題樣式 */
    .main-title {
        font-size: 2.8rem;
        font-weight: bold;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    /* 副標題樣式 */
    .sub-title {
        font-size: 1.3rem;
        color: #64748b;
        margin-bottom: 1.5rem;
    }
    
    /* 品牌標籤 */
    .brand-badge {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 0.4rem 1rem;
        border-radius: 2rem;
        font-weight: 600;
        display: inline-block;
        margin: 0.5rem 0;
        box-shadow: 0 4px 6px rgba(102, 126, 234, 0.3);
    }
    
    /* 步驟標籤 */
    .step-badge {
        background: #e0f2fe;
        color: #0369a1;
        padding: 0.3rem 0.8rem;
        border-radius: 1rem;
        font-weight: 600;
        display: inline-block;
        margin: 0.5rem 0;
    }
    
    /* 頁尾 */
    .footer {
        text-align: center;
        color: #94a3b8;
        padding: 2rem 0;
        margin-top: 3rem;
        border-top: 1px solid #e2e8f0;
    }
    
    /* 凱綸紫色強調 */
    .kailun-accent {
        color: #667eea;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 初始化 Session State ====================
if 'df_raw' not in st.session_state:
    st.session_state['df_raw'] = None
if 'show_roi_analysis' not in st.session_state:
    st.session_state['show_roi_analysis'] = False

# ==================== 資料庫定義 ====================
TARIFF_DB = {
    "高壓-三段式 (115年新制)": { 
        "cat": "高壓", 
        "logic": "3-Stage", 
        "desc": "尖峰 16-22, 適合一般工廠",
        "d_rates": {"reg_sum": 223.6, "reg_non": 166.9, "semi": 166.9, "sat_off": 44.7},
        "e_rates": {"sum_pk": 8.35, "sum_semi": 4.78, "sum_off": 1.98, "non_semi": 4.61, "non_off": 1.85}
    },
    "高壓-二段式": {
        "cat": "高壓", 
        "logic": "2-Stage", 
        "desc": "週六半日尖峰 (09-24)",
        "d_rates": {"reg_sum": 237.0, "reg_non": 173.2, "non_sum": 173.2, "sat_off": 44.7},
        "e_rates": {"sum_pk": 4.85, "sum_off": 1.92, "non_pk": 4.65, "non_off": 1.78, "sat_pk": 2.77}
    },
    "高壓-批次生產": {
        "cat": "高壓", 
        "logic": "Batch", 
        "desc": "離峰費率優惠",
        "d_rates": {"reg_sum": 223.6, "reg_non": 166.9, "non_sum": 166.9, "sat_off": 44.7},
        "e_rates": {"sum_pk": 8.50, "sum_off": 1.75, "non_pk": 8.15, "non_off": 1.65}
    },
    "特高壓-三段式 (115年新制)": {
        "cat": "特高壓", 
        "logic": "3-Stage", 
        "desc": "特高壓適用",
        "d_rates": {"reg_sum": 213.0, "reg_non": 159.0, "semi": 159.0, "sat_off": 42.9},
        "e_rates": {"sum_pk": 8.05, "sum_semi": 4.45, "sum_off": 1.85, "non_semi": 4.30, "non_off": 1.72}
    },
    "特高壓-二段式": {
        "cat": "特高壓", 
        "logic": "2-Stage", 
        "desc": "傳統尖離峰",
        "d_rates": {"reg_sum": 226.0, "reg_non": 165.0, "non_sum": 165.0, "sat_off": 42.9},
        "e_rates": {"sum_pk": 4.65, "sum_off": 1.78, "non_pk": 4.48, "non_off": 1.66, "sat_pk": 2.65}
    },
    "特高壓-批次生產": {
        "cat": "特高壓", 
        "logic": "Batch", 
        "desc": "大型產線",
        "d_rates": {"reg_sum": 213.0, "reg_non": 159.0, "non_sum": 159.0, "sat_off": 42.9},
        "e_rates": {"sum_pk": 8.15, "sum_off": 1.65, "non_pk": 7.82, "non_off": 1.56}
    }
}

DR_DB = {
    "日選 2小時": {"hours": 2, "rate": 2.47, "slots": [18, 19], "desc": "18-20"},
    "日選 4小時": {"hours": 4, "rate": 1.84, "slots": [16, 17, 18, 19], "desc": "16-20"},
    "日選 6小時": {"hours": 6, "rate": 1.69, "slots": [16, 17, 18, 19, 20, 21], "desc": "16-22"}
}

# ==================== 核心函數 ====================

def get_tou_series(row, logic):
    """時段判斷"""
    m, d, h, wd = int(row['Month']), int(row['Day']), int(row['Hour']), int(row['Weekday'])
    is_sum = (m == 5 and d >= 16) or (6 <= m <= 9) or (m == 10 and d <= 15)
    
    tou = 'Off'
    if logic == "3-Stage":
        if wd == 6:
            tou = 'Off'
        elif wd == 5:
            tou = 'Sat_Semi' if 8 <= h < 22 else 'Off'
        else:
            if is_sum:
                tou = 'Peak' if 16 <= h < 22 else ('Off' if h < 9 else 'Semi')
            else:
                tou = 'Semi' if 15 <= h < 24 else 'Off'
    elif logic == "Batch":
        if wd == 6:
            tou = 'Off'
        elif wd == 5:
            tou = 'Sat_Semi' if 9 <= h < 24 else 'Off'
        else:
            tou = 'Peak' if 16 <= h < 22 else 'Off'
    else:  # 2-Stage
        if wd == 6:
            tou = 'Off'
        elif wd == 5:
            tou = 'Sat_Semi' if 9 <= h < 24 else 'Off'
        else:
            is_pk = (9 <= h < 24) if is_sum else ((6 <= h < 11) or (14 <= h < 24))
            tou = 'Peak' if is_pk else 'Off'
    
    return pd.Series([tou, is_sum], index=['TOU', 'Is_Sum'])

def run_ems_simulation(df, pcs, kwh, eff, dod, soc_init, c_reg, use_dr, dr_opt_name, dr_target, use_peak, use_arb):
    """完整 EMS 模擬引擎"""
    bess_power = np.zeros(len(df))
    soc_state = np.zeros(len(df))
    
    current_kwh = kwh * soc_init
    min_kwh = kwh * (1 - dod)
    max_kwh = kwh
    
    dr_slots = []
    if use_dr and "日選" in dr_opt_name:
        dr_slots = DR_DB[dr_opt_name]['slots']
    
    loads = df['Load'].values
    months = df['Month'].values
    weekdays = df['Weekday'].values
    hours = df['Hour'].values
    tous = df['TOU'].values
    
    for i in range(len(df)):
        load = loads[i]
        tou = tous[i]
        
        is_dr_time = False
        if use_dr and (5 <= months[i] <= 10) and (weekdays[i] < 5) and (hours[i] in dr_slots):
            is_dr_time = True
        
        power = 0.0
        discharge_req = 0.0
        
        if is_dr_time:
            discharge_req = dr_target
        elif use_peak and (load > c_reg):
            discharge_req = load - c_reg
        elif use_arb and (tou in ['Peak', 'Semi']):
            discharge_req = pcs
        
        charge_req = 0.0
        if tou == 'Off' and not is_dr_time:
            charge_req = pcs
            headroom = c_reg - load
            if headroom < 0:
                charge_req = 0
            else:
                charge_req = min(charge_req, headroom)
        
        if discharge_req > 0:
            max_discharge_energy = max(0, current_kwh - min_kwh) * eff
            max_discharge_power = max_discharge_energy / 0.25
            real_power = min(discharge_req, pcs, max_discharge_power)
            current_kwh -= (real_power * 0.25 / eff)
            power = real_power
        elif charge_req > 0:
            max_charge_energy = max_kwh - current_kwh
            max_charge_power = max_charge_energy / 0.25
            real_power = min(charge_req, pcs, max_charge_power)
            current_kwh += (real_power * 0.25)
            power = -real_power
        
        bess_power[i] = power
        soc_state[i] = current_kwh / kwh
    
    return bess_power, soc_state

def calc_monthly_details(d_in, t_data, caps):
    """計算月度電費"""
    monthly_bills = {}
    er = t_data['e_rates']
    
    for m, g in d_in.groupby('Month'):
        is_s = g['Is_Sum'].iloc[0]
        
        # 基本電費簡化
        r_reg = t_data['d_rates'].get('reg_sum' if is_s else 'reg_non', 200.0)
        base = caps['reg'] * r_reg
        
        # 流動電費
        def get_rate(r):
            pre = "sum" if r.Is_Sum else "non"
            if r.TOU == 'Peak':
                k = f"{pre}_pk"
            elif r.TOU == 'Semi':
                k = f"{pre}_semi"
            elif r.TOU == 'Sat_Semi':
                k = "sat_pk"
            else:
                k = f"{pre}_off"
            return er.get(k, 2.0)
        
        eng = (g['Load'] * g.apply(get_rate, axis=1)).sum() / 4
        monthly_bills[m] = base + eng
    
    return monthly_bills

# ==================== 主頁面 ====================

# 凱綸品牌標題
st.markdown('<div class="main-title">⚡ 凱綸儲能效益評估系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">KaiLun Energy Solutions | 智慧儲能 × 數據決策 × 財務優化</div>', unsafe_allow_html=True)
st.markdown('<div class="brand-badge">Powered by 凱綸 KaiLun</div>', unsafe_allow_html=True)

# 頁面模式切換
page_mode = st.radio(
    "選擇分析模式",
    ["📊 完整效益分析", "🎯 策略 ROI 矩陣"],
    horizontal=True
)

# ==================== SIDEBAR ====================
st.sidebar.header("⚙️ 系統參數設定")

# A. 契約容量
st.sidebar.subheader("📊 用戶契約容量 (kW)")
with st.sidebar.expander("容量設定", expanded=True):
    c_reg = st.number_input("經常契約", value=2000, step=50)
    c_non = st.number_input("非夏月契約", value=0, step=50)
    c_semi = st.number_input("半尖峰契約", value=0, step=50)
    c_sat = st.number_input("週六半尖峰", value=0, step=50)
    c_off = st.number_input("離峰契約", value=300, step=50)
    caps = {'reg': c_reg, 'non': c_non, 'semi': c_semi, 'sat': c_sat, 'off': c_off}

# B. 電價設定
st.sidebar.subheader("💰 電價設定")
with st.sidebar.expander("費率參數", expanded=True):
    v_level = st.radio("電壓等級", ["高壓", "特高壓"], horizontal=True)
    opts = [k for k, v in TARIFF_DB.items() if v['cat'] == v_level]
    t_name = st.selectbox("電價方案", opts)
    t_data = TARIFF_DB[t_name]
    st.caption(f"💡 {t_data['desc']}")
    
    # 顯示詳細費率資訊
    st.markdown("---")
    st.markdown("**📊 基本電費 (元/kW/月)**")
    d_rates = t_data['d_rates']
    st.write(f"• 夏月經常: {d_rates.get('reg_sum', 0):.1f}")
    st.write(f"• 非夏月經常: {d_rates.get('reg_non', 0):.1f}")
    if 'semi' in d_rates:
        st.write(f"• 半尖峰: {d_rates.get('semi', 0):.1f}")
    if 'sat_off' in d_rates:
        st.write(f"• 週六/離峰優惠: {d_rates.get('sat_off', 0):.1f}")
    
    st.markdown("**⚡ 流動電費 (元/kWh)**")
    e_rates = t_data['e_rates']
    
    # 夏月費率
    st.write("**夏月 (5/16-10/15)**")
    if 'sum_pk' in e_rates:
        st.write(f"• 尖峰: {e_rates.get('sum_pk', 0):.2f}")
    if 'sum_semi' in e_rates:
        st.write(f"• 半尖峰: {e_rates.get('sum_semi', 0):.2f}")
    if 'sum_off' in e_rates:
        st.write(f"• 離峰: {e_rates.get('sum_off', 0):.2f}")
    
    # 非夏月費率
    st.write("**非夏月 (10/16-5/15)**")
    if 'non_pk' in e_rates:
        st.write(f"• 尖峰: {e_rates.get('non_pk', 0):.2f}")
    if 'non_semi' in e_rates:
        st.write(f"• 半尖峰: {e_rates.get('non_semi', 0):.2f}")
    if 'non_off' in e_rates:
        st.write(f"• 離峰: {e_rates.get('non_off', 0):.2f}")
    
    # 週六費率（如果有）
    if 'sat_pk' in e_rates:
        st.write("**週六**")
        st.write(f"• 半尖峰: {e_rates.get('sat_pk', 0):.2f}")

# C. 儲能系統
st.sidebar.subheader("🔋 儲能系統")
with st.sidebar.expander("規格參數", expanded=True):
    eq_mode = st.radio("設備模式", ["自訂規格", "固定規格 (125kW/261kWh)"])
    
    if eq_mode == "自訂規格":
        pcs = st.number_input("PCS (kW)", value=250.0, step=10.0)
        kwh = st.number_input("容量 (kWh)", value=750.0, step=10.0)
    else:
        eq_count = st.number_input("機台數量", value=2, step=1)
        pcs = 125.0 * eq_count
        kwh = 261.0 * eq_count
        st.caption(f"📦 總規格: {pcs:.0f} kW / {kwh:.0f} kWh")
    
    col_s1, col_s2, col_s3 = st.columns(3)
    eff = col_s1.slider("效率 (%)", 80, 100, 90) / 100.0
    soc_init = col_s2.slider("初始 SOC", 0, 100, 90) / 100.0
    dod = col_s3.slider("DOD (%)", 50, 100, 90) / 100.0

# D. 節電方案
st.sidebar.subheader("⚡ 節電方案")
use_arb = st.sidebar.checkbox("1. 價差套利", True)
use_peak = st.sidebar.checkbox("2. 契約優化", True)
use_dr = st.sidebar.checkbox("3. 需量反應 (DR)", False)

dr_target_kw = 0
dr_opt_name = list(DR_DB.keys())[0]
if use_dr:
    dr_opt_name = st.sidebar.selectbox("DR方案", list(DR_DB.keys()))
    dr_info = DR_DB[dr_opt_name]
    st.sidebar.caption(f"⏰ {dr_info['desc']} | {dr_info['rate']} 元/kWh")
    dr_target_kw = st.sidebar.number_input("抑低容量 (kW)", value=pcs, step=10.0)

# E. ETP
st.sidebar.subheader("🌐 電力交易 (ETP)")
use_etp = st.sidebar.checkbox("啟用 ETP", False)
etp_rev_total = 0

if use_etp:
    etp_type = st.sidebar.radio("商品", ["即時備轉", "補充備轉"])
    col_e1, col_e2 = st.sidebar.columns(2)
    etp_cap = col_e1.number_input("容量 (MW)", value=pcs/1000, step=0.1)
    etp_hours = col_e2.number_input("時數 (hr)", value=8000, step=100)
    etp_price = col_e1.number_input("結清價", value=182.0 if etp_type=="即時備轉" else 51.0)
    etp_perf_price = col_e2.number_input("效能費", value=100.0 if etp_type=="即時備轉" else 0.0)
    etp_rate = st.sidebar.slider("執行率 (%)", 0, 100, 100) / 100.0
    etp_rev_total = etp_cap * etp_rate * (etp_price + etp_perf_price) * etp_hours
    st.sidebar.success(f"💰 ${etp_rev_total:,.0f}")

# F. 財務
st.sidebar.subheader("📈 財務參數")
with st.sidebar.expander("成本設定", expanded=True):
    project_years = st.selectbox("年限 (年)", [10, 12, 15], index=2)
    capex_unit = st.number_input("Capex (元/kWh)", value=9000, step=500)
    total_capex = capex_unit * kwh
    opex_rate = st.number_input("維運費率 (%)", value=2.5, step=0.1)
    opex_yearly = total_capex * (opex_rate / 100)
    st.info(f"💰 Capex: ${total_capex:,.0f}\n\n🛠️ Opex: ${opex_yearly:,.0f}")

soh_list = [max(0.7, 0.95 - (0.025 * i)) for i in range(project_years)]

# ==================== 資料載入 ====================
st.markdown("---")
st.subheader("📁 負載資料")

col1, col2 = st.columns([1, 3])
mode = col1.radio("來源", ["智慧模擬", "上傳檔案"])

if mode == "智慧模擬":
    with col2:
        c1, c2 = st.columns(2)
        s_peak = c1.number_input("峰值 (kW)", value=c_reg * 1.1)
        s_base = c2.number_input("基載 (kW)", value=c_reg * 0.3)
        ind_type = st.selectbox("產業", ["傳統工廠 (08-18)", "24h 科技廠", "商業辦公"])
        
        if st.button("🚀 生成資料", type="primary"):
            with st.spinner("生成中..."):
                idx = pd.date_range('2026-01-01', '2026-12-31 23:45', freq='15min')
                sim = pd.DataFrame(index=idx)
                sim['Hour'] = sim.index.hour
                sim['Weekday'] = sim.index.weekday
                sim['Month'] = sim.index.month
                sim['Day'] = sim.index.day
                
                vals = []
                for i in range(len(sim)):
                    h, wd, m = sim['Hour'].iloc[i], sim['Weekday'].iloc[i], sim['Month'].iloc[i]
                    v = s_base
                    if "傳統" in ind_type:
                        v = s_peak if (8 <= h < 18 and wd < 5) else s_base
                    elif "24h" in ind_type:
                        v = s_peak * 0.95 if (8 <= h < 20 and wd < 5) else s_peak * 0.9
                    else:
                        v = s_peak if (9 <= h < 18 and wd < 5) else s_base * 0.4
                    season = 1.2 if 6 <= m <= 9 else 1.0
                    vals.append(max(0, v * season + np.random.normal(0, v * 0.05)))
                
                sim['Load'] = vals
                st.session_state['df_raw'] = sim
                st.success("✅ 完成！")
                st.rerun()
else:
    with col2:
        f = st.file_uploader("📁 上傳 .xlsx 或 .csv", type=['xlsx', 'csv'])
        if f:
            try:
                d = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
                d.iloc[:, 0] = pd.to_datetime(d.iloc[:, 0])
                d.columns = ['DateTime', 'Load']
                df = d.set_index('DateTime').sort_index().resample('15min').mean().interpolate()
                df['Hour'] = df.index.hour
                df['Weekday'] = df.index.weekday
                df['Month'] = df.index.month
                df['Day'] = df.index.day
                st.session_state['df_raw'] = df
                st.success(f"✅ {len(df):,} 筆")
            except Exception as e:
                st.error(f"❌ {e}")

# ==================== 主要分析 ====================
if st.session_state['df_raw'] is not None:
    df = st.session_state['df_raw'].copy()
    df[['TOU', 'Is_Sum']] = df.apply(lambda x: get_tou_series(x, t_data['logic']), axis=1)
    
    # 設定電價
    def get_price(row):
        er = t_data['e_rates']
        pre = "sum" if row['Is_Sum'] else "non"
        if row['TOU'] == 'Peak':
            return er.get(f"{pre}_pk", 5.0)
        elif row['TOU'] == 'Semi':
            return er.get(f"{pre}_semi", 3.0)
        else:
            return er.get(f"{pre}_off", 1.5)
    
    df['Price'] = df.apply(get_price, axis=1)
    
    # ========== 模式1: 完整效益分析 ==========
    if page_mode == "📊 完整效益分析":
        st.markdown("---")
        st.markdown('<div class="step-badge">完整效益分析模式</div>', unsafe_allow_html=True)
        
        # EMS 模擬
        with st.spinner("⚡ EMS 模擬中..."):
            bess_p, soc = run_ems_simulation(
                df, pcs, kwh, eff, dod, soc_init, 
                c_reg, use_dr, dr_opt_name, dr_target_kw, 
                use_peak, use_arb
            )
            df['BESS'] = bess_p
            df['SOC'] = soc
            df['Net_Load'] = (df['Load'] - df['BESS']).clip(lower=0)
        
        # 負載曲線
        st.subheader("📊 負載曲線分析")
        t1, t2 = st.tabs(["☀️ 夏月", "🌙 非夏月"])
        
        def plot_curve(is_sum):
            dd = df[(df['Is_Sum'] == is_sum) & (df['Weekday'] < 5)]
            if len(dd) == 0:
                st.warning("⚠️ 無數據")
                return
            
            avg = dd.groupby('Hour')[['Load', 'Net_Load', 'SOC']].mean().reset_index()
            
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            
            fig.add_trace(go.Scatter(x=avg['Hour'], y=[c_reg]*24, name='契約', 
                                    line=dict(color='red', dash='dash', width=2)), secondary_y=False)
            fig.add_trace(go.Scatter(x=avg['Hour'], y=avg['Load'], name='原負載',
                                    line=dict(color='#FF6B6B', width=3), fill='tozeroy'), secondary_y=False)
            fig.add_trace(go.Scatter(x=avg['Hour'], y=avg['Net_Load'], name='削峰後',
                                    line=dict(color='#4ECDC4', width=3)), secondary_y=False)
            fig.add_trace(go.Scatter(x=avg['Hour'], y=avg['SOC']*100, name='SOC',
                                    line=dict(color='orange', width=2)), secondary_y=True)
            
            fig.update_layout(height=500, hovermode='x unified', template="plotly_white",
                            xaxis=dict(title="時間 (小時)", tickmode='linear', dtick=2))
            fig.update_yaxes(title_text="負載 (kW)", secondary_y=False)
            fig.update_yaxes(title_text="SOC (%)", secondary_y=True, range=[0, 105])
            
            st.plotly_chart(fig, use_container_width=True)
        
        with t1:
            plot_curve(True)
        with t2:
            plot_curve(False)
        
        # 財務分析
        st.markdown("---")
        st.subheader("💰 財務效益")
        
        with st.spinner("💡 計算中..."):
            df_post = df.copy()
            df_post['Load'] = df['Net_Load']
            
            mon_pre = calc_monthly_details(df, t_data, caps)
            mon_post = calc_monthly_details(df_post, t_data, caps)
            
            bill_pre = sum(mon_pre.values())
            bill_post = sum(mon_post.values())
            save_bill = bill_pre - bill_post
            rate_save = (save_bill / bill_pre * 100) if bill_pre > 0 else 0
        
        # DR 收益
        inc_dr = 0
        if use_dr and "日選" in dr_opt_name:
            dr_info = DR_DB[dr_opt_name]
            mask = ((df['Month']>=5)&(df['Month']<=10)&(df['Weekday']<5)&(df['Hour'].isin(dr_info['slots'])))
            if mask.any():
                days = df[mask].index.normalize().nunique()
                avg_drop = min(pcs, max(0, df.loc[mask,'Load'].mean() - df.loc[mask,'Net_Load'].mean()))
                perf = min(avg_drop/dr_target_kw, 1.2) if dr_target_kw > 0 else 0
                ratio = 1.2 if perf>=0.95 else (1.0 if perf>=0.8 else (0.8 if perf>=0.6 else 0))
                inc_dr = dr_target_kw * perf * dr_info['hours'] * dr_info['rate'] * ratio * days
                st.info(f"📊 DR: {days}天 | {avg_drop:.0f}kW | {perf*100:.0f}%")
        
        total_benefit = save_bill + inc_dr + etp_rev_total
        
        col1, col2, col3 = st.columns(3)
        col1.metric("原電費", f"${bill_pre:,.0f}")
        col2.metric("新電費", f"${bill_post:,.0f}", delta=f"-${save_bill:,.0f}")
        col3.metric("節省率", f"{rate_save:.1f}%")
        
        st.divider()
        
        # 現金流
        cf = [-total_capex]
        for i in range(project_years):
            cf.append(total_benefit * soh_list[i] - opex_yearly)
        
        npv_val = npf.npv(0.05, cf)
        try:
            irr_val = npf.irr(cf) * 100
        except:
            irr_val = None
        
        cumsum = np.cumsum(cf)
        pb = next((i for i, v in enumerate(cumsum) if v >= 0), None)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("建置", f"${total_capex/10000:.0f}萬")
        c2.metric("年效益", f"${total_benefit:,.0f}")
        c3.metric("NPV", f"${npv_val:,.0f}")
        c4.metric("回收", f"{pb}年" if pb else "未回本")
        
        # 現金流圖
        st.write("#### 📊 現金流")
        
        fig_cf = go.Figure()
        fig_cf.add_trace(go.Bar(x=list(range(1,project_years+1)), y=cf[1:], 
                                name="年淨現金流", marker_color='#636EFA'))
        fig_cf.add_trace(go.Scatter(x=list(range(1,project_years+1)), y=cumsum[1:], 
                                    name="累計", mode='lines+markers', 
                                    line=dict(color='#EF553B', width=3)))
        fig_cf.add_hline(y=0, line_dash="dash", line_color="gray")
        fig_cf.update_layout(height=450, template="plotly_white", hovermode='x unified')
        st.plotly_chart(fig_cf, use_container_width=True)
        
        # 月度比較
        st.write("#### 📉 月度電費")
        
        months = list(range(1, 13))
        y_pre = [mon_pre.get(m, 0) for m in months]
        y_post = [mon_post.get(m, 0) for m in months]
        
        fig_bill = go.Figure()
        fig_bill.add_trace(go.Bar(y=months, x=y_pre, name='前', orientation='h', marker_color='#FF6666'))
        fig_bill.add_trace(go.Bar(y=months, x=y_post, name='後', orientation='h', marker_color='#66CC99'))
        fig_bill.update_layout(barmode='group', height=500, template="plotly_white")
        st.plotly_chart(fig_bill, use_container_width=True)
    
    # ========== 模式2: 策略 ROI 矩陣 ==========
    else:
        st.markdown("---")
        st.markdown('<div class="step-badge">策略 ROI 矩陣分析</div>', unsafe_allow_html=True)
        
        st.info("🔄 計算 14 種策略組合...")
        
        # 策略定義
        strategies = [
            ("電價套利-日選3h+即時", "日選 2小時"),
            ("電價套利-日選2h+即時", "日選 2小時"),
            ("電價套利-日選4h+即時", "日選 4小時"),
            ("電價套利-日選6h", "日選 6小時"),
            ("電價套利-日選4h", "日選 4小時"),
            ("電價套利-日選2h", "日選 2小時"),
            ("電價套利+即時", None),
            ("電價套利", None),
            ("削峰-日選6h", "日選 6小時"),
            ("削峰-日選4h", "日選 4小時"),
            ("削峰-日選2h", "日選 2小時"),
            ("削峰", None),
            ("義務時數型", None),
            ("需量競價", None),
        ]
        
        discount_rates = [0, 0.01, 0.02, 0.03, 0.04, 0.05]
        roi_matrix = []
        
        for strategy_name, dr_type in strategies:
            row_data = {'策略': strategy_name}
            
            # 簡化模擬
            dr_config = DR_DB.get(dr_type) if dr_type else None
            use_dr_temp = dr_type is not None
            use_arb_temp = "電價套利" in strategy_name
            use_peak_temp = "削峰" in strategy_name
            
            bess = np.zeros(len(df))
            for i in range(len(df)):
                if use_arb_temp and df['TOU'].iloc[i] == 'Peak':
                    bess[i] = pcs * 0.7
                if use_peak_temp and df['Load'].iloc[i] > c_reg:
                    bess[i] = min(pcs, df['Load'].iloc[i] - c_reg)
            
            df_s = df.copy()
            df_s['Load'] = (df['Load'] - bess).clip(lower=0)
            
            annual = (df['Load'] * df['Price']).sum()/4 - (df_s['Load'] * df['Price']).sum()/4
            
            for rate in discount_rates:
                cf = [-total_capex] + [annual - opex_yearly] * project_years
                npv = npf.npv(rate, cf)
                roi = (npv / total_capex * 100) if total_capex > 0 else 0
                row_data[f"{int(rate*100)}%"] = roi
            
            roi_matrix.append(row_data)
        
        roi_df = pd.DataFrame(roi_matrix)
        
        # ROI 熱力圖
        st.subheader("📊 策略 ROI 熱力圖")
        
        heatmap_data = roi_df.set_index('策略')
        
        fig_hm = go.Figure(data=go.Heatmap(
            z=heatmap_data.values,
            x=heatmap_data.columns,
            y=heatmap_data.index,
            colorscale=[
                [0, '#fef08a'], [0.3, '#bef264'], [0.5, '#86efac'], 
                [0.7, '#4ade80'], [1, '#22c55e']
            ],
            text=heatmap_data.values,
            texttemplate='%{text:.1f}%',
            textfont={"size": 10},
            colorbar=dict(title="ROI (%)")
        ))
        
        fig_hm.update_layout(
            title="策略 ROI 矩陣 (折現率 vs 策略)",
            xaxis_title="折現率",
            yaxis_title="策略",
            height=700,
            font=dict(size=11)
        )
        
        st.plotly_chart(fig_hm, use_container_width=True)
        
        # 最佳策略
        st.markdown("---")
        st.subheader("🎯 最佳策略")
        
        best_idx = roi_df['5%'].idxmax()
        best_strategy = roi_df.loc[best_idx, '策略']
        best_roi = roi_df.loc[best_idx, '5%']
        
        col1, col2 = st.columns(2)
        col1.metric("🏆 推薦", best_strategy)
        col2.metric("📈 ROI (5%)", f"{best_roi:.2f}%")
        
        # 匯出
        if st.button("📥 匯出 Excel"):
            csv = roi_df.to_csv(index=False)
            st.download_button("下載", csv, "roi_analysis.csv", "text/csv")

else:
    st.info("👈 請先生成/上傳資料")

# 頁尾
st.markdown("---")
st.markdown("""
<div class="footer">
    <p><span class="kailun-accent">⚡ 凱綸儲能效益評估系統</span> | KaiLun Energy Solutions © 2026</p>
    <p style="font-size: 0.9rem; margin-top: 0.5rem;">智慧儲能 × 數據決策 × 永續未來</p>
</div>
""", unsafe_allow_html=True)
