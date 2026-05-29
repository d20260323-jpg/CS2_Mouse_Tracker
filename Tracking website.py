import streamlit as st
import pandas as pd
import os
import plotly.express as px
import base64
from PIL import Image
#python -m streamlit run "C:\Users\donny.d.huang\PycharmProjects\PythonProject\AI\cs2_mouse_site\Tracking website.py"
# ==========================================
# 1. 页面配置与高清电竞风 CSS
# ==========================================
st.set_page_config(
    page_title="ZOWIE HUB | CS2 Pro Mouse Tracker",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 图片路径
ZOWIE_LOGO_PATH = r"C:\Users\donny.d.huang\PycharmProjects\PythonProject\AI\cs2_mouse_site\assets\zowie_logo.png"
HERO_MOUSE_PATH = r"C:\Users\donny.d.huang\PycharmProjects\PythonProject\AI\cs2_mouse_site\assets\hero_mouse.png"
EXCEL_PATH = r"C:\Users\donny.d.huang\.openclaw\workspace\memory\cs2_mouse_tracking.xlsx"

# 自定义 CSS 样式
st.markdown("""
<style>
    /* 全局背景 */
    .stApp {
        background-color: #050505; 
        color: #E0E0E0;
        font-family: 'Segoe UI', sans-serif;
    }

    /* 顶部 Hero 区域 */
    .hero-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: linear-gradient(135deg, #101010 0%, #1a1a1a 100%);
        padding: 30px;
        border-radius: 12px;
        border: 1px solid #222;
        margin-bottom: 35px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.5);
    }

    /* 标题样式 */
    h1 {
        font-size: 50px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        line-height: 1.1;
    }

    h2 {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border-left: 4px solid #E02020;
        padding-left: 15px;
        margin: 30px 0 20px 0 !important;
    }

    /* 指标卡片 */
    .metric-box {
        background-color: #151515;
        border: 1px solid #252525;
        border-radius: 8px;
        padding: 20px;
        text-align: center;
        transition: all 0.3s ease;
    }
    .metric-box:hover {
        border-color: #E02020;
        transform: translateY(-5px);
        box-shadow: 0 5px 15px rgba(224, 32, 32, 0.2);
    }

    /* 变动记录条目 */
    .change-log {
        background-color: #121212;
        border-left: 4px solid #333;
        padding: 12px 20px;
        margin-bottom: 10px;
        border-radius: 0 8px 8px 0;
    }
    .change-log:hover {
        border-left-color: #E02020;
        background-color: #1a1515;
    }

    /* 鼠标图片发光效果 */
    .hero-img {
        filter: drop-shadow(0 0 20px rgba(224, 32, 32, 0.4));
        max-width: 100%;
        height: auto;
    }
</style>
""", unsafe_allow_html=True)


# 辅助函数：将本地图片转为 Base64 以便在 HTML 中精准布局
def get_base64_image(path):
    if os.path.exists(path):
        with open(path, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode()
    return None


# ==========================================
# 2. 数据处理与缓存 (仅统计最后 23 行)
# ==========================================
@st.cache_data(ttl=300)
def load_data():
    if not os.path.exists(EXCEL_PATH):
        st.error(f"❌ 找不到数据文件: {EXCEL_PATH}")
        return None, None
    try:
        df_all = pd.read_excel(EXCEL_PATH)
        # 处理日期格式，兼容带秒数的情况
        df_all['QueryTime'] = pd.to_datetime(df_all['QueryTime'], errors='coerce', format='mixed')

        # 核心逻辑：截取最后 23 行作为当前快照
        df_latest = df_all.tail(1000).copy()

        # 提取品牌信息
        df_latest['Brand'] = df_latest['Mouse'].apply(lambda x: str(x).split()[0] if pd.notnull(x) else "Unknown")

        # 预处理 Changed 状态
        df_all['Changed'] = df_all['Changed'].astype(str).str.upper().str.strip()
        df_all['Brand'] = df_all['Mouse'].apply(lambda x: str(x).split()[0] if pd.notnull(x) else "Unknown")

        return df_all, df_latest
    except Exception as e:
        st.error(f"数据解析失败: {e}")
        return None, None


# ==========================================
# 3. 网页主结构
# ==========================================
def main():
    # --- A. 导航栏 (Logo) ---
    logo_b64 = get_base64_image(ZOWIE_LOGO_PATH)
    if logo_b64:
        st.markdown(f'<img src="data:image/png;base64,{logo_b64}" width="180">', unsafe_allow_html=True)
    else:
        st.markdown('<h2 style="color:#E02020; margin:0; border:none;">ZOWIE</h2>', unsafe_allow_html=True)

    # --- B. Hero 视觉区 ---
    st.markdown('<div class="hero-container">', unsafe_allow_html=True)
    h_col1, h_col2 = st.columns([1.5, 1])

    with h_col1:
        st.markdown("""
            <div style="padding-top:10px;">
                <p style="color:#E02020; font-weight:bold; letter-spacing:3px; margin-bottom:5px;">PROFESSIONAL CHOICE</p>
                <h1>CS2 PROS<br><span style="color:#E02020;">GEAR</span> TRACKER</h1>
                <p style="color:#888; font-size:18px; margin-top:15px; max-width:450px;">
                    实时同步全球顶尖选手的鼠标选择。基于最新 175 条核心样本数据分析。
                </p>
            </div>
        """, unsafe_allow_html=True)

    with h_col2:
        hero_b64 = get_base64_image(HERO_MOUSE_PATH)
        if hero_b64:
            st.markdown(
                f'<div style="text-align:right;"><img src="data:image/png;base64,{hero_b64}" class="hero-img"></div>',
                unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # 加载数据
    df_all, df_latest = load_data()
    if df_all is None: return

    # --- C. 关键指标 (修正了 := 语法错误) ---
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            f'<div class="metric-box"><small>TRACKED PLAYERS</small><br><span style="font-size:35px; color:#E02020; font-weight:bold;">{len(df_latest)}</span></div>',
            unsafe_allow_html=True)
    with c2:
        st.markdown(
            f'<div class="metric-card"><small>ACTIVE BRANDS</small><br><span style="font-size:35px; color:#E02020; font-weight:bold;">{df_latest["Brand"].nunique()}</span></div>',
            unsafe_allow_html=True)
    with c3:
        total_changes = len(df_all[df_all['Changed'] == 'YES'])
        st.markdown(
            f'<div class="metric-box"><small>TOTAL CHANGES</small><br><span style="font-size:35px; color:#E02020; font-weight:bold;">{total_changes}</span></div>',
            unsafe_allow_html=True)

    # --- D. 图表展示 ---
    st.write("")
    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.markdown("<h2>🔥 热门型号 (当前)</h2>", unsafe_allow_html=True)
        top_data = df_all.sort_values('QueryTime').groupby('Player').last().reset_index()['Mouse'].value_counts().head(5).reset_index()
        top_data.columns = ['Mouse', 'count']
        fig = px.bar(top_data, y='Mouse', x='count', orientation='h', text='count', color_discrete_sequence=['#E02020'])
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#E0E0E0",
                          xaxis=dict(visible=False), yaxis=dict(title="", autorange="reversed"), height=300,
                          margin=dict(l=0, r=0, t=0, b=0))
        st.plotly_chart(fig, use_container_width=True)

    with chart_col2:
        st.markdown("<h2>🏢 品牌占比</h2>", unsafe_allow_html=True)
        brand_data = df_latest['Brand'].value_counts().reset_index()
        brand_data.columns = ['Brand', 'count']
        fig_p = px.pie(brand_data, values='count', names='Brand', hole=0.5,
                       color_discrete_sequence=px.colors.sequential.Reds_r)
        fig_p.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#E0E0E0",
                            height=300, margin=dict(l=0, r=0, t=0, b=0))
        st.plotly_chart(fig_p, use_container_width=True)

    # --- D2. 品牌趋势折线图 ---
    st.markdown("<h2>📈 品牌使用趋势</h2>", unsafe_allow_html=True)

    # 处理数据：按时间 + 品牌统计每个时间点的使用数量
    df_trend = df_all.copy()
    df_trend['Date'] = df_trend['QueryTime'].dt.date  # 按天聚合
    brand_trend = df_trend.groupby(['Date', 'Brand'])['Player'].count().reset_index()
    brand_trend.columns = ['Date', 'Brand', 'Count']

    # 只展示 Top 5 品牌，避免图表太乱
    top_brands = df_trend['Brand'].value_counts().head(5).index.tolist()
    brand_trend = brand_trend[brand_trend['Brand'].isin(top_brands)]

    fig_line = px.line(
        brand_trend,
        x='Date',
        y='Count',
        color='Brand',
        markers=True,
        color_discrete_sequence=px.colors.sequential.Reds_r
    )
    fig_line.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font_color="#FFFFFF",
        height=350,
        margin=dict(l=0, r=0, t=0, b=0),
        xaxis=dict(gridcolor='#222', tickfont=dict(color='#FFFFFF')),
        yaxis=dict(gridcolor='#222', title=dict(text="使用人数",font = dict(color = '#FFFFFF' )), tickfont=dict(color='#FFFFFF')),
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#FFFFFF'))
    )
    st.plotly_chart(fig_line, use_container_width=True)

    # --- E. 变动快讯 ---
    st.markdown("<h2>🔄 实时变动动态</h2>", unsafe_allow_html=True)
    changed_list = df_all[df_all['Changed'] == 'YES'].sort_values('QueryTime', ascending=False).head(6)

    if not changed_list.empty:
        for _, row in changed_list.iterrows():
            # 兼容处理时间显示
            time_str = row['QueryTime'].strftime('%Y-%m-%d %H:%M') if pd.notnull(row['QueryTime']) else "N/A"
            st.markdown(f"""
                <div class="change-log">
                    <span style="color:#666; font-size:12px;">{time_str}</span><br>
                    选手 <b>{row['Player']}</b> 切换至 <span style="color:#E02020;">{row['Mouse']}</span>
                </div>
            """, unsafe_allow_html=True)
    else:
        st.info("目前监测中... 暂无近期设备变更记录。")


if __name__ == "__main__":
    main()