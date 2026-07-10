import streamlit as st
import pandas as pd
import os
import plotly.express as px
import base64
from PIL import Image
#python -m streamlit run "C:\Users\donny.d.huang\Desktop\Tracking website.py"
# ==========================================
# 1. 页面配置与高清电竞风 CSS
# ==========================================
st.set_page_config(
    page_title="ZOWIE HUB | CS2 Pro Mouse Tracker",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 图片路径
EXCEL_PATH = "https://raw.githubusercontent.com/d20260323-jpg/CS2_Mouse_Tracker/main/cs2_mouse_tracking.xlsx" #0708更改：路径换成github上数据库文件地址
ZOWIE_LOGO_PATH = "assets/zowie_logo.png"
HERO_MOUSE_PATH = "assets/hero_mouse.png"

# 品牌固定配色（同一品牌永远同一颜色，不随名次变化）
BRAND_COLORS = {
    'Logitech': '#8B0000',   # 深红
    'ZOWIE':    '#E02020',   # 正红
    'Razer':    '#00B14F',   # 雷蛇绿（雷蛇品牌色）
    'Pulsar':   '#FF6B35',   # 橙
    'VAXEE':    '#F4A261',   # 浅橙
    'Lamzu':    '#9B59B6',   # 紫
    'Finalmouse':'#3498DB',  # 蓝
    '其他':     '#555555',   # 灰
}

# ─── 明星选手名单（CS2 Major 冠军 / 2026 顶级选手，手动维护）───
# ⚠️ 名字拼写务必和你数据库 Player 字段一致；is_star 已忽略大小写
STAR_PLAYERS = {
    # —— NAVI（PGL Copenhagen 2024 冠军，CS2 首冠）——
    'Aleksib',
    'b1t',
    'jL',          # Copenhagen MVP
    'iM',
    'w0nderful',

    # —— Team Spirit（PW Shanghai 2024 冠军）——
    'donk',        # 史上最年轻 Major MVP
    'sh1ro',
    'chopper',
    'magixx',
    'zont1x',

    # —— Team Vitality（Austin 2025 + Budapest 2025 双冠）——
    'ZywOo',       # 四届年度最佳
    'apEX',
    'ropz',
    'flameZ',
    'mezii',

    # —— Team Falcons（IEM Cologne 2026 冠军，最新）——
    'NiKo',
    'm0NESY',
    'TeSeS',
    'kyxsan',
    'karrigan',

    # —— 其他顶级 / 名宿 ——
    's1mple',      # CS:GO GOAT
    'device',
    'FalleN',
    'stavn',
    'broky',
    'Twistzz',
    'frozen',
    'KSCERATO',
    'yuurih',
    'Spinx',
    'XANTARES',

# —— 人气 / 知名选手（补充）——
    'Twistzz',     # (已有)
    'rain',        # FaZe 老将，人气高
    'karrigan',    # (已有) 传奇指挥
    'Snax',        # 波兰传奇
    'Boombl4',     # 前 NAVI 队长，话题度高
    'sh1ro',       # (已有)
    'Ax1Le',       # Cloud9/前 Gambit 明星
    'niko',        # 注意：有个小写 niko（不是大写NiKo那个），是另一个选手，别搞混
    'kennyS',      # 法国 AWP 传奇，人气极高
    'coldzera',    # 巴西传奇，两届 Major MVP
    'olofmeister', # fnatic 传奇
    'dupreeh',     # 五冠王，Major 冠军最多
    'gla1ve',      # Astralis 指挥
    'Magisk',      # Astralis
}

def is_star(player):
    if not player:
        return False
    return str(player).strip().lower() in {s.lower() for s in STAR_PLAYERS}

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
def load_data():                     #0708:有做调整不会出现Github上运行报错的情况
    try:
        df_all = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        st.error(f"❌ 无法读取数据文件: {EXCEL_PATH}\n错误详情: {e}")
        return None, None

    try:
        # 处理日期格式，兼容带秒数的情况
        df_all['QueryTime'] = pd.to_datetime(df_all['QueryTime'], errors='coerce', format='mixed')

        # 核心逻辑：截取最后 23 行作为当前快照
        df_latest = df_all.sort_values('QueryTime').drop_duplicates('Player', keep='last').copy()

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

    # --- D. 图表展示（热门型号 + 品牌占比，共享时间滑块联动）---
    st.write("")

    # 共享时间滑块（横跨整宽，控制下面两个图）
    df_d = df_all.copy().sort_values('QueryTime')
    month_opts = pd.date_range(start=df_d['QueryTime'].min(), end=df_d['QueryTime'].max(), freq='ME')
    sel_month = st.select_slider(
        "📅 选择时间点",
        options=list(month_opts),
        value=month_opts[-1],
        format_func=lambda d: d.strftime('%Y-%m'),
        key='shared_month_slider'
    )

    # 用共享时间点重建那个月的快照（两个图共用同一份）
    snap_d = df_d[df_d['QueryTime'] <= sel_month].drop_duplicates('Player', keep='last')
    snap_d = snap_d[snap_d['QueryTime'] >= sel_month - pd.Timedelta(days=350)]

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.markdown(f"<h2>🔥 热门型号（{sel_month.strftime('%Y-%m')}）</h2>", unsafe_allow_html=True)
        top_data = snap_d['Mouse'].value_counts().head(5).reset_index()
        top_data.columns = ['Mouse', 'count']
        fig = px.bar(top_data, y='Mouse', x='count', orientation='h', text='count',
                     color_discrete_sequence=['#E02020'])
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                          font_color="#E0E0E0", xaxis=dict(visible=False),
                          yaxis=dict(title="", autorange="reversed"), height=300,
                          margin=dict(l=0, r=0, t=0, b=0))
        st.plotly_chart(fig, use_container_width=True)

    with chart_col2:

        st.markdown(f"<h2>🏢 品牌占比（{sel_month.strftime('%Y-%m')}）</h2>", unsafe_allow_html=True)
        brand_data = snap_d['Brand'].value_counts().reset_index()
        brand_data.columns = ['Brand', 'count']

        # 把占比 <3% 的小品牌合并成「其他」
        total = brand_data['count'].sum()
        brand_data['pct'] = brand_data['count'] / total
        big = brand_data[brand_data['pct'] >= 0.03].copy()
        small_sum = brand_data[brand_data['pct'] < 0.03]['count'].sum()
        if small_sum > 0:
            big = pd.concat([big[['Brand', 'count']],
                             pd.DataFrame([{'Brand': '其他', 'count': small_sum}])],
                            ignore_index=True)
        brand_data = big[['Brand', 'count']]

        # 按固定顺序排列（跟 BRAND_COLORS 的顺序一致，不随名次变）
        brand_order = list(BRAND_COLORS.keys())
        brand_data['_order'] = brand_data['Brand'].apply(
            lambda b: brand_order.index(b) if b in brand_order else 999
        )
        brand_data = brand_data.sort_values('_order').drop(columns='_order')

        fig_p = px.pie(brand_data, values='count', names='Brand', hole=0.5,
                       color='Brand',
                       color_discrete_map=BRAND_COLORS,
                       category_orders={'Brand': brand_order})
        fig_p.update_traces(sort=False)  # ← 关键：禁止plotly自己按大小重排扇形
        fig_p.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                            font_color="#E0E0E0", height=300, margin=dict(l=0, r=0, t=0, b=0))
        st.plotly_chart(fig_p, use_container_width=True)

    # --- D2. 品牌趋势折线图 ---
    st.markdown("<h2>📈 品牌使用趋势</h2>", unsafe_allow_html=True)

    # 处理数据：按时间 + 品牌统计每个时间点的使用数量
    # 处理数据：按月重建"存量"——每月底每个品牌有多少人在用（forward-fill）
    df_trend = df_all.copy()
    df_trend = df_trend.sort_values('QueryTime')

    # 生成每个月的月末时间点
    start = df_trend['QueryTime'].min()
    end = df_trend['QueryTime'].max()
    months = pd.date_range(start=start, end=end, freq='ME')

    records = []
    active_days = 350  # 只算最近 350 天还有记录的选手（和飞书口径一致）
    for m in months:
        # 截止到这个月底，每个选手最后一次的记录 = 他此刻在用的
        snap = df_trend[df_trend['QueryTime'] <= m].drop_duplicates('Player', keep='last')
        # 活跃过滤：只留最后一次记录在近 active_days 天内的选手
        snap = snap[snap['QueryTime'] >= m - pd.Timedelta(days=active_days)]
        # 数每个品牌多少人
        brand_counts = snap['Brand'].value_counts()
        for brand, cnt in brand_counts.items():
            records.append({'Date': m, 'Brand': brand, 'Count': cnt})

    brand_trend = pd.DataFrame(records)

    # 只展示 Top 5 品牌
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

    # --- D3. 鼠标型号趋势图 ---
    st.markdown("<h2>🔍 鼠标型号趋势</h2>", unsafe_allow_html=True)

    all_mice = sorted(df_all['Mouse'].dropna().unique().tolist())

    # 第一级：搜索框，缩小范围
    search = st.text_input("🔍 先搜索关键词缩小范围（如 superlight / zowie / razer，留空显示全部）", value="")

    # 根据搜索词过滤出候选型号
    if search.strip():
        filtered_mice = [m for m in all_mice if search.strip().lower() in m.lower()]
    else:
        filtered_mice = all_mice

    # 全选 / 清空
    c1, c2, _ = st.columns([1, 1, 4])
    with c1:
        if st.button("全选当前"):
            st.session_state['mice_selector'] = filtered_mice
    with c2:
        if st.button("清空"):
            st.session_state['mice_selector'] = []

    # 多选框：用 key 绑定 session_state，不用 default
    # 先确保 session 里存的值都在当前过滤范围内（换搜索词时清掉不在范围的）
    if 'mice_selector' in st.session_state:
        st.session_state['mice_selector'] = [m for m in st.session_state['mice_selector'] if m in filtered_mice]

    picked = st.multiselect(
        f"勾选型号（当前范围 {len(filtered_mice)} 个，勾多个会合并统计）",
        options=filtered_mice,
        key='mice_selector'  # ← 关键：用 key 绑定，不用 default
    )

    if picked:
        df_m = df_all.copy().sort_values('QueryTime')
        start = df_m['QueryTime'].min()
        end = df_m['QueryTime'].max()
        months = pd.date_range(start=start, end=end, freq='ME')
        active_days = 350

        records = []
        for mth in months:
            snap = df_m[df_m['QueryTime'] <= mth].drop_duplicates('Player', keep='last')
            snap = snap[snap['QueryTime'] >= mth - pd.Timedelta(days=active_days)]
            cnt = snap['Mouse'].isin(picked).sum()
            records.append({'Date': mth, 'Count': int(cnt)})

        mouse_trend = pd.DataFrame(records)
        label = "、".join(picked) if len(picked) <= 2 else f"{len(picked)} 个型号合计"

        fig2 = px.line(mouse_trend, x='Date', y='Count', markers=True, title=f"「{label}」使用人数趋势")
        fig2.update_traces(line_color='#E8112D')
        fig2.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                           font_color='white', xaxis_title='Date', yaxis_title='使用人数')
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("请先搜索、再勾选至少一个型号")

        # --- D4. 设置维度分布（eDPI / 回报率 / 重量）---
        st.markdown("<h2>📊 设置与硬件分布</h2>", unsafe_allow_html=True)

        # 时间滑块（这三个图共用）
        df_dist = df_all.copy().sort_values('QueryTime')
        month_opts_dist = pd.date_range(start=df_dist['QueryTime'].min(), end=df_dist['QueryTime'].max(), freq='ME')
        sel_month_dist = st.select_slider(
            "📅 选择时间点（下方三个分布联动）",
            options=list(month_opts_dist),
            value=month_opts_dist[-1],
            format_func=lambda d: d.strftime('%Y-%m'),
            key='dist_month_slider'
        )

        # 重建那个月的快照（每人最新一条 + 活跃过滤）
        snap_dist = df_dist[df_dist['QueryTime'] <= sel_month_dist].drop_duplicates('Player', keep='last')
        snap_dist = snap_dist[snap_dist['QueryTime'] >= sel_month_dist - pd.Timedelta(days=350)]

        dist_c1, dist_c2, dist_c3 = st.columns(3)

        # —— eDPI 分布 ——
        with dist_c1:
            st.markdown("<h2 style='font-size:20px;'>🎯 eDPI 分布</h2>", unsafe_allow_html=True)
            edpi = pd.to_numeric(snap_dist['eDPI'], errors='coerce').dropna()
            if len(edpi) > 0:
                bins = [0, 600, 800, 1000, 1200, 99999]
                labels = ['<600', '600-800', '800-1000', '1000-1200', '≥1200']
                edpi_binned = pd.cut(edpi, bins=bins, labels=labels, right=False)
                edpi_data = edpi_binned.value_counts().reindex(labels).reset_index()
                edpi_data.columns = ['区间', 'count']
                fig_e = px.bar(edpi_data, x='区间', y='count', text='count', color_discrete_sequence=['#E02020'])
                fig_e.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                    font_color="#E0E0E0", height=280, margin=dict(l=0, r=0, t=0, b=0),
                                    xaxis=dict(title=""), yaxis=dict(title="", gridcolor='#222'))
                fig_e.update_traces(textposition='outside')
                st.plotly_chart(fig_e, use_container_width=True)
            else:
                st.info("无 eDPI 数据")

        # —— 回报率分布 ——
        with dist_c2:
            st.markdown("<h2 style='font-size:20px;'>⚡ 回报率分布</h2>", unsafe_allow_html=True)
            hz = pd.to_numeric(snap_dist['HZ'], errors='coerce').dropna()
            if len(hz) > 0:
                # 固定档位，缺的补0，保证排版不变
                fixed_hz = [500, 1000, 2000, 4000, 8000]
                hz_counts = hz.astype(int).value_counts()
                hz_data = pd.DataFrame({
                    'Hz': [f'{h} Hz' for h in fixed_hz],
                    'count': [int(hz_counts.get(h, 0)) for h in fixed_hz]
                })
                fig_h = px.bar(hz_data, x='Hz', y='count', text='count', color_discrete_sequence=['#F39C12'])
                fig_h.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                    font_color="#E0E0E0", height=280, margin=dict(l=0, r=0, t=0, b=0),
                                    xaxis=dict(title=""), yaxis=dict(title="", gridcolor='#222'))
                fig_h.update_traces(textposition='outside')  # 回报率
                st.plotly_chart(fig_h, use_container_width=True)
            else:
                st.info("无回报率数据")

        # —— 重量分布 ——
        with dist_c3:
            st.markdown("<h2 style='font-size:20px;'>⚖️ 重量分布</h2>", unsafe_allow_html=True)
            weight = pd.to_numeric(snap_dist['weight'], errors='coerce').dropna()  # ← 注意列名
            if len(weight) > 0:
                bins_w = [0, 60, 70, 80, 999]
                labels_w = ['超轻<60g', '轻60-70g', '中70-80g', '重≥80g']
                w_binned = pd.cut(weight, bins=bins_w, labels=labels_w, right=False)
                w_data = w_binned.value_counts().reindex(labels_w).reset_index()
                w_data.columns = ['区间', 'count']
                fig_w = px.pie(w_data, values='count', names='区间', hole=0.4,
                               color_discrete_sequence=px.colors.sequential.Greens_r)
                fig_w.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                    font_color="#E0E0E0", height=280, margin=dict(l=0, r=0, t=0, b=0))
                st.plotly_chart(fig_w, use_container_width=True)
            else:
                st.info("无重量数据")

        # --- D5. 选手三年装备时间线 ---
        st.markdown("<h2>👤 选手装备演变（三年可追溯）</h2>", unsafe_allow_html=True)

        all_players = sorted(df_all['Player'].dropna().unique().tolist())

        # 给明星选手的显示名加👑（下拉框里显示带👑，但实际值还是原名）
        def player_label(name):
            return f"👑 {name}" if is_star(name) else name

        sel_player = st.selectbox(
            "选择选手（可下拉，也可打字搜索；👑=明星选手）",
            options=all_players,
            index=all_players.index("donk") if "donk" in all_players else 0,
            format_func=player_label,  # ← 关键：控制下拉里怎么显示
            key='player_timeline_select'
        )

        if sel_player:
            # 取该选手所有历史记录，按时间排序
            history = df_all[df_all['Player'] == sel_player].sort_values('QueryTime').copy()

            if not history.empty:
                # 上：时间线折线（eDPI 随

                # 下：变更明细表
                st.markdown(f"**{sel_player} 的装备变更记录（共 {len(history)} 次）**")
                show = history[['QueryTime', 'Mouse', 'DPI', 'Sens', 'eDPI', 'HZ']].copy()
                show['QueryTime'] = show['QueryTime'].dt.strftime('%Y-%m-%d')
                show.columns = ['日期', '鼠标型号', 'DPI', 'Sens', 'eDPI', '回报率(Hz)']
                st.dataframe(show, use_container_width=True, hide_index=True)
            else:
                st.info(f"暂无 {sel_player} 的记录")

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