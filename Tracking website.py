import streamlit as st
import pandas as pd
import os
import plotly.express as px
import base64
import re
from pathlib import Path

from PIL import Image

# python -m streamlit run "C:\Users\donny.d.huang\PycharmProjects\PythonProject\AI\cs2_mouse_site\Tracking website.py"
# ==========================================
# 1. 页面配置与高清电竞风 CSS
# ==========================================
st.set_page_config(
    page_title="ZOWIE HUB | CS2 Pro Mouse Tracker",
    layout="wide",
    initial_sidebar_state="collapsed"
)

import ast

def clean_physical_fields(df):
    """读取后统一清洗：直接覆盖原字段，全文使用干净数据。"""
    df = df.copy()

    # brand: '["Logitech"]' -> 'logitech'
    def parse_brand(raw):
        if pd.isna(raw):
            return raw
        try:
            val = ast.literal_eval(str(raw))
            if isinstance(val, list) and val:
                return str(val[0]).strip()
        except (ValueError, SyntaxError):
            pass
        return str(raw).strip()

    df['brand'] = df['brand'].apply(parse_brand)
    df['brand'] = df['brand'].astype(str).str.strip().str.lower().replace('nan', pd.NA)

    # 分类字段：统一小写去空格
    cat_fields = ['shape', 'size', 'material', 'hump_placement',
                  'front_flare', 'side_curvature', 'sensor_type']
    for f in cat_fields:
        df[f] = df[f].astype(str).str.strip().str.lower().replace('nan', pd.NA)

    # wireless: 1.0/0.0 -> 无线/有线
    df['wireless'] = df['wireless'].map({1.0: '无线', 0.0: '有线', 1: '无线', 0: '有线'})

    return df


# 图片路径
EXCEL_PATH = "https://raw.githubusercontent.com/d20260323-jpg/CS2_Mouse_Tracker/main/cs2_mouse_tracking.xlsx"  # EXCEL_PATH = "https://raw.githubusercontent.com/d20260323-jpg/CS2_Mouse_Tracker/main/cs2_mouse_tracking.xlsx"
ZOWIE_LOGO_PATH = "assets/zowie_logo.png"
HERO_MOUSE_PATH = "assets/hero_mouse.png"

# ── 表二：鼠标规格主表（每行一款鼠标 + 完整参数）──
SPEC_EXCEL_PATH = "mouseCatalog.xlsx"   # ← 改成你表二的真实文件名！

@st.cache_data(ttl=300)
def load_spec_table(path):
    """读取表二，只做规格对比需要的最小清洗（缺列也不报错）"""
    df = pd.read_excel(path)
    if 'wireless' in df.columns:
        df['wireless'] = df['wireless'].map(
            {1.0: '无线', 0.0: '有线', 1: '无线', 0: '有线'}
        ).fillna(df['wireless'])
    for f in ['size', 'shape']:
        if f in df.columns:
            df[f] = df[f].astype(str).str.strip().str.lower().replace('nan', pd.NA)
    return df

# 品牌固定配色（同一品牌永远同一颜色，不随名次变化）
BRAND_COLORS = {
    'Logitech': '#8B0000',  # 深红
    'ZOWIE': '#E02020',  # 正红
    'Razer': '#00B14F',  # 雷蛇绿（雷蛇品牌色）
    'Pulsar': '#FF6B35',  # 橙
    'VAXEE': '#F4A261',  # 浅橙
    'Lamzu': '#9B59B6',  # 紫
    'Finalmouse': '#3498DB',  # 蓝
    '其他': '#555555',  # 灰
}

# ─── 品牌标准化：统一大小写、简称、修掉选手名误判 ───
# 全网页品牌只认这一套标准写法，避免 ZOWIE/Zowie 分家、Fallen 被当品牌
_BRAND_ALIAS = {
    # 大小写归一
    'logitech': 'Logitech',
    'zowie': 'ZOWIE',
    'razer': 'Razer',
    'pulsar': 'Pulsar',
    'vaxee': 'VAXEE',
    'lamzu': 'Lamzu',
    'finalmouse': 'Finalmouse',
    'wlmouse': 'WLMouse',
    'dareu': 'Dareu',
    'hitscan': 'HITSCAN',
    'xtrfy': 'Xtrfy',
    'pwnage': 'Pwnage',
    'asus': 'ASUS',
    'corsair': 'Corsair',
    'glorious': 'Glorious',
    'steelseries': 'SteelSeries',
    'vxe': 'VXE',
    'atk': 'ATK',
    'teevolution': 'TEEVOLUTION',
    # 简称 -> 全称
    'endgame': 'Endgame Gear',
    'endgame gear': 'Endgame Gear',
    # 选手名/战队名误判为品牌 -> 归为未知（这些不是鼠标品牌）
    'fallen': None,
    'arbiter': None,
    'fnatic': None,
}


def canonical_brand(raw):
    """把任意品牌原始值标准化成统一写法；无法识别的选手名等返回 None。"""
    if pd.isna(raw):
        return None
    key = str(raw).strip().lower()
    if key in ('', 'nan', 'none', 'unknown'):
        return None
    if key in _BRAND_ALIAS:
        return _BRAND_ALIAS[key]  # 命中别名表（含误判归 None）
    # 未在别名表里的，原样返回（新品牌自动兼容）
    return str(raw).strip()


def resolve_brand(row):
    """统一的品牌解析：优先用数据表自带 brand 字段，缺失时回退 Mouse 首词。"""
    b = canonical_brand(row.get('brand'))
    if b is not None:
        return b
    mouse = row.get('Mouse')
    if pd.notnull(mouse):
        return canonical_brand(str(mouse).split()[0])
    return None


# ─── 明星选手名单（CS2 Major 冠军 / 2026 顶级选手，手动维护）───
# ⚠️ 名字拼写务必和你数据库 Player 字段一致；is_star 已忽略大小写
STAR_PLAYERS = {
    # —— NAVI（PGL Copenhagen 2024 冠军，CS2 首冠）——
    'Aleksib',
    'b1t',
    'jL',  # Copenhagen MVP
    'iM',
    'w0nderful',

    # —— Team Spirit（PW Shanghai 2024 冠军）——
    'donk',  # 史上最年轻 Major MVP
    'sh1ro',
    'chopper',
    'magixx',
    'zont1x',

    # —— Team Vitality（Austin 2025 + Budapest 2025 双冠）——
    'ZywOo',  # 四届年度最佳
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
    's1mple',  # CS:GO GOAT
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
    'Twistzz',  # (已有)
    'rain',  # FaZe 老将，人气高
    'karrigan',  # (已有) 传奇指挥
    'Snax',  # 波兰传奇
    'Boombl4',  # 前 NAVI 队长，话题度高
    'sh1ro',  # (已有)
    'Ax1Le',  # Cloud9/前 Gambit 明星
    'niko',  # 注意：有个小写 niko（不是大写NiKo那个），是另一个选手，别搞混
    'kennyS',  # 法国 AWP 传奇，人气极高
    'coldzera',  # 巴西传奇，两届 Major MVP
    'olofmeister',  # fnatic 传奇
    'dupreeh',  # 五冠王，Major 冠军最多
    'gla1ve',  # Astralis 指挥
    'Magisk',  # Astralis
}


def is_star(player):
    if not player:
        return False
    return str(player).strip().lower() in {s.lower() for s in STAR_PLAYERS}


BRIEFING_DIR = "市场简报"


def load_briefings():
    files = sorted(Path(BRIEFING_DIR).glob("*.md"), reverse=True)
    briefings = []

    for file in files:
        content = file.read_text(encoding="utf-8")

        # 提取第一个一级标题 / 二级标题作为主标题
        title_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
        if not title_match:
            title_match = re.search(r"^##\s+(.+)$", content, re.MULTILINE)

        title = title_match.group(1).strip() if title_match else file.stem

        # 从文件名提取日期
        date_match = re.search(r"(\d{8})", file.name)
        if date_match:
            date_raw = date_match.group(1)
            date = f"{date_raw[:4]}-{date_raw[4:6]}-{date_raw[6:]}"
        else:
            date = "Unknown Date"

        # 抽取前几个 ### 小标题作为摘要点
        section_titles = re.findall(r"^###\s+(.+)$", content, re.MULTILINE)
        preview = section_titles[:3]

        briefings.append({
            "title": title,
            "date": date,
            "file": str(file),
            "preview": preview,
            "content": content
        })

    return briefings


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

    /* 变动记录条目 */
    .change-log {
        background-color: #121212;
        border-left: 4px solid #333;
        padding: 12px 20px;
        margin-bottom: 10px;
        border-radius: 0 8px 8px 0;
    }
    
    /* 鼠标图片发光效果 */
    .hero-img {
        filter: drop-shadow(0 0 20px rgba(224, 32, 32, 0.4));
        max-width: 100%;
        height: auto;
    }

    .briefing-card {
    background: #151515;
    border: 1px solid #252525;
    border-radius: 12px;
    padding: 20px;
    min-height: 210px;
    transition: all 0.25s ease;
}

.briefing-date {
    color: #E02020;
    font-size: 13px;
    font-weight: 700;
    margin-bottom: 10px;
}

.briefing-title {
    color: white;
    font-size: 20px;
    font-weight: 800;
    line-height: 1.25;
    margin-bottom: 12px;
}

.briefing-preview {
    margin-top: 10px;
}

/* 下拉框（selectbox）hover 效果 */
.stSelectbox div[data-baseweb="select"] > div {
    transition: all 0.2s ease;
    cursor: pointer;
}
.stSelectbox div[data-baseweb="select"] > div:hover {
    border-color: #E02020 !important;
    box-shadow: 0 0 0 1px #E02020 !important;
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
# 2. 数据处理与缓存（仅统计最后 23 行）
# ==================================================

@st.cache_data(ttl=300)
def load_data():
    try:
        df_all = pd.read_excel(EXCEL_PATH)
        df_all = clean_physical_fields(df_all)

    except Exception as e:
        st.error(f"❌ 无法读取数据文件: {EXCEL_PATH}\n错误详情: {e}")
        return None, None

    try:
        # 处理日期格式，兼容带秒数的情况
        df_all['QueryTime'] = pd.to_datetime(
            df_all['QueryTime'],
            errors='coerce',
            format='mixed'
        )

        # 核心逻辑：截取最后 23 行作为当前快照
        df_latest = (
            df_all.sort_values('QueryTime')
            .drop_duplicates(subset='Player', keep='last')
            .copy()
        )

        # 提取品牌信息（统一走 resolve_brand：优先自带brand字段、修大小写/简称/误判）
        df_latest['Brand'] = df_latest.apply(resolve_brand, axis=1)

        # 预处理 Changed 状态
        df_all['Changed'] = df_all['Changed'].astype(str).str.upper().str.strip()

        df_all['Brand'] = df_all.apply(resolve_brand, axis=1)

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
                    实时同步全球顶尖选手的鼠标选择。基于三年核心样本数据分析。
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
            f'<div class="metric-box"><small>ACTIVE BRANDS</small><br><span style="font-size:35px; color:#E02020; font-weight:bold;">{df_latest["Brand"].nunique()}</span></div>',
            unsafe_allow_html=True)
    with c3:
        changed_col = df_all['Changed'].astype(str).str.strip().str.lower()

        total_changes = changed_col.isin([
            'both',
            'mouse',
            'settings'
        ]).sum()

        st.markdown(
            f'<div class="metric-box"><small>TOTAL CHANGES</small><br><span style="font-size:35px; color:#E02020; font-weight:bold;">{total_changes}</span></div>',
            unsafe_allow_html=True
        )

    st.markdown("<h2>📰 市场简报</h2>", unsafe_allow_html=True)

    briefings = load_briefings()  # 已按时间倒序，最新在最前

    # 弹窗：单篇全文
    @st.dialog("简报全文", width="large")
    def show_briefing(b):
        st.markdown(f"### {b['title']}")
        st.caption(f"📅 {b['date']}")
        st.markdown("---")
        st.markdown(b["content"])

    # 弹窗：往期报告列表（全部）
    @st.dialog("往期报告", width="large")
    def show_archive(all_briefings):
        st.markdown(f"共 {len(all_briefings)} 期简报（按时间倒序）")
        st.markdown("---")
        for b in all_briefings:
            with st.expander(f"📅 {b['date']}　{b['title']}"):
                st.markdown(b["content"])

    if briefings:
        latest = briefings[:3]  # 首页只展示最新3期
        cols = st.columns(3)

        for i, briefing in enumerate(latest):
            with cols[i % 3]:
                preview_html = ""
                for p in briefing["preview"]:
                    preview_html += f"<div style='color:#888; font-size:13px; margin-top:6px;'>• {p}</div>"

                st.markdown(f"""
                        <div class="briefing-card">
                            <div class="briefing-date">{briefing['date']}</div>
                            <div class="briefing-title">{briefing['title']}</div>
                            <div class="briefing-preview">
                                {preview_html}
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

                if st.button("查看全文", key=f"read_{i}", use_container_width=True):
                    show_briefing(briefing)

        # 往期报告按钮（只有超过3期才显示）
        if len(briefings) > 3:
            st.write("")
            _, mid, _ = st.columns([2, 1, 2])
            with mid:
                if st.button(f"📚 往期报告（共 {len(briefings)} 期）", use_container_width=True):
                    show_archive(briefings)
    else:
        st.info("暂无市场简报")

    # --- D. 图表展示（热门型号 + 品牌占比，共享时间滑块联动）---
    st.write("")

    # 共享时间滑块（横跨整宽，控制下面两个图）
    df_d = df_all.copy().sort_values('QueryTime')
    month_opts = pd.date_range(start=df_d['QueryTime'].min(), end=df_d['QueryTime'].max(), freq='ME')
    sel_month = st.selectbox(
        "📅 选择时间点",
        options=list(month_opts),
        index=len(month_opts) - 1,  # 默认选最后一个月
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
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

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
                            font_color="#E0E0E0", height=300, margin=dict(l=0, r=0, t=0, b=0),
                            legend=dict(font=dict(size=18, color="#FFFFFF")))
        st.plotly_chart(fig_p, use_container_width=True, config={'displayModeBar': False})

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
        color_discrete_map={
            'Logitech': '#00A2FF',  # 蓝（罗技代表色）
            'ZOWIE': '#E02020',  # 红（ZOWIE 主色）
            'Razer': '#3DDC84',  # 绿（雷蛇代表色）
            'VAXEE': '#FFB000',  # 黄
            'Pulsar': '#9D4EDD',  # 紫
        }
    )
    fig_line.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font_color="#FFFFFF",
        height=350,
        margin=dict(l=0, r=0, t=30, b=0),
        xaxis=dict(gridcolor='#222', tickfont=dict(color='#FFFFFF')),
        yaxis=dict(gridcolor='#222', title=None,
                   tickfont=dict(color='#FFFFFF')),
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#FFFFFF'))
    )

    fig_line.add_annotation(
        text="使用人数",
        xref="paper", yref="paper",
        x=0, y=1.05,  # 左上角，Y轴上方
        showarrow=False,
        font=dict(color="#FFFFFF", size=13),
        xanchor="left"
    )


    st.plotly_chart(fig_line, use_container_width=True, config={'displayModeBar': False})

    # --- D3. 鼠标型号趋势对比（多组：组内合并 + 组间对比）---
    import re

    def normalize_mouse_name(mouse):
        if pd.isna(mouse):
            return mouse

        mouse = str(mouse)

        # 去掉 Unreleased 标记
        mouse = re.sub(r"\s*\(Unreleased\)", "", mouse, flags=re.IGNORECASE)

        # 清理多余空格
        mouse = re.sub(r"\s+", " ", mouse).strip()

        return mouse

    def expand_mouse_keyword(kw):

        kw = kw.strip().lower()

        keyword_map = {

            "罗技": ["logitech", "g pro", "superlight", "superstrike"],
            "logi": ["logitech", "g pro", "superlight"],
            "雷蛇": ["razer", "viper", "deathadder"],
            "卓威": ["zowie"],  # 修复：原来["zowie","ZOWIE"]大写那半永远匹配不到
            "华硕": ["asus", "rog", "harpe", "keris"],
            "海盗船": ["corsair", "sabre"],
            "毒蝰": ["viper"],
            "蝰蛇": ["deathadder"],

        }

        return keyword_map.get(kw, [kw])

    # 注：Brand 已在 load_data 中通过 resolve_brand 统一设好（含大小写/简称/误判修正），
    # 此处不再重算，避免用更弱的逻辑把干净的 Brand 覆盖回脏值。

    st.markdown("<h2>🔍 鼠标型号趋势对比</h2>", unsafe_allow_html=True)

    df_m = df_all.copy()
    df_m['Mouse_Normalized'] = df_m['Mouse'].apply(normalize_mouse_name)
    df_m = df_m.sort_values('QueryTime')

    all_mice = sorted(df_m['Mouse_Normalized'].dropna().unique().tolist())
    end_all = df_m['QueryTime'].max()

    # 用递增id标识每个组（而不是列表下标），避免删组后checkbox状态串到别的组
    if 'trend_groups' not in st.session_state:
        latest_snap = df_m.drop_duplicates('Player', keep='last')
        latest_snap = latest_snap[latest_snap['QueryTime'] >= end_all - pd.Timedelta(days=350)]
        top1 = latest_snap['Mouse'].value_counts().head(1).index.tolist()
        st.session_state['trend_groups'] = [{'id': 0, 'name': '', 'mice': top1}]
        st.session_state['trend_group_id_counter'] = 1

    cc1, cc2, _ = st.columns([1, 1, 4])
    with cc1:
        if st.button("➕ 添加对比鼠标"):
            new_id = st.session_state['trend_group_id_counter']
            st.session_state['trend_groups'].append({'id': new_id, 'name': '', 'mice': []})
            st.session_state['trend_group_id_counter'] += 1
            st.rerun()

    with cc2:
        if st.button("🗑 清空所有组"):
            st.session_state['trend_groups'] = []
            st.rerun()

    group_to_delete = None

    for group in st.session_state['trend_groups']:
        gid = group['id']

        # ★修复1：在画标题之前，先直接从session_state读这个组当前真实勾选状态，
        # 而不是用上一轮遗留的group['mice']——这样标题里的"已选X个"才不会慢一拍。
        # session_state里的checkbox值在这一轮rerun开始时就已经是最新的了。
        synced_mice = set()
        for m in all_mice:
            ckey = f"mouse_{gid}_{m}"
            if ckey in st.session_state:
                if st.session_state[ckey]:
                    synced_mice.add(m)
            elif m in group['mice']:
                # 这个mouse这一轮还没渲染过checkbox（比如还没搜索到），保留原有勾选
                synced_mice.add(m)
        group['mice'] = list(synced_mice)

        display_index = st.session_state['trend_groups'].index(group) + 1
        if group['name'].strip():
            display_name = group['name']  # 用户手动输入的优先
        elif group['mice']:
            display_name = group['mice'][0]  # 没输入就用已选的第一个型号名
        else:
            display_name = f"鼠标{display_index}"  # 都没有才用占位名

        # ★修复2：expander用固定的key（绑定gid），不再依赖会变化的标题文字，
        # 这样用户手动折叠/展开的状态不会因为标题里数字变化或加了新组而被重置。
        expander_key = f"expander_{gid}"
        if expander_key not in st.session_state:
            st.session_state[expander_key] = True  # 只在第一次出现时给默认展开

        with st.expander(
                f"[点击收起]\n📊 {display_name}（已选 {len(group['mice'])} 个型号）",
                expanded=st.session_state[expander_key],
                key=expander_key
        ):

            group['name'] = st.text_input(
                "鼠标",
                value=group['name'],
                placeholder="输入品牌或型号，例如 罗技 / Logitech / 雷蛇 / Razer / ZOWIE / Superlight",
                key=f"gname_{gid}"
            )

            kw = group['name'].strip()

            if kw:
                search_terms = expand_mouse_keyword(kw)

                filtered_mice = [
                    m for m in all_mice
                    if any(term in m.lower() for term in search_terms)
                ]
            else:
                filtered_mice = []

            st.caption(f"找到 {len(filtered_mice)} 个型号")
            if len(filtered_mice) > 100:
                st.caption("⚠️ 结果过多，仅显示前100个，建议输入更具体的关键词缩小范围")

            c1, c2, c3 = st.columns(3)

            with c1:
                if st.button("全选当前结果", key=f"select_all_{gid}"):
                    for mouse in filtered_mice:
                        st.session_state[f"mouse_{gid}_{mouse}"] = True
                    group["mice"] = list(set(group["mice"]) | set(filtered_mice))
                    st.rerun()

            with c2:
                if st.button("清空本组鼠标", key=f"clear_group_{gid}"):
                    for mouse in all_mice:
                        st.session_state[f"mouse_{gid}_{mouse}"] = False
                    group["mice"] = []
                    st.rerun()

            selected_now = set(group["mice"])

            for mouse in filtered_mice[:100]:
                key = f"mouse_{gid}_{mouse}"

                if key not in st.session_state:
                    st.session_state[key] = mouse in selected_now

                checked = st.checkbox(mouse, key=key)

                if checked:
                    selected_now.add(mouse)
                else:
                    selected_now.discard(mouse)

            group["mice"] = list(selected_now)

            if group["mice"]:
                st.caption(
                    "已选：" +
                    "、".join(group["mice"][:8]) +
                    ("..." if len(group["mice"]) > 8 else "")
                )
                hidden_selected = set(group["mice"]) - set(filtered_mice)
                if hidden_selected:
                    st.caption(f"（其中 {len(hidden_selected)} 个不在当前搜索结果中，但仍保留在本组）")

            if st.button(f"删除 {display_name}", key=f"gdel_{gid}"):
                group_to_delete = gid

    if group_to_delete is not None:
        st.session_state['trend_groups'] = [
            g for g in st.session_state['trend_groups'] if g['id'] != group_to_delete
        ]
        st.rerun()

    active_groups = [g for g in st.session_state['trend_groups'] if g['mice']]

    if active_groups:
        months = pd.date_range(start=df_m['QueryTime'].min(), end=end_all, freq='ME')
        active_days = 350

        records = []
        for mth in months:
            snap = df_m[df_m['QueryTime'] <= mth].drop_duplicates('Player', keep='last')
            snap = snap[snap['QueryTime'] >= mth - pd.Timedelta(days=active_days)]

            for g in active_groups:
                line_name = g['name'].strip() if g['name'].strip() else "未命名鼠标"
                cnt = snap['Mouse_Normalized'].isin(g['mice']).sum()
                records.append({'Date': mth, '鼠标': line_name, 'Count': int(cnt)})

        trend_df = pd.DataFrame(records)

        fig2 = px.line(
            trend_df,
            x='Date',
            y='Count',
            color='鼠标',
            markers=True
        )

        fig2.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            height=420,
            margin=dict(t=30),  # ← 顶部留空间给横排标题
            xaxis_title='Date',
            yaxis_title=None,  # ← 关掉竖排"使用人数"
            legend=dict(
                bgcolor='rgba(0,0,0,0)',
                font=dict(color='#FFFFFF'),
                orientation='h',
                y=-0.3
            )
        )

        fig2.add_annotation(  # ← 横排"使用人数"放Y轴上方
            text="使用人数",
            xref="paper", yref="paper",
            x=0, y=1.05,
            showarrow=False,
            font=dict(color="#FFFFFF", size=13),
            xanchor="left"
        )

        st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("请先输入鼠标关键词并选择型号")

    # --- D4. 设置维度分布（eDPI / 回报率 / 重量）---
    st.markdown("<h2>📊 设置与硬件分布</h2>", unsafe_allow_html=True)

    # 时间滑块（这三个图共用）
    df_dist = df_all.copy().sort_values('QueryTime')
    month_opts_dist = pd.date_range(start=df_dist['QueryTime'].min(), end=df_dist['QueryTime'].max(), freq='ME')
    sel_month_dist = st.selectbox(
        "📅 选择时间点（下方三个分布联动）",
        options=list(month_opts_dist),
        index=len(month_opts_dist) - 1,  # 默认选最后一个月
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
            st.plotly_chart(fig_e, use_container_width=True, config={'displayModeBar': False})
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
            st.plotly_chart(fig_h, use_container_width=True, config={'displayModeBar': False})
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
            st.plotly_chart(fig_w, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("无重量数据")

    # --- D4-A. 物理属性分布：形状 / 尺寸 / 连接方式（ID设计师视角）---
    st.markdown("<h2>🖱️ 鼠标物理属性分布</h2>", unsafe_allow_html=True)
    st.caption("职业选手当前在用鼠标的物理设计特征分布，供硬件设计参考")

    # 时间滑块（这三个物理属性图共用）
    df_phys_all = df_all.copy()
    df_phys_all['QueryTime'] = pd.to_datetime(df_phys_all['QueryTime'], errors='coerce')
    df_phys_all = df_phys_all.dropna(subset=['QueryTime'])
    month_opts_phys = pd.date_range(start=df_phys_all['QueryTime'].min(),
                                    end=df_phys_all['QueryTime'].max(), freq='ME')
    sel_month_phys = st.selectbox(
        "📅 选择时间点（下方三个分布联动）",
        options=list(month_opts_phys),
        index=len(month_opts_phys) - 1,  # 默认选最后一个月
        format_func=lambda d: d.strftime('%Y-%m'),
        key='phys_month_slider'
    )

    # 复用和其他图一致的口径：当前活跃选手快照（每人最新一条 + 350天活跃过滤）
    df_phys = df_phys_all[df_phys_all['QueryTime'] <= sel_month_phys].copy()
    df_phys = df_phys.sort_values('QueryTime').drop_duplicates('Player', keep='last')
    df_phys = df_phys[df_phys['QueryTime'] >= sel_month_phys - pd.Timedelta(days=350)]

    phys_c1, phys_c2, phys_c3 = st.columns(3)

    # —— 形状分布 ——
    with phys_c1:
        n_shape = df_phys['shape'].notna().sum()
        st.markdown(f"<h2 style='font-size:20px;'>🔷 形状分布（{n_shape}人）</h2>", unsafe_allow_html=True)
        shape_data = df_phys['shape'].dropna().value_counts().reset_index()
        shape_data.columns = ['形状', 'count']
        # 英文值转中文显示，看着更友好
        shape_label = {'symmetrical': '对称', 'ergonomic': '人体工学', 'hybrid': '混合'}
        shape_data['形状'] = shape_data['形状'].map(lambda x: shape_label.get(x, x))
        if len(shape_data) > 0:
            fig_shape = px.pie(shape_data, values='count', names='形状', hole=0.4,
                               color_discrete_sequence=['#E02020', '#F39C12', '#3498DB'])
            fig_shape.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                    font_color="#E0E0E0", height=280, margin=dict(l=0, r=0, t=0, b=0))
            st.plotly_chart(fig_shape, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("无形状数据")

    # —— 尺寸分布 ——
    with phys_c2:
        n_size = df_phys['size'].notna().sum()
        st.markdown(f"<h2 style='font-size:20px;'>📏 尺寸分布（{n_size}人）</h2>", unsafe_allow_html=True)
        size_data = df_phys['size'].dropna().value_counts().reset_index()
        size_data.columns = ['尺寸', 'count']
        size_label = {'large': '大', 'medium': '中', 'small': '小'}
        size_order = ['小', '中', '大']
        size_data['尺寸'] = size_data['尺寸'].map(lambda x: size_label.get(x, x))
        if len(size_data) > 0:
            fig_size = px.bar(size_data, x='尺寸', y='count', text='count',
                              color_discrete_sequence=['#E02020'])
            fig_size.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                   font_color="#E0E0E0", height=280, margin=dict(l=0, r=0, t=0, b=0),
                                   xaxis=dict(title="", categoryorder='array', categoryarray=size_order),
                                   yaxis=dict(title="", gridcolor='#222'))
            fig_size.update_traces(textposition='outside')
            st.plotly_chart(fig_size, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("无尺寸数据")

    # —— 连接方式分布 ——
    with phys_c3:
        n_wl = df_phys['wireless'].notna().sum()
        st.markdown(f"<h2 style='font-size:20px;'>📡 连接方式（{n_wl}人）</h2>", unsafe_allow_html=True)
        wl_data = df_phys['wireless'].dropna().value_counts().reset_index()
        wl_data.columns = ['连接方式', 'count']
        if len(wl_data) > 0:
            fig_wl = px.pie(wl_data, values='count', names='连接方式', hole=0.4,
                            color_discrete_sequence=['#E02020', '#555555'])
            fig_wl.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                 font_color="#E0E0E0", height=280, margin=dict(l=0, r=0, t=0, b=0))
            st.plotly_chart(fig_wl, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("无连接方式数据")

    # --- D5. 选手三年装备时间线 ---
    st.markdown("<h2>👤 选手装备演变</h2>", unsafe_allow_html=True)

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
            # 数值列格式化：DPI/eDPI/HZ 去小数，Sens 保留2位
            for col in ['DPI', 'eDPI', 'HZ']:
                show[col] = show[col].round(0).astype('Int64')
            show['Sens'] = show['Sens'].map(lambda x: f"{x:g}" if pd.notnull(x) else "")
            show.columns = ['日期', '鼠标型号', 'DPI', 'Sens', 'eDPI', '回报率(Hz)']
            st.table(show.set_index('日期'))
        else:
            st.info(f"暂无 {sel_player} 的记录")

        # --- D6+D8 合并：自由交叉分析（数字表 + 热力图 双视图）---
        st.markdown("<h2>🔀 自由交叉分析</h2>", unsafe_allow_html=True)
        st.caption("自选两个维度：上方看精确人数（表格），下方看扎堆/空白（热力图）。如 品牌×回报率：看哪个品牌爱用高刷")

        # 当前快照（每人最新一条 + 活跃过滤），与其他图口径一致
        df_ct = df_all.copy().sort_values('QueryTime').drop_duplicates('Player', keep='last')
        df_ct = df_ct[df_ct['QueryTime'] >= df_ct['QueryTime'].max() - pd.Timedelta(days=350)].copy()

        # 数值字段分档，做成可交叉的分类字段
        edpi_labels = ['<600', '600-800', '800-1000', '1000-1200', '≥1200']
        df_ct['eDPI档'] = pd.cut(pd.to_numeric(df_ct['eDPI'], errors='coerce'),
                                 bins=[0, 600, 800, 1000, 1200, 99999],
                                 labels=edpi_labels)
        df_ct['重量档'] = pd.cut(pd.to_numeric(df_ct['weight'], errors='coerce'),
                                 bins=[0, 55, 60, 65, 70, 80, 999],
                                 labels=['≤55g', '55-60g', '60-65g', '65-70g', '70-80g', '>80g'])
        df_ct['回报率'] = pd.cut(pd.to_numeric(df_ct['HZ'], errors='coerce'),
                                 bins=[0, 500, 1000, 2000, 4000, 99999],
                                 labels=['≤500Hz', '1000Hz', '2000Hz', '4000Hz', '8000Hz'])
        df_ct['形状'] = df_ct['shape'].map({'symmetrical': '对称', 'ergonomic': '人体工学', 'hybrid': '混合'})
        df_ct['尺寸'] = df_ct['size'].map({'large': '大', 'medium': '中', 'small': '小'})
        df_ct['连接方式'] = df_ct['wireless']  # 已清洗成"无线/有线"
        df_ct['长度'] = pd.cut(pd.to_numeric(df_ct['length'], errors='coerce'),
                                 bins=[0, 118, 120, 122, 124, 126, 128, 130],
                                 labels=['≤118mm', '118-120mm', '120-122mm', '122-124gmm', '124-126mm', '126-128mm','128-130mm'])
        df_ct['宽度'] = pd.cut(pd.to_numeric(df_ct['width'], errors='coerce'),
                                   bins=[0, 62, 64, 66, 68, 70],
                                   labels=['≤62mm', '62-64mm', '64-66mm', '66-68mm', '68-70mm'])
        df_ct['高度'] = pd.cut(pd.to_numeric(df_ct['height'], errors='coerce'),
                                   bins=[0, 38, 40, 42, 44],
                                   labels=['≤38mm', '38-40mm', '40-42mm', '42-44mm'])
        df_ct['隆起位置'] = df_ct['hump_placement'].map({
            'back - aggressive': '后-明显',
            'back - minimal': '后-轻微',
            'back - moderate': '后-适中',
            'center': '居中',
        })
        df_ct['前端外扩'] = df_ct['front_flare'].map({
            'flat': '平直',
            'outward - slight': '外扩-轻微',
            'outward - moderate': '外扩-中等',
            'outward - aggressive': '外扩-强烈',
        })
        _sc = df_ct['side_curvature'].astype(str).str.replace('–', '-', regex=False).str.strip()
        df_ct['侧面弧度'] = _sc.map({
            'flat': '平直',
            'inward': '内凹',
            'inward - aggressive': '内凹-强烈',
        })
        # ===== 新增维度：选手DPI / 回报率上限 / 传感器规格 / 传感器型号 / 微动 / 滚轮 =====
        # 选手实际DPI（分档）
        df_ct['DPI档'] = pd.cut(pd.to_numeric(df_ct['dpi'], errors='coerce'),
                                bins=[0, 400, 800, 1600, 3200, 99999],
                                labels=['≤400', '401-800', '801-1600', '1601-3200', '>3200'])
        # 鼠标支持的回报率上限（取值少，直接分类）
        df_ct['回报率上限'] = pd.to_numeric(df_ct['polling_rate'], errors='coerce').astype('Int64').astype(str) \
            .replace('<NA>', pd.NA).map(lambda x: f'{x}Hz' if pd.notna(x) else x)

        # 传感器规格三件套（分档）
        df_ct['传感器最大DPI'] = pd.cut(pd.to_numeric(df_ct['sensor_dpi'], errors='coerce'),
                                        bins=[0, 20000, 26000, 32000, 40000, 99999],
                                        labels=['≤20k', '20-26k', '26-32k', '32-40k', '>40k'])
        df_ct['最大追踪速度'] = pd.cut(pd.to_numeric(df_ct['sensor_tracking_speed'], errors='coerce'),
                                       bins=[0, 400, 650, 750, 888, 9999],
                                       labels=['≤400', '400-650', '650-750', '750-888', '>888'])
        df_ct['最大加速度'] = pd.cut(pd.to_numeric(df_ct['acceleration'], errors='coerce'),
                                     bins=[0, 40, 50, 70, 88, 999],
                                     labels=['≤40G', '40-50G', '50-70G', '70-88G', '>88G'])
        # 传感器类型（分类）
        df_ct['传感器类型'] = df_ct['sensor_type'].astype(str).str.strip().str.lower().map(
            {'optical': '光学', 'laser': '激光'})

        # 取值很碎的字段：只保留最常见的 TOP N，其余归"其他"，避免交叉表炸成几十列
        def topn_category(series, n=12):
            s = series.astype(str).str.strip().replace('nan', pd.NA)
            top = s.value_counts().head(n).index
            return s.where(s.isin(top), other='其他').where(s.notna(), other=pd.NA)

        df_ct['传感器型号'] = topn_category(df_ct['sensor'], n=12)
        df_ct['微动'] = topn_category(df_ct['switch'], n=12)
        df_ct['滚轮'] = topn_category(df_ct['scroll'], n=12)


        # 可选交叉维度（显示名 → 字段名）
        field_map = {
            '品牌': 'Brand',
            '战队': 'Team',
            '回报率': '回报率',
            'eDPI档': 'eDPI档',
            '重量档': '重量档',
            '形状': '形状',
            '尺寸': '尺寸',
            '连接方式': '连接方式',
            '长度': '长度',
            '宽度': '宽度',
            '高度': '高度',
            '隆起位置': '隆起位置',
            '前端外扩': '前端外扩',
            '侧面弧度': '侧面弧度',
            'DPI档': 'DPI档',
            '回报率上限': '回报率上限',
            '传感器最大DPI': '传感器最大DPI',
            '最大追踪速度': '最大追踪速度',
            '最大加速度': '最大加速度',
            '传感器类型': '传感器类型',
            '传感器型号': '传感器型号',
            '微动': '微动',
            '滚轮': '滚轮',
        }
        field_names = list(field_map.keys())

        c1, c2 = st.columns(2)
        with c1:
            row_sel = st.selectbox("行维度", field_names, index=0, key='ct_row')
        with c2:
            col_sel = st.selectbox("列维度", field_names, index=2, key='ct_col')

        if row_sel == col_sel:
            st.warning("行和列请选不同的维度")
        else:
            row_col = field_map[row_sel]
            col_col = field_map[col_sel]
            row_data = df_ct[row_col].astype(str).replace('nan', '未知').fillna('未知')
            col_data = df_ct[col_col].astype(str).replace('nan', '未知').fillna('未知')

            # 各维度的"正确顺序"（有内在大小的维度按逻辑排，不按字母）
            order_map = {
                '回报率': ['≤500Hz', '1000Hz', '2000Hz', '4000Hz', '8000Hz'],
                '重量档': ['≤55g', '55-60g', '60-65g', '65-70g', '70-80g', '>80g'],
                'eDPI档': edpi_labels,   # 直接引用，保证和pd.cut用的完全一致
                '尺寸': ['小', '中', '大'],
                '连接方式': ['有线', '无线'],
                '长度': ['≤118mm', '118-120mm', '120-122mm', '122-124gmm', '124-126mm', '126-128mm','128-130mm'],
                '宽度': ['≤62mm', '62-64mm', '64-66mm', '66-68mm', '68-70mm'],
                '高度': ['≤38mm', '38-40mm', '40-42mm', '42-44mm'],
                '隆起位置': ['居中', '后-轻微', '后-适中', '后-明显'],
                '前端外扩': ['平直', '外扩-轻微', '外扩-中等', '外扩-强烈'],
                '侧面弧度': ['平直', '内凹', '内凹-强烈'],
                'DPI档': ['≤400', '401-800', '801-1600', '1601-3200', '>3200'],
                '回报率上限': ['1000Hz', '4000Hz', '8000Hz'],
                '传感器最大DPI': ['≤20k', '20-26k', '26-32k', '32-40k', '>40k'],
                '最大追踪速度': ['≤400', '400-650', '650-750', '750-888', '>888'],
                '最大加速度': ['≤40G', '40-50G', '50-70G', '70-88G', '>88G'],
                '传感器类型': ['光学', '激光'],
            }

            def apply_order(df_table, is_rows):
                """按维度逻辑顺序重排行或列，未知放最后"""
                sel = row_sel if is_rows else col_sel
                if sel in order_map:
                    desired = [x for x in order_map[sel] if x in (df_table.index if is_rows else df_table.columns)]
                    # 把"未知"和其他没列到的值追加到末尾
                    rest = [x for x in (df_table.index if is_rows else df_table.columns)
                            if x not in desired and x != '合计']
                    final = desired + rest
                    if '合计' in (df_table.index if is_rows else df_table.columns):
                        final = final + ['合计']
                    return df_table.reindex(index=final) if is_rows else df_table.reindex(columns=final)
                return df_table

            # —— 上：精确数字表（带合计，按维度逻辑顺序排）——
            ct = pd.crosstab(row_data, col_data, margins=True, margins_name='合计')
            ct = apply_order(ct, is_rows=True)
            ct = apply_order(ct, is_rows=False)
            st.markdown("**📋 精确人数表**")
            st.table(ct)

            # —— 下：热力图（不含合计行列，看扎堆/空白，同样按逻辑顺序排）——
            ct_heat = pd.crosstab(row_data, col_data)  # 无 margins
            ct_heat = apply_order(ct_heat, is_rows=True)
            ct_heat = apply_order(ct_heat, is_rows=False)
            st.markdown("**🗺️ 分布热力图**（颜色越深=越扎堆，空白格=潜在机会）")

            import plotly.graph_objects as go
            fig_ct = go.Figure(data=go.Heatmap(
                z=ct_heat.values,
                x=[str(c) for c in ct_heat.columns],
                y=[str(i) for i in ct_heat.index],
                text=ct_heat.values,
                texttemplate="%{text}",
                textfont={"size": 15, "color": "#FFFFFF"},
                colorscale=[[0, '#1a1a1a'], [0.01, '#3d0000'], [0.5, '#a01515'], [1, '#E02020']],
                showscale=True,
                colorbar=dict(title="选手数"),
                hovertemplate=f"{row_sel}=%{{y}}<br>{col_sel}=%{{x}}<br>选手数=%{{z}}<extra></extra>"
            ))
            fig_ct.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', height=400,
                xaxis_title=col_sel, yaxis_title=row_sel,
                xaxis=dict(type='category'), yaxis=dict(type='category'),
                margin=dict(l=0, r=0, t=10, b=0)
            )
            st.plotly_chart(fig_ct, use_container_width=True, config={'displayModeBar': False})
            st.caption("⚠️ 空白格有两种含义：真空白(市场机会) 或 不合理组合(如超重+超小)，需结合设计经验判断")

        # --- D7. 鼠标规格对比器（搜索逐个加入 → 参数并排对比）---
        st.markdown("<h2>🔧 鼠标规格对比</h2>", unsafe_allow_html=True)
        st.caption("搜索并逐个添加鼠标，横向对比物理规格参数，供硬件设计/竞品分析参考")

        # 规格数据源 = 表二（只含有具体参数的鼠标）
        spec_cols = ['weight', 'size', 'shape', 'wireless', 'length', 'width',
                     'height', 'material', 'sensor', 'switch']

        try:
            df_spec_raw = load_spec_table(SPEC_EXCEL_PATH)
        except Exception as e:
            st.error(f"❌ 无法读取规格表(表二): {SPEC_EXCEL_PATH} — {e}")
            df_spec_raw = pd.DataFrame(columns=['Mouse'] + spec_cols)

        # 只保留表二里真实存在的规格列，避免列名对不上
        spec_cols = [c for c in spec_cols if c in df_spec_raw.columns]
        # 同型号可能有多行 → 按型号去重取第一条
        df_spec = (df_spec_raw.dropna(subset=['Mouse'])
                   .drop_duplicates('Mouse')
                   .set_index('Mouse'))

        # 初始化对比清单
        if 'compare_mice' not in st.session_state:
            st.session_state['compare_mice'] = []

        # 候选型号 = 表二里的鼠标（保证每个选出来都有参数）
        all_mouse_options = sorted(df_spec.index.dropna().unique().tolist())

        # 清掉之前可能残留、但表二里没有的旧选择，避免 multiselect 报错
        st.session_state['compare_mice'] = [
            m for m in st.session_state['compare_mice'] if m in all_mouse_options
        ]
        selected = st.multiselect(
            "🔍 搜索并选择鼠标型号（可输入关键词筛选，支持多选对比）",
            options=all_mouse_options,
            default=st.session_state['compare_mice'],
            key='spec_multiselect',
            placeholder="输入关键词，如 superlight / EC2 / viper"
        )
        st.session_state['compare_mice'] = selected

        if st.session_state['compare_mice']:
            # 中文参数名映射
            row_labels = {
                'weight': '重量(g)', 'size': '尺寸', 'shape': '形状',
                'wireless': '连接方式', 'length': '长(mm)', 'width': '宽(mm)',
                'height': '高(mm)', 'material': '材质', 'sensor': '传感器', 'switch': '微动',
            }
            # 值的中英转换（复用已清洗字段）
            val_map = {
                'symmetrical': '对称', 'ergonomic': '人体工学', 'hybrid': '混合',
                'large': '大', 'medium': '中', 'small': '小',
            }

            def fmt(mouse, col):
                if mouse not in df_spec.index:
                    return '—'
                v = df_spec.at[mouse, col]
                if pd.isna(v):
                    return '—'
                v = val_map.get(v, v)  # 中文化
                # 数字去掉多余小数
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                return v

            # 组装对比表：行=参数，列=各款鼠标
            table = {}
            for mouse in st.session_state['compare_mice']:
                table[mouse] = [fmt(mouse, c) for c in spec_cols]
            compare_df = pd.DataFrame(table, index=[row_labels[c] for c in spec_cols])

            st.dataframe(compare_df, use_container_width=True)
            st.caption("“—”表示该型号此项参数在数据库中暂缺")
        else:
            st.info("搜索并添加鼠标后，这里会显示参数对比表")

    # --- E. 变动快讯（显示具体变更内容）---
    st.markdown("<h2>🔄 最近十次变动动态</h2>", unsafe_allow_html=True)

    # 清洗 + 转时间
    df_all['Changed'] = df_all['Changed'].astype(str).str.strip().str.upper()
    df_all['QueryTime'] = pd.to_datetime(df_all['QueryTime'], errors='coerce')

    # 要对比的设置字段：列名 -> 显示名
    SETTING_FIELDS = {
        'DPI': 'DPI',
        'polling_rate': '回报率',
        'Sens': '灵敏度',
        'eDPI': 'eDPI',
    }

    # 为每位选手按时间排序，取出"上一条"的各设置值，用于差异对比
    df_all = df_all.sort_values(['Player', 'QueryTime'])
    for col in SETTING_FIELDS:
        if col in df_all.columns:
            df_all[f'prev_{col}'] = df_all.groupby('Player')[col].shift(1)

    def diff_settings(row):
        """列出该行相对上一条记录，具体变了哪些设置"""
        changes = []
        for col, label in SETTING_FIELDS.items():
            if col not in df_all.columns:
                continue
            old, new = row.get(f'prev_{col}'), row.get(col)
            if pd.notnull(old) and pd.notnull(new) and str(old) != str(new):
                changes.append(f'{label} {old}→{new}')
        return changes

    def describe_change(row):
        ctype = str(row['Changed']).strip().upper()
        mouse = row.get('Mouse', '未知')
        if ctype == 'MOUSE':
            return f'选手 <b>{row["Player"]}</b> 切换至 <span style="color:#E02020;">{mouse}</span>'
        if ctype == 'BOTH':
            detail = '；'.join(diff_settings(row))
            tail = f'（{detail}）' if detail else '（设备与设置同时变动）'
            return f'选手 <b>{row["Player"]}</b> 切换至 <span style="color:#E02020;">{mouse}</span> {tail}'
        if ctype == 'SETTINGS':
            detail = '；'.join(diff_settings(row))
            tail = f'：{detail}' if detail else '（DPI / 回报率等）'
            return f'选手 <b>{row["Player"]}</b> 更新了设置{tail}'
        if ctype == 'NEW':
            return f'新增选手 <b>{row["Player"]}</b>（<span style="color:#E02020;">{mouse}</span>）'
        return f'选手 <b>{row["Player"]}</b> 数据有更新'

    # 排序取最近 10 条（差异列已算好，这里再按时间倒序）
    CHANGE_TYPES = ['MOUSE', 'BOTH', 'SETTINGS', 'NEW']
    changed_list = (
        df_all[df_all['Changed'].isin(CHANGE_TYPES)]
        .sort_values('QueryTime', ascending=False)
        .head(10)
    )

    if not changed_list.empty:
        for _, row in changed_list.iterrows():
            time_str = row['QueryTime'].strftime('%Y-%m-%d %H:%M') if pd.notnull(row['QueryTime']) else "N/A"
            st.markdown(f"""
                <div class="change-log">
                    <span style="color:#666; font-size:12px;">{time_str}</span><br>
                    {describe_change(row)}
                </div>
            """, unsafe_allow_html=True)
    else:
        st.info("目前监测中... 暂无近期设备变更记录。")




if __name__ == "__main__":
    main()
