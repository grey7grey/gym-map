"""
工会合作健身房地图 — Streamlit + 高德地图 API + folium

功能：
  1. 读取 xlsx 门店列表（门店名称 / 服务 / 详细地址）
  2. 调用高德地图地理编码 API 将地址转为经纬度
  3. 侧边栏设置"我的位置"（🏠家 / 🏢公司 一键定位，或临时文字地址，坐标已烘焙无需 Key）
  4. 计算直线距离，侧边栏列出离我最近的 Top 5 门店
  5. 右侧用 folium 渲染交互式地图，所有门店大头针可点击查看店名与服务

依赖安装：
  pip install streamlit pandas openpyxl requests folium streamlit-folium

运行：
  streamlit run app.py
"""

import os
import time
import math
import re
import importlib.util

import pandas as pd
import requests
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from dotenv import load_dotenv

load_dotenv()  # 读取同目录 .env 里的 AMAP_KEY（不上 Git）


# ---------------- 坐标系转换 ----------------
def gcj02_to_wgs84(lng: float, lat: float):
    """高德 GCJ02（火星坐标）→ WGS84（OpenStreetMap 瓦片坐标系）。

    门店坐标 / 家 / 公司 / 临时地址均来自高德地理编码，皆为 GCJ02；
    而 folium 默认瓦片是 WGS84，绘制前必须转换，否则所有标记会整体向东南偏移。
    这里用一次反馈迭代做近似逆变换，误差 < 1m 级，足以用于地图显示。
    """
    a = 6378245.0
    ee = 0.00669342162296594323

    def _t_lat(x, y):
        ret = (-100.0 + 2.0 * x + 3.0 * y + 0.2 * y * y + 0.1 * x * y
               + 0.2 * math.sqrt(abs(x)))
        ret += (20.0 * math.sin(6.0 * x * math.pi) + 20.0 * math.sin(2.0 * x * math.pi)) * 2.0 / 3.0
        ret += (20.0 * math.sin(y * math.pi) + 40.0 * math.sin(y / 3.0 * math.pi)) * 2.0 / 3.0
        ret += (160.0 * math.sin(y / 12.0 * math.pi) + 320 * math.sin(y * math.pi / 30.0)) * 2.0 / 3.0
        return ret

    def _t_lng(x, y):
        ret = (300.0 + x + 2.0 * y + 0.1 * x * x + 0.1 * x * y
               + 0.1 * math.sqrt(abs(x)))
        ret += (20.0 * math.sin(6.0 * x * math.pi) + 20.0 * math.sin(2.0 * x * math.pi)) * 2.0 / 3.0
        ret += (20.0 * math.sin(x * math.pi) + 40.0 * math.sin(x / 3.0 * math.pi)) * 2.0 / 3.0
        ret += (150.0 * math.sin(x / 12.0 * math.pi) + 300.0 * math.sin(x / 30.0 * math.pi)) * 2.0 / 3.0
        return ret

    if not (73.66 < lng < 135.05 and 3.86 < lat < 53.55):
        return lng, lat  # 境外不转换
    dlat = _t_lat(lng - 105.0, lat - 35.0)
    dlng = _t_lng(lng - 105.0, lat - 35.0)
    radlat = lat / 180.0 * math.pi
    magic = math.sin(radlat)
    magic = 1 - ee * magic * magic
    sqrtmagic = math.sqrt(magic)
    dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * math.pi)
    dlng = (dlng * 180.0) / (a / sqrtmagic * math.cos(radlat) * math.pi)
    mglat = lat + dlat
    mglng = lng + dlng
    # 逆变换近似：WGS84 = GCJ02 - (GCJ02 - WGS84)
    return lng - (mglng - lng), lat - (mglat - lat)


# ---------------- 配置 ----------------
# 脚本与 xlsx 放在同一目录，自动定位表格
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "公司工会合作健身房列表.xlsx")
# 上海市中心（人民广场附近）(纬度 lat, 经度 lng)
SHANGHAI_CENTER = (31.2304, 121.4737)
AMAP_GEO_URL = "https://restapi.amap.com/v3/geocode/geo"

# 家/公司坐标已烘焙（经预解析，启动不再调用高德）；仅临时外出地址才实时解析
HOME_ADDR = "上海市浦东新区德平路25弄"          # 显示用
HOME_COORD = (31.254035, 121.557767)            # (lat, lng)
WORK_ADDR = "上海市浦东新区银城路167号"          # 显示用
WORK_COORD = (31.243060, 121.513486)            # (lat, lng)

# 当前位置（默认「家」），用 session_state 保存以便在定位后联动更新
if "cur_lat" not in st.session_state:
    st.session_state.cur_lat = HOME_COORD[0]
if "cur_lng" not in st.session_state:
    st.session_state.cur_lng = HOME_COORD[1]


# ---------------- 数据加载 ----------------
@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    """读取表格并归一化列名，过滤掉无实体地址的线上券。"""
    df = pd.read_excel(path)
    rename = {}
    for col in df.columns:
        s = str(col)
        if any(k in s for k in ("门店", "名称", "项目")):
            rename[col] = "name"
        elif "服务" in s:
            rename[col] = "service"
        elif "地址" in s:
            rename[col] = "address"
    df = df.rename(columns=rename)[["name", "service", "address"]]
    # 过滤"详见第三方服务网络"、全城通兑类无实体地址的行
    mask = df["address"].notna() & ~df["address"].astype(str).str.contains("第三方|详见|全城|所有门店", na=False)
    df = df[mask].reset_index(drop=True)
    return df


# ---------------- 高德地理编码 ----------------
def _address_variants(addr: str):
    """生成地理编码候选地址：原样 → 去括号内容 → 削掉楼层/铺位后缀。"""
    variants = [addr]
    s = re.sub(r"[（(][^（）()]*[）)]", "", addr).strip()  # 去 (地铁口) 等
    if s and s != addr:
        variants.append(s)
    m = re.match(r"^(.*?(?:路|道|街|号)[^\d]*\d+号?)", s)  # 取到"号"为止，丢弃 L5楼18B号铺
    if m and m.group(1).strip() != addr:
        variants.append(m.group(1).strip())
    seen, out = set(), []
    for v in variants:
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


@st.cache_data(show_spinner=False)
def geocode(address: str, api_key: str, city: str = "上海"):
    """调用高德地理编码 API，返回 (lng, lat)；失败返回 None。会对地址做去括号/削楼层兜底重试。"""
    def _call(addr: str, city_arg: str):
        params = {"address": addr, "key": api_key}
        if city_arg:
            params["city"] = city_arg
        for attempt in range(5):  # 退避重试应对 QPS 限流
            try:
                r = requests.get(AMAP_GEO_URL, params=params, timeout=10)
                data = r.json()
                if data.get("status") == "1" and data.get("geocodes"):
                    lng, lat = data["geocodes"][0]["location"].split(",")
                    return float(lng), float(lat)
                info = data.get("info", "")
                if info in ("QPS_EXCEED", "INVALID_USER_KEY", "USERKEY_PLAT_NOMATCH"):
                    time.sleep(0.4 * (2 ** attempt))
                    continue
                return None
            except Exception:
                time.sleep(0.4 * (2 ** attempt))
                continue
        return None

    # 依次尝试各候选地址（原样 → 去括号 → 削楼层），带城市偏好失败再不带城市
    for variant in _address_variants(address):
        c = _call(variant, city) or _call(variant, "")
        if c:
            return c
    return None


# ---------------- 距离计算（Haversine 直线距离，单位 km） ----------------
def haversine(lat1: float, lng1: float, lat2: float, lng2: float) -> float:
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lng2 - lng1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


# ---------------- 页面 ----------------
st.set_page_config(page_title="工会合作健身房地图", layout="wide")

# ===== 侧边栏 =====

# Key 走 .env 自动读取，仅临时外出地址解析时才需要（家/公司坐标已烘焙，无需调用）
api_key = os.getenv("AMAP_KEY", "")

st.sidebar.header("📍 我的位置")

# 家/公司坐标已烘焙，直接写入 session_state，启动零网络调用
c1, c2 = st.sidebar.columns(2)
if c1.button("🏠 家", use_container_width=True, help=HOME_ADDR):
    st.session_state.cur_lat, st.session_state.cur_lng = HOME_COORD
    st.rerun()
if c2.button("🏢 公司", use_container_width=True, help=WORK_ADDR):
    st.session_state.cur_lat, st.session_state.cur_lng = WORK_COORD
    st.rerun()

# 临时位置：外出不在家/公司时，手动输入文字地址并点「定位」转成坐标（才调用高德）
st.sidebar.markdown("**或临时输入一个地址**")
tmp_addr = st.sidebar.text_input(
    "临时地址",
    value=st.session_state.get("tmp_addr", ""),
    key="tmp_addr",
    placeholder="例如：上海市浦东新区世纪大道100号",
    label_visibility="collapsed",
)
if st.sidebar.button("🔍 定位这个地址", key="addr_btn", use_container_width=True):
    if not api_key:
        st.sidebar.error("临时地址转坐标需要高德 Key：请在项目根目录 .env 里配置有效 AMAP_KEY。")
    elif not tmp_addr.strip():
        st.sidebar.warning("请先输入一个地址再点定位。")
    else:
        with st.spinner("正在解析地址…"):
            c = geocode(tmp_addr.strip(), api_key)
        if c:
            st.session_state.cur_lat, st.session_state.cur_lng = c[1], c[0]  # (lng, lat)
            st.rerun()
        else:
            st.sidebar.error("地址解析失败，请写得更具体些（带区 / 路 / 号）。")

# ===== 侧边栏：筛选大标题 + 搜索框 + 快捷关键词 =====
st.sidebar.header("🔍 筛选")
search_kw = st.sidebar.text_input(
    "搜索门店",
    key="search_kw",
    placeholder="🔍 搜索（名称 / 服务，留空显示全部）",
    help="同时匹配「门店名称」和「提供服务」两列，留空则显示全部。",
    label_visibility="collapsed",
)
def _set_kw(val: str):
    # 回调在 widget 实例化前运行，可安全写入 session_state
    st.session_state.search_kw = val


qc1, qc2 = st.sidebar.columns(2)
qc1.button("⚡ 24/7", key="q_24", use_container_width=True, on_click=_set_kw, args=("24/7",))
qc2.button("⚡ 游泳", key="q_swim", use_container_width=True, on_click=_set_kw, args=("游泳",))

# ===== 主界面 =====
st.title("🏋️ 工会合作健身房地图")
st.caption("读取门店列表 → 高德地理编码 → 计算直线距离 → 交互式地图")

df = load_data(EXCEL_PATH)
st.write(f"可用门店：**{len(df)}** 家")

# 当前位置统一从 session_state 读取（地址按钮 / GPS / 手动输入都写回这里）
cur_lat = st.session_state.cur_lat
cur_lng = st.session_state.cur_lng

# 优先使用内置坐标（store_coords.py 由 geocode_once.py 生成），无需 Key、无需联网
# 每次刷新都重新加载文件，改完坐标后按 F5 即生效，无需手动重启进程
def _load_store_coords():
    path = os.path.join(os.path.dirname(__file__), "store_coords.py")
    try:
        spec = importlib.util.spec_from_file_location("store_coords_dyn", path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.STORE_COORDS, True
    except Exception:
        return {}, False


STORE_COORDS, HAVE_BUILTIN = _load_store_coords()

if HAVE_BUILTIN:
    df["lng"] = df["name"].map(lambda n: STORE_COORDS.get(n, {}).get("lng"))
    df["lat"] = df["name"].map(lambda n: STORE_COORDS.get(n, {}).get("lat"))
    df_geo = df[df["lat"].notna()].copy()
    failed = [n for n in df["name"] if n not in STORE_COORDS]
else:
    if not api_key:
        st.warning("请在左侧侧边栏填入「高德地图 API Key」后再解析地址；"
                   "或复制 .env.example 为 .env 并填入 Key，再运行 `python geocode_once.py` 生成内置坐标，之后不再需要 Key。")
        st.stop()
    with st.spinner("正在通过高德 API 解析门店地址…"):
        lngs, lats, failed = [], [], []
        for _, row in df.iterrows():
            c = geocode(row["address"], api_key)
            if c:
                lngs.append(c[0])
                lats.append(c[1])
            else:
                lngs.append(None)
                lats.append(None)
                failed.append(row["name"])
            time.sleep(0.3)  # 限速 ~3 次/秒，避免触发高德免费 Key 限流
    df["lng"], df["lat"] = lngs, lats
    df_geo = df[df["lat"].notna()].copy()

if failed:
    st.warning(f"有 {len(failed)} 家门店地址解析失败：{', '.join(failed[:8])}"
               + (" …" if len(failed) > 8 else ""))

# 计算到当前位置的直线距离
df_geo["distance_km"] = df_geo.apply(
    lambda r: haversine(cur_lat, cur_lng, r["lat"], r["lng"]), axis=1
)
df_geo = df_geo.sort_values("distance_km").reset_index(drop=True)

# 关键词筛选（同时匹配门店名称 / 提供服务）；留空 = 全部
kw = (st.session_state.get("search_kw") or "").strip()
df_view = df_geo
if kw:
    pat = re.escape(kw)
    df_view = df_geo[
        df_geo["name"].astype(str).str.contains(pat, case=False, na=False)
        |         df_geo["service"].astype(str).str.contains(pat, case=False, na=False)
    ].copy()

if kw:
    st.info(f"🔎 正在筛选「{kw}」，共 **{len(df_view)}** 家匹配，已按距当前位置由近到远排序。")

# ===== 侧边栏：离我最近的 Top 5 =====
st.sidebar.header("🏆 离我最近TOP5")
if kw:
    st.sidebar.caption(f"筛选「{kw}」：命中 **{len(df_view)}** 家")
if len(df_view) == 0:
    st.sidebar.warning("没有匹配的门店，换个关键词试试。")
else:
    for i, r in df_view.head(5).iterrows():
        st.sidebar.markdown(f"**{i + 1}. {r['name']}**")
        st.sidebar.caption(f"📍 {r['distance_km']:.2f} 公里 | {r['address']}")

# ===== 右侧：folium 交互式地图 =====
# 注意：高德坐标(GCJ02)需转 WGS84 才能对齐 OpenStreetMap 瓦片，否则整体向东南偏
cur_lng_wgs, cur_lat_wgs = gcj02_to_wgs84(cur_lng, cur_lat)
# 默认 zoom=14 看清家门口（家附近商圈级），有筛选时再用 fit_bounds 缩到筛选范围
m = folium.Map(location=[cur_lat_wgs, cur_lng_wgs], zoom_start=14)

# 我的位置（红色大头针，独立于 cluster，置顶）
folium.Marker(
    [cur_lat_wgs, cur_lng_wgs],
    tooltip="我的位置",
    icon=folium.Icon(color="red", icon="home"),
).add_to(m)

# 有筛选时，地图只显示筛选门店，更聚焦；否则显示全部
df_markers = df_view if kw else df_geo
# 把 GCJ02 门店坐标转为 WGS84 再绘制（对齐 OpenStreetMap 瓦片，修正整体东南偏）
# 函数 gcj02_to_wgs84(lng, lat) 返回 (lng_wgs, lat_wgs)，zip 解包后要按顺序赋给 wlng / wlat
df_markers = df_markers.copy()
df_markers["wlng"], df_markers["wlat"] = zip(*df_markers.apply(
    lambda r: gcj02_to_wgs84(r["lng"], r["lat"]), axis=1))

# 用 MarkerCluster 自动聚类：缩小时多个标记合成一个带数字的大圆点，
# 放大后逐个散开；Top5 的红圆数字号仍在 cluster 内（放大后能看清）。
marker_cluster = MarkerCluster(
    name="门店",
    show=True,
    options={
        "maxClusterRadius": 60,        # 像素：同位置 60px 内聚成一团
        "disableClusteringAtZoom": 16, # 放到 16 级以上不再聚合，逐个看
    },
).add_to(m)

# 各门店：Top5 用红色数字徽标高亮，其余用蓝色 i 图标；点击看店名 + 服务 + 地址 + 距离
def numbered_icon(num: int):
    html = (
        f'<div style="background-color:#e74c3c;color:#fff;border-radius:50%;'
        f'width:28px;height:28px;display:flex;align-items:center;justify-content:center;'
        f'font-weight:bold;font-size:15px;border:2px solid #fff;'
        f'box-shadow:0 0 4px rgba(0,0,0,.4);">{num}</div>'
    )
    return folium.DivIcon(html=html, icon_size=(28, 28), icon_anchor=(14, 14))


for idx, r in df_markers.iterrows():
    popup_html = (
        f"<b>{r['name']}</b><br>"
        f"🛎 服务：{r['service']}<br>"
        f"📍 地址：{r['address']}<br>"
        f"📏 距离：{r['distance_km']:.2f} km"
    )
    if idx < 5:
        icon = numbered_icon(idx + 1)
    else:
        icon = folium.Icon(color="blue", icon="info-sign")
    folium.Marker(
        [r["wlat"], r["wlng"]],
        popup=folium.Popup(popup_html, max_width=320),
        tooltip=r["name"],
        icon=icon,
    ).add_to(marker_cluster)

# 有筛选时把地图缩放到筛选范围；无筛选时保持 zoom 14 看清家门口
if kw:
    all_lats = list(df_markers["wlat"]) + [cur_lat_wgs]
    all_lngs = list(df_markers["wlng"]) + [cur_lng_wgs]
    m.fit_bounds([[min(all_lats), min(all_lngs)], [max(all_lats), max(all_lngs)]])

# 用 streamlit 新版 st.iframe 渲染 folium —— 写到本地临时 html，iframe 引用，streamlit 不 diff 内部节点，
# 避免 streamlit-folium 在筛选/重定位触发重渲染时偶发 React NotFoundError (removeChild)
# 代价：不能像 st_folium 那样把点击坐标回传 streamlit（本应用未用到，无影响）
_map_cache = os.path.join(os.path.dirname(__file__), ".cache", "map.html")
os.makedirs(os.path.dirname(_map_cache), exist_ok=True)
m.save(_map_cache)
st.iframe(_map_cache, height=620)
