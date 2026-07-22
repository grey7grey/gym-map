# 工会合作健身房地图

基于 **Streamlit + 高德地图 + folium** 的上海工会合作健身房可视化地图。输入「我的位置」（家 / 公司 / 临时地址），自动计算直线距离，侧边栏列出离你最近的门店，并在地图上交互式展示全部门店位置。

> 坐标已预先解析并「烘焙」进 `store_coords.py`，**打开应用即显示全部门店、搜索、Top5，无需任何网络调用**。仅当你手动输入「临时文字地址」定位时，才会调用高德地理编码 API（需要 Key）。

## 功能

- **一键定位**：🏠 家（德平路25弄）/ 🏢 公司（银城路167号）按钮，坐标已内置于 `app.py`，点击即定位、零延迟。
- **临时地址**：外出不在家/公司时，输入任意文字地址点「定位这个地址」，实时调用高德解析为坐标。
- **搜索 / 筛选**：搜索框同时匹配「门店名称」和「提供服务」两列；另提供 ⚡24/7、⚡游泳 两个快捷筛选。
- **离我最近 TOP5**：按当前位置由近到远列出最近的 5 家（店名 + 距离 + 地址）。
- **交互式地图**：folium 渲染，门店按距离聚类（MarkerCluster），Top5 用红色数字徽标高亮；有筛选时地图自动缩放聚焦命中门店。

## 目录结构

```
.
├── app.py                      # 主应用（Streamlit 入口）
├── store_coords.py             # 烘焙的门店坐标（295 家，由 geocode_once.py 生成）
├── 公司工会合作健身房列表.xlsx  # 唯一数据源（门店 / 服务 / 地址）
├── geocode_once.py             # 一次性把 Excel 地址解析为坐标，生成 store_coords.py
├── run_geocode.bat             #  Windows 下一键跑 geocode_once.py
├── .env.example                # 高德 Key 模板（复制为 .env 后填真实 Key）
├── .gitignore                  # 排除 .env / .workbuddy / 缓存等
└── README.md
```

## 环境准备

需要 Python 3.10+，依赖：

```bash
pip install streamlit pandas openpyxl requests folium streamlit-folium python-dotenv
```

## 运行

```bash
# 1. 配置高德 Key（仅「临时文字地址」定位需要；门店坐标已烘焙，不配也能跑地图/搜索/Top5）
cp .env.example .env
#   然后编辑 .env，把 AMAP_KEY 改成你的高德 Web 服务 Key
#   申请地址：https://lbs.amap.com/

# 2. 启动
streamlit run app.py --server.port 8501
```

浏览器打开 http://localhost:8501 即可。

> 默认开页定位 = 家。改完 `store_coords.py` / Excel 后按 F5 即生效（应用每次刷新动态重载坐标文件，无需手动重启进程）。

## 常见问题

- **地图/搜索/Top5 不显示门店？** 多半是旧的 streamlit 进程堆积导致缓存未刷新。杀掉 8501 端口上的旧进程后重新 `streamlit run` 即可。
- **只想改门店数据？** 编辑 `公司工会合作健身房列表.xlsx`，再跑一遍 `geocode_once.py` 重新生成 `store_coords.py`（会调用高德解析新地址，需要有效的 `AMAP_KEY`）。
- **家/公司地址换了？** 直接改 `app.py` 顶部的 `HOME_ADDR` / `HOME_COORD` 与 `WORK_ADDR` / `WORK_COORD` 常量即可（当前为固定值，已烘焙，启动不再联网获取）。

## 版本控制

本项目已 `git init`（main 分支）。密钥 `.env` 与 `.workbuddy/` 已被 `.gitignore` 排除，不会提交。
