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
├── requirements.txt            # 云端部署依赖清单（Streamlit Cloud 自动安装）
├── .gitignore                  # 排除 .env / .workbuddy / 缓存等
└── README.md
```

## 环境准备

需要 Python 3.10+。本项目配套的 Python 已经装好，并且 **`streamlit` 命令已经加进你的用户级 PATH**——所以平常直接输入 `streamlit` 就能用，不用写一长串路径。

如果你在别的电脑 / 终端里输入 `streamlit` 提示「不是内部或外部命令」，说明那台机器没配，可选下面任一方式：
- 用全路径调用：`C:\Users\prote\.workbuddy\binaries\python\envs\default\Scripts\python.exe -m streamlit run app.py --server.port 8501`
- 或自行安装依赖：`pip install streamlit pandas openpyxl requests folium streamlit-folium python-dotenv`

## 配置高德 Key（可选）

只有当你想用「临时文字地址」定位时才需要 Key；门店坐标已烘焙，**不配置也能正常看地图 / 搜索 / TOP5**。

```bash
cp .env.example .env
# 然后用记事本打开 .env，把 AMAP_KEY 改成你的高德 Web 服务 Key
# 申请地址：https://lbs.amap.com/
```

## 怎么打开这个网页（日常）

应用地址固定为 **http://localhost:8501**，不会变。

- **服务正在运行时**：直接打开浏览器，输入上面的地址即可。服务在后台运行，关掉命令行窗口也不会停。
- **打不开（多半是电脑重启过、服务已停止）**：按下面「启动 / 重启服务」操作。

## 启动 / 重启服务（一步一步来，适合零基础）

1. 打开命令提示符：按键盘 `Win + R`，输入 `cmd`，回车。会弹出一个黑色窗口（这就是「命令行」）。
2. 进入项目目录：在黑窗口里**粘贴**下面这行，按回车：
   ```
   cd /d "D:\ai空间\健身房"
   ```
3. 启动应用：再**粘贴**下面这行，按回车：
   ```
   streamlit run app.py --server.port 8501
   ```
4. 看到 `You can view your Streamlit app in your browser.` 和 `Local URL: http://localhost:8501` 就成功了。
5. 打开浏览器，访问 http://localhost:8501 。
6. **保持这个黑色窗口开着**（可以最小化，但不要关闭）。关掉它，服务就停了。

> 如果第 3 步提示「端口 8501 被占用」，说明旧的服务还在跑——不用再起一个，直接刷新浏览器 http://localhost:8501 即可。

## 验证全局命令是否生效

新开一个 cmd，输入 `streamlit --version`，能显示版本号（例如 `Streamlit, version 1.x.x`）就说明全局 `streamlit` 命令已配置好。

> 默认开页定位 = 家。改完 `store_coords.py` / Excel 后按 **F5** 即生效（应用每次刷新会动态重载坐标文件，无需手动重启进程）。

## 为什么命令变短了

本项目用的 Python 装在隔离环境（C 盘的 WorkBuddy 目录），原本必须写一长串完整路径才能调用。现在已经把 `streamlit` 加进你的用户 PATH，所以新开的命令行里直接输入 `streamlit` 就能用，不用再写长路径。

## 常见问题

- **地图/搜索/Top5 不显示门店？** 多半是旧的 streamlit 进程堆积导致缓存未刷新。杀掉 8501 端口上的旧进程后重新 `streamlit run` 即可。
- **只想改门店数据？** 编辑 `公司工会合作健身房列表.xlsx`，再跑一遍 `geocode_once.py` 重新生成 `store_coords.py`（会调用高德解析新地址，需要有效的 `AMAP_KEY`）。
- **家/公司地址换了？** 直接改 `app.py` 顶部的 `HOME_ADDR` / `HOME_COORD` 与 `WORK_ADDR` / `WORK_COORD` 常量即可（当前为固定值，已烘焙，启动不再联网获取）。

## 部署到 Streamlit Community Cloud（公网访问）

把应用发布成一个任何人都能打开的公网网址。以下流程假设你已经有一个 GitHub 账号。

### 1. 准备（本地，本项目已完成）

- `app.py` 在仓库根目录（Streamlit Cloud 默认入口，无需额外配置）。
- `requirements.txt` 已就绪，云端会自动安装依赖。
- 坐标已烘焙进 `store_coords.py`，**云端启动即显示门店，无需任何网络调用**。
- `.env` / `.cache/` / 虚拟环境 已被 `.gitignore` 排除，不会上传。

### 2. 初始化并提交 Git（若还没做）

```bash
cd "D:\ai空间\健身房"
git init            # 本项目已执行过，可跳过
git add .
git commit -m "健身房地图应用"
```

### 3. 推送到 GitHub

1. 在 GitHub 上新建一个**公开（Public）**仓库（例如 `gym-map`），**不要**勾选自动生成 README / .gitignore。
2. 本地关联并推送：

```bash
git remote add origin https://github.com/你的用户名/gym-map.git
git branch -M main
git push -u origin main
```

> 推送前可用 `git status` 确认没有把 `.env` 提交进去（被 `.gitignore` 排除就不会出现）。

### 4. 在 Streamlit Cloud 部署

1. 打开 https://share.streamlit.io/ ，用 **GitHub 账号登录**（授权一次）。
2. 点 **New app** → 选择刚推送的仓库 `gym-map`、分支 `main`、入口文件 `app.py` → **Deploy**。
3. 在部署设置里填入 **AMAP_KEY**（在 app 的 **Settings → Secrets** 面板，粘贴 TOML 格式）：
   ```toml
   AMAP_KEY = "你的高德Web服务Key"
   ```
   应用代码读取顺序：优先 `st.secrets["AMAP_KEY"]`（云端），本地则回退到 `.env`。**不要把 Key 写进代码或提交到 Git。**
4. 部署完成后，Streamlit Cloud 会给出一个形如 `https://xxxx.streamlit.app` 的公网网址，分享即可。

> 注意：Streamlit Cloud 默认运行在海外节点。高德地理编码 API 是公网 HTTPS，可正常访问；门店坐标已烘焙，地图本身不依赖网络。

## 版本控制

本项目已 `git init`（main 分支）。密钥 `.env` 与 `.workbuddy/` 已被 `.gitignore` 排除，不会提交。
