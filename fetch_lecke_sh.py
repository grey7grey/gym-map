"""
上海乐刻健身（非私教馆）门店地址列表采集脚本

数据源：高德地图 Web 服务 POI 搜索 API（与本项目 app.py / geocode_once.py 同一套 Key）
输出：
  - 乐刻上海地址列表.xlsx   （含品牌分类、区、地址、经纬度）
  - 乐刻上海非私教馆地址.csv （仅乐刻健身主品牌，可直接用）

用法（Key 走本地 .env，不进对话）：
  1) 在项目目录新建 .env，写入：  AMAP_KEY=你的高德Web服务Key
  2) 运行：  python fetch_lecke_sh.py

品牌分类规则（针对"乐刻"关键词返回的全部 POI）：
  乐刻健身(主品牌) : 名称含 "乐刻健身" 或 "乐刻运动健身"   -> 保留
  乐刻私教馆(FEELINGME): 名称含 "私教" 或 "FEELINGME"      -> 排除（用户明确不要）
  其他子品牌       : 闪电熊猫 / FitTribe飞踹 / YOGAPOD小瑜荚 / LOVEFITT拉飞 / Recore无限核子 -> 单独列出、不计入主列表
"""

import os
import re
import csv
import json
import time
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")
OUT_XLSX = os.path.join(HERE, "乐刻上海地址列表.xlsx")
OUT_CSV = os.path.join(HERE, "乐刻上海非私教馆地址.csv")

AMAP_TEXT_URL = "https://restapi.amap.com/v3/place/text"
RAW_CACHE = os.path.join(HERE, ".lecke_raw.json")  # 原始 POI 缓存，避免重复打 API

# 从 .env 读取 Key（避免 Key 出现在对话/代码里）
def load_key():
    key = os.getenv("AMAP_KEY")
    if key:
        return key
    try:
        with open(ENV_PATH, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("AMAP_KEY"):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    except FileNotFoundError:
        pass
    return None


# 子品牌 / 非门店 识别（均不计入"乐刻健身"主品牌）
PRIVATE_HINTS = ["私教", "FEELINGME"]                       # 私教馆
IRON_HINTS = ["铁馆", "LOVEFIT", "拉飞"]                    # LOVEFITT 专业铁馆（独立子品牌）
OTHER_SUB_HINTS = ["闪电熊猫", "FitTribe", "飞踹", "YOGAPOD",
                   "小瑜荚", "LOVEFITT", "Recore", "无限核子"]  # 其他子品牌
NON_STORE_HINTS = ["研训中心", "培训中心"]                  # 教练培训中心，非会员门店


def classify(name: str):
    """返回 (brand_type, keep_as_main)。

    判定顺序很重要：先排除私教馆，再排除 铁馆/其他子品牌/培训中心，
    最后凡名称含「乐刻健身/乐刻运动健身/乐刻运动/乐刻·健身」或仅含「乐刻」的，
    都算作 乐刻健身 主品牌（含 精品馆/升级馆/铁馆之外的常规馆）。
    """
    n = name
    if any(h.lower() in n.lower() for h in PRIVATE_HINTS):
        return "乐刻私教馆(FEELINGME)", False
    if any(h.lower() in n.lower() for h in IRON_HINTS):
        return "乐刻铁馆(LOVEFITT子品牌)", False
    if any(h.lower() in n.lower() for h in OTHER_SUB_HINTS):
        return "其他子品牌", False
    if any(h in n for h in NON_STORE_HINTS):
        return "乐刻研训中心(非门店)", False
    if (("乐刻健身" in n) or ("乐刻运动健身" in n) or ("乐刻运动" in n)
            or ("乐刻·健身" in n) or ("乐刻" in n and "健身" in n)):
        return "乐刻健身(主品牌)", True
    if "乐刻" in n:
        return "乐刻健身(主品牌)", True
    return "其他", False


def fetch_all(key: str, keyword: str, city: str = "上海"):
    """分页拉取某关键词下全部 POI，去重后返回列表。"""
    out = {}
    page = 1
    while True:
        params = {
            "key": key,
            "keywords": keyword,
            "city": city,
            "citylimit": "true",
            "offset": 25,
            "page": page,
            "extensions": "all",
        }
        r = requests.get(AMAP_TEXT_URL, params=params, timeout=15)
        data = r.json()
        if data.get("status") != "1":
            print(f"  [警告] 第{page}页返回异常: {data.get('info')} {data.get('infocode')}")
            break
        pois = data.get("pois", [])
        if not pois:
            break
        for p in pois:
            pid = p.get("id")
            if pid and pid not in out:
                out[pid] = p
        print(f"  [{keyword}] 第{page}页拿到 {len(pois)} 条，累计 {len(out)} 条")
        if len(pois) < 25:
            break
        if page >= 50:   # 安全阀
            break
        page += 1
        time.sleep(0.25)  # 限速，避免触发高德限流
    return list(out.values())


def main():
    key = load_key()
    if not key:
        raise SystemExit(
            "❌ 未找到高德 Key。请在项目目录新建 .env 写入 AMAP_KEY=你的高德Web服务Key，"
            "或先 setx AMAP_KEY \"你的Key\" 后重启终端。Key 不会出现在对话里。"
        )

    print("🔍 正在通过高德 POI 搜索枚举上海『乐刻』相关门店…")
    # 用较宽的关键词，把所有乐刻系门店都抓回来，再按名称过滤
    if os.path.exists(RAW_CACHE):
        print("  检测到本地缓存，直接复用原始 POI（如需重新拉取请删除 .lecke_raw.json）")
        with open(RAW_CACHE, encoding="utf-8") as f:
            pois = json.load(f)
    else:
        pois = fetch_all(key, "乐刻", "上海")
        with open(RAW_CACHE, "w", encoding="utf-8") as f:
            json.dump(pois, f, ensure_ascii=False)
        print(f"  已缓存 {len(pois)} 条原始 POI 到 {RAW_CACHE}")
    print(f"✅ 共获取去重后 {len(pois)} 条 POI")

    rows = []
    for p in pois:
        name = p.get("name", "")
        address = p.get("address", "")
        adname = p.get("adname", "")        # 区
        location = p.get("location", "")     # lng,lat
        btype, keep = classify(name)
        lng, lat = ("", "")
        if location:
            lng, lat = (location.split(",") + ["", ""])[:2]
        rows.append({
            "name": name,
            "brand_type": btype,
            "is_main": keep,
            "district": adname,
            "address": address,
            "lng": lng,
            "lat": lat,
        })

    # 主品牌（非私教馆）列表
    main_rows = [r for r in rows if r["is_main"]]
    print(f"🏋️ 乐刻健身主品牌（非私教馆）门店：{len(main_rows)} 家")
    print("（其余子品牌/私教馆已单独保留在 xlsx 中）")

    # ---- 导出 xlsx（全部 + 分类）----
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "全部乐刻系门店"
    hdr = ["门店名称", "品牌类型", "是否主品牌(非私教馆)", "区", "详细地址", "经度", "纬度"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in sorted(rows, key=lambda x: (x["brand_type"], x["district"], x["name"])):
        ws.append([r["name"], r["brand_type"], "是" if r["is_main"] else "否",
                   r["district"], r["address"], r["lng"], r["lat"]])
    wb.save(OUT_XLSX)
    print(f"💾 已写入：{OUT_XLSX}")

    # ---- 导出 csv（仅非私教馆主品牌，方便直接使用）----
    with open(OUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["门店名称", "区", "详细地址", "经度", "纬度"])
        for r in sorted(main_rows, key=lambda x: (x["district"], x["name"])):
            w.writerow([r["name"], r["district"], r["address"], r["lng"], r["lat"]])
    print(f"💾 已写入：{OUT_CSV}")


if __name__ == "__main__":
    main()
