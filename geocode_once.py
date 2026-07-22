"""
一次性把门店地址解析成坐标，写入 store_coords.py，供 app.py 直接读取（不再依赖运行时 Key / 联网）。

用法（Key 走环境变量，不进对话）：
    1) 在命令提示符执行：  setx AMAP_KEY "你的高德Web服务Key"
    2) 关闭并重新打开命令提示符（让环境变量生效）
    3) 执行本脚本：        python geocode_once.py
    4) 生成的 store_coords.py 即被 app.py 自动读取

说明：免费 Key 有 QPS 限制，脚本已加 0.3s 间隔 + 指数退避重试，确保所有地址都能解析。
"""
import os
import sys
import time
import re
import pandas as pd
import requests
from dotenv import load_dotenv

# 行缓冲：print 立刻显示（避免看上去像卡住）
try:
    sys.stdout.reconfigure(line_buffering=True)
except Exception:
    pass

load_dotenv()  # 自动读取同目录 .env 里的 AMAP_KEY（不上 Git）

HERE = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(HERE, "公司工会合作健身房列表.xlsx")
OUT = os.path.join(HERE, "store_coords.py")
URL = "https://restapi.amap.com/v3/geocode/geo"

API_KEY = os.getenv("AMAP_KEY")
if not API_KEY:
    raise SystemExit("❌ 未找到环境变量 AMAP_KEY，请先在 .env 里配置 AMAP_KEY=你的Key 后重试。")
print(f"✅ 已从 .env 加载 AMAP_KEY（长度 {len(API_KEY)}）", flush=True)


def load_stores():
    df = pd.read_excel(EXCEL)
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
    mask = df["address"].notna() & ~df["address"].astype(str).str.contains(
        "第三方|详见|全城|所有门店", na=False)
    return df[mask].reset_index(drop=True)


def _address_variants(addr: str):
    """生成地理编码候选地址：原样 → 去括号内容 → 削掉楼层/铺位后缀。"""
    variants = [addr]
    s = re.sub(r"[（(][^（）()]*[）)]", "", addr).strip()
    if s and s != addr:
        variants.append(s)
    m = re.match(r"^(.*?(?:路|道|街|号)[^\d]*\d+号?)", s)
    if m and m.group(1).strip() != addr:
        variants.append(m.group(1).strip())
    seen, out = set(), []
    for v in variants:
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def geocode(address: str, city: str = "上海"):
    last_err = ""
    for variant in _address_variants(address):
        for city_arg in (city, ""):
            params = {"address": variant, "key": API_KEY}
            if city_arg:
                params["city"] = city_arg
            for attempt in range(5):  # 退避重试应对 QPS 限流
                try:
                    r = requests.get(URL, params=params, timeout=10)
                    d = r.json()
                    if d.get("status") == "1" and d.get("geocodes"):
                        lng, lat = d["geocodes"][0]["location"].split(",")
                        return float(lng), float(lat)
                    last_err = d.get("info", "")
                    if last_err in ("QPS_EXCEED", "INVALID_USER_KEY", "USERKEY_PLAT_NOMATCH"):
                        time.sleep(0.4 * (2 ** attempt))
                        continue
                    break  # 该候选明确失败，换下一个候选
                except Exception:
                    time.sleep(0.4 * (2 ** attempt))
                    continue
    return None


def main():
    df = load_stores()
    total = len(df)
    print(f"📋 共 {total} 家待解析（带 0.3s 限速，预计 {total * 0.5 // 60 + 1} 分钟）", flush=True)
    print("─" * 50, flush=True)
    coords = {}
    failed = []
    t0 = time.time()
    for i, (_, row) in enumerate(df.iterrows(), 1):
        c = geocode(row["address"])
        if c:
            coords[row["name"]] = {
                "lng": c[0], "lat": c[1], "address": row["address"]}
        else:
            failed.append(row["name"])
        # 每 5 家打一行进度，最后一家必打
        if i % 5 == 0 or i == total:
            elapsed = time.time() - t0
            rate = i / elapsed if elapsed > 0 else 0
            eta = (total - i) / rate if rate > 0 else 0
            print(
                f"  [{i:>3}/{total}] ok={len(coords):>3}  fail={len(failed):>2}  "
                f"已用 {elapsed:5.0f}s  预计还 {eta:5.0f}s",
                flush=True,
            )
        time.sleep(0.3)  # 限速 ~3 次/秒，避免触发限流

    with open(OUT, "w", encoding="utf-8") as f:
        f.write("# 自动生成：门店坐标（高德地理编码），由 geocode_once.py 生成\n")
        f.write("# 坐标格式 {门店名: {lng, lat, address}}\n")
        f.write("STORE_COORDS = ")
        f.write(repr(coords))
        f.write("\n")

    print("─" * 50, flush=True)
    print(f"✅ 成功解析 {len(coords)} 家，失败 {len(failed)} 家", flush=True)
    if failed:
        print("失败门店：" + ", ".join(failed), flush=True)
    print(f"已写入：{OUT}", flush=True)


if __name__ == "__main__":
    main()
