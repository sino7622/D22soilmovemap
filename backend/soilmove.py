# backend/soilmove.py
import os
from datetime import datetime

import pandas as pd
import requests
import simplekml
import urllib3
from urllib3.exceptions import InsecureRequestWarning

from pyproj import Transformer

# 停用不安全請求警告
urllib3.disable_warnings(InsecureRequestWarning)

# ====== Render / Linux 友善：固定輸出到 /tmp ======
OUT_DIR = os.getenv("SOILMOVE_OUT_DIR", "/tmp/soilmove")
os.makedirs(OUT_DIR, exist_ok=True)

# 讓 app.py 的 download 路由可以抓到「最新版」檔案
EXCEL_PATH = os.path.join(OUT_DIR, "全台土資場清單_latest.xlsx")
KML_PATH = os.path.join(OUT_DIR, "全台土資場分佈圖_latest.kml")

# TWD97 / TM2(121) -> WGS84
# 若資料是 EPSG:3826（常見台灣 TM2），會轉成經緯度
_TWD97_TO_WGS84 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)


def _now_str() -> str:
    # 若你已在 Render 設 TZ=Asia/Taipei，這裡會自動是台灣時間
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _looks_like_lnglat(lng: float, lat: float) -> bool:
    # 粗範圍：台灣附近
    return (118.0 < lng < 125.0) and (20.0 < lat < 26.5)


def _looks_like_twd97_tm2(x: float, y: float) -> bool:
    # 台灣 TM2(121) 常見公尺座標粗範圍
    # x: 約 100000~400000, y: 約 2000000~3200000
    return (100000.0 <= x <= 400000.0) and (2000000.0 <= y <= 3200000.0)


def _normalize_coords(row) -> pd.Series:
    """
    來源 x/y 可能是：
      (A) 已是經緯度 (lng/lat)
      (B) 經緯度但 x/y 反了
      (C) TWD97 TM2(3826) 公尺座標
      (D) TWD97 TM2 但 x/y 反了
    回傳：lng, lat, coord_status
    """
    try:
        v1 = float(row.get("x") or 0)
        v2 = float(row.get("y") or 0)

        # (A) 已是經緯度
        if _looks_like_lnglat(v1, v2):
            return pd.Series([v1, v2, "原始經緯度"])

        # (B) 反轉後是經緯度
        if _looks_like_lnglat(v2, v1):
            return pd.Series([v2, v1, "已修正(X/Y反轉→經緯度)"])

        # (C) 看起來是 TWD97 TM2 → 轉 WGS84
        if _looks_like_twd97_tm2(v1, v2):
            lng, lat = _TWD97_TO_WGS84.transform(v1, v2)
            if _looks_like_lnglat(lng, lat):
                return pd.Series([lng, lat, "TWD97(TM2 3826)→WGS84 轉換"])

        # (D) TWD97 但 x/y 反了
        if _looks_like_twd97_tm2(v2, v1):
            lng, lat = _TWD97_TO_WGS84.transform(v2, v1)
            if _looks_like_lnglat(lng, lat):
                return pd.Series([lng, lat, "已修正(X/Y反轉)+TWD97→WGS84 轉換"])

        return pd.Series([0.0, 0.0, "座標異常/未能辨識座標系統"])
    except Exception:
        return pd.Series([0.0, 0.0, "轉換錯誤"])


def update_all() -> dict:
    """
    抓取全台土資場資料 → 清洗 → 產 Excel / KML → 回傳前端可用 JSON
    回傳格式對齊 app.py：
      {"updated","count","data","excel_url","kml_url"}
    """
    url = "https://www.soilmove.tw/soilmove/dumpsiteGisQueryList"
    base_url = "https://www.soilmove.tw/soilmove/dumpsiteGisQuery"

    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "X-Requested-With": "XMLHttpRequest",
        "Referer": base_url,
    }

    session = requests.Session()

    # 1) 先 GET 一次，讓對方站點 session 正常
    session.get(base_url, headers=headers, timeout=20, verify=False)

    # 2) POST 拿資料
    r = session.post(url, headers=headers, data={"city": ""}, timeout=30, verify=False)
    r.raise_for_status()
    data = r.json()

    if not data:
        return {
            "updated": _now_str(),
            "count": 0,
            "data": [],
            "excel_url": "/download/excel",
            "kml_url": "/download/kml",
        }

    df = pd.DataFrame(data)

    # （可選）印出 sample 看看來源座標長什麼樣（在 Render logs）
    # print("sample x,y:", df[["x", "y"]].head(5).to_dict(orient="records"))

    # 3) 座標清洗：產出 lng/lat
    df[["lng", "lat", "coord_status"]] = df.apply(_normalize_coords, axis=1)

    # 4) Excel（中文欄位 + 固定輸出到 latest）
    column_mapping = {
        "dumpname": "名稱",
        "lng": "經度",
        "lat": "緯度",
        "city": "縣市",
        "remain": "B1~B7剩餘填埋量",
        "coord_status": "轉換狀態",
        "typename": "類型",
        "id": "ID",
        "controlId": "流向編號",
        "x": "原始X",
        "y": "原始Y",
        "area": "面積",
        "maxbury": "B1~B7核准填埋量",
        "applydate": "申報日期",
    }

    keep_cols = [c for c in column_mapping.keys() if c in df.columns]
    df_excel = df[keep_cols].rename(columns=column_mapping)
    df_excel.to_excel(EXCEL_PATH, index=False)

    # 5) KML（覆蓋寫入 latest）
    kml = simplekml.Kml()
    style = simplekml.Style()
    style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

    count_kml = 0
    for _, row in df.iterrows():
        lng = float(row.get("lng") or 0)
        lat = float(row.get("lat") or 0)

        # 過濾有效經緯度（台灣範圍）
        if not _looks_like_lnglat(lng, lat):
            continue

        name = str(row.get("dumpname") or "未命名")
        pnt = kml.newpoint(name=name, coords=[(lng, lat)])

        description = (
            f"<b>縣市：</b> {row.get('city', '')}<br/>"
            f"<b>類型：</b> {row.get('typename', '')}<br/>"
            f"<b>流向編號：</b> {row.get('controlId', '')}<br/>"
            f"<b>申報日期：</b> {row.get('applydate', '')}<br/>"
            f"<hr/>"
            f"<b>B1~B7 剩餘填埋量：</b> {row.get('remain', '')} ㎥<br/>"
            f"<b>B1~B7 核准填埋量：</b> {row.get('maxbury', '')} ㎥<br/>"
            f"<b>面積：</b> {row.get('area', '')} 公頃<br/>"
            f"<hr/>"
            f"<b>經度：</b> {lng}<br/>"
            f"<b>緯度：</b> {lat}<br/>"
            f"<b>座標狀態：</b> {row.get('coord_status', '')}<br/>"
        )

        pnt.description = description
        pnt.style = style
        count_kml += 1

    kml.save(KML_PATH)

    # 6) 回傳給前端的資料（保留原欄位 + lng/lat + 其他 popup 欄位）
    out_cols = [
        "dumpname", "city", "typename", "controlId", "applydate",
        "remain", "maxbury", "area", "lng", "lat"
    ]
    # 確保欄位都存在（缺的補空）
    for c in out_cols:
        if c not in df.columns:
            df[c] = ""

    df_out = df[out_cols].copy()

    # 只回傳有效點
    df_out = df_out[(df_out["lng"].astype(float) > 0) & (df_out["lat"].astype(float) > 0)]
    payload = df_out.to_dict(orient="records")

    return {
        "updated": _now_str(),
        "count": int(len(payload)),
        "data": payload,
        "excel_url": "/download/excel",
        "kml_url": "/download/kml",
    }
