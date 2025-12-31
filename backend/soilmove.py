import os
from datetime import datetime

import pandas as pd
import requests
import simplekml
import urllib3
from urllib3.exceptions import InsecureRequestWarning

# 停用不安全請求警告
urllib3.disable_warnings(InsecureRequestWarning)

# ====== Render / Linux 友善：固定輸出到 /tmp ======
OUT_DIR = os.getenv("SOILMOVE_OUT_DIR", "/tmp/soilmove")
os.makedirs(OUT_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(OUT_DIR, "全台土資場清單_latest.xlsx")
KML_PATH = os.path.join(OUT_DIR, "全台土資場分佈圖_latest.kml")


def _now_str():
    # 若 Render 有設 TZ=Asia/Taipei，這裡就會是台灣時間
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _normalize_coords(row) -> pd.Series:
    try:
        v1 = float(row.get("x") or 0)
        v2 = float(row.get("y") or 0)

        # v1 是經度
        if 118 < v1 < 125 and 20 < v2 < 26:
            return pd.Series([v1, v2, "原始正確"])

        # v2 是經度（反轉修正）
        if 118 < v2 < 125 and 20 < v1 < 26:
            return pd.Series([v2, v1, "已自動修正(X/Y反轉)"])

        # 只判斷經度範圍（緯度不合理）
        if 118 < v1 < 125:
            return pd.Series([v1, v2, "座標疑似異常(緯度不合理)"])
        if 118 < v2 < 125:
            return pd.Series([v2, v1, "座標疑似異常(緯度不合理, 已反轉)"])

        return pd.Series([0, 0, "座標異常"])
    except Exception:
        return pd.Series([0, 0, "轉換錯誤"])


def update_all() -> dict:
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

    try:
        # 1) 先 GET 一次（讓 session 正常）
        session.get(base_url, headers=headers, timeout=20, verify=False, allow_redirects=True)

        # 2) POST 拿資料
        r = session.post(url, headers=headers, data={"city": ""}, timeout=30, verify=False, allow_redirects=True)
        r.raise_for_status()

        try:
            data = r.json()
        except Exception:
            # 有時候對方回 HTML/非 JSON
            data = []

    except Exception:
        # 任何抓取錯誤：回空結果（前端不會掛）
        return {
            "updated": _now_str(),
            "count": 0,
            "data": [],
            "excel_url": "/download/excel",
            "kml_url": "/download/kml",
        }

    if not data:
        return {
            "updated": _now_str(),
            "count": 0,
            "data": [],
            "excel_url": "/download/excel",
            "kml_url": "/download/kml",
        }

    df = pd.DataFrame(data)

    # 3) 座標清洗
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

    # 覆蓋寫入 latest（指定 engine 更穩）
    df_excel.to_excel(EXCEL_PATH, index=False, engine="openpyxl")

    # 5) KML（覆蓋寫入 latest）
    kml = simplekml.Kml()
    style = simplekml.Style()
    style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

    for _, row in df.iterrows():
        lng = float(row.get("lng") or 0)
        lat = float(row.get("lat") or 0)

        if not (118 < lng < 125 and 20 < lat < 26):
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
        )

        pnt.description = description
        pnt.style = style

    kml.save(KML_PATH)

    # 6) 回傳前端資料（保持你原版欄位）
    out_cols = [
        "dumpname", "city", "typename", "controlId", "applydate",
        "remain", "maxbury", "area", "lng", "lat"
    ]
    out_cols = [c for c in out_cols if c in df.columns]
    df_out = df[out_cols].copy()

    # 只回傳有效座標
    df_out = df_out[(df_out["lng"] > 0) & (df_out["lat"] > 0)]

    # ✅ NaN → 空字串，避免前端出現 nan / JSON 非標準
    df_out = df_out.where(pd.notnull(df_out), "")

    payload = df_out.to_dict(orient="records")

    return {
        "updated": _now_str(),
        "count": int(len(payload)),
        "data": payload,
        "excel_url": "/download/excel",
        "kml_url": "/download/kml",
    }
