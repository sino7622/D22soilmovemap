import os
from datetime import datetime

import pandas as pd
import requests
import simplekml
import urllib3
from urllib3.exceptions import InsecureRequestWarning

# 停用不安全請求警告（你現在的寫法是正確的）
urllib3.disable_warnings(InsecureRequestWarning)

# ====== Render / Linux 友善：固定輸出到 /tmp ======
OUT_DIR = os.getenv("SOILMOVE_OUT_DIR", "/tmp/soilmove")
os.makedirs(OUT_DIR, exist_ok=True)

# 讓 app.py 的 download 路由可以抓到「最新版」檔案
EXCEL_PATH = os.path.join(OUT_DIR, "全台土資場清單_latest.xlsx")
KML_PATH = os.path.join(OUT_DIR, "全台土資場分佈圖_latest.kml")


def _now_str():
    # 台灣時間顯示（Render 是 UTC；你若要 +8，可在這裡調）
    # 先用伺服器時間，避免引入額外套件
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _normalize_coords(row) -> pd.Series:
    """
    soilmove 來源資料 x/y 可能反轉：
    - 目標：lng(118~125), lat(20~26)
    - x/y 來源格式可能是字串/None
    """
    try:
        v1 = float(row.get("x") or 0)
        v2 = float(row.get("y") or 0)

        # v1 是經度
        if 118 < v1 < 125 and 20 < v2 < 26:
            return pd.Series([v1, v2, "原始正確"])

        # v2 是經度（反轉修正）
        if 118 < v2 < 125 and 20 < v1 < 26:
            return pd.Series([v2, v1, "已自動修正(X/Y反轉)"])

        # 只判斷經度範圍（若緯度不合理，仍標為異常）
        if 118 < v1 < 125:
            return pd.Series([v1, v2, "座標疑似異常(緯度不合理)"])
        if 118 < v2 < 125:
            return pd.Series([v2, v1, "座標疑似異常(緯度不合理, 已反轉)"])

        return pd.Series([0, 0, "座標異常"])
    except Exception:
        return pd.Series([0, 0, "轉換錯誤"])


def update_all() -> dict:
    """
    抓取全台土資場資料 → 清洗 → 產 Excel / KML → 回傳前端可用 JSON
    回傳格式需對齊 app.py：{"updated","count","data","excel_url","kml_url"}
    """
    url = "https://www.soilmove.tw/soilmove/dumpsiteGisQueryList"
    base_url = "https://www.soilmove.tw/soilmove/dumpsiteGisQuery"

    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
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
        # 不要丟例外，回空結果讓前端可顯示
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

    # 覆蓋寫入 latest
    df_excel.to_excel(EXCEL_PATH, index=False)

    # 5) KML（覆蓋寫入 latest）
    kml = simplekml.Kml()
    style = simplekml.Style()
    style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

    count = 0
    for _, row in df.iterrows():
        lng = float(row.get("lng") or 0)
        lat = float(row.get("lat") or 0)

        # 過濾台灣有效座標
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
        count += 1

    kml.save(KML_PATH)

    # 6) 回傳給前端的資料（直接用原欄位 + lng/lat）
    #    注意：確保是乾淨的 Python 型別（避免 numpy type 造成 jsonify 問題）
    out_cols = [
        "dumpname", "city", "typename", "controlId", "applydate",
        "remain", "maxbury", "area", "lng", "lat"
    ]
    out_cols = [c for c in out_cols if c in df.columns]
    df_out = df[out_cols].copy()

    # 只回傳有效座標（前端就不用再過濾）
    df_out = df_out[(df_out["lng"] > 0) & (df_out["lat"] > 0)]
    payload = df_out.to_dict(orient="records")

    return {
        "updated": _now_str(),
        "count": int(len(payload)),
        "data": payload,
        "excel_url": "/download/excel",
        "kml_url": "/download/kml",
    }
