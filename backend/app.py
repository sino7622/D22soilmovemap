import os
from flask import Flask, render_template, jsonify, send_file
from flask_cors import CORS

from services.soilmove import update_all, EXCEL_PATH, KML_PATH

app = Flask(__name__)
CORS(app)  # ✅ 允許跨網域（GitHub Pages / 其他前端網域呼叫 API）

# 啟動時先給一個預設狀態（尚未更新）
latest = {"updated": "尚未更新", "count": 0, "data": []}


@app.get("/healthz")
def healthz():
    """Render 免費方案常睡眠：用 GET 先喚醒服務。"""
    return jsonify({"ok": True})


@app.get("/")
def index():
    # 仍保留原本頁面（如果你後續要「前後端分離」，也可以改成只回 API）
    return render_template("index.html", info=latest)


@app.get("/api/update")
def api_update():
    """抓取最新資料並輸出 Excel/KML，同時回傳前端可用的點位資料。"""
    global latest
    latest = update_all()
    return jsonify(latest)


@app.get("/api/data")
def api_data():
    """回傳目前快取資料（不觸發更新）。"""
    return jsonify(latest)


@app.get("/download/excel")
def download_excel():
    # ✅ 避免檔案暫存消失造成 500
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        return jsonify({"error": "excel not ready"}), 404

    return send_file(
        EXCEL_PATH,
        as_attachment=True,
        download_name="全台土資場清單_latest.xlsx"
    )


@app.get("/download/kml")
def download_kml():
    # ✅ 避免檔案暫存消失造成 500
    if not KML_PATH or not os.path.exists(KML_PATH):
        return jsonify({"error": "kml not ready"}), 404

    return send_file(
        KML_PATH,
        as_attachment=True,
        download_name="全台土資場分佈圖_latest.kml"
    )


# ✅ Render 會用 gunicorn 啟動：不要 app.run()
