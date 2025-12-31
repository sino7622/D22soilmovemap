from flask import Flask, jsonify, send_file, send_from_directory
import os
from soilmove import update_all, EXCEL_PATH, KML_PATH

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")
STATIC_DIR = os.path.join(FRONTEND_DIR, "static")

app = Flask(
    __name__,
    static_folder=STATIC_DIR,
    static_url_path="/static"
)

latest = {"updated": "尚未更新", "count": 0, "data": []}

# ===== 首頁 =====
@app.route("/")
def index():
    return send_from_directory(FRONTEND_DIR, "index.html")

# ===== 健康檢查（Render 會用）=====
@app.route("/healthz")
def healthz():
    return jsonify({"status": "ok"})

# ===== API =====
@app.route("/api/update")
def api_update():
    global latest
    latest = update_all()
    return jsonify(latest)

@app.route("/api/data")
def api_data():
    return jsonify(latest)

# ===== 下載 =====
@app.route("/download/excel")
def download_excel():
    return send_file(EXCEL_PATH, as_attachment=True)

@app.route("/download/kml")
def download_kml():
    return send_file(KML_PATH, as_attachment=True)
