import os
from flask import Flask, jsonify, send_file, send_from_directory

from soilmove import update_all, EXCEL_PATH, KML_PATH

# backend/app.py
BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))           # .../backend
BASE_DIR = os.path.dirname(BACKEND_DIR)                            # .../
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")                  # .../frontend
STATIC_DIR = os.path.join(FRONTEND_DIR, "static")                  # .../frontend/static

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")

latest = {"updated": "尚未更新", "count": 0, "data": []}

@app.route("/healthz")
def healthz():
    return jsonify({"ok": True})

@app.route("/")
def index():
    # 若找不到 index.html，直接回傳清楚的錯誤訊息（避免只看到 500）
    index_path = os.path.join(FRONTEND_DIR, "index.html")
    if not os.path.exists(index_path):
        return jsonify({
            "error": "index.html not found",
            "expected": index_path,
            "cwd": os.getcwd(),
            "files_in_frontend": os.listdir(FRONTEND_DIR) if os.path.exists(FRONTEND_DIR) else None
        }), 500
    return send_from_directory(FRONTEND_DIR, "index.html")

@app.route("/api/update")
def api_update():
    global latest
    latest = update_all()
    return jsonify(latest)

@app.route("/api/data")
def api_data():
    return jsonify(latest)

@app.route("/download/excel")
def download_excel():
    return send_file(EXCEL_PATH, as_attachment=True, download_name="全台土資場清單_latest.xlsx")

@app.route("/download/kml")
def download_kml():
    return send_file(KML_PATH, as_attachment=True, download_name="全台土資場分佈圖_latest.kml")

if __name__ == "__main__":
    app.run(debug=True)
