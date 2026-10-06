"""smm-review の「審査書類一式ファイル出力」用 PDF 変換 API（Cloud Run）。

画面（Firebase Hosting: https://smm-review-tool.web.app）からは、Hosting の rewrite で
同じオリジンの /api/convert-sheets として呼ばれる。手元の開発用ページ（localhost）からの
直接呼び出しに備えて、許可したオリジンにだけ CORS を返す。

- POST /api/convert-sheets : multipart の file（.xls / .xlsx / .xlsm）→ PDF
    成功: 200 application/pdf、ヘッダー X-Pattern（A/B）と X-Conversion-Warnings（URLエンコードしたJSON配列）
    失敗: 4xx/5xx {"error": "メッセージ"}
- GET /healthz : 死活確認

見積ファイル（顧客情報を含む）は一時フォルダーで変換したらすぐ削除し、保存しない。
ログにもファイル名や内容は出さない。
"""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from urllib.parse import quote

from flask import Flask, Response, jsonify, request

MAX_MB = 30  # Cloud Run の HTTP/1 リクエスト上限（32MiB）より少し小さくする
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
CONVERT_TIMEOUT = int(os.environ.get("CONVERT_TIMEOUT", "100"))
ALLOWED_ORIGINS = {
    "https://smm-review-tool.web.app",
    "https://smm-review-tool.firebaseapp.com",
    # 手元での動作確認用（npm run serve / 起動.bat / file://）
    "http://localhost:8080", "http://127.0.0.1:8080",
    "http://localhost:8765", "http://127.0.0.1:8765",
    "null",
}
HERE = Path(__file__).resolve().parent

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger("converter")


@app.after_request
def cors(resp: Response) -> Response:
    origin = request.headers.get("Origin")
    if origin in ALLOWED_ORIGINS:
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Access-Control-Expose-Headers"] = "X-Pattern, X-Conversion-Warnings"
        resp.headers["Vary"] = "Origin"
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    return resp


@app.errorhandler(413)
def too_large(_):
    return jsonify(error=f"ファイルが{MAX_MB}MBを超えています。"), 413


@app.get("/healthz")
def healthz():
    return jsonify(ok=True)


@app.route("/api/convert-sheets", methods=["OPTIONS"])
def preflight():
    resp = Response(status=204)
    if request.headers.get("Origin") in ALLOWED_ORIGINS:
        resp.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        resp.headers["Access-Control-Max-Age"] = "600"
    return resp


@app.post("/api/convert-sheets")
def convert_sheets():
    part = request.files.get("file")
    if part is None or not part.filename:
        return jsonify(error="Excelファイルを選択してください。"), 400
    name = part.filename.replace("\\", "/").split("/")[-1]
    suffix = Path(name).suffix.lower()
    if suffix not in EXCEL_EXTENSIONS or name.startswith("~$"):
        return jsonify(error=".xls、.xlsx、.xlsmファイルを選択してください。"), 400

    tmp = Path(tempfile.mkdtemp(prefix="smm_"))
    try:
        source = tmp / f"input{suffix}"
        output = tmp / "output.pdf"
        part.save(source)
        try:
            proc = subprocess.run(
                [sys.executable, str(HERE / "convert.py"), str(source), str(output)],
                capture_output=True, text=True, timeout=CONVERT_TIMEOUT,
            )
        except subprocess.TimeoutExpired:
            log.warning("conversion timed out")
            return jsonify(error="変換に時間がかかりすぎたため中止しました。もう一度お試しください。"), 504
        try:
            result = json.loads(proc.stdout.strip().splitlines()[-1])
        except Exception:
            log.error("converter crashed: rc=%s stderr=%s", proc.returncode, proc.stderr[-2000:])
            return jsonify(error="PDFを作成できませんでした。"), 500
        if "error" in result:
            if result.get("kind") == "input":
                return jsonify(error=result["error"]), 422
            log.error("conversion failed: %s", result["error"])
            return jsonify(error=f"PDFを作成できませんでした: {result['error']}"), 500
        resp = Response(output.read_bytes(), mimetype="application/pdf")
        resp.headers["X-Pattern"] = result["pattern"]
        resp.headers["X-Conversion-Warnings"] = quote(json.dumps(result["warnings"], ensure_ascii=False), safe="")
        log.info("converted pattern=%s size=%d", result["pattern"], output.stat().st_size)
        return resp
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
