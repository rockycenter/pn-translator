#!/usr/bin/env python3
"""
料号翻译 Web 工具 — Flask 应用
上传周报，一键翻译并下载（Mapping 表已内置）。
"""

import os
import tempfile
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify
from translate import translate_report

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB
UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "pn_translator"
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)

# 内置 Mapping 文件路径
BUILTIN_MAPPING = Path(__file__).parent / "mapping.xlsx"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/translate", methods=["POST"])
def translate():
    report = request.files.get("report")
    if not report:
        return jsonify({"error": "请上传周报文件"}), 400

    report_path = UPLOAD_FOLDER / report.filename
    output_path = UPLOAD_FOLDER / f"translated_{report.filename}"
    report.save(report_path)

    try:
        stats = translate_report(
            str(report_path),
            str(BUILTIN_MAPPING),
            str(output_path),
        )
    except Exception as e:
        return jsonify({"error": f"翻译失败: {str(e)}"}), 500

    total_translated = sum(s.get("translated", 0) for s in stats.values())
    total_unchanged = sum(s.get("unchanged", 0) for s in stats.values())
    total_not_found = sum(s.get("not_found", 0) for s in stats.values())

    return jsonify({
        "success": True,
        "filename": f"translated_{report.filename}",
        "stats": {
            "translated": total_translated,
            "unchanged": total_unchanged,
            "not_found": total_not_found,
            "by_sheet": {k: v for k, v in stats.items()},
        },
    })


@app.route("/download/<filename>")
def download(filename):
    file_path = UPLOAD_FOLDER / filename
    if not file_path.exists():
        return "File not found", 404
    return send_file(file_path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    print("\n" + "=" * 56)
    print("  料号翻译工具已启动！")
    print("  Mapping 已内置，只需上传周报即可")
    print("  打开浏览器访问: http://127.0.0.1:5050")
    print("=" * 56 + "\n")
    app.run(debug=True, host="127.0.0.1", port=5050)
