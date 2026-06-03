"""
Flask web app — MathType to LaTeX Converter
Upload .docx → convert OLE equations to LaTeX → download result
"""

import os
import sys
import tempfile
import shutil

from flask import (
    Flask,
    flash,
    redirect,
    render_template_string,
    request,
    url_for,
    send_file,
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from convert_mathtype_to_latex import (
    process_docx,
    remove_paragraphs_containing,
    format_mcq_answers_in_docx,
    remove_level_tags_in_docx,
    convert_tables_to_text,
)

app = Flask(__name__)
app.secret_key = os.urandom(24).hex()

UPLOAD_FOLDER = tempfile.mkdtemp(prefix="mathtype_uploads_")
OUTPUT_FOLDER = tempfile.mkdtemp(prefix="mathtype_outputs_")

HTML = """\
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MathType → LaTeX Converter</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', system-ui, sans-serif; background: #f5f7fa; min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 20px; }
.card { background: #fff; border-radius: 16px; box-shadow: 0 4px 24px rgba(0,0,0,.08); padding: 36px 40px; max-width: 640px; width: 100%; }
h1 { font-size: 22px; margin-bottom: 4px; color: #1a1a2e; }
p.sub { color: #666; font-size: 14px; margin-bottom: 24px; }
label { display: block; font-weight: 600; font-size: 14px; margin: 14px 0 4px; color: #333; }
input[type="file"] { width: 100%; padding: 8px 0; }
input[type="text"] { width: 100%; padding: 10px 12px; border: 1px solid #d0d5dd; border-radius: 8px; font-size: 14px; }
.checkbox-group { margin: 12px 0; }
.checkbox-group label { display: inline-flex; align-items: center; gap: 6px; font-weight: 400; margin: 0 16px 6px 0; cursor: pointer; }
.checkbox-group input[type="checkbox"] { width: 16px; height: 16px; }
button { background: #1a1a2e; color: #fff; border: none; padding: 12px 28px; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; margin-top: 18px; width: 100%; transition: background .2s; }
button:hover { background: #16213e; }
button:disabled { opacity: .6; cursor: not-allowed; }
.alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; }
.alert-success { background: #d1fae5; color: #065f46; }
.alert-error { background: #fee2e2; color: #991b1b; }
hr { border: none; border-top: 1px solid #e5e7eb; margin: 18px 0; }
</style>
</head>
<body>
<div class="card">
  <h1>📐 MathType → LaTeX</h1>
  <p class="sub">Upload file .docx, chuyển công thức MathType thành LaTeX</p>

  {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}
      {% for category, msg in messages %}
        <div class="alert alert-{{ category }}">{{ msg }}</div>
      {% endfor %}
    {% endif %}
  {% endwith %}

  <form method="post" action="/" enctype="multipart/form-data">
    <label for="file">Chọn file .docx</label>
    <input type="file" name="file" accept=".docx" required>

    <label for="keyword">Xóa dòng chứa keyword (tùy chọn)</label>
    <input type="text" name="keyword" placeholder="Ví dụ: Hướng dẫn">

    <div class="checkbox-group">
      <label><input type="checkbox" name="delete_mucdo"> Xóa [Mức độ 1], [Mức độ 2]...</label>
      <label><input type="checkbox" name="format_mcq"> Format MCQ (B. C. D. xuống dòng)</label>
      <label><input type="checkbox" name="convert_tables"> Chuyển bảng → văn bản</label>
    </div>

    <button type="submit">🔄 Convert</button>
  </form>

  <hr>
  <p style="font-size:13px; color:#888; text-align:center;">Chạy trên server, không cần Word hay MathType</p>
</div>
</body>
</html>
"""


def clean_temp():
    for d in (UPLOAD_FOLDER, OUTPUT_FOLDER):
        for f in os.listdir(d):
            try:
                os.remove(os.path.join(d, f))
            except Exception:
                pass


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML)

    if "file" not in request.files:
        flash("Vui lòng chọn file.", "error")
        return render_template_string(HTML)

    f = request.files["file"]
    if not f or not f.filename.endswith(".docx"):
        flash("Chỉ chấp nhận file .docx.", "error")
        return render_template_string(HTML)

    in_path = os.path.join(UPLOAD_FOLDER, f.filename)
    out_name = f.filename.replace(".docx", "-latex.docx")
    out_path = os.path.join(OUTPUT_FOLDER, out_name)
    f.save(in_path)

    keyword = request.form.get("keyword", "").strip()
    delete_mucdo = request.form.get("delete_mucdo") == "on"
    format_mcq = request.form.get("format_mcq") == "on"
    convert_tables = request.form.get("convert_tables") == "on"

    try:
        import io
        import contextlib

        tmp = in_path
        temps = []

        if convert_tables:
            tmp2 = os.path.join(OUTPUT_FOLDER, "_tmp_tables.docx")
            temps.append(tmp2)
            convert_tables_to_text(tmp, tmp2)
            tmp = tmp2

        if keyword:
            tmp2 = os.path.join(OUTPUT_FOLDER, "_tmp_kw.docx")
            temps.append(tmp2)
            remove_paragraphs_containing(tmp, tmp2, keyword)
            if tmp != in_path:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
            tmp = tmp2

        if delete_mucdo:
            tmp2 = os.path.join(OUTPUT_FOLDER, "_tmp_level.docx")
            temps.append(tmp2)
            remove_level_tags_in_docx(tmp, tmp2)
            if tmp != in_path:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
            tmp = tmp2

        if format_mcq:
            tmp2 = os.path.join(OUTPUT_FOLDER, "_tmp_mcq.docx")
            temps.append(tmp2)
            format_mcq_answers_in_docx(tmp, tmp2)
            if tmp != in_path:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
            tmp = tmp2

        process_docx(tmp, out_path)

        for t in temps:
            try:
                if os.path.exists(t):
                    os.remove(t)
            except Exception:
                pass

        flash("✅ Conversion thành công! Nhấn nút bên dưới để tải file.", "success")

        return render_template_string(
            HTML
            + '<a href="/download/{{name}}" style="display:block;text-align:center;padding:10px 0;background:#059669;color:#fff;border-radius:8px;text-decoration:none;font-weight:600;">📥 Tải file: {{name}}</a>'.replace(
                "{{name}}", out_name
            )
        )

    except Exception as e:
        flash(f"Lỗi: {e}", "error")
        return render_template_string(HTML)


@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        flash("File không tồn tại hoặc đã bị xóa.", "error")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
