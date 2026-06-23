from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile

from process import process_report

app = Flask(__name__)
CORS(app)

# 定义本地模板文件夹路径
TEMPLATE_BASE = os.path.join(os.path.dirname(__file__), "templates")

@app.route("/")
def index():
    return jsonify({"service": "bank-report-api", "status": "ok"})

@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200

@app.route("/generate", methods=["POST"])
def generate():
    try:
        # 从表单文本获取模板ID（gd001、gd002 不带后缀）
        template_id = request.form.get("templateId")
        report_date = request.form.get("reportDate", "")
        photos = []
        for i in range(4):
            key = f"photo_{i}"
            if key in request.files:
                photos.append(request.files[key])

        print(f"[generate] templateId={template_id}, date={report_date}, photos={len(photos)}")

        # 校验模板参数
        if not template_id:
            return jsonify({"success": False, "error": "missing templateId"}), 400
        # 拼接本地模板文件路径
        template_path = os.path.join(TEMPLATE_BASE, f"{template_id}.docx")
        if not os.path.exists(template_path):
            return jsonify({"success": False, "error": f"template {template_id}.docx not found"}), 400

        # 校验照片数量（固定4张）
        if len(photos) != 4:
            msg = f"need 4 photos, got {len(photos)}"
            return jsonify({"success": False, "error": msg}), 400

        if not report_date:
            return jsonify({"success": False, "error": "missing date"}), 400

        temp_dir = tempfile.mkdtemp()
        output_path = os.path.join(temp_dir, "output.docx")

        # 保存上传照片到临时目录
        photo_paths = []
        for i, photo in enumerate(photos):
            p = os.path.join(temp_dir, f"photo_{i}.jpg")
            photo.save(p)
            photo_paths.append(p)

        print("[generate] files saved, processing...")

        # 调用原有生成报告逻辑，传入本地模板路径
        success = process_report(template_path, photo_paths, report_date, output_path)

        if not success:
            return jsonify({"success": False, "error": "process failed"}), 500

        print("[generate] done, sending file...")

        filename = f"report_{report_date.replace('-', '')}.docx"
        return send_file(
            output_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print(f"[generate] error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)