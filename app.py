"""PI/CI/PL 검토 툴 — Flask 웹 서버"""
import json
import time
import uuid
from pathlib import Path

from flask import Flask, Response, abort, jsonify, render_template, request, send_file

import worker
from config import JOBS_DIR, MAX_CONTENT_BYTES, SECRET_KEY

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_BYTES
app.config["SECRET_KEY"] = SECRET_KEY

worker.start_worker()


# ---------------------------------------------------------------------------
# 라우트
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    output_dir = job_dir / "output"
    output_dir.mkdir(exist_ok=True)

    single = request.files.get("file")
    folder_files = request.files.getlist("files")

    if single and single.filename and single.filename.lower().endswith(".zip"):
        zip_path = job_dir / "upload.zip"
        single.save(str(zip_path))
        worker.enqueue(job_id, zip_path, output_dir)

    elif folder_files:
        upload_dir = job_dir / "upload"
        upload_dir.mkdir(exist_ok=True)
        for f in folder_files:
            if not f.filename:
                continue
            # 경로 순회 공격 방지 후 저장
            safe_parts = [p for p in Path(f.filename).parts if p not in ("", ".", "..")]
            if not safe_parts:
                continue
            dest = upload_dir.joinpath(*safe_parts)
            dest.parent.mkdir(parents=True, exist_ok=True)
            f.save(str(dest))
        worker.enqueue(job_id, upload_dir, output_dir)

    else:
        import shutil
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": "파일이 없습니다. ZIP 또는 폴더를 업로드해주세요."}), 400

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id: str):
    state = worker.job_states.get(job_id)
    if state is None:
        return jsonify({"error": "작업을 찾을 수 없습니다."}), 404

    if request.headers.get("Accept") == "text/event-stream":
        def _stream():
            last_pct = -1
            for _ in range(7200):   # 최대 2시간
                s = worker.job_states.get(job_id, {})
                pct = s.get("progress", 0)
                if pct != last_pct:
                    last_pct = pct
                    yield f"data: {json.dumps(s)}\n\n"
                if s.get("status") in ("done", "error"):
                    return
                time.sleep(0.5)
        return Response(
            _stream(),
            mimetype="text/event-stream",
            headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
        )

    return jsonify(state)


@app.route("/download/<job_id>/excel")
def download_excel(job_id: str):
    state = worker.job_states.get(job_id)
    if not state or state.get("status") != "done":
        abort(404)
    path = Path(state["excel_path"])
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name="pi_ci_pl_review.xlsx")


@app.route("/download/<job_id>/pdf")
def download_pdf(job_id: str):
    state = worker.job_states.get(job_id)
    if not state or state.get("status") != "done":
        abort(404)
    path = Path(state["pdf_path"])
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name="pi_ci_pl_report.pdf")


@app.route("/health")
def health():
    return jsonify({"ok": True})


if __name__ == "__main__":
    from config import HOST, PORT
    app.run(host=HOST, port=PORT, debug=False, threaded=True)
