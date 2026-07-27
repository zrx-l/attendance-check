"""
考勤核对服务 - Flask 后端（PythonAnywhere 兼容版）
从 FastAPI 迁移到 Flask，支持 PythonAnywhere 免费版 WSGI 部署
"""

from flask import Flask, request, jsonify, send_file, render_template_string
from pathlib import Path
import subprocess
import sys
import shutil
import uuid
import threading
from datetime import datetime

app = Flask(__name__)

# ===== 目录配置 =====
BASE_DIR = Path(__file__).parent
TEMP_DIR = BASE_DIR / "temp_uploads"
OUTPUT_DIR = BASE_DIR / "output"
TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
HTML_DIR = BASE_DIR / "web"
HTML_DIR.mkdir(exist_ok=True)

# ===== 内存任务存储 =====
tasks = {}  # task_id -> {status, logs, result_file, error, created_at}


def _log(task_id: str, message: str):
    """向任务日志追加一行"""
    if task_id in tasks:
        timestamp = datetime.now().strftime("%H:%M:%S")
        tasks[task_id]["logs"].append(f"[{timestamp}] {message}")


def _run_task_async(task_id: str, start_date: str, end_date: str, department: str,
                    template_path: Path, checkin_path: Path, leave_path: Path | None, remote_path: Path | None):
    """后台线程执行核对任务"""
    try:
        tasks[task_id]["status"] = "running"
        _log(task_id, "开始读取 BI 考勤模板...")

        script_path = BASE_DIR / "check_attendance_manual.py"
        cmd = [
            "python3", str(script_path),
            "--start", start_date,
            "--end", end_date,
            "--dept", department,
            "--template", str(template_path),
            "--checkin", str(checkin_path),
        ]
        if leave_path:
            cmd.extend(["--leave", str(leave_path)])
        if remote_path:
            cmd.extend(["--remote", str(remote_path)])

        _log(task_id, f"执行核对脚本: check_attendance_manual.py --start {start_date} --end {end_date} --dept {department}")

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=str(BASE_DIR),
            encoding='utf-8',
            errors='replace'
        )

        for line in process.stdout:
            line = line.strip()
            if line:
                _log(task_id, line)

        process.wait()
        # 清理临时文件
        for p in [template_path, checkin_path]:
            if p.exists():
                p.unlink()
        if leave_path and leave_path.exists():
            leave_path.unlink()
        if remote_path and remote_path.exists():
            remote_path.unlink()

        if process.returncode != 0:
            _log(task_id, f"脚本执行失败，退出码: {process.returncode}")
            tasks[task_id]["status"] = "error"
            tasks[task_id]["error"] = f"脚本执行失败，退出码: {process.returncode}"
        else:
            _log(task_id, "核对完成！")

            # 直接构造文件名
            start = datetime.strptime(start_date, "%Y-%m-%d")
            end = datetime.strptime(end_date, "%Y-%m-%d")
            start_str = f"{start.year}.{start.month}.{start.day}"
            end_str = f"{end.year}.{end.month}.{end.day}"
            filename = f"{department}考勤（{start_str}-{end_str}）.xlsx"
            output_file = OUTPUT_DIR / filename

            if output_file.exists():
                tasks[task_id]["status"] = "done"
                tasks[task_id]["result_file"] = output_file.name
                _log(task_id, f"结果已保存: {output_file.name}")
            else:
                # 兜底：取最新文件
                xlsx_files = list(OUTPUT_DIR.glob("*.xlsx"))
                if xlsx_files:
                    output_file = max(xlsx_files, key=lambda f: f.stat().st_mtime)
                    tasks[task_id]["status"] = "done"
                    tasks[task_id]["result_file"] = output_file.name
                    _log(task_id, f"使用最新文件: {output_file.name}")
                else:
                    tasks[task_id]["status"] = "error"
                    tasks[task_id]["error"] = "未生成结果文件"
                    _log(task_id, "未找到结果文件")

    except Exception as e:
        _log(task_id, f"异常: {str(e)}")
        tasks[task_id]["status"] = "error"
        tasks[task_id]["error"] = str(e)


# ===== 网页服务 =====

@app.route("/")
def serve_index():
    """提供主页"""
    web_file = HTML_DIR / "index.html"
    if web_file.exists():
        return render_template_string(web_file.read_text(encoding="utf-8"))
    return "<h1>页面文件未找到</h1><p>请确认 index.html 已放入 web/ 目录</p>"


# ===== API 接口 =====

@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "考勤核对服务", "version": "3.0-pa"})


@app.route("/upload_and_check", methods=["POST"])
def upload_and_check():
    """上传文件并启动后台核对任务（休假和远程为可选）"""
    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")
    department = request.form.get("department", "工程造价一部")

    try:
        datetime.strptime(start_date, "%Y-%m-%d")
        datetime.strptime(end_date, "%Y-%m-%d")
    except (ValueError, TypeError):
        return jsonify({"detail": "日期格式错误，应为 YYYY-MM-DD"}), 400

    task_id = str(uuid.uuid4())[:8]
    date_prefix = start_date.replace("-", "") + "_" + end_date.replace("-", "")

    template_path = TEMP_DIR / f"template_{date_prefix}_{task_id}.xlsx"
    checkin_path = TEMP_DIR / f"checkin_{date_prefix}_{task_id}.xlsx"
    leave_path = None
    remote_path = None

    try:
        # 保存模板
        template_file = request.files.get("template")
        if template_file:
            template_file.save(str(template_path))
        else:
            return jsonify({"detail": "未上传模板文件"}), 400

        # 保存打卡日报
        checkin_file = request.files.get("checkin")
        if checkin_file:
            checkin_file.save(str(checkin_path))
        else:
            return jsonify({"detail": "未上传打卡文件"}), 400

        # 保存休假申请（可选）
        leave_file = request.files.get("leave")
        if leave_file and leave_file.filename:
            leave_path = TEMP_DIR / f"leave_{date_prefix}_{task_id}.xls"
            leave_file.save(str(leave_path))

        # 保存远程办公（可选）
        remote_file = request.files.get("remote")
        if remote_file and remote_file.filename:
            remote_path = TEMP_DIR / f"remote_{date_prefix}_{task_id}.xls"
            remote_file.save(str(remote_path))

    except Exception as e:
        return jsonify({"detail": f"保存文件失败: {str(e)}"}), 500

    tasks[task_id] = {
        "status": "pending",
        "logs": [f"[系统] 任务创建成功，ID: {task_id}"],
        "result_file": None,
        "error": None,
        "created_at": datetime.now().isoformat(),
        "start_date": start_date,
        "end_date": end_date,
        "department": department,
    }

    thread = threading.Thread(
        target=_run_task_async,
        args=(task_id, start_date, end_date, department, template_path, checkin_path, leave_path, remote_path),
        daemon=True
    )
    thread.start()

    return jsonify({"task_id": task_id, "status": "pending"})


@app.route("/task/<task_id>")
def get_task_status(task_id: str):
    """查询任务状态和日志"""
    if task_id not in tasks:
        return jsonify({"detail": "任务不存在"}), 404
    t = tasks[task_id]
    return jsonify({
        "task_id": task_id,
        "status": t["status"],
        "logs": t["logs"],
        "result_file": t.get("result_file"),
        "error": t.get("error"),
    })


@app.route("/download/<filename>")
def download_file(filename: str):
    """下载结果文件"""
    safe_filename = Path(filename).name
    file_path = OUTPUT_DIR / safe_filename

    if not file_path.exists():
        possible_files = list(OUTPUT_DIR.glob(f"*{safe_filename}"))
        if possible_files:
            file_path = possible_files[0]
        else:
            available = [f.name for f in OUTPUT_DIR.glob("*.xlsx")]
            return jsonify({"detail": f"文件不存在: {safe_filename}。可用文件: {available[:5] if available else '无'}"}), 404

    return send_file(str(file_path), as_attachment=True, download_name=safe_filename)


@app.route("/api/info")
def api_info():
    return jsonify({
        "service": "考勤核对服务",
        "version": "3.0-pa",
        "mode": "PythonAnywhere",
        "endpoints": {
            "index": "/",
            "health": "/health",
            "upload": "/upload_and_check",
            "task_status": "/task/{task_id}",
            "download": "/download/{filename}",
        }
    })


if __name__ == "__main__":
    print("=" * 50)
    print("   考勤核对服务 v3.0 (PythonAnywhere Flask)")
    print("=" * 50)
    print(f"   本地地址: http://localhost:8000")
    print(f"   启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)
    app.run(host="0.0.0.0", port=8000)
