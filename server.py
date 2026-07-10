"""
考勤核对服务 - FastAPI 后端（一体化版本）
同时提供网页界面和API接口，cpolar 只需穿透一个端口
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import subprocess
import sys
import shutil
import uuid
import threading
import uvicorn
from datetime import datetime

app = FastAPI(title="考勤核对服务")

# CORS — 允许所有来源
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===== 目录配置 =====
BASE_DIR = Path(__file__).parent
TEMP_DIR = BASE_DIR / "temp_uploads"
TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR = BASE_DIR / "output"
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
                    template_path: Path, checkin_path: Path, leave_path: Path, remote_path: Path | None):
    """后台线程执行核对任务"""
    try:
        tasks[task_id]["status"] = "running"
        _log(task_id, "开始读取 BI 考勤模板...")

        script_path = BASE_DIR / "check_attendance_manual.py"
        cmd = [
            sys.executable, str(script_path),
            "--start", start_date,
            "--end", end_date,
            "--dept", department,
            "--template", str(template_path),
            "--checkin", str(checkin_path),
            "--leave", str(leave_path)
        ]
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
        for p in [template_path, checkin_path, leave_path]:
            if p.exists():
                 p.unlink()
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

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    """提供主页"""
    web_file = HTML_DIR / "index.html"
    if web_file.exists():
        return HTMLResponse(content=web_file.read_text(encoding="utf-8"))
    fallback = Path(r"C:\Users\CD2900\WorkBuddy\2026-05-18-task-1\attendance_dashboard.html")
    if fallback.exists():
        return HTMLResponse(content=fallback.read_text(encoding="utf-8"))
    return HTMLResponse(content="<h1>页面文件未找到</h1><p>请确认 index.html 已放入 web/ 目录</p>")


# ===== API 接口 =====

@app.get("/health")
async def health():
    return {"status": "ok", "service": "考勤核对服务", "version": "3.0"}


@app.post("/upload_and_check")
async def upload_and_check(
    start_date: str = Form(...),
    end_date: str = Form(...),
    department: str = Form(default="工程造价一部"),
    template: UploadFile = File(...),
    checkin: UploadFile = File(...),
    leave: UploadFile = File(...),
    remote: UploadFile | None = File(None)
):
    """上传四个源文件并启动后台核对任务"""
    try:
        datetime.strptime(start_date, "%Y-%m-%d")
        datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError:
        return JSONResponse(status_code=400, content={"detail": "日期格式错误，应为 YYYY-MM-DD"})
    
    task_id = str(uuid.uuid4())[:8]
    date_prefix = start_date.replace("-", "") + "_" + end_date.replace("-", "")

    template_path = TEMP_DIR / f"template_{date_prefix}_{task_id}.xlsx"
    checkin_path = TEMP_DIR / f"checkin_{date_prefix}_{task_id}.xlsx"
    leave_path = TEMP_DIR / f"leave_{date_prefix}_{task_id}.xls"
    remote_path = None

    try:
        # 保存模板
        with open(template_path, "wb") as f:
            shutil.copyfileobj(template.file, f)
        # 保存打卡日报
        with open(checkin_path, "wb") as f:
            shutil.copyfileobj(checkin.file, f)
        # 保存休假申请
        with open(leave_path, "wb") as f:
            shutil.copyfileobj(leave.file, f)
    except Exception as e:
        return JSONResponse(status_code=500, content={"detail": f"保存文件失败: {str(e)}"})

    # 保存远程办公（可选）
    if remote:
        remote_path = TEMP_DIR / f"remote_{date_prefix}_{task_id}.xls"
        try:
            with open(remote_path, "wb") as f:
                shutil.copyfileobj(remote.file, f)
        except Exception as e:
            return JSONResponse(status_code=500, content={"detail": f"保存远程文件失败: {str(e)}"})

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

    return {"task_id": task_id, "status": "pending"}

@app.get("/task/{task_id}")
async def get_task_status(task_id: str):
    """查询任务状态和日志"""
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="任务不存在")
    t = tasks[task_id]
    return {
        "task_id": task_id,
        "status": t["status"],
        "logs": t["logs"],
        "result_file": t.get("result_file"),
        "error": t.get("error"),
    }


@app.get("/download/{filename}")
async def download_file(filename: str):
    """下载结果文件"""
    safe_filename = Path(filename).name
    file_path = OUTPUT_DIR / safe_filename
    
    if not file_path.exists():
        possible_files = list(OUTPUT_DIR.glob(f"*{safe_filename}"))
        if possible_files:
            file_path = possible_files[0]
        else:
            available = [f.name for f in OUTPUT_DIR.glob("*.xlsx")]
            raise HTTPException(
                status_code=404, 
                detail=f"文件不存在: {safe_filename}。可用文件: {available[:5] if available else '无'}"
            )
    
    return FileResponse(file_path, filename=safe_filename)


@app.get("/api/info")
async def api_info():
    return {
        "service": "考勤核对服务",
        "version": "3.0",
        "mode": "integrated",
        "endpoints": {
            "index": "/",
            "health": "/health",
            "upload": "/upload_and_check",
            "task_status": "/task/{task_id}",
            "download": "/download/{filename}",
        }
    }


if __name__ == "__main__":
    print("=" * 50)
    print("   考勤核对服务 v3.0")
    print("=" * 50)
    print(f"   本地地址: http://localhost:8000")
    print(f"   启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("-" * 50)
    print("   说明: 浏览器打开上述地址即可使用")
    print("   网页和API都在同一个服务中")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8000)