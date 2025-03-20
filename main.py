from fastapi import FastAPI, File, UploadFile, Request, Form, Query
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import schedule
from sqlalchemy import create_engine, text, inspect
from sqlalchemy.orm import sessionmaker
from typing import List, Optional
import uvicorn
import os, datetime, time
import shutil
# from uuid import uuid4
from threading import Thread
import subprocess
import sys
import tempfile
import zipfile
import asyncio
import concurrent.futures
from services.report import report_maker
from services.databases import tasklist, comparelist, adminuser, db, engine
db.metadata.create_all(engine)
Session = sessionmaker(bind=engine)

app = FastAPI(max_request_size=100*1024*1024)

# 掛載靜態文件目錄
app.mount("/static", StaticFiles(directory="static"), name="static")

# 設定模板目錄
templates = Jinja2Templates(directory="templates")

# 設定指定的資料夾（確保它存在）
UPLOAD_FOLDER = "data/uploadproject"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def start_gunicorn():
    """启动 Gunicorn 服务"""
    uvicorn.run("main:app", host="0.0.0.0", port=8080, reload=False)
    # subprocess.Popen(
    #     [ "gunicorn", "-c", "gunicorn.conf.py", "main:app" ],
    #     stdout=sys.stdout,
    #     stderr=sys.stderr,
    #     text=True
    # )

def run_script():
    """运行定时脚本"""
    try:
        result = subprocess.run(
            ["python", "-m", "services.report"],
            stdout=sys.stdout,
            stderr=sys.stderr,
            text=True,
            check=True
        )
    except subprocess.CalledProcessError as e:
        print(f"Script failed with return code {e.returncode}", file=sys.stderr)

def run_scheduler():
    """调度器循环运行定时任务"""
    schedule.every(2).minutes.do(run_script)
    while True:
        try:
            schedule.run_pending()
            time.sleep(1)
        except Exception as e:
            print(f"Scheduler encountered an error: {e}", file=sys.stderr)

#====================================================================================
@app.get("/")
async def read_root(request: Request):   # 首頁
    return templates.TemplateResponse("index.html", {"request": request, "title": "Nessus&AppScan報告產出系統"})

@app.post("/upload/")
async def upload_file(
    files: List[UploadFile] = File(...),
    uploadtaskname: Optional[str] = Form(None)
):   
    # # 產生唯一檔名，避免覆蓋
    # unique_folder_path = os.path.join(UPLOAD_FOLDER, uuid4().hex)
    # os.makedirs(unique_folder_path, exist_ok=True)
    # file_path = os.path.join(unique_folder_path, os.path.basename(file.filename))

    nowtime = datetime.datetime.now().strftime("%Y%m%d")  # 當前日期
    session = Session()

    try:
        # 檢查用戶是否提供自定義任務名稱
        if uploadtaskname:
            # 檢查資料庫中是否已存在該任務名稱
            if session.query(tasklist).filter_by(taskname=uploadtaskname).first():
                return {"error": f"Task name '{uploadtaskname}' already exists."}

        # 查詢當天所有自動編號的任務
        existtask = session.query(tasklist).filter(
            tasklist.exectime == nowtime,
            tasklist.taskname.like(f"{nowtime}_%")
        ).all()

        # 計算新的自動編號（僅基於當天的自動編號任務）
        tmpno = 0 if len(existtask) == 0 else max([int(t.taskname.split('_')[1]) for t in existtask])
        tmpfilename = f"{nowtime}_{tmpno + 1}"  # 自動生成任務名稱

        # 如果未提供自定義名稱，則使用自動生成的名稱
        uploadtaskname = uploadtaskname or tmpfilename

        # 檢查是否有上傳 AppScan 檔案
        if len(files) == 0:
            return {"error": "No AppScan files uploaded."}

        # 建立儲存 AppScan 檔案的資料夾
        temppath = os.path.join(os.getcwd(), 'data/uploadproject', uploadtaskname)
        os.makedirs(temppath, exist_ok=True)

        # 儲存 AppScan 檔案
        for i in files:
            if os.path.splitext(i.filename)[1] == '.scan':
                tmp_file_name = f"{os.path.splitext(os.path.basename(i.filename))[0]}.scan"
                os.makedirs(os.path.join(temppath, "appscan"), exist_ok=True)
                tempfilename = os.path.join(temppath, "appscan", tmp_file_name)
            elif os.path.splitext(i.filename)[1] == '.nessus':
                tmp_file_name = f"{os.path.splitext(os.path.basename(i.filename))[0]}.nessus"
                os.makedirs(os.path.join(temppath, "nessus"), exist_ok=True)
                tempfilename = os.path.join(temppath, "nessus", tmp_file_name)
            with open(tempfilename, "wb") as buffer:
                shutil.copyfileobj(i.file, buffer)

        # 新增任務到資料庫
        t = tasklist(taskname=uploadtaskname, exectime=nowtime)
        session.add(t)
        session.commit()  # 提交新增操作

        # 查詢剛剛新增的任務記錄
        t2 = session.query(tasklist).filter_by(taskname=uploadtaskname).all()
        # 更新任務狀態為 'prepare'
        t2[0].status = 'prepare'
        session.commit()  # 提交更新操作

        # 使用线程直接執行資料分析
        Thread(target=run_script, daemon=True).start()

        return
    except Exception as e:
        print(f"Error: {e}")
        session.rollback()  # 確保回滾任何未完成的交易
        return {"error": str(e)}
    finally:
        session.close()  # 確保資源被正確釋放

@app.get("/download")
async def download_file(     # 下載報告
    name: str = Query(..., description="要下載的報告名稱"),
    company_name: str = Query(..., description="公司名稱"),
    project_name: str = Query(..., description="專案名稱"),
    report_name: str = Query(..., description="報告名稱"),
    file_no: str = Query(..., description="文編預留位置"),
    scanner_ip: str = Query(..., description="scanner_ip"),
    date_start: str = Query(..., description="date_start"),
    date_end: str = Query(..., description="date_end"),
    type: str = Query(..., description="要下載的報告為單一(single)或多個(multiple)")
    ):
    ## http://localhost:8080/download?name=MyTaskName3&company_name=Tech+Corp&project_name=AI+Research&report_name=Quarterly+Analysis&file_no=A12345&scanner_ip=192.168.1.100&date_start=2025-01-01&date_end=2025-03-12&type=single
    ## 確認資料是否已經完成解析
    session = Session()
    result = session.query(tasklist).filter(tasklist.taskname == name).first()
    if result is None:
        return {"error": f"任務 '{name}' 不存在"}
    elif result.status != "ok":
        return {"error": f"任務 '{name}' 尚未完成解析"}
    ## 產出報告
    task_folder_path = os.path.join(os.getcwd(), "data/uploadproject", name)
    task_folder_web_path = os.path.join(task_folder_path, "appscan")
    task_folder_va_path = os.path.join(task_folder_path, "nessus")
    title_word = {
        "company_name": company_name,
        "project_name": project_name,
        "report_name": report_name,
        "file_no": file_no,
        "scanner_ip": scanner_ip,
        "date_start": date_start,
        "date_end": date_end
    }
    download_paths = []
    data = {
        "word_title": title_word,
        "report_type": type,
        "task_folder_path": task_folder_path
    }
    
    report_maker.report_generater()

    if os.path.isdir(task_folder_web_path):
        ## 下載報告
        word = os.path.join(task_folder_web_path, f"{name}.docx")
        excel = os.path.join(task_folder_web_path, f"{name}.xlsx")
        if type == "single":
            download_paths.extend([word, excel])
        else:
            for root, _, files in os.walk(task_folder_web_path):
                download_paths.extend([os.path.join(root, f) for f in files if f.endswith((".docx", ".xlsx"))])
            for file in [word, excel]:
                try:
                    download_paths.remove(file)
                except ValueError:  # 只捕獲 list.remove() 可能拋出的錯誤
                    pass
    ## nessus
    if os.path.isdir(task_folder_va_path):
        for root, _, files in os.walk(os.path.join(task_folder_path, "nessus")):
            download_paths.extend([os.path.join(root, f) for f in files if f.endswith((".docx", ".xlsx"))])
    temp_dir = tempfile.mkdtemp()  # 創建臨時 ZIP 檔案
    zip_path = os.path.join(temp_dir, f"{name}.zip")
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for ff in download_paths:
            zipf.write(ff, arcname=f"{os.path.basename(ff)}")
    ## 傳送檔案給使用者
    return FileResponse(zip_path, filename=f"{name}.zip", media_type="application/octet-stream")

@app.get("/show")
async def show(request: Request):
    # 建立 session
    session = Session()
    # 使用 SQLAlchemy 的 text 執行原生 SQL 指令
    result = session.query(tasklist.id, tasklist.taskname, tasklist.exectime, tasklist.status).all()
    # 取得所有列資料
    tasks = [
        {"id": r[0], "taskname": r[1], "exectime": r[2], "status": r[3]}
        for r in result
    ]
    session.close()
    return tasks

@app.get("/del")
async def delete_task(name: str = Query(..., description="Task name to delete")):
    try:
        # 建立 session
        session = Session()
        # 刪除該任務資料夾
        folder_path = os.path.join(os.getcwd(), "data/uploadproject", name)
        shutil.rmtree(folder_path)  
        # 執行刪除指令
        session.query(tasklist).filter(tasklist.taskname == name).delete()    
        # 提交變更
        session.commit()
        session.close()
        return {"message": f"Task '{name}' deleted successfully"}   
    except FileNotFoundError:
        return {"message": f"任務 '{name}' 不存在"}
    except PermissionError:
        return {"message": f"沒有權限刪除 '{name}'，請確保沒有檔案被佔用"}
    except Exception as e:
        return {"message": f"刪除資料夾時發生錯誤: {e}"}
    finally:
        session.close()  # 確保資源被正確釋放

# 啟動伺服器
# 使用 `uvicorn filename:app --reload`

if __name__ == "__main__":
    # uvicorn.run("main:app", host="0.0.0.0", port=8080, reload=True)

    try:
        # 使用线程同时运行 Gunicorn 和调度器
        gunicorn_thread = Thread(target=start_gunicorn)
        gunicorn_thread.daemon = True
        gunicorn_thread.start()

        # 启动任务调度器
        run_scheduler()
    
    except KeyboardInterrupt:
        print("Shutting down gracefully...")
    
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)

    