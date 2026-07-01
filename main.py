"""
main.py
=======
FastAPI app: POST {main_brand, competitors, start_date, end_date}
→ fetch CMS data for all topics → process → generate PPTX → save to ./report.

Endpoints:
    POST /generate                  Start background pipeline, returns {task_id}
    GET  /tasks                     List all tasks (in-memory)
    GET  /task/{task_id}            Poll task status
    GET  /task/{task_id}/download   Download PPTX when done
    GET  /reports                   List saved reports
    GET  /reports/{filename}/download  Download report by filename
    DELETE /reports/{filename}      Delete a report
    POST /reports/upload            Upload a PPTX manually
    GET  /data                      List cached datasets
    DELETE /data/{key}              Delete a cached dataset
    GET  /project/topics            List topics in a CMS project
    GET  /health                    Health check

Run:
    uvicorn main:app --reload
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import traceback
import uuid
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

from src.cms.login import login, get_project_info
from src.cms.export_data import fetch_data_export
from src.process_data import process_data
from src.report_generator import generate_report
import src.database as db

# ── Config ────────────────────────────────────────────────────────────────────

TEMPLATE_PATH  = Path(os.getenv("TEMPLATE_PATH", "Template_full_fixed.pptx")).resolve()
CMS_PROJECT_ID = os.getenv("CMS_PROJECT_ID", "")

# ── App ───────────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(_app: FastAPI):
    db.init_db()
    yield


app = FastAPI(
    lifespan=lifespan,
    title="PPTX Report Generator",
    description=(
        "Generate social listening PPTX reports from Kompa CMS data.\n\n"
        "**Flow:** `POST /generate` → poll `GET /task/{task_id}` → download via "
        "`GET /task/{task_id}/download` or `GET /reports/{filename}/download`."
    ),
    version="3.0.0",
    openapi_tags=[
        {"name": "Pipeline",  "description": "Tạo report, theo dõi và tải kết quả."},
        {"name": "Reports",   "description": "Quản lý file PPTX trong `./report`."},
        {"name": "Data",      "description": "Quản lý dataset thô đã cache trong `./data`."},
        {"name": "Project",   "description": "Tra cứu thông tin CMS project."},
        {"name": "System",    "description": "Health check và thông tin hệ thống."},
    ],
)

# ── Paths ─────────────────────────────────────────────────────────────────────

REPORT_DIR = Path("report")
DATA_DIR   = Path("data")


def _report_filename(req: "GenerateRequest", ts: datetime) -> str:
    safe_brand = req.main_brand.replace(" ", "_")
    return f"{safe_brand}_{req.start_date}_{req.end_date}_{int(ts.timestamp())}.pptx"


def _save_report(pptx_bytes: bytes, filename: str) -> str:
    REPORT_DIR.mkdir(exist_ok=True)
    out_path = REPORT_DIR / filename
    out_path.write_bytes(pptx_bytes)
    return str(out_path.resolve())


# ── Data cache helpers ────────────────────────────────────────────────────────

def _cache_key(project_id: str, main_brand: str, competitors: list[str], start_date: str, end_date: str) -> str:
    payload = json.dumps({
        "project_id": project_id,
        "main_brand": main_brand,
        "competitors": sorted(competitors),
        "start_date": start_date,
        "end_date": end_date,
    }, sort_keys=True)
    return hashlib.sha256(payload.encode()).hexdigest()[:16]


def _load_cached_df(key: str):
    import pandas as pd
    csv_file = DATA_DIR / f"{key}.csv"
    if csv_file.exists():
        return pd.read_csv(csv_file)
    return None


def _save_data_cache(key: str, df, req: "GenerateRequest", project_id: str) -> None:
    import pandas as pd
    DATA_DIR.mkdir(exist_ok=True)
    df.to_csv(DATA_DIR / f"{key}.csv", index=False)
    (DATA_DIR / f"{key}.meta.json").write_text(
        json.dumps({
            "key": key,
            "project_id": project_id,
            "main_brand": req.main_brand,
            "competitors": req.competitors,
            "start_date": req.start_date,
            "end_date": req.end_date,
            "record_count": len(df),
            "created_at": datetime.now().isoformat(),
        }, ensure_ascii=False),
        encoding="utf-8",
    )


# ── Request / Response models ─────────────────────────────────────────────────

class GenerateRequest(BaseModel):
    project_id: str | None = Field(
        default=None,
        description="CMS project `_id`. Mặc định dùng env `CMS_PROJECT_ID`.",
        examples=["6650a1b2c3d4e5f600000001"],
    )
    main_brand: str = Field(..., description="Tên brand chính.", examples=["MSN"])
    competitors: list[str] = Field(
        default=[],
        description="Danh sách brand đối thủ.",
        examples=[["KIDO", "MWG", "Hòa Phát"]],
    )
    start_date: str = Field(..., description="Ngày bắt đầu (YYYY-MM-DD).", examples=["2026-06-22"])
    end_date:   str = Field(..., description="Ngày kết thúc (YYYY-MM-DD).", examples=["2026-06-28"])

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "project_id": "6650a1b2c3d4e5f600000001",
                    "main_brand": "MSN",
                    "competitors": ["KIDO", "MWG", "Hòa Phát"],
                    "start_date": "2026-06-22",
                    "end_date":   "2026-06-28",
                }
            ]
        }
    }


class TaskOut(BaseModel):
    status:           str        = Field(description="`pending` | `running` | `done` | `failed`")
    step:             str | None = Field(None, description="Bước đang chạy (khi `running`).")
    file_path:        str | None = Field(None, description="Đường dẫn tuyệt đối file PPTX (khi `done`).")
    filename:         str | None = Field(None, description="Tên file PPTX (khi `done`).")
    error:            str | None = Field(None, description="Mô tả lỗi (khi `failed`).")
    duration_seconds: float | None = Field(None, description="Thời gian chạy tính từ lúc tạo task (giây).")

    model_config = {
        "json_schema_extra": {
            "examples": [
                {"status": "pending",  "step": None,           "file_path": None, "filename": None, "error": None, "duration_seconds": None},
                {"status": "running",  "step": "fetch_data",   "file_path": None, "filename": None, "error": None, "duration_seconds": None},
                {"status": "done",     "step": None,           "file_path": "/app/report/report_MSN_2026-06-22_2026-06-28.pptx", "filename": "report_MSN_2026-06-22_2026-06-28.pptx", "error": None, "duration_seconds": 142.3},
                {"status": "failed",   "step": None,           "file_path": None, "filename": None, "error": "No CMS topics found matching brands: ['MSN']", "duration_seconds": 8.1},
            ]
        }
    }


class ReportItem(BaseModel):
    filename:    str            = Field(examples=["report_MSN_2026-06-22_2026-06-28.pptx"])
    main_brand:  str | None     = Field(None, examples=["MSN"])
    competitors: list[str]      = Field(default=[], examples=[["KIDO", "MWG"]])
    start_date:  str | None     = Field(None, examples=["2026-06-22"])
    end_date:    str | None     = Field(None, examples=["2026-06-28"])
    size_bytes:  int            = Field(examples=[2048576])
    created_at:  str            = Field(examples=["2026-06-29T10:30:00"])
    file_exists: bool           = Field(default=True)


class DatasetItem(BaseModel):
    key:          str       = Field(examples=["a1b2c3d4e5f60001"])
    project_id:   str       = Field(examples=["6650a1b2c3d4e5f600000001"])
    main_brand:   str       = Field(examples=["MSN"])
    competitors:  list[str] = Field(examples=[["KIDO", "MWG"]])
    start_date:   str       = Field(examples=["2026-06-22"])
    end_date:     str       = Field(examples=["2026-06-28"])
    record_count: int       = Field(examples=[15420])
    created_at:   str       = Field(examples=["2026-06-29T10:15:00"])
    size_bytes:   int       = Field(examples=[8388608])


class TopicItem(BaseModel):
    topicId: str = Field(examples=["topic_abc123"])
    topic:   str = Field(examples=["MSN"])


# ── Background pipeline ───────────────────────────────────────────────────────

def _run_pipeline(task_id: str, req: GenerateRequest) -> None:
    try:
        project_id = req.project_id or CMS_PROJECT_ID
        if not project_id:
            raise ValueError(
                "project_id is required — set CMS_PROJECT_ID env or pass in request body"
            )

        start_dt = datetime.strptime(req.start_date, "%Y-%m-%d")
        end_dt   = datetime.strptime(req.end_date,   "%Y-%m-%d").replace(
                       hour=23, minute=59, second=59)

        # ── Check data cache ──────────────────────────────────────────────────
        db.update_task(task_id, status="running", step="check_cache")
        key = _cache_key(project_id, req.main_brand, req.competitors, req.start_date, req.end_date)
        df  = _load_cached_df(key)

        if df is not None:
            logger.info("[%s] Cache hit: key=%s, %d rows", task_id, key, len(df))
        else:
            all_brands = [req.main_brand] + req.competitors

            # 1 ── Login
            db.update_task(task_id, step="login")
            access_token, refresh_token = login()

            # 2 ── Get project info
            db.update_task(task_id, step="get_project_info")
            info = get_project_info(
                project_id=project_id,
                access_token=access_token,
                refresh_token=refresh_token,
                selected_topics=all_brands,
            )
            if not info["topics"]:
                raise ValueError(
                    f"No CMS topics found matching brands: {all_brands}. "
                    "Check that topic names in the CMS project match the brand names."
                )
            logger.info("[%s] Topics: %s", task_id, [t["topic"] for t in info["topics"]])

            # 3 ── Build date range (current period + equal-length previous period)
            days       = (end_dt.date() - start_dt.date()).days + 1
            prev_start = start_dt - timedelta(days=days)
            from_date  = prev_start.strftime("%Y-%m-%d 00:00:00")
            to_date    = end_dt.strftime("%Y-%m-%d %H:%M:%S")

            # 4 ── Export data
            db.update_task(task_id, step="fetch_data")
            topic_ids        = [t["topicId"] for t in info["topics"]]
            topic_id_to_name = {t["topicId"]: t["topic"] for t in info["topics"]}
            logger.info("[%s] Exporting %d topics: %s", task_id, len(topic_ids), list(topic_id_to_name.values()))

            df = fetch_data_export(
                access_token=access_token,
                refresh_token=refresh_token,
                topic_ids=topic_ids,
                topic_id_to_name=topic_id_to_name,
                from_date=from_date,
                to_date=to_date,
                project_info=info,
            )
            logger.info("[%s] Export complete: %d rows", task_id, len(df))

            _save_data_cache(key, df, req, project_id)
            logger.info("[%s] Data cached: key=%s", task_id, key)

        # 5 ── Validate
        db.update_task(task_id, step="process_data")
        if df.empty:
            raise ValueError("No data returned from CMS for the given date range and topics.")
        logger.info("[%s] DataFrame: %d rows", task_id, len(df))

        # 6 ── Process data
        data = process_data(df, req.main_brand, req.competitors, start_dt, end_dt)

        # 7 ── Generate PPTX
        db.update_task(task_id, step="generate_report")
        if not TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")

        pptx_bytes = generate_report(
            data, TEMPLATE_PATH, api_key=os.getenv("DEEPINFRA_API_KEY")
        )

        # 8 ── Save
        filename  = _report_filename(req, datetime.now())
        file_path = _save_report(pptx_bytes, filename)

        db.save_report(req.main_brand, req.competitors, req.start_date, req.end_date, filename)
        logger.info("[%s] Report saved to DB: %s", task_id, filename)

        db.update_task(task_id, status="done", step=None, file_path=file_path, filename=filename)
        logger.info("[%s] Done → %s", task_id, file_path)

    except Exception as exc:
        logger.error("[%s] Pipeline failed:\n%s", task_id, traceback.format_exc())
        db.update_task(task_id, status="failed", step=None, error=str(exc))


# ── Endpoints — Pipeline ──────────────────────────────────────────────────────

_PPTX_MEDIA = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
_ERR_404    = {"description": "Không tìm thấy"}
_ERR_400    = {"description": "Request không hợp lệ"}


@app.post(
    "/generate",
    tags=["Pipeline"],
    summary="Tạo report PPTX",
    responses={
        200: {
            "description": "Task được tạo hoặc report đã tồn tại.",
            "content": {
                "application/json": {
                    "examples": {
                        "task_created": {
                            "summary": "Task mới được tạo",
                            "value": {"task_id": "550e8400-e29b-41d4-a716-446655440000"},
                        },
                        "already_exists": {
                            "summary": "Report đã tồn tại",
                            "value": {
                                "already_exists": True,
                                "filename": "report_MSN_2026-06-22_2026-06-28.pptx",
                                "message": "Report đã tồn tại. Dùng GET /reports/report_MSN_2026-06-22_2026-06-28.pptx để tải về.",
                            },
                        },
                    }
                }
            },
        }
    },
)
def generate(req: GenerateRequest, background_tasks: BackgroundTasks):
    """
    Khởi chạy pipeline trong background. Trả về `task_id` ngay lập tức.

    - Nếu report **đã tồn tại** trong `./report`, trả về `already_exists: true` kèm tên file.
    - Nếu **chưa có**, tạo task mới → poll `GET /task/{task_id}` để theo dõi tiến độ.
    - Data thô sẽ được **cache** vào `./data`; lần sau cùng params sẽ bỏ qua bước fetch.
    """
    # 1. Report đã tồn tại?
    existing = db.find_report(req.main_brand, req.competitors, req.start_date, req.end_date)
    if existing:
        fname = existing["report_name"]
        if (REPORT_DIR / fname).exists():
            return {
                "already_exists": True,
                "filename":       fname,
                "message":        f"Report đã tồn tại. Dùng GET /reports/{fname}/download để tải về.",
            }
        # Record in DB but file deleted — fall through to regenerate

    # 2. Atomic check + create — ngăn duplicate pipeline khi 2 request đến cùng lúc
    task_id, created = db.claim_task(
        str(uuid.uuid4()), req.main_brand, req.competitors, req.start_date, req.end_date
    )
    if not created:
        return {"task_id": task_id, "already_running": True}

    background_tasks.add_task(_run_pipeline, task_id, req)
    return {"task_id": task_id}


@app.get(
    "/tasks",
    tags=["Pipeline"],
    summary="Danh sách tất cả tasks",
    responses={
        200: {
            "description": "Map task_id → trạng thái task (lưu trong DB, không mất khi restart).",
            "content": {
                "application/json": {
                    "example": {
                        "550e8400-e29b-41d4-a716-446655440000": {
                            "status": "done",
                            "step": None,
                            "file_path": "/app/report/report_MSN_2026-06-22_2026-06-28.pptx",
                            "filename": "report_MSN_2026-06-22_2026-06-28.pptx",
                            "error": None,
                        },
                        "661f9511-f3ac-52e5-b827-557766551111": {
                            "status": "running",
                            "step": "fetch_data",
                            "file_path": None,
                            "filename": None,
                            "error": None,
                        },
                    }
                }
            },
        }
    },
)
def list_tasks():
    """Trả về toàn bộ tasks đã lưu trong DB."""
    return db.list_tasks()


@app.get(
    "/task/{task_id}",
    tags=["Pipeline"],
    summary="Trạng thái task",
    response_model=TaskOut,
    responses={
        200: {
            "description": "Trạng thái hiện tại của task.",
            "content": {
                "application/json": {
                    "examples": {
                        "running": {
                            "summary": "Đang chạy",
                            "value": {"status": "running", "step": "fetch_data", "file_path": None, "filename": None, "error": None},
                        },
                        "done": {
                            "summary": "Hoàn thành",
                            "value": {"status": "done", "step": None, "file_path": "/app/report/report_MSN_2026-06-22_2026-06-28.pptx", "filename": "report_MSN_2026-06-22_2026-06-28.pptx", "error": None},
                        },
                        "failed": {
                            "summary": "Thất bại",
                            "value": {"status": "failed", "step": None, "file_path": None, "filename": None, "error": "No CMS topics found matching brands: ['MSN']"},
                        },
                    }
                }
            },
        },
        404: _ERR_404,
    },
)
def get_task(task_id: str):
    """
    Poll trạng thái task theo `task_id`.

    **Các bước** khi `status=running`:
    `check_cache` → `login` → `get_project_info` → `fetch_data` →
    `process_data` → `generate_report`

    Khi `status=done`, dùng `GET /task/{task_id}/download` để tải file.
    """
    task = db.get_task(task_id)
    if task is None:
        raise HTTPException(status_code=404, detail=f"Task '{task_id}' not found")
    return task


@app.get(
    "/task/{task_id}/download",
    tags=["Pipeline"],
    summary="Tải PPTX của task",
    responses={
        200: {"description": "File PPTX.", "content": {_PPTX_MEDIA: {}}},
        404: _ERR_404,
        410: {"description": "Task đã done nhưng file bị xoá khỏi server."},
        425: {"description": "Task chưa hoàn thành (`status` != `done`)."},
    },
)
def download_task(task_id: str):
    """Tải file PPTX sau khi task có `status=done`."""
    task = db.get_task(task_id)
    if task is None:
        raise HTTPException(status_code=404, detail=f"Task '{task_id}' not found")
    if task["status"] != "done":
        raise HTTPException(status_code=425, detail=f"Task not done yet (status: {task['status']})")
    file_path = Path(task["file_path"])
    if not file_path.exists():
        raise HTTPException(status_code=410, detail="File no longer exists on server")
    return FileResponse(path=str(file_path), media_type=_PPTX_MEDIA, filename=task["filename"])


# ── Endpoints — Reports ───────────────────────────────────────────────────────
# QUAN TRỌNG: route tĩnh (upload) phải đăng ký TRƯỚC route động ({filename})
# để Starlette không match "upload" vào {filename} và trả 405.

@app.get(
    "/reports",
    tags=["Reports"],
    summary="Danh sách report đã tạo",
    responses={
        200: {
            "description": "Danh sách file PPTX trong `./report`, mới nhất trước.",
            "content": {
                "application/json": {
                    "example": {
                        "reports": [
                            {"filename": "report_MSN_2026-06-22_2026-06-28.pptx", "size_bytes": 2048576, "created_at": "2026-06-29T10:30:00"},
                            {"filename": "report_MSN_2026-06-15_2026-06-21.pptx", "size_bytes": 1987432, "created_at": "2026-06-22T09:00:00"},
                        ]
                    }
                }
            },
        }
    },
)
def list_reports():
    """Liệt kê tất cả report (từ SQLite + file `.pptx` chưa có trong DB)."""
    REPORT_DIR.mkdir(exist_ok=True)
    db_records = db.list_reports()
    result = []
    seen: set[str] = set()

    for r in db_records:
        fname = r["report_name"]
        seen.add(fname)
        fpath = REPORT_DIR / fname
        exists = fpath.exists()
        result.append({
            "filename":    fname,
            "main_brand":  r["main_brand"],
            "competitors": r["competitors"],
            "start_date":  r["start_date"],
            "end_date":    r["end_date"],
            "size_bytes":  fpath.stat().st_size if exists else 0,
            "created_at":  r["created_at"],
            "file_exists": exists,
        })

    # Include files on disk not tracked in DB (e.g., manually uploaded)
    for f in sorted(REPORT_DIR.glob("*.pptx"), key=lambda f: f.stat().st_mtime, reverse=True):
        if f.name not in seen:
            result.append({
                "filename":    f.name,
                "main_brand":  None,
                "competitors": [],
                "start_date":  None,
                "end_date":    None,
                "size_bytes":  f.stat().st_size,
                "created_at":  datetime.fromtimestamp(f.stat().st_mtime).isoformat(),
                "file_exists": True,
            })

    return {"reports": result}


@app.post(
    "/reports/upload",
    tags=["Reports"],
    summary="Upload report thủ công",
    responses={
        200: {
            "description": "Upload thành công.",
            "content": {
                "application/json": {
                    "example": {
                        "filename":  "report_MSN_2026-06-22_2026-06-28.pptx",
                        "file_path": "/app/report/report_MSN_2026-06-22_2026-06-28.pptx",
                    }
                }
            },
        },
        400: _ERR_400,
    },
)
async def upload_report(file: UploadFile = File(..., description="File .pptx cần upload.")):
    """Upload thủ công một file `.pptx` vào `./report`."""
    if not file.filename.endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file .pptx")
    REPORT_DIR.mkdir(exist_ok=True)
    out_path = REPORT_DIR / file.filename
    out_path.write_bytes(await file.read())
    return {"filename": file.filename, "file_path": str(out_path.resolve())}


@app.get(
    "/reports/{filename}",
    tags=["Reports"],
    summary="Tải report theo tên file",
    responses={
        200: {"description": "File PPTX.", "content": {_PPTX_MEDIA: {}}},
        400: _ERR_400,
        404: _ERR_404,
    },
)
def download_report_by_filename(filename: str):
    """
    Tải file PPTX trực tiếp theo tên file.

    Dùng khi `POST /generate` trả về `already_exists: true`.
    """
    if not filename.endswith(".pptx") or "/" in filename or ".." in filename:
        raise HTTPException(status_code=400, detail="Tên file không hợp lệ")
    file_path = REPORT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"File '{filename}' không tồn tại")
    return FileResponse(path=str(file_path), media_type=_PPTX_MEDIA, filename=filename)


@app.delete(
    "/reports/{filename}",
    tags=["Reports"],
    summary="Xoá report",
    responses={
        200: {
            "description": "Xoá thành công.",
            "content": {"application/json": {"example": {"deleted": "report_MSN_2026-06-22_2026-06-28.pptx"}}},
        },
        400: _ERR_400,
        404: _ERR_404,
    },
)
def delete_report(filename: str):
    """Xoá file `.pptx` khỏi `./report` và xoá record trong DB."""
    if not filename.endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file .pptx")
    file_path = REPORT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"File '{filename}' không tồn tại")
    file_path.unlink()
    db.delete_report_by_name(filename)
    return {"deleted": filename}


# ── Endpoints — Data ──────────────────────────────────────────────────────────

@app.get(
    "/data",
    tags=["Data"],
    summary="Danh sách dataset đã cache",
    responses={
        200: {
            "description": "Danh sách dataset trong `./data`, mới nhất trước.",
            "content": {
                "application/json": {
                    "example": {
                        "datasets": [
                            {
                                "key":          "a1b2c3d4e5f60001",
                                "project_id":   "6650a1b2c3d4e5f600000001",
                                "main_brand":   "MSN",
                                "competitors":  ["KIDO", "MWG"],
                                "start_date":   "2026-06-22",
                                "end_date":     "2026-06-28",
                                "record_count": 15420,
                                "created_at":   "2026-06-29T10:15:00",
                                "size_bytes":   8388608,
                            }
                        ]
                    }
                }
            },
        }
    },
)
def list_data():
    """Liệt kê tất cả dataset đã cache trong `./data`."""
    DATA_DIR.mkdir(exist_ok=True)
    meta_files = sorted(DATA_DIR.glob("*.meta.json"), key=lambda f: f.stat().st_mtime, reverse=True)
    datasets = []
    for mf in meta_files:
        meta = json.loads(mf.read_text(encoding="utf-8"))
        # Support both new CSV cache and legacy JSON cache
        for ext in (".csv", ".json"):
            data_file = DATA_DIR / f"{meta['key']}{ext}"
            if data_file.exists():
                meta["size_bytes"] = data_file.stat().st_size
                break
        else:
            meta["size_bytes"] = 0
        datasets.append(meta)
    return {"datasets": datasets}


@app.delete(
    "/data/{key}",
    tags=["Data"],
    summary="Xoá cached dataset",
    responses={
        200: {
            "description": "Xoá thành công.",
            "content": {
                "application/json": {
                    "example": {"deleted": ["a1b2c3d4e5f60001.json", "a1b2c3d4e5f60001.meta.json"]}
                }
            },
        },
        404: _ERR_404,
    },
)
def delete_data(key: str):
    """
    Xoá dataset cache khỏi `./data` (hỗ trợ cả `.csv` và `.json` legacy).

    `key` lấy từ trường `key` trong kết quả `GET /data`.
    """
    csv_file  = DATA_DIR / f"{key}.csv"
    json_file = DATA_DIR / f"{key}.json"
    meta_file = DATA_DIR / f"{key}.meta.json"
    if not csv_file.exists() and not json_file.exists() and not meta_file.exists():
        raise HTTPException(status_code=404, detail=f"Dataset '{key}' không tồn tại")
    deleted = []
    for f in [csv_file, json_file, meta_file]:
        if f.exists():
            f.unlink()
            deleted.append(f.name)
    return {"deleted": deleted}


# ── Endpoints — Project ───────────────────────────────────────────────────────

@app.get(
    "/project/topics",
    tags=["Project"],
    summary="Danh sách topics trong CMS project",
    responses={
        200: {
            "description": "Topics của project.",
            "content": {
                "application/json": {
                    "example": {
                        "projectId":   "6650a1b2c3d4e5f600000001",
                        "projectName": "Masan Consumer",
                        "topics": [
                            {"topicId": "topic_abc111", "topic": "MSN"},
                            {"topicId": "topic_abc222", "topic": "KIDO"},
                        ],
                    }
                }
            },
        },
        400: _ERR_400,
        500: {"description": "Lỗi kết nối CMS."},
    },
)
def list_topics(project_id: str | None = None):
    """
    Lấy danh sách topics trong CMS project.

    Dùng giá trị `topic` trong kết quả để điền vào `main_brand` / `competitors`
    khi gọi `POST /generate`.
    """
    pid = project_id or CMS_PROJECT_ID
    if not pid:
        raise HTTPException(status_code=400, detail="project_id required (or set CMS_PROJECT_ID env)")
    try:
        access_token, refresh_token = login()
        info = get_project_info(project_id=pid, access_token=access_token, refresh_token=refresh_token)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
    return {
        "projectId":   info["projectId"],
        "projectName": info["projectName"],
        "topics":      info["topics"],
    }


# ── Endpoints — System ────────────────────────────────────────────────────────

@app.get(
    "/health",
    tags=["System"],
    summary="Health check",
    responses={
        200: {
            "description": "Trạng thái server.",
            "content": {
                "application/json": {
                    "example": {
                        "status":          "ok",
                        "template":        "/app/Template_full_fixed.pptx",
                        "template_exists": True,
                        "cms_project_id":  "6650a1b2c3d4e5f600000001",
                    }
                }
            },
        }
    },
)
def health():
    """Kiểm tra server còn sống và template tồn tại."""
    return {
        "status":           "ok",
        "template":         str(TEMPLATE_PATH),
        "template_exists":  TEMPLATE_PATH.exists(),
        "cms_project_id":   CMS_PROJECT_ID or "(not set)"
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
