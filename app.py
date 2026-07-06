from __future__ import annotations

import shutil
import uuid
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from generate_answer_cards import generate, read_students

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PDF = BASE_DIR / "template" / "answer_card_template.pdf"
JOBS_DIR = BASE_DIR / "jobs"
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(title="答题卡批量生成器", version="1.0.0")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    index_path = STATIC_DIR / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=500, detail="找不到前端页面 static/index.html")
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


def _validate_excel_filename(filename: str | None) -> None:
    if not filename:
        raise HTTPException(status_code=400, detail="请上传 Excel 文件。")
    suffix = Path(filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise HTTPException(status_code=400, detail="目前仅支持 .xlsx / .xlsm 格式。")


def _find_output_files(output_dir: Path) -> dict[str, Path]:
    merged_candidates = sorted(output_dir.glob("批量答题卡_*人合并版.pdf"))
    single_zip = output_dir / "批量答题卡_单人版.zip"
    preview = output_dir / "第一页预览.png"

    if not merged_candidates:
        raise RuntimeError("生成完成但找不到合并版 PDF。")
    if not single_zip.exists():
        raise RuntimeError("生成完成但找不到单人版 ZIP。")

    result = {
        "merged": merged_candidates[-1],
        "zip": single_zip,
    }
    if preview.exists():
        result["preview"] = preview
    return result


@app.post("/api/generate")
async def generate_cards(
    excel: UploadFile = File(...),
    exam_name: str = Form(""),
) -> dict[str, Any]:
    _validate_excel_filename(excel.filename)
    if not TEMPLATE_PDF.exists():
        raise HTTPException(status_code=500, detail="服务器缺少模板 PDF：template/answer_card_template.pdf")

    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_DIR / job_id
    upload_dir = job_dir / "upload"
    output_dir = job_dir / "output"
    upload_dir.mkdir(parents=True, exist_ok=True)

    excel_path = upload_dir / "学生信息.xlsx"
    try:
        with excel_path.open("wb") as f:
            shutil.copyfileobj(excel.file, f)

        students = read_students(excel_path)
        generate(
            excel_path=excel_path,
            template_pdf=TEMPLATE_PDF,
            output_dir=output_dir,
            exam_name=exam_name.strip() or None,
            make_preview=True,
        )
        files = _find_output_files(output_dir)
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        await excel.close()

    response: dict[str, Any] = {
        "job_id": job_id,
        "student_count": len(students),
        "merged_pdf_url": f"/download/{job_id}/merged",
        "single_zip_url": f"/download/{job_id}/zip",
        "merged_pdf_name": files["merged"].name,
        "single_zip_name": files["zip"].name,
    }
    if "preview" in files:
        response["preview_url"] = f"/download/{job_id}/preview"
    return response


@app.get("/download/{job_id}/{kind}")
def download(job_id: str, kind: str) -> FileResponse:
    if not job_id.replace("-", "").isalnum():
        raise HTTPException(status_code=400, detail="无效的任务 ID。")

    output_dir = JOBS_DIR / job_id / "output"
    if not output_dir.exists():
        raise HTTPException(status_code=404, detail="该生成任务不存在或已被清理。")

    try:
        files = _find_output_files(output_dir)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    if kind == "merged":
        path = files["merged"]
        return FileResponse(path, media_type="application/pdf", filename=path.name)
    if kind == "zip":
        path = files["zip"]
        return FileResponse(path, media_type="application/zip", filename=path.name)
    if kind == "preview" and "preview" in files:
        path = files["preview"]
        return FileResponse(path, media_type="image/png", filename=path.name)

    raise HTTPException(status_code=404, detail="找不到对应下载文件。")
