from __future__ import annotations

import argparse
import io
import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
import qrcode
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

from config import CONFIG


@dataclass
class Student:
    index: int
    name: str
    student_id: str


def top_to_bottom_y(page_height: float, top_y: float) -> float:
    """把“距离页面顶部”的 y 坐标转成 ReportLab 的底部原点坐标。"""
    return page_height - top_y


def rect_top_to_reportlab(page_height: float, x: float, y_top: float, w: float, h: float) -> tuple[float, float, float, float]:
    """矩形：左上坐标 -> ReportLab 左下坐标。"""
    return x, page_height - y_top - h, w, h


def safe_filename(text: str) -> str:
    """生成 Windows/macOS/Linux 都相对安全的文件名片段。"""
    text = str(text).strip()
    return re.sub(r'[\\/:*?"<>|\s]+', '_', text)


def normalize_header(value: Any) -> str:
    return str(value).strip().replace(" ", "").replace("\n", "") if value is not None else ""


def find_column(headers: list[str], candidates: list[str], field_name: str) -> int:
    normalized = [normalize_header(h) for h in headers]
    candidate_set = {normalize_header(c) for c in candidates}
    for idx, header in enumerate(normalized):
        if header in candidate_set:
            return idx
    raise ValueError(
        f"Excel 中找不到 {field_name} 列。当前表头：{headers}；可接受列名：{candidates}"
    )


def read_students(excel_path: Path) -> list[Student]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel 是空的。")

    headers = list(rows[0])
    name_col = find_column(headers, CONFIG["excel_columns"]["name"], "姓名")
    id_col = find_column(headers, CONFIG["excel_columns"]["student_id"], "识别号")

    students: list[Student] = []
    for i, row in enumerate(rows[1:], start=1):
        if row is None:
            continue
        name = row[name_col] if name_col < len(row) else None
        student_id = row[id_col] if id_col < len(row) else None
        if name is None and student_id is None:
            continue
        if name is None or student_id is None:
            raise ValueError(f"第 {i + 1} 行缺少姓名或识别号：{row}")

        students.append(Student(index=i, name=str(name).strip(), student_id=str(student_id).strip()))

    if not students:
        raise ValueError("Excel 中没有可用学生数据。")
    return students


def make_qr_image(content: str, border: int = 1) -> Image.Image:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=border,
    )
    qr.add_data(str(content))
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def make_overlay_pdf(
    overlay_path: Path,
    page_width: float,
    page_height: float,
    student: Student,
    exam_name: str | None = None,
) -> None:
    font_name = CONFIG["font_name"]
    font_size = CONFIG["font_size"]
    pdfmetrics.registerFont(UnicodeCIDFont(font_name))

    c = canvas.Canvas(str(overlay_path), pagesize=(page_width, page_height))

    # 1) 白色遮罩：盖掉模板中原姓名、原考号、原二维码。
    c.setFillColorRGB(1, 1, 1)
    c.setStrokeColorRGB(1, 1, 1)
    for mask in CONFIG["white_masks"]:
        x, y, w, h = rect_top_to_reportlab(page_height, mask["x"], mask["y"], mask["w"], mask["h"])
        c.rect(x, y, w, h, stroke=0, fill=1)

    # 2) 考试名称与横线。
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)
    c.setFont(font_name, font_size)

    exam_cfg = CONFIG["exam_title"]
    c.drawString(exam_cfg["x"], top_to_bottom_y(page_height, exam_cfg["baseline_y"]), exam_cfg["label"])

    # 默认画手写横线；如果提供 exam_name，也保留横线作为视觉基线，并把文字写在横线上方。
    c.setLineWidth(exam_cfg["line_width"])
    c.line(
        exam_cfg["line_start_x"],
        top_to_bottom_y(page_height, exam_cfg["line_y"]),
        exam_cfg["line_end_x"],
        top_to_bottom_y(page_height, exam_cfg["line_y"]),
    )
    if exam_name:
        c.drawString(
            exam_cfg["value_x"],
            top_to_bottom_y(page_height, exam_cfg["value_baseline_y"]),
            str(exam_name),
        )

    # 3) 学生姓名与识别号。
    text_cfg = CONFIG["student_text"]
    c.setFont(font_name, font_size)
    c.drawString(
        text_cfg["name_x"],
        top_to_bottom_y(page_height, text_cfg["name_baseline_y"]),
        f'{text_cfg["name_prefix"]}{student.name}',
    )
    c.drawString(
        text_cfg["id_x"],
        top_to_bottom_y(page_height, text_cfg["id_baseline_y"]),
        f'{text_cfg["id_prefix"]}{student.student_id}',
    )

    # 4) 二维码：内容为识别号。
    qr_cfg = CONFIG["qr"]
    qr_img = make_qr_image(student.student_id, border=qr_cfg["border"])
    img_buffer = io.BytesIO()
    qr_img.save(img_buffer, format="PNG")
    img_buffer.seek(0)
    c.drawImage(
        ImageReader(img_buffer),
        qr_cfg["x"],
        page_height - qr_cfg["y"] - qr_cfg["size"],
        width=qr_cfg["size"],
        height=qr_cfg["size"],
        preserveAspectRatio=True,
        mask="auto",
    )

    # 5) 重绘四个视觉定位小黑块，放在最上层，防止被遮罩或文字影响。
    loc_cfg = CONFIG["locator_squares"]
    size = loc_cfg["size"]
    original_size = loc_cfg["original_size"]
    shift = (size - original_size) / 2
    c.setFillColorRGB(0, 0, 0)
    for item in loc_cfg["items"]:
        x_top = item["x"] - shift
        y_top = item["y"] - shift
        x, y, w, h = rect_top_to_reportlab(page_height, x_top, y_top, size, size)
        c.rect(x, y, w, h, stroke=0, fill=1)

    c.save()


def merge_template_and_overlay(template_pdf: Path, overlay_pdf: Path, output_pdf: Path) -> None:
    base_reader = PdfReader(str(template_pdf))
    overlay_reader = PdfReader(str(overlay_pdf))

    writer = PdfWriter()
    page = base_reader.pages[0]
    page.merge_page(overlay_reader.pages[0])
    writer.add_page(page)

    with output_pdf.open("wb") as f:
        writer.write(f)


def render_preview(pdf_path: Path, preview_path: Path, zoom: float = 2.0) -> bool:
    """生成第一页预览图。PyMuPDF 为可选依赖，如果失败不影响主流程。"""
    try:
        import fitz  # PyMuPDF

        doc = fitz.open(str(pdf_path))
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        pix.save(str(preview_path))
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"[提示] 预览图生成失败，但 PDF 已正常生成：{exc}")
        return False


def generate(
    excel_path: Path,
    template_pdf: Path,
    output_dir: Path,
    exam_name: str | None = None,
    make_preview: bool = True,
) -> None:
    if not excel_path.exists():
        raise FileNotFoundError(f"找不到 Excel：{excel_path}")
    if not template_pdf.exists():
        raise FileNotFoundError(f"找不到模板 PDF：{template_pdf}")

    students = read_students(excel_path)

    output_dir.mkdir(parents=True, exist_ok=True)
    single_dir = output_dir / "单人版"
    temp_dir = output_dir / "_temp_overlay"

    if single_dir.exists():
        shutil.rmtree(single_dir)
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    single_dir.mkdir(parents=True, exist_ok=True)
    temp_dir.mkdir(parents=True, exist_ok=True)

    template_reader = PdfReader(str(template_pdf))
    first_page = template_reader.pages[0]
    page_width = float(first_page.mediabox.width)
    page_height = float(first_page.mediabox.height)

    merged_writer = PdfWriter()
    generated_files: list[Path] = []

    for idx, student in enumerate(students, start=1):
        overlay_path = temp_dir / f"overlay_{idx:03d}.pdf"
        single_pdf = single_dir / f"{idx:02d}_{safe_filename(student.student_id)}_{safe_filename(student.name)}.pdf"

        make_overlay_pdf(
            overlay_path=overlay_path,
            page_width=page_width,
            page_height=page_height,
            student=student,
            exam_name=exam_name,
        )
        merge_template_and_overlay(template_pdf, overlay_path, single_pdf)
        generated_files.append(single_pdf)

        # 合并版也从单人 PDF 读取，确保与单人版完全一致。
        single_reader = PdfReader(str(single_pdf))
        merged_writer.add_page(single_reader.pages[0])

    merged_pdf = output_dir / f"批量答题卡_{len(students)}人合并版.pdf"
    with merged_pdf.open("wb") as f:
        merged_writer.write(f)

    zip_path = output_dir / "批量答题卡_单人版.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for pdf in generated_files:
            zf.write(pdf, arcname=pdf.name)

    if make_preview and generated_files:
        preview_path = output_dir / "第一页预览.png"
        render_preview(generated_files[0], preview_path)

    shutil.rmtree(temp_dir, ignore_errors=True)

    print("生成完成")
    print(f"学生数量：{len(students)}")
    print(f"合并版 PDF：{merged_pdf}")
    print(f"单人版目录：{single_dir}")
    print(f"单人版 ZIP：{zip_path}")
    if make_preview:
        print(f"第一页预览：{output_dir / '第一页预览.png'}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="批量生成答题卡 PDF")
    parser.add_argument("--excel", default="input/学生信息.xlsx", help="学生信息 Excel 路径")
    parser.add_argument("--template", default="template/answer_card_template.pdf", help="答题卡模板 PDF 路径")
    parser.add_argument("--out", default="output", help="输出目录")
    parser.add_argument("--exam-name", default="", help="考试名称；为空时保留横线供手写")
    parser.add_argument("--no-preview", action="store_true", help="不生成第一页 PNG 预览")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    generate(
        excel_path=Path(args.excel),
        template_pdf=Path(args.template),
        output_dir=Path(args.out),
        exam_name=args.exam_name.strip() or None,
        make_preview=not args.no_preview,
    )
