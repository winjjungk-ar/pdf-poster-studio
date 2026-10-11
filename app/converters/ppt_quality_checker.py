from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import math
import tempfile
import html
import shutil

import fitz
from PIL import Image, ImageChops, ImageStat


@dataclass
class PptQualityPageResult:
    page_no: int
    score: float
    diff_percent: float
    pdf_image: str
    ppt_image: str
    diff_image: str


@dataclass
class PptQualityReport:
    source_pdf: str
    pptx_path: str
    report_html: str
    average_score: float
    pages: list[PptQualityPageResult]


def _render_pdf_pages(
    source_pdf_path: str,
    output_dir: Path,
    page_indexes: list[int],
    dpi: int = 150,
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(source_pdf_path)
    paths: list[Path] = []

    try:
        zoom = max(dpi, 72) / 72.0

        for idx in page_indexes:
            if idx < 0 or idx >= len(doc):
                continue

            page = doc.load_page(idx)
            pix = page.get_pixmap(
                matrix=fitz.Matrix(zoom, zoom),
                alpha=False,
                colorspace=fitz.csRGB,
            )

            out = output_dir / f"pdf_page_{idx + 1:04d}.png"
            pix.save(str(out))
            paths.append(out)

    finally:
        doc.close()

    return paths


def _export_ppt_slides_to_png(
    pptx_path: str,
    output_dir: Path,
) -> list[Path]:
    """
    Export PPTX slides to PNG using installed Microsoft PowerPoint.

    Requires:
      pip install pywin32
      Microsoft PowerPoint installed
    """
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is not installed. Run: pip install pywin32"
        ) from exc

    output_dir.mkdir(parents=True, exist_ok=True)

    powerpoint = None
    presentation = None

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1

        presentation = powerpoint.Presentations.Open(
            str(Path(pptx_path).resolve()),
            WithWindow=False,
        )

        # 2 = ppSaveAsPNG
        presentation.SaveAs(str(output_dir.resolve()), 18)

    finally:
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass

        try:
            if powerpoint is not None:
                powerpoint.Quit()
        except Exception:
            pass

    files = sorted(output_dir.glob("*.PNG")) + sorted(output_dir.glob("*.png"))
    return files


def _resize_to_match(a: Image.Image, b: Image.Image) -> tuple[Image.Image, Image.Image]:
    if a.size == b.size:
        return a.convert("RGB"), b.convert("RGB")

    b = b.resize(a.size, Image.Resampling.LANCZOS)
    return a.convert("RGB"), b.convert("RGB")


def _compare_images(pdf_img_path: Path, ppt_img_path: Path, diff_path: Path) -> tuple[float, float]:
    pdf_img = Image.open(pdf_img_path)
    ppt_img = Image.open(ppt_img_path)

    pdf_img, ppt_img = _resize_to_match(pdf_img, ppt_img)

    diff = ImageChops.difference(pdf_img, ppt_img)
    stat = ImageStat.Stat(diff)

    # Mean RGB difference: 0~255
    mean_diff = sum(stat.mean[:3]) / 3.0
    diff_percent = min(100.0, max(0.0, (mean_diff / 255.0) * 100.0))
    score = max(0.0, 100.0 - diff_percent)

    # Enhance diff for visibility.
    enhanced = diff.point(lambda p: min(255, p * 4))
    enhanced.save(diff_path)

    return score, diff_percent


def create_ppt_quality_report(
    source_pdf_path: str,
    pptx_path: str,
    output_dir: str,
    page_indexes: list[int] | None = None,
    dpi: int = 150,
) -> PptQualityReport:
    source_pdf = Path(source_pdf_path)
    pptx = Path(pptx_path)
    report_dir = Path(output_dir)
    report_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(source_pdf))
    try:
        if page_indexes is None:
            page_indexes = list(range(len(doc)))
    finally:
        doc.close()

    pdf_dir = report_dir / "pdf_pages"
    ppt_dir = report_dir / "ppt_slides"
    diff_dir = report_dir / "diff"
    diff_dir.mkdir(parents=True, exist_ok=True)

    pdf_images = _render_pdf_pages(str(source_pdf), pdf_dir, page_indexes, dpi=dpi)
    ppt_images = _export_ppt_slides_to_png(str(pptx), ppt_dir)

    count = min(len(pdf_images), len(ppt_images))
    results: list[PptQualityPageResult] = []

    for i in range(count):
        diff_path = diff_dir / f"diff_{i + 1:04d}.png"
        score, diff_percent = _compare_images(pdf_images[i], ppt_images[i], diff_path)

        results.append(
            PptQualityPageResult(
                page_no=page_indexes[i] + 1 if i < len(page_indexes) else i + 1,
                score=score,
                diff_percent=diff_percent,
                pdf_image=str(pdf_images[i]),
                ppt_image=str(ppt_images[i]),
                diff_image=str(diff_path),
            )
        )

    average_score = (
        sum(item.score for item in results) / len(results)
        if results
        else 0.0
    )

    report_html = report_dir / "ppt_quality_report.html"
    _write_html_report(
        report_html=report_html,
        source_pdf=str(source_pdf),
        pptx_path=str(pptx),
        average_score=average_score,
        results=results,
    )

    return PptQualityReport(
        source_pdf=str(source_pdf),
        pptx_path=str(pptx),
        report_html=str(report_html),
        average_score=average_score,
        pages=results,
    )


def _write_html_report(
    report_html: Path,
    source_pdf: str,
    pptx_path: str,
    average_score: float,
    results: list[PptQualityPageResult],
) -> None:
    rows = []

    for item in results:
        rows.append(
            f"""
            <tr>
              <td>{item.page_no}</td>
              <td>{item.score:.2f}</td>
              <td>{item.diff_percent:.2f}%</td>
              <td><img src="{html.escape(Path(item.pdf_image).as_posix())}" width="220"></td>
              <td><img src="{html.escape(Path(item.ppt_image).as_posix())}" width="220"></td>
              <td><img src="{html.escape(Path(item.diff_image).as_posix())}" width="220"></td>
            </tr>
            """
        )

    body = "\n".join(rows)

    report_html.write_text(
        f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>PPT Quality Report</title>
<style>
body {{
  font-family: Arial, sans-serif;
  margin: 20px;
}}
table {{
  border-collapse: collapse;
  width: 100%;
}}
th, td {{
  border: 1px solid #ccc;
  padding: 8px;
  vertical-align: top;
}}
th {{
  background: #f3f3f3;
}}
.score {{
  font-size: 20px;
  font-weight: bold;
}}
</style>
</head>
<body>
<h1>PPT Quality Report</h1>
<p><b>Source PDF:</b> {html.escape(source_pdf)}</p>
<p><b>PPTX:</b> {html.escape(pptx_path)}</p>
<p class="score">Average Score: {average_score:.2f}</p>

<table>
<thead>
<tr>
  <th>Page</th>
  <th>Score</th>
  <th>Diff</th>
  <th>PDF Render</th>
  <th>PPT Render</th>
  <th>Difference</th>
</tr>
</thead>
<tbody>
{body}
</tbody>
</table>
</body>
</html>
""",
        encoding="utf-8",
    )
