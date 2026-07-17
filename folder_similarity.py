from __future__ import annotations

import argparse
import base64
import csv
import json
import mimetypes
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment
from PIL import Image as PILImage

from similarity import format_duration, parse_percent


DEFAULT_INPUT_DIR = Path(__file__).parent / "inputs" / "测试"
DEFAULT_OUTPUT = Path(__file__).parent / "outputs" / "测试_图片相似度结果.xlsx"
DEFAULT_RUN_LOG = Path(__file__).parent / "outputs" / "测试_图片相似度结果.log"
DEFAULT_METADATA_EXCEL = Path(__file__).parent / "inputs" / "一品红表格 - 打标核对.xlsx"
DEFAULT_METADATA_SHEET = "最终结果"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
OUTPUT_HEADERS = [
    "1级分类",
    "主题标签",
    "ASIN",
    "图片链接",
    "图片",
    "产品链接",
    "品牌",
    "pcs",
    "标题",
    "近30天销量",
    "上架时间",
    "价格",
    "数据来源",
    "价格趋势图",
    "综合相似度",
    "颜色",
    "风格",
    "元素",
    "排版",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare competitor images in a folder against one target image.")
    parser.add_argument("--input-dir", default=str(DEFAULT_INPUT_DIR))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--run-log", default=str(DEFAULT_RUN_LOG))
    parser.add_argument("--model", default=None)
    parser.add_argument("--workers", type=int, default=3)
    parser.add_argument("--image-detail", default=None)
    parser.add_argument("--request-timeout", type=float, default=60.0)
    parser.add_argument("--max-retries", type=int, default=1)
    parser.add_argument("--similarity-threshold", type=float, default=80.0)
    parser.add_argument("--metadata-excel", default=str(DEFAULT_METADATA_EXCEL))
    parser.add_argument("--metadata-sheet", default=DEFAULT_METADATA_SHEET)
    return parser.parse_args()


def image_to_data_url(path: Path) -> str:
    mime_type = mimetypes.guess_type(path.name)[0] or "image/png"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{data}"


def find_target_image(input_dir: Path) -> Path:
    for path in input_dir.iterdir():
        if path.is_file() and "目标" in path.stem and path.suffix.lower() in IMAGE_EXTENSIONS:
            return path
    raise FileNotFoundError(f"No target image found in {input_dir}. Expected a file name containing '目标'.")


def list_candidate_images(input_dir: Path, target_image: Path) -> list[Path]:
    images = [
        path
        for path in input_dir.iterdir()
        if path.is_file() and path != target_image and path.suffix.lower() in IMAGE_EXTENSIONS
    ]
    return sorted(images, key=lambda path: natural_key(path.stem))


def natural_key(text: str) -> list[Any]:
    parts = re.split(r"(\d+)", text)
    return [int(part) if part.isdigit() else part.lower() for part in parts]


def extract_asin_from_stem(stem: str) -> str:
    match = re.search(r"(B0[A-Z0-9]{8}|B[A-Z0-9]{9})", stem, flags=re.IGNORECASE)
    return match.group(1).upper() if match else stem


def load_manifest(input_dir: Path) -> dict[str, dict[str, str]]:
    manifest_path = input_dir / "manifest.tsv"
    if not manifest_path.exists():
        return {}
    with manifest_path.open("r", encoding="utf-8-sig", newline="") as file:
        return {
            row["filename"]: row
            for row in csv.DictReader(file, delimiter="\t")
            if row.get("filename")
        }


def load_excel_metadata(excel_path: Path, sheet_name: str) -> dict[int, dict[str, Any]]:
    if not excel_path.exists():
        return {}
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    if sheet_name not in workbook.sheetnames:
        return {}
    worksheet = workbook[sheet_name]
    rows = worksheet.iter_rows(values_only=True)
    headers = [str(value).strip() if value is not None else "" for value in next(rows)]
    metadata: dict[int, dict[str, Any]] = {}
    for row_number, row_values in enumerate(rows, start=2):
        metadata[row_number] = dict(zip(headers, row_values))
    return metadata


def metadata_for_candidate(
    candidate: Path,
    manifest: dict[str, dict[str, str]],
    excel_metadata: dict[int, dict[str, Any]],
) -> dict[str, Any]:
    row: dict[str, Any] = {}
    manifest_row = manifest.get(candidate.name, {})
    row_number = manifest_row.get("row")
    if row_number and row_number.isdigit():
        row.update(excel_metadata.get(int(row_number), {}))
    if manifest_row:
        row.setdefault("ASIN", manifest_row.get("asin"))
        row.setdefault("图片链接", manifest_row.get("original_url") or manifest_row.get("used_url_size"))
    row.setdefault("ASIN", extract_asin_from_stem(candidate.stem))
    row.setdefault("图片链接", candidate.name)
    return row


def build_compare_prompt(threshold: float) -> str:
    return f"""你是资深商品图片视觉相似度评估专家。
现在给你两张商品图片：
- 第一张是目标图；
- 第二张是竞品图。

请只评估商品正面印刷图案，不要因为商品类型、摆放角度、盘子/纸巾数量、包装组合、拍摄背景、阴影、页面文字而扣分。

请从四个维度分别给出 0-100 的相似度：
1. color相似度：比较主要视觉配色是否接近，忽略白色摄影背景。
2. style相似度：比较插画、矢量、水彩、写实、特效等视觉风格是否接近。
3. elements相似度：比较主体元素、主题元素、主设计文案和艺术字内容是否一致或语义接近，例如一品红/圣诞花、松枝/冬青/圣诞叶可以视为高度相近；如果有“Merry Christmas”等主文案，也在这个维度中一起判断。
4. format相似度：比较散排、居中、环绕、铺满、文案居中等构图关系是否接近。

综合相似度公式：
综合相似度 = elements相似度*45% + style相似度*20% + color相似度*15% + format相似度*20%

如果两张图片本质上是同一设计、同一主图或只是分辨率不同，应给出接近或等于100%的分数。
相似阈值为 {threshold}%。

只输出合法 JSON，不要输出解释性段落或 Markdown。格式如下：
{{
  "color相似度": 98,
  "style相似度": 95,
  "elements相似度": 95,
  "format相似度": 96,
  "相似理由": "两张图在主体元素、风格、颜色、版式和主文案上高度一致"
}}
"""


def normalize_similarity_payload(data: dict[str, Any], threshold: float) -> dict[str, str]:
    color = parse_percent(data.get("color相似度", data.get("颜色相似度", data.get("color_similarity"))))
    style = parse_percent(data.get("style相似度", data.get("风格相似度", data.get("style_similarity"))))
    elements = parse_percent(data.get("elements相似度", data.get("元素相似度", data.get("elements_similarity"))))
    fmt = parse_percent(data.get("format相似度", data.get("版式相似度", data.get("format_similarity"))))
    overall = elements * 0.45 + style * 0.20 + color * 0.15 + fmt * 0.20
    return {
        "综合相似度": f"{overall:.1f}%",
        "相似理由": str(data.get("相似理由", data.get("reason", ""))).strip(),
        "color相似度": f"{round(color)}%",
        "style相似度": f"{round(style)}%",
        "elements相似度": f"{round(elements)}%",
        "format相似度": f"{round(fmt)}%",
        "是否相似": "是" if overall >= threshold else "否",
    }


def compare_pair(
    client: Any,
    model: str,
    target_url: str,
    candidate_url: str,
    image_detail: str,
    threshold: float,
) -> dict[str, str]:
    from tagger import extract_json

    response = client.chat.completions.create(
        model=model,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": build_compare_prompt(threshold)},
                    {"type": "image_url", "image_url": {"url": target_url, "detail": image_detail}},
                    {"type": "image_url", "image_url": {"url": candidate_url, "detail": image_detail}},
                ],
            }
        ],
    )
    return normalize_similarity_payload(extract_json(response.choices[0].message.content or ""), threshold)


def percent_sort_value(value: Any) -> float:
    try:
        return parse_percent(value)
    except Exception:
        return -1.0


def make_excel_thumbnail(image_path: Path, thumbnail_dir: Path, size: int = 220) -> Path:
    thumbnail_dir.mkdir(parents=True, exist_ok=True)
    thumbnail_path = thumbnail_dir / f"{image_path.stem}.png"
    with PILImage.open(image_path) as source:
        source.thumbnail((size, size))
        canvas = PILImage.new("RGB", (size, size), "white")
        x = (size - source.width) // 2
        y = (size - source.height) // 2
        canvas.paste(source.convert("RGB"), (x, y))
        canvas.save(thumbnail_path, "PNG")
    return thumbnail_path


def add_image(ws: Any, image_path: Path, cell: str, thumbnail_dir: Path, size: int = 220) -> None:
    thumbnail_path = make_excel_thumbnail(image_path, thumbnail_dir, size=size)
    image = ExcelImage(str(thumbnail_path))
    image.width = size
    image.height = size
    ws.add_image(image, cell)


def process_candidate(
    candidate: Path,
    client: Any,
    model: str,
    target_url: str,
    image_detail: str,
    threshold: float,
) -> dict[str, Any]:
    started = time.perf_counter()
    try:
        similarities = compare_pair(client, model, target_url, image_to_data_url(candidate), image_detail, threshold)
        error = ""
    except Exception as exc:
        similarities = {
            "综合相似度": "",
            "相似理由": "",
            "color相似度": "",
            "style相似度": "",
            "elements相似度": "",
            "format相似度": "",
            "是否相似": "",
        }
        error = str(exc)
    return {
        "path": candidate,
        "similarities": similarities,
        "error": error,
        "elapsed": time.perf_counter() - started,
    }


def main() -> int:
    from dotenv import load_dotenv
    from openai import OpenAI

    args = parse_args()
    load_dotenv(Path(__file__).parent / ".env")

    input_dir = Path(args.input_dir)
    output_path = Path(args.output)
    run_log_path = Path(args.run_log)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    run_log_path.parent.mkdir(parents=True, exist_ok=True)
    run_log_path.write_text("", encoding="utf-8")

    target_image = find_target_image(input_dir)
    candidates = list_candidate_images(input_dir, target_image)
    if not candidates:
        raise RuntimeError(f"No candidate images found in {input_dir}")
    manifest = load_manifest(input_dir)
    excel_metadata = load_excel_metadata(Path(args.metadata_excel), args.metadata_sheet)

    api_key = os.getenv("OPENAI_API_KEY") or os.getenv("TEXT_API_KEY")
    if not api_key:
        raise RuntimeError("Missing OPENAI_API_KEY or TEXT_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL") or os.getenv("LLM_BASE_URL")
    model = args.model or os.getenv("OPENAI_MODEL") or "gpt-5.5"
    image_detail = args.image_detail or os.getenv("OPENAI_IMAGE_MODE") or "high"
    client = OpenAI(api_key=api_key, base_url=base_url, timeout=args.request_timeout, max_retries=args.max_retries)

    target_url = image_to_data_url(target_image)
    run_started = time.perf_counter()

    results: dict[Path, dict[str, Any]] = {}
    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {
            executor.submit(process_candidate, candidate, client, model, target_url, image_detail, args.similarity_threshold): candidate
            for candidate in candidates
        }
        done = 0
        for future in as_completed(futures):
            candidate = futures[future]
            result = future.result()
            results[candidate] = result
            done += 1
            with run_log_path.open("a", encoding="utf-8") as log:
                log.write(
                    f"{candidate.name}\t{json.dumps(result['similarities'], ensure_ascii=False)}"
                    f"\t运行时间: {format_duration(result['elapsed'])}"
                    + (f"\t错误: {result['error']}" if result["error"] else "")
                    + "\n"
                )
            print(f"[{done}/{len(candidates)}] {candidate.name} {json.dumps(result['similarities'], ensure_ascii=False)}")

    wb = Workbook()
    ws = wb.active
    ws.title = "相似度结果"
    for col, header in enumerate(OUTPUT_HEADERS, start=1):
        ws.cell(row=1, column=col).value = header
    column_widths = {
        "A": 14,
        "B": 18,
        "C": 18,
        "D": 46,
        "E": 34,
        "F": 42,
        "G": 18,
        "H": 10,
        "I": 46,
        "J": 14,
        "K": 16,
        "L": 12,
        "M": 16,
        "N": 22,
        "O": 14,
        "P": 14,
        "Q": 14,
        "R": 14,
        "S": 14,
    }
    for column, width in column_widths.items():
        ws.column_dimensions[column].width = width

    sorted_candidates = sorted(
        candidates,
        key=lambda candidate: percent_sort_value(results[candidate]["similarities"].get("综合相似度")),
        reverse=True,
    )
    thumbnail_dir = output_path.parent / "_excel_image_cache" / output_path.stem

    for row_index, candidate in enumerate(sorted_candidates, start=2):
        result = results[candidate]
        similarities = result["similarities"]
        source_row = metadata_for_candidate(candidate, manifest, excel_metadata)
        for col_index, header in enumerate(OUTPUT_HEADERS, start=1):
            if header == "图片":
                continue
            ws.cell(row=row_index, column=col_index).value = source_row.get(header, "")
        ws.cell(row=row_index, column=3).value = source_row.get("ASIN") or extract_asin_from_stem(candidate.stem)
        ws.row_dimensions[row_index].height = 170
        ws.cell(row=row_index, column=5).value = candidate.name
        ws.cell(row=row_index, column=5).alignment = Alignment(horizontal="center", vertical="center")
        add_image(ws, candidate, f"E{row_index}", thumbnail_dir)
        ws.cell(row=row_index, column=15).value = similarities["综合相似度"]

    wb.save(output_path)
    with run_log_path.open("a", encoding="utf-8") as log:
        log.write(f"本次运行总时间: {format_duration(time.perf_counter() - run_started)}\n")
    print(f"Target image: {target_image}")
    print(f"Output: {output_path}")
    print(f"Run log: {run_log_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
