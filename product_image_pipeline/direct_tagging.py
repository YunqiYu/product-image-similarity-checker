"""直接五要素打标入口：保持源表顺序，不经过相似度排序。"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from .input_adapter import (
    CandidateRecord,
    download_record,
    read_candidate_records,
    sample_candidate_records,
    write_manifest,
)
from .pipeline import create_workspace


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def parse_args() -> argparse.Namespace:
    """解析直接打标所需的源表与运行参数。"""
    parser = argparse.ArgumentParser(description="直接为 Excel 图片链接补齐五要素标签。")
    parser.add_argument("--source-excel", required=True)
    parser.add_argument("--source-sheet", required=True)
    parser.add_argument("--job-id", required=True)
    parser.add_argument("--jobs-dir", default="outputs/jobs")
    parser.add_argument("--image-url-column", default="图片链接")
    parser.add_argument("--max-candidates", type=int, default=1000)
    parser.add_argument("--workers", type=int, default=3)
    parser.add_argument("--tagging-attempts", type=int, default=3)
    parser.add_argument("--tagging-model", default=None)
    return parser.parse_args()


def format_duration(seconds: float) -> str:
    """将耗时格式化为时分秒。"""
    total_seconds = int(round(seconds))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def download_candidates(records: list[CandidateRecord], images_dir: Path) -> list[CandidateRecord]:
    """并发下载候选原图，下载失败的记录保留在 manifest 中。"""
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(download_record, record, images_dir, 60.0) for record in records]
        downloaded = [future.result() for future in as_completed(futures)]
    # 下载完成顺序不固定，写出结果前按源表行号恢复业务排序。
    downloaded.sort(key=lambda record: record.source_row)
    write_manifest(downloaded, images_dir / "manifest.tsv")
    return [record for record in downloaded if not record.error]


def main() -> int:
    """下载源表图片并直接写出五要素打标结果。"""
    args = parse_args()
    started = time.perf_counter()
    source_excel = Path(args.source_excel)
    if not source_excel.exists():
        raise FileNotFoundError(f"找不到源 Excel：{source_excel}")

    workspace = create_workspace(Path(args.jobs_dir), args.job_id)
    records = sample_candidate_records(
        read_candidate_records(source_excel, args.source_sheet, args.image_url_column),
        args.max_candidates,
        selection_mode="first",
    )
    downloaded = download_candidates(records, workspace.downloaded_images)
    if not downloaded:
        raise RuntimeError("没有成功下载的图片，任务终止。")

    # 标签阶段沿用既有 JSON 协议，但不写入任何相似度数据。
    items = [
        {
            "source_row": record.source_row,
            "asin": "",
            "image_url": record.image_url,
            "image_name": record.filename,
            "overall_similarity": "",
            "similarities": {},
            "error": "",
        }
        for record in downloaded
    ]
    workspace.intermediate_json.write_text(
        json.dumps({"schema_version": 1, "candidate_count": len(downloaded), "result_count": len(items), "items": items}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    command = [
        sys.executable,
        "-m",
        "product_image_pipeline.tagging_pipeline",
        "--input-json",
        str(workspace.intermediate_json),
        "--input-dir",
        str(workspace.downloaded_images),
        "--source-excel",
        str(source_excel),
        "--source-sheet",
        args.source_sheet,
        "--output",
        str(workspace.final_excel),
        "--run-log",
        str(workspace.root / "tagging.log"),
        "--workers",
        str(args.workers),
        "--tagging-attempts",
        str(args.tagging_attempts),
    ]
    if args.tagging_model:
        command.extend(["--model", args.tagging_model])
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)

    workspace.intermediate_json.unlink(missing_ok=True)
    workspace.run_log.write_text(
        f"任务状态: 成功\n下载成功图片: {len(downloaded)}\n最终结果: {workspace.final_excel}\n本次运行总时间: {format_duration(time.perf_counter() - started)}\n",
        encoding="utf-8",
    )
    print(f"直接打标完成，最终 Excel: {workspace.final_excel}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
