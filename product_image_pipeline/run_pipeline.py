"""一键任务入口：串联图片准备、相似度排序和五要素打标，最终只保留业务 Excel。"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

from .input_adapter import PipelineInput
from .pipeline import prepare_similarity_job


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def format_duration(seconds: float) -> str:
    """将运行秒数格式化为时分秒，便于写入总任务日志。"""
    total_seconds = int(round(seconds))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def parse_args() -> argparse.Namespace:
    """解析公司系统上传文件或本地调试所需的统一任务参数。"""
    parser = argparse.ArgumentParser(description="一键执行商品图片相似度排序和五要素打标。")
    parser.add_argument("--target-image", required=True)
    parser.add_argument("--source-excel", required=True)
    parser.add_argument("--source-sheet", required=True)
    parser.add_argument("--job-id", required=True)
    parser.add_argument("--jobs-dir", default="outputs/jobs")
    parser.add_argument("--image-url-column", default="图片链接")
    parser.add_argument("--max-candidates", type=int, default=1000)
    parser.add_argument("--max-results", type=int, default=300)
    parser.add_argument("--random-seed", type=int, default=None)
    parser.add_argument("--selection-mode", choices=("random", "first"), default="random", help="候选抽样方式；测试源表前 N 条时使用 first。")
    parser.add_argument("--weights", default=None, help="四维相似度权重 JSON。")
    parser.add_argument("--similarity-model", default=None)
    parser.add_argument("--tagging-model", default=None)
    parser.add_argument("--similarity-workers", type=int, default=3)
    parser.add_argument("--tagging-workers", type=int, default=3)
    parser.add_argument("--tagging-attempts", type=int, default=3)
    parser.add_argument("--exclude-first-candidate", action="store_true", help="将源表第一条图片仅作为目标图，不参与竞品排序。")
    parser.add_argument("--keep-intermediate", action="store_true", help="成功后保留内部 similarity_rank.json，仅供开发排查。")
    return parser.parse_args()


def run_module(arguments: list[str]) -> None:
    """以当前虚拟环境执行子模块，失败时立即停止后续业务输出。"""
    subprocess.run([sys.executable, "-m", *arguments], cwd=PROJECT_ROOT, check=True)


def main() -> int:
    """执行一次完整任务，成功时仅向业务侧留下最终 Excel。"""
    args = parse_args()
    started = time.perf_counter()
    pipeline_input = PipelineInput(
        target_image_path=Path(args.target_image),
        source_excel_path=Path(args.source_excel),
        sheet_name=args.source_sheet,
        image_url_column=args.image_url_column,
    )
    workspace, records = prepare_similarity_job(
        pipeline_input,
        Path(args.jobs_dir),
        args.job_id,
        max_candidates=args.max_candidates,
        random_seed=args.random_seed,
        selection_mode=args.selection_mode,
    )
    if args.exclude_first_candidate and records:
        # 第一条源数据用于下载目标图时，避免它作为候选与自身产生 100% 相似度。
        target_record = min(records, key=lambda record: record.source_row)
        (workspace.downloaded_images / target_record.filename).unlink(missing_ok=True)
        records = [record for record in records if record is not target_record]
    if not records:
        raise RuntimeError("没有成功下载的候选图片，任务终止。")

    ranking_arguments = [
        "product_image_pipeline.similarity_ranker",
        "--input-dir", str(workspace.downloaded_images),
        "--output", str(workspace.intermediate_json),
        "--run-log", str(workspace.root / "similarity_rank.log"),
        "--metadata-excel", str(pipeline_input.source_excel_path),
        "--metadata-sheet", pipeline_input.sheet_name,
        "--max-candidates", str(args.max_candidates),
        "--max-results", str(args.max_results),
        "--workers", str(args.similarity_workers),
    ]
    if args.weights:
        ranking_arguments.extend(["--weights", args.weights])
    if args.similarity_model:
        ranking_arguments.extend(["--model", args.similarity_model])
    run_module(ranking_arguments)

    tagging_arguments = [
        "product_image_pipeline.tagging_pipeline",
        "--input-json", str(workspace.intermediate_json),
        "--input-dir", str(workspace.downloaded_images),
        "--source-excel", str(pipeline_input.source_excel_path),
        "--source-sheet", pipeline_input.sheet_name,
        "--output", str(workspace.final_excel),
        "--run-log", str(workspace.root / "tagging.log"),
        "--workers", str(args.tagging_workers),
        "--tagging-attempts", str(args.tagging_attempts),
    ]
    if args.tagging_model:
        tagging_arguments.extend(["--model", args.tagging_model])
    run_module(tagging_arguments)

    if not args.keep_intermediate:
        workspace.intermediate_json.unlink(missing_ok=True)
    workspace.run_log.write_text(
        f"任务状态: 成功\n"
        f"下载成功图片: {len(records)}\n"
        f"最终结果: {workspace.final_excel}\n"
        f"本次运行总时间: {format_duration(time.perf_counter() - started)}\n",
        encoding="utf-8",
    )
    print(f"一键任务完成，最终 Excel: {workspace.final_excel}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
