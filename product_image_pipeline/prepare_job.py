"""数据准备入口：将上传的目标图和源 Excel 转换为相似度排序可直接使用的任务工作区。"""

from __future__ import annotations

import argparse
from pathlib import Path

from .input_adapter import PipelineInput
from .pipeline import prepare_similarity_job


def parse_args() -> argparse.Namespace:
    """解析本地测试或后端任务适配层传入的上传文件参数。"""
    parser = argparse.ArgumentParser(description="准备商品图片相似度排序任务。")
    parser.add_argument("--target-image", required=True)
    parser.add_argument("--source-excel", required=True)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--job-id", required=True)
    parser.add_argument("--jobs-dir", default="outputs/jobs")
    parser.add_argument("--image-url-column", default="图片链接")
    parser.add_argument("--max-candidates", type=int, default=1000)
    parser.add_argument("--random-seed", type=int, default=None)
    parser.add_argument("--selection-mode", choices=("random", "first"), default="random")
    return parser.parse_args()


def main() -> int:
    """创建工作区并下载随机抽样后的候选原图。"""
    args = parse_args()
    pipeline_input = PipelineInput(
        target_image_path=Path(args.target_image),
        source_excel_path=Path(args.source_excel),
        sheet_name=args.sheet,
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
    print(f"任务工作区: {workspace.root}")
    print(f"下载成功图片: {len(records)}")
    print(f"图片目录: {workspace.downloaded_images}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
