"""任务工作区管理：为公司系统的每次上传任务创建隔离的中间文件目录。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from shutil import copy2

from .input_adapter import CandidateRecord, PipelineInput, prepare_uploaded_excel


@dataclass(frozen=True)
class JobWorkspace:
    """单个任务的内部文件路径，业务前台只应拿到 final_excel。"""

    root: Path

    @property
    def downloaded_images(self) -> Path:
        return self.root / "images"

    @property
    def intermediate_json(self) -> Path:
        """阶段一内部排序结果，供阶段二直接读取，不对业务前台暴露。"""
        return self.root / "similarity_rank.json"

    @property
    def final_excel(self) -> Path:
        return self.root / "final_tagged.xlsx"

    @property
    def run_log(self) -> Path:
        return self.root / "run.log"


def create_workspace(base_dir: Path, job_id: str) -> JobWorkspace:
    """按任务 ID 创建隔离目录，避免并发任务相互覆盖。"""
    workspace = JobWorkspace(base_dir / job_id)
    workspace.downloaded_images.mkdir(parents=True, exist_ok=True)
    return workspace


def prepare_similarity_job(
    pipeline_input: PipelineInput,
    base_dir: Path,
    job_id: str,
    *,
    max_candidates: int = 1000,
    random_seed: int | None = None,
    selection_mode: str = "random",
) -> tuple[JobWorkspace, list[CandidateRecord]]:
    """为上传任务准备目标图、本地候选图和 manifest，供相似度排序模块直接读取。"""
    workspace = create_workspace(base_dir, job_id)
    target_suffix = pipeline_input.target_image_path.suffix or ".png"
    target_copy = workspace.downloaded_images / f"目标图片{target_suffix}"
    copy2(pipeline_input.target_image_path, target_copy)
    records = prepare_uploaded_excel(
        pipeline_input,
        workspace.downloaded_images,
        max_candidates=max_candidates,
        random_seed=random_seed,
        selection_mode=selection_mode,
    )
    return workspace, records
