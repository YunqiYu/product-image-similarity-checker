"""任务工作区管理：为公司系统的每次上传任务创建隔离的中间文件目录。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class JobWorkspace:
    """单个任务的内部文件路径，业务前台只应拿到 final_excel。"""

    root: Path

    @property
    def downloaded_images(self) -> Path:
        return self.root / "images"

    @property
    def intermediate_excel(self) -> Path:
        return self.root / "similarity_rank.xlsx"

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
