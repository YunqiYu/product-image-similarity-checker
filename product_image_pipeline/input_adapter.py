"""将上传的 Excel 或图片列表转换为流水线可处理的标准候选记录。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PipelineInput:
    """描述一次任务上传后的文件位置和 Excel 字段约定。"""

    target_image_path: Path
    source_excel_path: Path
    sheet_name: str
    image_url_column: str = "图片链接"
