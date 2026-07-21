"""集中读取模型名称和图片细节等级等运行配置。"""

from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass(frozen=True)
class ModelSettings:
    """保存相似度排序和五要素打标分别使用的模型配置。"""

    base_url: str | None
    api_key: str | None
    similarity_model: str
    tagging_model: str
    image_detail: str

    @classmethod
    def from_environment(cls) -> "ModelSettings":
        """兼容旧环境变量，并优先读取拆分后的模型变量。"""
        return cls(
            base_url=os.getenv("OPENAI_BASE_URL") or os.getenv("LLM_BASE_URL"),
            api_key=os.getenv("OPENAI_API_KEY") or os.getenv("TEXT_API_KEY"),
            similarity_model=os.getenv("SIMILARITY_MODEL") or os.getenv("OPENAI_MODEL") or "qwen3.7-plus",
            tagging_model=os.getenv("TAGGING_MODEL") or "gpt-5.5",
            image_detail=os.getenv("OPENAI_IMAGE_MODE") or "high",
        )
