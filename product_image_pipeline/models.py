"""定义图片相似度与打标签任务共用的数据结构。"""

from __future__ import annotations

import json
from dataclasses import dataclass


@dataclass(frozen=True)
class SimilarityWeights:
    """保存四个相似度维度的权重，并负责校验和计算综合分数。"""

    color: float = 20.0
    style: float = 30.0
    elements: float = 20.0
    format: float = 30.0

    def __post_init__(self) -> None:
        values = (self.color, self.style, self.elements, self.format)
        if any(value < 0 or value > 100 for value in values):
            raise ValueError("相似度权重必须在 0 到 100 之间。")
        if abs(sum(values) - 100) > 1e-6:
            raise ValueError("颜色、风格、元素、排版四项权重之和必须为 100。")

    @classmethod
    def from_json(cls, value: str) -> "SimilarityWeights":
        """解析前端或命令行传入的 JSON 权重配置。"""
        data = json.loads(value)
        if not isinstance(data, dict):
            raise ValueError("权重配置必须是 JSON 对象。")
        return cls(
            color=float(data.get("color", 0)),
            style=float(data.get("style", 0)),
            elements=float(data.get("elements", 0)),
            format=float(data.get("format", 0)),
        )

    def score(self, *, color: float, style: float, elements: float, format: float) -> float:
        """根据本次任务权重计算 0 到 100 的综合相似度。"""
        return (
            color * self.color / 100
            + style * self.style / 100
            + elements * self.elements / 100
            + format * self.format / 100
        )

    def as_dict(self) -> dict[str, float]:
        """返回方便写入任务记录和日志的权重字典。"""
        return {"color": self.color, "style": self.style, "elements": self.elements, "format": self.format}
