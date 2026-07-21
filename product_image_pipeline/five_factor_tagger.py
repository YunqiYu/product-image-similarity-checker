"""使用视觉模型为单张商品图片生成五要素标签。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .vision_client import extract_json


FIVE_FACTOR_PROMPT = """你是资深商品图片标签识别专家。请只识别商品正面印刷图案，忽略拍摄背景、包装数量和阴影。
请只输出合法 JSON：
{
  "color": ["主要颜色"],
  "style": ["视觉风格"],
  "format": "构图与排版",
  "elements": ["主体元素与主题元素"],
  "wordArt": "主文案或艺术字；没有则为空字符串"
}
"""


def load_tagging_prompt(prompt_path: Path | None = None) -> str:
    """优先读取可维护的标签 Prompt，缺失时使用内置默认 Prompt。"""
    if prompt_path and prompt_path.exists():
        return prompt_path.read_text(encoding="utf-8")
    return FIVE_FACTOR_PROMPT


def normalize_five_factors(payload: dict[str, Any]) -> dict[str, Any]:
    """规范化模型返回，保证最终 Excel 始终有五个字段。"""
    def list_value(key: str) -> list[str]:
        value = payload.get(key, [])
        if value is None:
            return []
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        return [str(value).strip()] if str(value).strip() else []

    def text_value(key: str) -> str:
        value = payload.get(key)
        return str(value).strip() if value is not None else ""

    return {
        "color": list_value("color"),
        "style": list_value("style"),
        "format": text_value("format"),
        "elements": list_value("elements"),
        "wordArt": text_value("wordArt"),
    }


def tag_image(client: Any, model: str, image_data_url: str, image_detail: str, prompt: str) -> dict[str, Any]:
    """调用视觉模型，并返回一张图片的五要素标签。"""
    response = client.chat.completions.create(
        model=model,
        response_format={"type": "json_object"},
        messages=[{"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url", "image_url": {"url": image_data_url, "detail": image_detail}},
        ]}],
    )
    content = response.choices[0].message.content or "{}"
    return normalize_five_factors(extract_json(content))
