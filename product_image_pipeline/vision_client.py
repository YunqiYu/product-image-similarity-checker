"""视觉模型通用调用：图片编码、JSON 解析和 OpenAI 兼容客户端创建。"""

from __future__ import annotations

import base64
import json
import mimetypes
import re
from pathlib import Path
from typing import Any

from .settings import ModelSettings


def image_to_data_url(path: Path) -> str:
    """将本地图片转换为模型接口可接收的 data URL。"""
    mime_type = mimetypes.guess_type(path.name)[0] or "image/png"
    return f"data:{mime_type};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


def extract_json(text: str) -> dict[str, Any]:
    """兼容模型返回的纯 JSON 或 Markdown JSON 代码块。"""
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```(?:json)?\s*", "", stripped)
        stripped = re.sub(r"\s*```$", "", stripped)
    try:
        data = json.loads(stripped)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", stripped, flags=re.S)
        if not match:
            raise
        data = json.loads(match.group(0))
    if not isinstance(data, dict):
        raise ValueError("模型返回内容不是 JSON 对象。")
    return data


def create_client(settings: ModelSettings, *, timeout: float, max_retries: int) -> Any:
    """创建 OpenAI 兼容客户端，供排序和打标签共用。"""
    from openai import OpenAI

    return OpenAI(api_key=settings.api_key, base_url=settings.base_url, timeout=timeout, max_retries=max_retries)
