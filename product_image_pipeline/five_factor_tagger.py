"""五要素打标核心：调用视觉模型、兼容响应格式并校验原始提示词协议。"""

from __future__ import annotations

import json
import time
from typing import Any

from .vision_client import extract_json


FIVE_FACTOR_PROMPT = """你是资深商品图片视觉标签专家。请仅识别商品正面印刷图案，忽略摄影背景、阴影、包装数量和商品摆放角度。

只输出合法 JSON，不要输出解释或 Markdown。字段如下：
{
  "color": ["主要颜色"],
  "style": ["视觉风格"],
  "format": "构图和排版方式",
  "elements": ["主体元素和主题元素"],
  "wordArt": "主文案或艺术字；没有则为空字符串"
}
"""

ALLOWED_STYLES = ("插画", "矢量", "水彩", "特效", "写实")
ALLOWED_FORMATS = (
    "图案-散排",
    "图案-居中",
    "图案-环绕",
    "图案-铺满",
    "文案-上下",
    "文案-环绕",
    "文案-居中",
)


class TaggingError(RuntimeError):
    """模型调用或返回结果不满足五要素协议时抛出。"""


def extract_message_text(content: Any) -> str | None:
    """兼容标准接口和部分兼容服务商的分段文本返回格式。"""
    if isinstance(content, str):
        return content.strip() or None
    if not isinstance(content, list):
        return None

    text_parts: list[str] = []
    for part in content:
        if isinstance(part, dict):
            text = part.get("text") or part.get("content")
        else:
            text = getattr(part, "text", None) or getattr(part, "content", None)
        if isinstance(text, str) and text.strip():
            text_parts.append(text.strip())
    return "\n".join(text_parts) or None


def _text_list(value: Any, field_name: str, *, minimum: int = 1, maximum: int | None = None) -> list[str]:
    """校验数组型标签，避免将空值或非文本写入最终 Excel。"""
    if not isinstance(value, list) or len(value) < minimum or (maximum is not None and len(value) > maximum):
        raise TaggingError(f"{field_name} 必须是至少包含 {minimum} 项的数组：{value!r}")
    result = [item.strip() for item in value if isinstance(item, str) and item.strip()]
    if len(result) != len(value):
        raise TaggingError(f"{field_name} 只能包含非空文本：{value!r}")
    return result


def parse_five_factors(raw: str) -> dict[str, Any]:
    """按 config 中原始 V4 五要素协议解析并校验模型 JSON。"""
    try:
        result = extract_json(raw)
    except (json.JSONDecodeError, ValueError) as error:
        raise TaggingError(f"模型未返回可解析的 JSON：{raw[:300]!r}") from error

    # 部分兼容模型会将提示词中的 WordArt 统一转成小写，字段语义保持一致。
    if "wordart" in result and "WordArt" not in result:
        result["WordArt"] = result.pop("wordart")

    expected_keys = {"color", "style", "format", "elements", "WordArt"}
    if set(result) != expected_keys:
        raise TaggingError(
            f"模型 JSON 字段必须严格为 {sorted(expected_keys)}，实际为：{sorted(result)}"
        )
    format_label = result["format"]
    color = result["color"]
    word_art = result["WordArt"]
    if not isinstance(color, str) or not color.strip():
        raise TaggingError(f"color 必须是非空文本：{color!r}")
    if not isinstance(format_label, str) or not format_label.strip():
        raise TaggingError(f"format 必须是非空文本：{format_label!r}")
    if not isinstance(word_art, str):
        raise TaggingError(f"WordArt 必须是文本：{word_art!r}")
    if format_label not in ALLOWED_FORMATS:
        raise TaggingError(f"format 必须是 V4 固定标签之一：{format_label!r}")
    style = _text_list(result["style"], "style", maximum=2)
    if any(item not in ALLOWED_STYLES for item in style) or len(set(style)) != len(style):
        raise TaggingError(f"style 必须是不重复的 V4 固定风格：{style!r}")

    return {
        "color": color.strip(),
        "style": style,
        "format": format_label.strip(),
        "elements": _text_list(result["elements"], "elements", maximum=4),
        "wordArt": word_art.strip(),
    }


def _response_content(response: Any) -> str:
    """从模型响应中取出可解析的正文，并在异常时输出简短诊断。"""
    try:
        choice = response.choices[0]
        message = choice.message
        extra_fields = getattr(message, "model_extra", None) or {}
        content = extract_message_text(message.content) or extract_message_text(extra_fields.get("content"))
    except (AttributeError, IndexError, TypeError) as error:
        raise TaggingError(f"API 返回格式不符合预期：{response!r}") from error
    if content:
        return content
    diagnostic = {
        "finish_reason": getattr(choice, "finish_reason", None),
        "content_type": type(getattr(message, "content", None)).__name__,
        "refusal_present": bool(getattr(message, "refusal", None)),
        "extra_fields": sorted(str(key) for key in extra_fields),
    }
    raise TaggingError(f"API 返回中没有可用的文本内容：{json.dumps(diagnostic, ensure_ascii=False)}")


def tag_image(
    client: Any,
    model: str,
    image_data_url: str,
    image_detail: str,
    prompt: str,
    *,
    attempts: int = 3,
    retry_delay_seconds: float = 2.0,
) -> dict[str, Any]:
    """调用视觉模型打标，遇到接口或标签校验失败时按次重试。"""
    if attempts < 1:
        raise ValueError("attempts 必须大于 0。")
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            response = client.chat.completions.create(
                model=model,
                temperature=0,
                max_tokens=1000,
                response_format={"type": "json_object"},
                messages=[{"role": "user", "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": image_data_url, "detail": image_detail}},
                ]}],
            )
            return parse_five_factors(_response_content(response))
        except Exception as error:
            last_error = error
            if attempt < attempts:
                time.sleep(retry_delay_seconds * attempt)
    raise TaggingError(f"五要素打标失败（已重试 {attempts} 次）：{last_error}") from last_error
