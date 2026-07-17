# Product Image Similarity Checker

商品图片相似度判断工具。当前核心流程是：给定一张目标图片，将其与竞品图片逐张对比，输出综合相似度，并生成带图片的 Excel 结果。

备注：一品红已测试，业务反馈效果可以，执行 py 为 `folder_similarity.py`。

## 功能

- 读取文件夹中的目标图与竞品图。
- 使用视觉模型进行图片相似度判断。
- 内部按颜色、风格、元素、排版四个维度计算综合相似度。
- 输出 Excel，并保留原始表结构中的业务字段。
- Excel 中会填入图片，方便人工核对。

## 安装

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
```

在 `.env` 中配置模型服务：

```text
TEXT_API_KEY=your_api_key_here
LLM_BASE_URL=https://api.example.com/v1
OPENAI_MODEL=gpt-5.5
OPENAI_IMAGE_MODE=high
```

## 运行

默认执行文件：

```powershell
python folder_similarity.py
```

指定输入目录与输出文件：

```powershell
python folder_similarity.py `
  --input-dir "inputs\测试2" `
  --output "outputs\测试2_图片相似度结果.xlsx" `
  --run-log "outputs\测试2_图片相似度结果.log"
```

输入目录中需要包含一张文件名带有 `目标` 的图片，其余图片会作为竞品图参与对比。

## 输出列

当前 Excel 输出结构为：

```json
[
  "1级分类",
  "主题标签",
  "ASIN",
  "图片链接",
  "图片",
  "产品链接",
  "品牌",
  "pcs",
  "标题",
  "近30天销量",
  "上架时间",
  "价格",
  "数据来源",
  "价格趋势图",
  "综合相似度",
  "颜色",
  "风格",
  "元素",
  "排版"
]
```

其中 `综合相似度` 为模型判断结果；`颜色`、`风格`、`元素`、`排版` 保留原 Excel 中的文字标签内容。
