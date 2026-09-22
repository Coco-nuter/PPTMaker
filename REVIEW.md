# PPTMaker 启动与使用

## 1. 环境要求

- Windows 10/11。
- 已安装并激活 Microsoft PowerPoint 桌面版。
- 可用的 DeepSeek/OpenAI 兼容接口、API Key 和模型名称。

## 2. 安装工具并获取项目

在 PowerShell 中执行：

```powershell
winget install --id Git.Git -e
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
uv python install 3.12

git clone https://github.com/Coco-nuter/PPTMaker.git
Set-Location PPTMaker
uv sync
```

如果安装后找不到 `git` 或 `uv`，关闭并重新打开 PowerShell。

## 3. 配置模型

复制配置模板：

```powershell
Copy-Item .env.example .env
notepad .env
```

在 `.env` 中填写：

```dotenv
OPENAI_API_KEY=你的API密钥
OPENAI_BASE_URL=服务商提供的OpenAI兼容接口地址
OPENAI_MODEL=deepseek-v4-flash
LLM_TIMEOUT_SECONDS=120

PREVIEW_BACKEND=powerpoint
PREVIEW_TIMEOUT_SECONDS=120
PREVIEW_WIDTH=1920
PREVIEW_HEIGHT=1080
```

不要修改 `.env.example`，不要提交 `.env` 或密钥文件。

## 4. 启动

在项目根目录执行：

```powershell
uv run streamlit run app.py
```

浏览器通常会自动打开；未打开时访问终端显示的本地地址，一般是 `http://localhost:8501`。

## 5. 使用流程

1. 填写自然语言要求、PPT 主题、受众、用途和页数。
2. 点击“生成大纲”。
3. 检查每页标题、布局和内容；可重新规划或放弃。
4. 点击“确认大纲”。
5. 等待 PowerPoint 完成 PNG 预览导出。
6. 检查全部预览页并下载 PPTX。

生成文件位于：

```text
workspace/<project_id>/output/sample_deck.pptx
workspace/<project_id>/preview/preview-*/png/
```

## 6. 验证与排错

运行基础检查：

```powershell
uv run pytest -q
uv run ruff check .
```

- 提示缺少 `OPENAI_API_KEY`、`OPENAI_MODEL`：检查 `.env` 是否位于项目根目录，修改后重启应用。
- 提示认证失败：确认 API Key、接口地址和模型名称属于同一服务商。
- 无法生成预览：确认安装的是 Microsoft PowerPoint 桌面版，并关闭 PowerPoint 的弹窗或受保护视图提示后重试。
- 修改代码或配置后结果未更新：停止当前进程，再重新运行启动命令。

## 7. 进度

已完成：

- 调研网上开源 PPT 生成项目并整理产品、架构、测试和 MVP 实施方案。
- 建立 Python 3.12、uv、pytest、Ruff 工程环境。
- 定义并校验 `DeckSpec 2.0`，支持 `cover`、`section`、`bullets`、`two_column`、`metrics`、`timeline`、`process`、`comparison`、`closing` 九种布局。
- 接入 DeepSeek/OpenAI 兼容接口，根据文字要求生成合法大纲并选择布局。
- 使用 `python-pptx` 绘制可编辑的文字、卡片、色块、节点、连接线和箭头。
- 使用 PowerPoint COM 从最终 PPTX 导出 1920×1080 PNG 预览。
- 完成 Streamlit 的文字输入、大纲确认、重新规划、生成、预览和下载流程。
- 完成离线 Mock 测试、真实 DeepSeek 测试和 PowerPoint 集成测试。

未完成或未继续扩展：

- 尚未增加更多版式变体，现有页面主要依赖九种固定布局。
- 尚未实现图片占位、用户图片、AI 背景图或图片搜索。
- 尚未实现多个 TXT、PDF、DOCX 等文档上传、解析和内容合并。
- 尚未实现多轮局部修改、版本回退、自动 QA 和复杂图表。

## 8. 设计取舍

最初先调研网上可运行的开源项目和常见方案，让 AI 汇总出一条完整技术路线，再根据题目范围和本机环境做裁剪：

1. 去掉 LibreOffice 渲染方案，改用 Microsoft PowerPoint COM。项目只面向 Windows，PowerPoint 导出的 PNG 与用户最终打开的 PPTX 更一致，中文字体、换行和原生对象表现也更可靠。
2. 取消用户上传图片和文件素材。题目要求用户只提供文字，先完成“文字 → 大纲 → 可编辑 PPTX → 真实预览”的最小闭环，避免把时间消耗在上传、解析、版权和文件安全上。
3. 取消独立自动 QA、哈希绑定、版本管理等重复检查体系，但保留必要检查：模型结果经过 Pydantic 校验，PPTX 保存后重新打开，预览检查页数、尺寸和 PowerPoint 进程清理。
4. 将系统拆为数据合同、模型规划、确定性渲染、真实预览和 Streamlit 工作流。DeepSeek 负责故事线、布局类型和内容，程序负责坐标、字号、颜色、边距及对象绘制，避免模型直接生成不可控的绘图代码。
5. 按阶段逐步实现，每一步先查看测试和真实产物；符合预期后继续，不符合预期就修改技术方案并重做。例如预览链路从 LibreOffice 调整为 PowerPoint COM，数据合同从三种基础布局扩展为九种独立布局。

## 9. 已知缺陷

1. **版式多样性不足**：虽然支持九种语义布局，但每种布局主要只有一套确定性模板，同类型页面较多时容易显得重复。
2. **缺少图片能力**：当前没有图片占位、用户图片或 AI 背景图，难以生成以产品、人物、场景和视觉氛围为重点的 PPT。
3. **运行环境依赖较强**：真实预览依赖 Windows、Microsoft PowerPoint 和 COM；未安装 PowerPoint、存在 Office 弹窗或字体不一致时，预览可能失败或出现排版差异。

## 10. 优先级

如果再给 2 小时，我会优先扩展版式多样性：为常用布局增加 2～3 套可确定选择的原生形状变体，并根据内容数量和类型自动选择。它不需要改变当前数据合同和工作流，风险较低，能够最快改善所有生成结果的观感。

后续优先顺序：

1. 增加版式变体、背景色块和装饰组合。
2. 评估 AI 多元化背景图和图片占位，并先更新产品范围、数据合同和架构。
3. 增加多个 TXT 或文档上传、内容提取、去重和合并规划，最终生成一份 PPT。

图片生成和文档上传都会扩大当前“纯文字输入”MVP 的范围，需要额外处理文件安全、来源冲突、版权、缓存和失败恢复，因此不适合作为两小时内的第一项修改。
