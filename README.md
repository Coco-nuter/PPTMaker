# AI PPT Agent

一个 Windows 本地运行的 AI PPT Demo：用户用大白话描述需求，DeepSeek 返回结构化 `DeckSpec`，程序使用 `python-pptx` 绘制可编辑 PowerPoint 原生对象，再由 Microsoft PowerPoint 导出真实 PNG 预览。

## MVP 流程

```text
用户文字要求
→ DeepSeek/OpenAI 兼容接口生成 DeckSpec
→ Pydantic 校验
→ 用户确认大纲
→ python-pptx 绘制文字、卡片、色块、节点和连接线
→ 可编辑 PPTX
→ PowerPoint COM 导出 PNG
→ 页面预览与 PPTX 下载
```

当前已经可以完成自然语言规划、大纲确认、九种布局 PPTX 生成、PowerPoint 真实预览和下载。`DeckSpec` 2.0 的全部布局均由 `python-pptx` 使用原生文字框、卡片、色块、节点、连接线和箭头确定性绘制。

## MVP 边界

MVP 只接受文字输入，不上传或解析文件，不使用用户图片，也不生成、搜索图片。MVP 不包含 `PatchPlan`、多轮局部修改、版本管理、版本回退、自动 QA 系统、预览哈希绑定或图表布局。

[历史调研方案](AI_PPT_Agent_开源项目与技术方案.md)仅作为背景参考，其中的素材解析、图片和多轮修改方案不属于当前开发范围。

## 技术栈

| 工具 | 用途 |
|---|---|
| Python 3.12 + [uv](https://docs.astral.sh/uv/) | 运行时、虚拟环境和依赖锁定 |
| [Streamlit](https://docs.streamlit.io/) | 文字输入、大纲确认、预览和下载 |
| OpenAI Python SDK + Pydantic | 调用 DeepSeek/OpenAI 兼容接口并校验 `DeckSpec` |
| [python-pptx](https://python-pptx.readthedocs.io/) | 绘制原生可编辑 PPTX 对象 |
| Microsoft PowerPoint + [pywin32](https://pypi.org/project/pywin32/) | 通过 COM 将最终 PPTX 逐页导出为 PNG |
| pytest + Ruff | 自动化测试和静态检查 |

项目不使用 MarkItDown，也不需要 LibreOffice。

## Windows 快速开始

安装 Git、uv 和 Python 3.12：

```powershell
winget install --id Git.Git -e
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
uv python install 3.12
uv sync
```

本机需要安装带有效许可的 Microsoft PowerPoint。复制环境配置：

```powershell
Copy-Item .env.example .env
```

在未提交的 `.env` 中配置：

```dotenv
OPENAI_API_KEY=your-api-key
OPENAI_BASE_URL=your-openai-compatible-endpoint
OPENAI_MODEL=your-deepseek-model
LLM_TIMEOUT_SECONDS=120
PREVIEW_BACKEND=powerpoint
PREVIEW_TIMEOUT_SECONDS=120
PREVIEW_WIDTH=1920
PREVIEW_HEIGHT=1080
```

不要把真实密钥写入 `.env.example` 或提交到 Git。

启动应用：

```powershell
uv run streamlit run app.py
```

页面会收集主题、受众、用途、页数和补充要求。模型生成大纲后，用户可以确认、重新规划或放弃；只有确认后才会生成 PPTX 和 PowerPoint PNG 预览。

## 开发与测试

```powershell
uv run pytest -q
uv run ruff check .
```

独立生成 `DeckSpec`：

```powershell
uv run python src/plan_deck.py "生成一份6页的研究生开题汇报，简洁蓝白风" --output output/planned_deck.json
```

生成现有示例 PPTX 和预览：

```powershell
uv run python src/generate_sample_pptx.py
uv run python src/generate_sample_preview.py
```

真实 Provider 和 PowerPoint 测试默认跳过，显式测试命令见 [docs/TESTING.md](docs/TESTING.md)。普通测试不访问网络，也不会自动启动 PowerPoint。

## 项目结构

```text
.
├─ app.py                   # Streamlit 文字输入、大纲确认与生成 UI
├─ src/
│  ├─ app_workflow.py      # UI 输入与规划边界
│  ├─ config.py            # 环境配置
│  ├─ llm.py               # DeepSeek/OpenAI 兼容接口适配器
│  ├─ models.py            # DeckSpec 数据合同
│  ├─ planner.py           # 文字要求到 DeckSpec
│  ├─ pptx_renderer.py     # 原生 PPTX 渲染
│  └─ preview.py           # PowerPoint COM 真实预览
├─ prompts/plan_deck.md
├─ tests/
├─ docs/
├─ pyproject.toml
├─ uv.lock
└─ .env.example
```

## 文档

- [产品定义](docs/PRODUCT.md)
- [架构设计](docs/ARCHITECTURE.md)
- [测试策略](docs/TESTING.md)
- [当前 MVP 计划](docs/exec-plans/active/mvp.md)
- [仓库长期规则](AGENTS.md)
- [历史调研参考](AI_PPT_Agent_开源项目与技术方案.md)：不是当前 MVP 范围。

生成的 `workspace/`、`output/`、`.env` 和本地密钥文件不得提交。
