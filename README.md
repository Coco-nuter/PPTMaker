# AI PPT Agent

通过大白话和用户素材生成可编辑 PPTX，并支持大纲确认、多轮局部修改、版本回退和最终文件真实预览。

## 当前状态

项目已完成 Python 3.12、uv、DeckSpec 数据合同、三种布局的确定性 PPTX 渲染、`PPTX → Microsoft PowerPoint COM → PNG` 真实预览链路，以及基于固定 `sample_deck.json` 的无 LLM Streamlit 最小闭环。当前也可通过独立命令行调用 OpenAI Responses API，把自然语言要求规划成经过 Pydantic 二次校验的 `DeckSpec`；该能力尚未接入 Streamlit。

当前运行依赖为 Streamlit、OpenAI Python SDK、Pydantic、pydantic-settings、python-pptx，以及仅在 Windows 安装的 pywin32；开发依赖为 pytest 和 Ruff。MarkItDown 尚未引入。

MVP 的目标闭环是：

```text
自然语言/素材 → 可确认大纲 → DeckSpec → 可编辑 PPTX
              → PPTX 实际渲染预览 → 大白话局部修改 → 新版本/回退
```

## 文档导航

- [产品定义](docs/PRODUCT.md)：用户、范围、用户流程、功能和验收标准。
- [架构设计](docs/ARCHITECTURE.md)：组件、数据合同、存储、依赖和技术边界。
- [测试策略](docs/TESTING.md)：测试分层、测试矩阵、质量门禁和完成定义。
- [MVP 执行计划](docs/exec-plans/active/mvp.md)：按依赖排序、可独立测试的开发任务。
- [原始调研方案](AI_PPT_Agent_开源项目与技术方案.md)：开源项目调研、工具安装和方案推导。
- [仓库规则](AGENTS.md)：对开发者和自动化 Agent 长期有效的约束。

## 计划采用的 MVP 技术栈

| 工具 | 用途 |
|---|---|
| Python 3.12 + [uv](https://docs.astral.sh/uv/) | 运行时、虚拟环境和依赖锁定 |
| [Streamlit](https://docs.streamlit.io/) | 聊天、上传、大纲、预览和下载界面 |
| OpenAI Python SDK + Pydantic | 结构化生成 `DeckSpec` 和 `PatchPlan` |
| [MarkItDown](https://github.com/microsoft/markitdown) | PDF、DOCX、PPTX、XLSX 等素材解析 |
| [python-pptx](https://python-pptx.readthedocs.io/) | 生成原生可编辑 PPTX |
| Microsoft PowerPoint + [pywin32](https://pypi.org/project/pywin32/) | 通过 COM 将最终 PPTX 逐页导出为 PNG |
| pytest + Ruff | 自动化测试和静态检查 |

## Windows 开发环境准备

以下命令可在 PowerShell 中执行；它们只准备工具，不会生成业务代码。

```powershell
winget install --id Git.Git -e
winget install --id Microsoft.VisualStudioCode -e
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

重新打开 PowerShell 后验证：

```powershell
git --version
uv --version
uv python install 3.12
```

本机还需要安装带有效许可的 Microsoft PowerPoint。复制预览配置；默认导出尺寸为 1920×1080：

```powershell
Copy-Item .env.example .env
```

人工测试自然语言规划前，在未提交的 `.env` 中填写 `OPENAI_API_KEY` 和支持结构化输出的 `OPENAI_MODEL`；官方端点可保留示例中的 `OPENAI_BASE_URL`。不要把真实密钥写入 `.env.example`。

## 当前可用的开发命令

在仓库根目录执行：

```powershell
uv sync
uv run pytest -q
uv run ruff check .
uv run python src/generate_sample_pptx.py
uv run python src/generate_sample_preview.py
uv run streamlit run app.py
```

Streamlit 页面读取并校验 `tests/fixtures/sample_deck.json`，展示三页大纲。每次点击“生成 PPT”都会创建新的 `workspace/<project_id>/`，生成可编辑 PPTX、展示 PowerPoint 导出的全部 PNG，并提供 PPTX 下载。

独立执行自然语言规划（只生成 DeckSpec JSON，不启动 Streamlit、不生成 PPTX）：

```powershell
uv run python src/plan_deck.py "生成一份6页的研究生开题汇报，简洁蓝白风" --output output/planned_deck.json
```

真实调用需要有效密钥，结果写入 `output/planned_deck.json`。普通 `pytest` 全部使用 Mock，不会访问网络。

只运行 PowerPoint 真实集成测试：

```powershell
uv run pytest -m "integration and powerpoint" -q
```

## 当前仓库结构

```text
.
├─ AGENTS.md
├─ app.py                   # 无 LLM Streamlit 最小闭环
├─ README.md
├─ docs/
│  ├─ PRODUCT.md
│  ├─ ARCHITECTURE.md
│  ├─ TESTING.md
│  └─ exec-plans/active/mvp.md
├─ src/config.py            # 环境配置
├─ src/demo_workflow.py     # 独立项目创建与生成编排
├─ src/llm.py               # OpenAI 结构化输出适配器
├─ src/models.py            # DeckSpec 数据合同
├─ src/planner.py           # 自然语言到 DeckSpec 规划
├─ src/plan_deck.py         # 规划命令行入口
├─ src/pptx_renderer.py     # 三布局 PPTX 渲染
├─ src/preview.py           # PowerPoint COM 真实预览
├─ prompts/plan_deck.md     # DeckSpec 规划提示词
├─ tests/                   # 自动化测试与 fixtures
├─ pyproject.toml           # 依赖、Ruff 和 pytest 配置
├─ uv.lock                  # 可重复安装的依赖锁文件
└─ .env.example             # 不含真实密钥的配置示例
```

## 安全

- 不要提交 `.env`、API 密钥、用户上传材料、生成的 PPTX/PNG 或运行时工作区。
- 当前目录中的任何包含“密钥”或“secret”字样的本地文件均应保持未跟踪状态。
- 用户文件只能在对应项目工作区内读写，不能用未经校验的文件名拼接任意路径。

## 开始开发

按 [MVP 执行计划](docs/exec-plans/active/mvp.md) 继续当前阶段。每个任务都列出了产物、独立测试和完成条件；不要跳过前置质量门禁。
