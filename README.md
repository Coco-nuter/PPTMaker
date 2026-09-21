# AI PPT Agent

通过大白话和用户素材生成可编辑 PPTX，并支持大纲确认、多轮局部修改、版本回退和最终文件真实预览。

## 当前状态

项目处于初始化规划阶段。当前仓库只有调研与实施文档，尚未创建业务代码、Python 包或可运行 Demo。

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
| [LibreOffice](https://www.libreoffice.org/download/) | 无界面地将 PPTX 转成 PDF |
| [PyMuPDF](https://pymupdf.readthedocs.io/) | 将 PDF 页面渲染为预览 PNG |
| pytest + Ruff | 自动化测试和静态检查 |

## Windows 开发环境准备

以下命令可在 PowerShell 中执行；它们只准备工具，不会生成业务代码。

```powershell
winget install --id Git.Git -e
winget install --id Microsoft.VisualStudioCode -e
winget install --id TheDocumentFoundation.LibreOffice -e

powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

重新打开 PowerShell 后验证：

```powershell
git --version
uv --version
uv python install 3.12
& "C:\Program Files\LibreOffice\program\soffice.exe" --version
```

## 计划中的运行命令

下面的命令会在对应 MVP 任务完成后生效；当前尚不可运行。

```powershell
uv sync
uv run pytest
uv run ruff check .
uv run streamlit run app.py
```

默认开发地址计划为 `http://localhost:8501`。

## 计划中的仓库结构

```text
.
├─ AGENTS.md
├─ README.md
├─ docs/
│  ├─ PRODUCT.md
│  ├─ ARCHITECTURE.md
│  ├─ TESTING.md
│  └─ exec-plans/active/mvp.md
├─ prompts/                 # 后续创建：模型提示词
├─ src/                     # 后续创建：领域与基础设施模块
├─ templates/               # 后续创建：主题和布局资源
├─ tests/                   # 后续创建：自动化测试与 fixtures
├─ app.py                   # 后续创建：Streamlit 入口
└─ workspace/               # 运行时创建且不提交
```

## 安全

- 不要提交 `.env`、API 密钥、用户上传材料、生成的 PPTX/PDF/PNG 或运行时工作区。
- 当前目录中的任何包含“密钥”或“secret”字样的本地文件均应保持未跟踪状态。
- 用户文件只能在对应项目工作区内读写，不能用未经校验的文件名拼接任意路径。

## 开始开发

按 [MVP 执行计划](docs/exec-plans/active/mvp.md) 从 M0 开始。每个任务都列出了产物、独立测试和完成条件；不要跳过前置质量门禁。

