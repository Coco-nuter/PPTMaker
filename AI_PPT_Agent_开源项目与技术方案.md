# 大白话生成 PPT Agent：Windows 可运行 Demo 实施方案

> 目标：用户用自然语言描述需求，可上传 PDF、Word、PPT、Excel、Markdown 和图片；程序先生成可确认的大纲，再生成可编辑 `.pptx`；用户继续用大白话修改指定页面，最后直接用 PowerPoint/WPS 打开。
>
> 调研与方案更新时间：2026-09-21。链接优先使用官方仓库和官方文档。

## 1. Demo 最终效果

启动程序后，浏览器中应能完成：

1. 输入：“根据这些材料，做一份 8 页的研究生开题汇报，简洁蓝白风，突出研究问题和技术路线。”
2. 上传 PDF、DOCX、PPTX、XLSX、Markdown 或图片。
3. 程序给出标题、受众、页数和逐页大纲；用户可确认或修改。
4. 程序生成可编辑 PPTX，并显示由该 PPTX 实际渲染得到的页面预览。
5. 用户继续输入：“第 3 页压缩成三点；第 5 页改成时间轴；整体主色换成深蓝。”
6. 程序只修改命中的页面/主题，保留其他内容，生成新版本并支持回退。
7. 用户下载 `.pptx`，用 PowerPoint 或 WPS 直接打开和编辑文字、形状、表格与图表。

第一版暂不做多人协作、复杂动画、自由画布和任意现有 PPTX 的无损编辑。先把“自然语言 → 素材理解 → 大纲 → 可编辑 PPTX → 多轮局部修改”完整跑通。

## 2. 总体方案

```text
自然语言 + 用户素材
        │
        ▼
素材解析（MarkItDown）── 图片/文件保存到项目目录
        │
        ▼
需求理解 + 大纲规划（LLM + Pydantic 结构化输出）
        │
        ▼
DeckSpec：PPT 的唯一事实源，页面具有稳定 ID
        │
        ├── 用户确认/修改大纲
        │
        ├── 大白话修改 → PatchPlan → 修改 DeckSpec
        │
        ▼
模板布局器 + python-pptx → 原生可编辑 PPTX
        │
        ▼
Microsoft PowerPoint COM → 逐页 PNG
        │
        ▼
规则检查/视觉检查 → 预览、回退、下载
```

核心原则：LLM 负责理解、规划和选择布局，不直接随意输出大量绝对坐标；程序用可测试的布局函数生成 PPT。多轮修改时，LLM 只返回结构化修改计划，不重新编写整份 PPT。

## 3. 为什么这样选

| 选择 | 原因 |
|---|---|
| Python 单体应用 | Windows 开发最简单，第一版不用维护前后端、Node、数据库和消息队列 |
| [Streamlit](https://docs.streamlit.io/get-started/installation) | 很快做出聊天、文件上传、预览、按钮和下载界面 |
| [MarkItDown](https://github.com/microsoft/markitdown) | 可把 PDF、PPTX、DOCX、XLSX、图片等转换成适合 LLM 阅读的 Markdown |
| [Pydantic](https://docs.pydantic.dev/) + 结构化输出 | 强制模型生成合法 DeckSpec/PatchPlan，避免解析一段不稳定的自然语言 |
| [python-pptx](https://python-pptx.readthedocs.io/en/latest/) | 无需安装 PowerPoint即可生成原生 PPTX；文字、形状、表格、图片和常用图表可编辑 |
| Microsoft PowerPoint + [pywin32](https://pypi.org/project/pywin32/) | 用 PowerPoint 自身渲染最终 PPTX，并逐页导出真实 PNG |
| 显式工作流 | Demo 不需要 LangGraph；状态机更容易调试、测试和控制成本 |

这吸收了以下开源项目的优点：

- [Presenton](https://github.com/presenton/presenton)：自然语言/文档/模板生成、Docker/桌面运行、可编辑 PPTX。
- [CreatPPT](https://github.com/seekskyworld/CreatPPT)：使用 DeckSpec 作为编辑器和 PPTX 导出器之间的共同数据源。
- [PPTAgent](https://github.com/icip-cas/PPTAgent)：生成后渲染、视觉复核、再修正。
- [AIPPT](https://github.com/LRriver/AIPPT)：先审大纲、单页重做、历史版本和多模型分工。

## 4. Windows 工具与部署

以下命令在 PowerShell 中运行。建议使用 Windows 10/11 x64。

### 4.1 Git：下载代码与版本管理

功能：克隆项目、保存代码版本；不是程序运行时核心依赖，但强烈建议安装。

```powershell
winget install --id Git.Git -e
git --version
```

### 4.2 VS Code：开发编辑器

功能：编辑 Python、Markdown、JSON 和环境变量文件。

```powershell
winget install --id Microsoft.VisualStudioCode -e
code --version
```

### 4.3 uv + Python 3.12：Python 环境和依赖管理

功能：安装 Python、创建虚拟环境、锁定依赖并运行程序。相比手动管理 `venv + pip` 更省事。[uv Windows 安装说明](https://github.com/astral-sh/uv/blob/main/docs/reference/installer.md)

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# 关闭并重新打开 PowerShell，然后执行
uv --version
uv python install 3.12
uv python list
```

### 4.4 Microsoft PowerPoint：PPTX 真实渲染

功能：通过 PowerPoint COM 将生成的 PPTX 逐页导出 PNG。预览必须来自最终 PPTX，而不是前端自己画一个近似页面。

本机需要 Windows 和已安装、已许可的桌面版 Microsoft PowerPoint。COM 渲染代码放入独立子进程，主进程负责超时和失败清理。

### 4.5 Python 包：Demo 的业务组件

| 包 | 功能 |
|---|---|
| `streamlit` | Web 界面、聊天、上传、预览、下载 |
| `openai` | 调用 OpenAI API；也可通过 `base_url` 接兼容服务 |
| `pydantic` / `pydantic-settings` | DeckSpec、PatchPlan 校验和配置读取 |
| `markitdown[all]` | PDF、DOCX、PPTX、XLSX、图片等素材解析 |
| `python-pptx` | 生成原生可编辑 PPTX |
| `pywin32`（仅 Windows） | 调用 PowerPoint COM 逐页导出 PNG |
| `pillow` | 图片读取、缩放、裁剪和格式转换 |
| `python-dotenv` | 从 `.env` 读取密钥与路径 |
| `httpx` | 下载用户提供的远程素材；后续可接图库 API |
| `pytest` / `ruff` | 测试与代码检查 |

## 5. 从零创建项目

### 步骤 1：初始化目录和依赖

```powershell
New-Item -ItemType Directory -Path ppt-agent-demo
Set-Location ppt-agent-demo

uv init --python 3.12
uv add streamlit openai pydantic pydantic-settings python-pptx pillow python-dotenv httpx
uv add "pywin32>=311; sys_platform == 'win32'"
uv add "markitdown[all]"
uv add --dev pytest ruff

uv run streamlit hello
```

若 `streamlit hello` 能打开浏览器页面，Python 环境已经可用。

### 步骤 2：建立目录结构

```text
ppt-agent-demo/
├─ app.py                       # Streamlit 入口
├─ pyproject.toml
├─ uv.lock
├─ .env                         # 密钥；禁止提交 Git
├─ .gitignore
├─ prompts/
│  ├─ plan_deck.md              # 生成大纲/DeckSpec 的提示词
│  └─ edit_deck.md              # 把修改要求转成 PatchPlan
├─ src/
│  ├─ config.py                 # 环境变量
│  ├─ models.py                 # Pydantic 数据模型
│  ├─ llm.py                    # 模型调用和结构化输出
│  ├─ ingest.py                 # 素材保存、解析和来源编号
│  ├─ planner.py                # 生成初始 DeckSpec
│  ├─ patcher.py                # 应用多轮修改
│  ├─ layouts.py                # 可测试的页面布局函数
│  ├─ pptx_renderer.py          # DeckSpec → PPTX
│  ├─ preview.py                # PPTX → PowerPoint COM → PNG
│  ├─ qa.py                     # 溢出、密度、缺图等检查
│  └─ storage.py                # 项目、版本、产物保存
├─ templates/
│  ├─ themes/
│  │  ├─ blue_business.json
│  │  ├─ academic.json
│  │  └─ dark_tech.json
│  └─ fonts/
├─ workspace/                   # 运行时项目；禁止提交 Git
└─ tests/
   ├─ test_models.py
   ├─ test_patches.py
   └─ test_renderer.py
```

`.gitignore` 至少包含：

```gitignore
.env
.venv/
__pycache__/
workspace/
.streamlit/secrets.toml
```

### 步骤 3：配置模型和 PowerPoint 预览

`.env` 示例：

```dotenv
OPENAI_API_KEY=替换成你的密钥
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-5.6

PREVIEW_BACKEND=powerpoint
PREVIEW_TIMEOUT_SECONDS=120
PREVIEW_WIDTH=1920
PREVIEW_HEIGHT=1080
WORKSPACE_DIR=workspace
MAX_SOURCE_CHARS=60000
```

模型名应替换为账号实际可用且支持结构化输出的模型。密钥只放在服务端 `.env`，不要写进浏览器代码或提交到 Git。

官方 OpenAI Python SDK 可用 Pydantic 定义结构化输出；[官方文档](https://developers.openai.com/api/docs/guides/structured-outputs)给出了 `client.responses.parse(..., text_format=YourModel)` 的模式。第三方“OpenAI-compatible”接口不一定支持 Responses API 或 Pydantic parse，接入前应先验证；不支持时退回 JSON Schema/JSON Object 后再用 Pydantic 校验。

### 步骤 4：定义稳定的 DeckSpec

第一版不让模型输出每个文本框的坐标，只让它选择程序支持的布局和内容。最小数据模型：

```python
from typing import Literal
from pydantic import BaseModel, Field

LayoutName = Literal[
    "cover", "section", "bullets", "two_column",
    "image_text", "metrics", "timeline", "chart", "closing"
]

class SourceRef(BaseModel):
    id: str
    file_name: str
    locator: str = ""

class SlideSpec(BaseModel):
    id: str                         # 创建后永不改变，例如 slide_005
    layout: LayoutName
    title: str
    subtitle: str = ""
    bullets: list[str] = Field(default_factory=list, max_length=6)
    left_points: list[str] = Field(default_factory=list)
    right_points: list[str] = Field(default_factory=list)
    metrics: list[dict] = Field(default_factory=list)
    timeline: list[dict] = Field(default_factory=list)
    chart: dict | None = None
    asset_ids: list[str] = Field(default_factory=list)
    source_ids: list[str] = Field(default_factory=list)
    speaker_notes: str = ""

class ThemeSpec(BaseModel):
    name: str = "blue_business"
    primary: str = "1F4E79"
    accent: str = "2F80ED"
    background: str = "F7F9FC"
    text: str = "172B4D"
    font_family: str = "Microsoft YaHei"

class DeckSpec(BaseModel):
    id: str
    title: str
    audience: str
    purpose: str
    language: str = "zh-CN"
    aspect_ratio: Literal["16:9", "4:3"] = "16:9"
    theme: ThemeSpec
    sources: list[SourceRef] = Field(default_factory=list)
    slides: list[SlideSpec]
```

DeckSpec 是唯一事实源。PPTX、预览图、大纲和历史版本全部从它生成。

### 步骤 5：实现素材解析

`ingest.py` 的职责：

1. 将上传文件保存到 `workspace/<project_id>/uploads/`。
2. 为每个文件和图片生成稳定 `asset_id/source_id`。
3. 使用 MarkItDown 转为 Markdown，写入 `materials/`。
4. 保留文件名、页码/幻灯片号等定位信息。
5. 对超长材料按标题、段落或页切块；第一版先截取总字符数，后续再加向量检索。

核心调用形式：

```python
from markitdown import MarkItDown

converter = MarkItDown(enable_plugins=False)
result = converter.convert(str(file_path))
markdown_text = result.text_content
```

图片不要塞进模型返回的 JSON；保存为本地文件，只在 DeckSpec 中记录 `asset_id`。用户上传的图片优先于在线搜索或 AI 生图。

### 步骤 6：大白话生成大纲和 DeckSpec

`planner.py` 将以下内容交给模型：

- 用户原始要求；
- 受众、用途、页数、风格等已知条件；
- 素材摘要及来源 ID；
- 支持的 9 种布局和每种布局的容量限制；
- DeckSpec 的 Pydantic schema；
- 禁止编造数字、每页只表达一个核心结论等规则。

OpenAI 主路径示例：

```python
import os
from openai import OpenAI
from src.models import DeckSpec

client = OpenAI(
    api_key=os.environ["OPENAI_API_KEY"],
    base_url=os.environ.get("OPENAI_BASE_URL"),
)

def generate_deck_spec(user_request: str, materials: str) -> DeckSpec:
    response = client.responses.parse(
        model=os.environ["OPENAI_MODEL"],
        input=[
            {
                "role": "system",
                "content": "你是PPT策划师。先设计故事线，再输出合法DeckSpec。不得编造素材中没有的数据。",
            },
            {
                "role": "user",
                "content": f"用户要求：\n{user_request}\n\n素材：\n{materials}",
            },
        ],
        text_format=DeckSpec,
    )
    if response.output_parsed is None:
        raise RuntimeError("模型未返回合法 DeckSpec")
    return response.output_parsed
```

程序先把 `slides[].title/layout/一句话摘要` 显示成大纲。用户确认后才渲染 PPTX，避免花费时间生成一份故事线错误的文件。

### 步骤 7：实现模板布局器

`layouts.py` 为每种布局写一个确定性函数：

| 布局 | 必需内容 | 容量限制 |
|---|---|---|
| `cover` | 标题、副标题 | 标题建议不超过 24 个中文字符 |
| `section` | 章节名、一句话 | 内容极少 |
| `bullets` | 标题、要点 | 3～6 点，每点尽量不超过两行 |
| `two_column` | 标题、左右内容 | 每侧 2～4 点 |
| `image_text` | 标题、图片、短文 | 图片约占 45%～60% |
| `metrics` | 标题、指标卡 | 2～4 个指标 |
| `timeline` | 标题、阶段 | 3～6 个节点 |
| `chart` | 标题、图表数据、结论 | 第一版支持柱、线、饼图 |
| `closing` | 总结、行动项 | 1 个结论 + 3 个行动项以内 |

所有布局统一使用 16:9 画布、主题颜色、字体、边距和网格。不要让每页单独发明设计规则。

`pptx_renderer.py` 的处理顺序：

1. 创建 `Presentation()`，设置 16:9。
2. 加载 ThemeSpec。
3. 按 `slide.layout` 调用对应布局函数。
4. 从 `asset_id` 解析本地图片路径，按比例裁剪，禁止拉伸。
5. 添加来源脚注与演讲者备注。
6. 保存为 `deck_vNNN.pptx`。

第一版建议完全使用原生文本框、形状、表格和图表。这样用户下载后可以继续编辑，而不是得到“一页一张图片”的假 PPT。

### 步骤 8：生成真实预览

先确认本机已安装可用的 Microsoft PowerPoint。`preview.py` 在独立子进程中执行 COM 工作线程：

```python
import pythoncom
import win32com.client

pythoncom.CoInitialize()
powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
presentation = powerpoint.Presentations.Open(
    str(pptx_path), ReadOnly=True, Untitled=False, WithWindow=False
)
try:
    for index in range(1, presentation.Slides.Count + 1):
        presentation.Slides.Item(index).Export(
            str(output_dir / f"slide_{index:03d}.png"), "PNG", 1920, 1080
        )
finally:
    presentation.Close()
    powerpoint.Quit()
    pythoncom.CoUninitialize()
```

生产实现还需要由主进程设置超时，并先写临时目录、全部校验后再提交成功目录。Streamlit 展示这些 PNG，以发现字体替换、文本溢出和位置偏差。

### 步骤 9：实现多轮大白话修改

修改流程不是“把原提示词再发一遍”，而是：

```text
用户修改要求 + 当前 DeckSpec + 页面 ID
                  │
                  ▼
              PatchPlan
                  │
       校验目标、字段和操作权限
                  │
                  ▼
          应用到 DeckSpec 副本
                  │
       Pydantic 校验 + 保存新版本
                  │
                  ▼
          重新导出与真实预览
```

第一版只开放有限操作：

```python
from typing import Literal
from pydantic import BaseModel, Field

class PatchOperation(BaseModel):
    op: Literal[
        "update_theme", "update_slide", "insert_slide",
        "delete_slide", "move_slide"
    ]
    slide_id: str | None = None
    payload: dict = Field(default_factory=dict)

class PatchPlan(BaseModel):
    summary: str
    operations: list[PatchOperation]
```

示例：

```json
{
  "summary": "压缩第3页并将第5页改成时间轴，同时切换为深蓝主题",
  "operations": [
    {
      "op": "update_slide",
      "slide_id": "slide_003",
      "payload": {"layout": "bullets", "bullets": ["要点A", "要点B", "要点C"]}
    },
    {
      "op": "update_slide",
      "slide_id": "slide_005",
      "payload": {"layout": "timeline", "timeline": []}
    },
    {
      "op": "update_theme",
      "payload": {"primary": "17365D", "accent": "2F75B5"}
    }
  ]
}
```

应用 Patch 后重新构建整个文件是可以接受的，因为未修改页面仍由原 DeckSpec 原样生成；禁止再次让模型重写所有页面。每次修改保存 `versions/v001.json`、`v002.json`，回退就是加载旧版本并重新导出。

### 步骤 10：实现 Streamlit 页面

`app.py` 建议布局：

- 左侧：新建项目、上传素材、页数、主题、画面比例、当前版本。
- 中间上方：聊天历史与 `st.chat_input()`。
- 中间下方：大纲或页面 PNG 网格。
- 右侧/顶部按钮：确认大纲、生成 PPT、重新检查、回退版本、下载 PPTX。

状态至少包含：

```python
st.session_state.project_id
st.session_state.deck_spec
st.session_state.stage       # intake / outline / rendering / ready
st.session_state.version
st.session_state.chat_history
st.session_state.preview_paths
```

按钮行为必须明确：

- 第一次输入且无 DeckSpec：调用规划器。
- 已有 DeckSpec 时继续聊天：调用 PatchPlan 生成器。
- “确认大纲”：渲染 PPTX。
- “回退”：加载指定 JSON 版本并重新渲染。
- “下载”：只下载最新且已通过校验的文件。

### 步骤 11：加入第一版质量检查

`qa.py` 先做确定性检查，不必一开始就调用视觉模型：

- 标题过长、单页要点过多、单条文本过长。
- 空页面、重复标题、缺失素材、无效 `asset_id/source_id`。
- 图表数据为空、指标卡数量超限、时间轴节点超限。
- 中文字体是否存在；图片是否被拉伸。
- PowerPoint 是否成功导出全部 PNG；PNG 页数和尺寸是否符合 DeckSpec 与配置。
- PPTX ZIP 是否损坏，能否由 `python-pptx` 再次打开。

第二阶段再把页面 PNG 发给视觉模型，检查遮挡、对比度、密度、对齐和跨页风格；模型只返回问题列表和修复建议，最终修改仍通过 PatchPlan 完成。

### 步骤 12：运行 Demo

```powershell
uv run pytest
uv run ruff check .
uv run streamlit run app.py
```

浏览器访问 `http://localhost:8501`。

局域网临时演示：

```powershell
uv run streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

仅在可信局域网使用，并在 Windows 防火墙中只开放需要的网络范围。正式部署不要直接暴露 Streamlit 开发服务。

每个项目应产生：

```text
workspace/<project_id>/
├─ uploads/                 # 原始素材
├─ materials/               # 提取后的 Markdown
├─ assets/                  # 图片等素材
├─ versions/
│  ├─ v001.json
│  └─ v002.json
├─ output/
│  ├─ deck_v001.pptx
│  └─ deck_v002.pptx
└─ preview/
   ├─ v001/slide_001.png
   └─ v002/slide_001.png
```

## 6. 可交付 Demo 的开发顺序

| 里程碑 | 实现内容 | 完成标志 |
|---|---|---|
| M0 环境 | uv、Python、Microsoft PowerPoint、Streamlit | `streamlit hello` 和 PowerPoint COM 集成测试成功 |
| M1 单次生成 | 固定主题、9 种布局、DeckSpec、PPTX 导出 | 一句话生成 6～10 页可编辑 PPTX |
| M2 素材输入 | MarkItDown、文件保存、来源 ID | PDF/DOCX/PPTX 上传后能生成基于素材的内容 |
| M3 真实预览 | PPTX → PowerPoint COM → PNG | 页面预览来自最终 PPTX，页数和尺寸一致 |
| M4 多轮修改 | PatchPlan、稳定页面 ID、版本保存与回退 | 能准确执行“改第3页/换主题/加一页” |
| M5 质量检查 | 文本容量、资源、页数、文件完整性检查 | 不合格文件不允许直接下载 |
| M6 视觉增强 | 图片检索/生图、VLM 复核、自动修复 | 视觉质量提升且失败时可降级 |

建议严格按顺序完成。不要在 M1 还不稳定时先做多 Agent、向量数据库或复杂前端。

## 7. 快速验证现成项目

在自研之前，可用 [Presenton](https://github.com/presenton/presenton) 建立质量基线。它已经支持 Prompt/文档输入、模板、可编辑 PPTX、Docker 和桌面运行。

### Docker Desktop 部署

功能：隔离运行完整 Presenton，不污染本机 Python 环境。

```powershell
winget install --id Docker.DockerDesktop -e
```

安装后启动 Docker Desktop；首次使用可能要求启用 WSL2 并重启 Windows。然后执行：

```powershell
New-Item -ItemType Directory -Path presenton-demo
Set-Location presenton-demo

docker run -it --name presenton -p 5001:80 `
  -v "${PWD}\app_data:/app_data" `
  ghcr.io/presenton/presenton:latest
```

访问 `http://localhost:5001`，在页面中配置模型密钥。用同一批提示词和素材与自研 Demo 比较：内容准确性、视觉效果、PPTX 可编辑性、生成耗时和修改体验。

不建议直接复制其所有架构；先跑通并确认哪些能力需要复用或重写。

## 8. 第二阶段升级工具

第一版完成后再逐项引入。

### 8.1 PptxGenJS：更强的新建 PPTX 渲染器

功能：在 JavaScript/TypeScript 中生成文字、形状、图片、表格、图表和母版；适合把渲染器拆成独立服务。[官方仓库](https://github.com/gitbrent/PptxGenJS)

```powershell
winget install --id OpenJS.NodeJS.LTS -e
node --version
npm --version

New-Item -ItemType Directory -Path pptx-renderer
Set-Location pptx-renderer
npm init -y
npm install pptxgenjs
```

升级方式：Python Agent 输出相同 DeckSpec，通过 HTTP 或子进程传给 Node 渲染器。先保持 DeckSpec 不变，再替换底层导出器。

### 8.2 React Konva：自由画布编辑器

功能：在浏览器中拖拽、缩放、旋转、框选和对齐页面元素。[官方文档](https://konvajs.org/docs/react/)

```powershell
npm create vite@latest ppt-editor -- --template react-ts
Set-Location ppt-editor
npm install
npm install konva react-konva
npm run dev
```

画布只能读写 DeckSpec，不能创建第二套私有数据结构，否则网页预览与 PPTX 会逐渐不一致。

### 8.3 FastAPI：把单体拆成服务

功能：提供上传、项目、生成、修改、预览和下载 API；适合替换 Streamlit 后端。

```powershell
uv add fastapi "uvicorn[standard]" python-multipart
uv run uvicorn src.api:app --reload --host 127.0.0.1 --port 8000
```

### 8.4 office-kit/Open XML SDK：修改已有 PPTX

- [@office-kit/pptx](https://github.com/office-kit/pptx)：TypeScript 读取、修改、写回 PPTX；适合模板填充，当前 0.x 需锁定版本。
- [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)：微软维护的底层 OOXML SDK，适合需要更强格式保真和校验的场景，但开发成本高。

第一版只生成自己的 PPTX；等 PatchPlan、版本和质量检查稳定后，再做“上传现有 PPTX 并无损修改”。

### 8.5 数据库与任务队列

当单机目录不够用时再加入：

- PostgreSQL：用户、项目、DeckSpec 元数据、版本索引。
- S3/MinIO：原始文件、图片、PPTX、PDF 和预览图。
- Redis + RQ/Celery：生成、渲染和视觉检查等长任务。

Demo 阶段直接使用 `workspace/` 文件夹即可，避免过早增加部署复杂度。

## 9. 验收清单

### 功能

- [ ] 仅输入一句自然语言也能生成 PPT。
- [ ] 可上传至少 PDF、DOCX、PPTX、XLSX、Markdown 和图片。
- [ ] 生成前可确认或修改逐页大纲。
- [ ] 能用大白话修改指定页面、换主题、插入/删除/移动页面。
- [ ] 每次修改生成版本，可回退。
- [ ] 能下载 `.pptx`，PowerPoint/WPS 可直接打开。

### 文件质量

- [ ] PPTX 打开时无“需要修复文件”提示。
- [ ] 文字、形状、表格和常用图表可编辑。
- [ ] 页面无明显文字溢出、图片变形、元素遮挡和字体乱码。
- [ ] 预览 PNG 来自最终 PPTX 的实际渲染。
- [ ] 未指定修改的页面内容保持不变。
- [ ] 关键数字和结论能追溯到用户素材；缺乏依据时不编造。

### 工程质量

- [ ] `.env` 和用户素材不进入 Git。
- [ ] DeckSpec、PatchPlan 和环境配置都有 Pydantic 校验。
- [ ] 每种布局至少有一个渲染测试。
- [ ] 异常任务不会覆盖上一个成功版本。
- [ ] API 密钥只在服务端使用。

## 10. 最终推荐

第一版就使用：

`Streamlit + OpenAI Python SDK/Pydantic + MarkItDown + python-pptx + PowerPoint COM/pywin32 + 本地版本目录`

这条路线组件少、Windows 原生可运行，最适合快速做出真正能演示的产品闭环。Demo 稳定后，保持 DeckSpec/PatchPlan 协议不变，逐步替换为：

`React Konva + FastAPI + PptxGenJS + PostgreSQL/对象存储 + 视觉模型质检 + office-kit/Open XML SDK`

最重要的不是一开始堆很多 Agent，而是保证三件事：DeckSpec 始终是唯一事实源、多轮修改始终是可验证的局部 Patch、用户看到的预览始终来自最终导出的 PPTX。
