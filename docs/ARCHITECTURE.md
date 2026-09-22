# 架构设计

## 1. 架构目标

MVP 需要在 Windows 上低成本部署和调试，同时为后续前后端拆分、渲染器替换和现有 PPTX 编辑保留边界。

关键质量属性：

- 可重复：相同 `DeckSpec` 和主题应产生结构一致的 PPTX。
- 可验证：所有模型输出、文件路径和可下载产物都经过校验。
- 可修改：用户指令被翻译为局部 `PatchPlan`。
- 可恢复：失败不覆盖上一成功版本。
- 可替换：模型提供商、UI 和 PPTX 渲染器不进入领域合同。
- Windows-first：开发和 Demo 不依赖 WSL、Linux shell 或容器。

## 2. MVP 技术决策

| 领域 | 决策 | 说明 |
|---|---|---|
| 应用形态 | Python 单体 | 减少第一版部署与调试复杂度 |
| Python | 3.12 | Windows 支持成熟，类型与依赖生态完整 |
| 依赖管理 | uv | 负责 Python 安装、虚拟环境、锁文件和命令执行 |
| UI | Streamlit | 快速提供聊天、上传、状态、预览和下载 |
| 数据模型 | Pydantic | 校验配置、`DeckSpec`、`PatchPlan` 和 QA 报告 |
| 模型访问 | 提供商适配器 | 首个实现可用 OpenAI Python SDK，但领域层不依赖 SDK 类型 |
| 素材解析 | MarkItDown | 将常见 Office/PDF 文件转成 Markdown |
| PPTX 生成 | python-pptx | 第一版生成原生可编辑对象 |
| PPTX 预览 | Microsoft PowerPoint COM + pywin32 | 独立子进程逐页调用 `Slide.Export` 生成 PNG |
| 持久化 | 本地项目目录 | JSON 版本和文件产物；MVP 不使用数据库 |
| 测试 | pytest + Ruff | 单元、集成、端到端和静态检查 |

## 3. 系统上下文

```text
用户
 ├─ 自然语言需求
 ├─ 上传材料
 ├─ 大纲确认
 ├─ 修改指令
 └─ 下载/回退
        │
        ▼
Streamlit 应用
 ├─ 项目与会话
 ├─ 工作流协调
 └─ UI 状态
        │
        ├────────► 模型提供商
        │            ├─ 生成 DeckSpec
        │            └─ 生成 PatchPlan
        │
        ├────────► 本地工作区
        │            ├─ 上传文件
        │            ├─ 解析材料
        │            ├─ JSON 版本
        │            └─ PPTX/PNG
        │
        └────────► PowerPoint COM 子进程
                     └─ PPTX → 逐页 PNG
```

## 4. 模块边界

计划中的模块及职责：

```text
app.py
└─ 仅处理 Streamlit 组件、用户事件和视图状态

src/demo_workflow.py
└─ 固定 DeckSpec Demo 的项目隔离、快照、PPTX 渲染和预览编排

src/config.py
└─ 读取并校验环境变量、路径、模型和限制

src/models.py
└─ 领域合同：DeckSpec、SlideSpec、ThemeSpec、PatchPlan、QAReport

src/ingest.py
└─ 保存上传、解析文件、建立 source/asset 注册表

src/llm.py
└─ 模型提供商接口、结构化调用、错误映射和可测试替身

src/planner.py
└─ 用户请求 + 材料 → DeckSpec；不负责渲染

src/patcher.py
└─ 修改指令 → PatchPlan；验证并原子应用到 DeckSpec 副本

src/layouts.py
└─ 确定性页面布局函数与容量规则

src/pptx_renderer.py
└─ DeckSpec + 资源注册表 → PPTX

src/preview.py
└─ PPTX → PowerPoint COM → PNG；管理独立子进程、超时、清理与错误

src/qa.py
└─ 结构、内容容量、引用、文件完整性和预览一致性检查

src/storage.py
└─ 安全项目路径、不可变版本、当前指针和产物提交
```

依赖方向：

```text
UI/工作流
  ├─ planner/patcher
  ├─ renderer/preview/qa
  └─ storage
        │
        ▼
领域模型（models）

基础设施（llm、MarkItDown、python-pptx、pywin32/PowerPoint COM）
只能通过各自模块进入，不反向污染领域模型。
```

## 5. 核心数据合同

### 5.1 DeckSpec

`DeckSpec` 是唯一事实源，至少包含：

- deck ID、标题、受众、目的、语言和画面比例；
- 主题颜色、字体和主题名；
- 来源注册表；
- 有顺序的页面列表；
- 每页稳定 ID、布局意图、内容、资源引用、来源引用和备注。

约束：

- deck、slide、source、asset ID 创建后不可复用。
- 页面显示序号由列表顺序计算，不作为 ID。
- 模型不能写本地绝对路径，只能引用已注册 ID。
- 页面只允许使用注册的布局名称和字段。
- 任何读入的 JSON 都必须经过当前 schema 校验。
- schema 发生不兼容变化时增加 `schema_version` 并提供迁移。

### 5.2 PatchPlan

MVP 仅允许以下操作：

- `update_theme`
- `update_slide`
- `insert_slide`
- `delete_slide`
- `move_slide`

应用规则：

1. 解析模型输出并校验 schema。
2. 校验目标 ID 存在、操作字段允许、顺序合法。
3. 深拷贝当前 DeckSpec。
4. 在副本上应用全部操作。
5. 对结果执行完整 DeckSpec 校验和 QA 前置检查。
6. 全部成功后写入新版本；任何失败都丢弃副本。

### 5.3 QAReport

QA 结果应结构化，至少包含：

- `status`: pass/warn/fail；
- issue code、严重级别、slide ID、说明；
- 可自动修复标记；
- 检查器版本；
- 被检查的 DeckSpec 版本和 PPTX 哈希。

## 6. 端到端工作流

### 6.1 创建流程

```text
创建项目
→ 保存上传文件
→ 解析材料和注册来源/资源
→ 生成并校验 DeckSpec 草稿
→ 展示大纲
→ 用户确认
→ 保存 v001 DeckSpec
→ 生成临时 PPTX
→ 文件 QA
→ 用 PowerPoint 导出逐页 PNG
→ 预览 QA
→ 原子提交 v001 产物
→ 允许下载
```

### 6.2 修改流程

```text
用户修改要求
→ 读取当前成功版本
→ 生成并校验 PatchPlan
→ 在 DeckSpec 副本上原子应用
→ 保存候选 vNNN
→ 重建 PPTX/预览/QA
→ 成功：提交并移动 current 指针
→ 失败：保留旧 current，候选标记失败
```

### 6.3 回退流程

回退不修改历史 JSON。系统选择既有成功版本作为基础，重新生成一个新的版本或移动明确的当前指针；具体策略在实现前由执行计划任务确定并通过测试固定。

## 7. 布局与渲染

### 7.1 支持的 MVP 布局

- `cover`
- `section`
- `bullets`
- `two_column`
- `image_text`
- `metrics`
- `timeline`
- `chart`
- `closing`

每个布局必须定义：

- 必需字段；
- 最大内容容量；
- 字号下限；
- 图片槽位与裁剪规则；
- 支持的图表类型；
- 来源脚注区域；
- 确定性的降级策略。

### 7.2 渲染约束

- MVP 默认 16:9。
- 坐标、边距、字号和颜色来自布局与主题，不来自自由文本提示。
- 图片保持纵横比，通过裁剪填充，不直接拉伸。
- 缺失可选图片时使用确定性无图布局；不生成破损占位符。
- 内容超过布局容量时 QA 失败或由明确规则降级；禁止静默缩到不可读字号。
- 优先使用原生文本、形状、表格和图表。
- 生成文件必须能被 `python-pptx` 再次打开。

## 8. 预览架构

最终预览链路固定为：

```text
当前版本 PPTX
→ 启动独立 Python 子进程
→ 初始化 COM 并启动独立 PowerPoint 实例
→ 逐页调用 `Slide.Export` 生成 PNG
→ 比较 PPTX/DeckSpec/PNG 页数与图片尺寸
→ UI 展示 PNG
```

外部进程要求：

- 仅在 Windows 启用，启动时验证 Microsoft PowerPoint COM 可用。
- 主进程对子进程设置超时并捕获 stdout/stderr；超时时终止本次创建的 PowerPoint 进程。
- COM 线程调用 `CoInitialize`/`CoUninitialize`，并在 `finally` 中关闭 Presentation、退出 PowerPoint。
- 以只读、无窗口方式打开最终 PPTX，按配置的宽高逐页导出。
- 每次转换先写独立临时目录；全部验证通过后再原子移动到新的成功目录。
- 输出文件必须位于当前项目工作区。
- 文件名固定为 `slide_NNN.png`，PNG 数量必须等于 PPTX 页数，尺寸必须符合配置。

### 8.1 无 LLM Streamlit 最小闭环

当前阶段使用固定 `sample_deck.json` 验证 UI 到真实产物的链路：

```text
fixture → DeckSpec 校验 → 大纲展示 → 新 project_id
→ workspace/<project_id>/deck.json + PPTX + PNG
→ Streamlit 预览与 PPTX 下载
```

- `app.py` 只保存 `project_id`、`deck_spec`、`stage` 和产物路径等视图状态。
- `demo_workflow.py` 负责创建隔离项目并调用渲染器与预览后端。
- 每次点击都创建新项目；开始新任务前清空旧下载路径。
- 只有 PPTX 与全部真实 PNG 成功后才能进入 `ready`。
- PowerPoint 不可用或任一阶段失败时进入 `error`，不展示伪造预览或下载入口。

## 9. 本地存储

计划目录：

```text
workspace/<project_id>/
├─ project.json
├─ uploads/
├─ materials/
├─ assets/
├─ versions/
│  ├─ v001/
│  │  ├─ deck.json
│  │  ├─ qa.json
│  │  └─ status.json
│  └─ v002/
├─ output/
│  ├─ deck_v001.pptx
│  └─ deck_v002.pptx
└─ preview/
   └─ v001/slide_001.png
```

存储规则：

- `project_id` 由程序生成，不接受用户提供的目录名。
- 上传文件名只用于显示；磁盘名使用生成的 ID 和安全扩展名。
- 写 JSON 和最终产物使用临时文件后原子替换。
- 版本目录不可就地修改。
- `current` 状态只指向成功版本。
- 整个 `workspace/` 不进入 Git。

## 10. 模型提供商边界

模型适配器至少提供：

- `generate_deck(request, materials) -> DeckSpec`
- `generate_patch(instruction, current_deck) -> PatchPlan`

要求：

- 领域层只接收 Pydantic 模型或领域异常。
- SDK 响应、token 统计和 HTTP 错误在适配器内转换。
- 支持超时、有限重试和不可重试错误分类。
- 测试默认使用确定性 fake provider，不依赖网络和真实密钥。
- 第三方 OpenAI-compatible 服务只有通过契约测试后才视为受支持。

### 10.1 当前 DeckSpec 规划实现

当前阶段只实现无素材的“自然语言要求 → DeckSpec”，不接入 Streamlit：

```text
用户要求 → planner 加载 prompts/plan_deck.md
        → OpenAI Responses API structured output
        → SDK 按 DeckSpec 解析
        → 转回普通数据并再次执行 Pydantic 校验
        → 合法 DeckSpec
```

- `src/config.py` 从环境读取 `OPENAI_API_KEY`、`OPENAI_BASE_URL`、`OPENAI_MODEL` 和 `LLM_TIMEOUT_SECONDS`；密钥使用 `SecretStr`，不得进入日志或异常正文。
- `src/llm.py` 是 SDK 边界，使用 `responses.parse(..., text_format=DeckSpec)`，并把超时、连接失败、拒绝、空结果和非法结构转换为领域异常。
- `src/planner.py` 负责提示词、输入长度、显式页数和最终二次校验；当前不读取素材、不写项目版本。
- 模型名不硬编码，由部署环境选择支持 Responses API 结构化输出的模型；通过显式 smoke test 验证具体账户与模型组合。

## 11. 安全与隐私

- API 密钥来自 `.env` 或进程环境，绝不进入 DeckSpec、日志或前端状态。
- 不记录完整用户材料和模型请求正文到默认日志。
- 上传扩展名、MIME、大小和实际解析结果都需要校验。
- 拒绝绝对路径、父目录跳转和符号链接逃逸。
- 远程 URL 下载默认关闭；启用后必须限制协议、大小、超时和重定向。
- 不执行上传文件中的宏、脚本或嵌入对象。
- 用户材料、生成产物和本地密钥均由 `.gitignore` 排除。

## 12. 质量门禁

可下载版本必须满足：

1. DeckSpec schema 有效。
2. 所有 source/asset 引用有效。
3. 内容容量规则无 fail 级问题。
4. PPTX 成功生成，ZIP/OOXML 基本完整，可再次打开。
5. PowerPoint COM 成功导出全部 PNG，且进程正确退出。
6. DeckSpec、PPTX 和 PNG 页数一致，PNG 尺寸符合配置。
7. QAReport 不含 fail 级问题。

详细测试要求见 [TESTING.md](TESTING.md)。

## 13. 演进边界

MVP 后允许替换：

- Streamlit → React + FastAPI；
- python-pptx → PptxGenJS 或其他渲染服务；
- 本地目录 → PostgreSQL + 对象存储；
- 同步调用 → Redis + 任务队列；
- 固定布局 → React Konva 可视编辑；
- 新建 PPTX → office-kit/Open XML SDK 增量编辑。

替换条件：

- `DeckSpec`/`PatchPlan` 有版本化合同；
- 现有 golden fixtures 可在新实现上通过；
- 最终预览仍来自导出的 PPTX；
- 未降低可编辑性、来源追溯和失败回退能力。

## 14. 相关文档

- 产品范围：[PRODUCT.md](PRODUCT.md)
- 测试策略：[TESTING.md](TESTING.md)
- MVP 执行计划：[exec-plans/active/mvp.md](exec-plans/active/mvp.md)
