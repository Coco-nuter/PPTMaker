# 架构设计

## 1. 架构目标

MVP 是 Windows 本地 Python 单体。它只接收用户文字，让 DeepSeek 生成受控的视觉结构，再由确定性代码绘制可编辑 PPTX，并用 Microsoft PowerPoint 渲染真实预览。

关键属性：

- 可验证：模型输出必须通过 Pydantic 校验。
- 可重复：同一 `DeckSpec` 应产生结构一致的 PPTX。
- 可编辑：文字和视觉元素使用 PowerPoint 原生对象。
- 预览真实：PNG 必须来自最终 PPTX。
- 可清理：PowerPoint 在成功、失败和超时后都能退出。
- Windows-first：运行和开发不依赖 WSL、容器或 Linux 工具链。

## 2. 固定数据流

```text
用户文字要求
→ Streamlit 收集主题、受众、用途、页数和补充要求
→ Planner 调用 DeepSeek/OpenAI 兼容接口
→ Pydantic 校验 DeckSpec
→ Streamlit 展示大纲
→ 用户确认
→ python-pptx 绘制原生对象
→ 保存可编辑 PPTX
→ PowerPoint COM 子进程逐页导出 PNG
→ Streamlit 显示预览并提供 PPTX 下载
```

用户重新规划时创建新的候选 `DeckSpec`。确认前不生成文件；失败候选不替换已有成功产物。

## 3. 技术决策

| 领域 | 决策 | 说明 |
|---|---|---|
| 应用形态 | Python 3.12 单体 | Windows 上快速部署和调试 |
| 依赖管理 | uv | 管理 Python、虚拟环境和锁文件 |
| UI | Streamlit | 文字输入、大纲确认、预览和下载 |
| 数据合同 | Pydantic `DeckSpec` | 约束内容、主题和视觉结构 |
| 模型 | DeepSeek/OpenAI 兼容接口 | 使用 OpenAI Python SDK 适配，配置由环境提供 |
| PPTX | python-pptx | 绘制原生可编辑对象 |
| 预览 | PowerPoint COM + pywin32 | 独立子进程调用 `Slide.Export` |
| 运行产物 | 隔离的本地 workspace | 每次确认生成独立目录，不提供版本系统 |
| 测试 | pytest + Ruff | Mock 单元测试与显式真实集成测试 |

MVP 不使用 MarkItDown、LibreOffice、图片生成服务、图片搜索、数据库或任务队列。

## 4. 模块边界

```text
app.py
└─ Streamlit 组件、用户事件和 session state

src/app_workflow.py
└─ 输入校验、planner 请求构造和大纲工作流

src/config.py
└─ 模型、预览、超时和尺寸配置

src/models.py
└─ DeckSpec、SlideSpec、ThemeSpec 及视觉结构合同

src/llm.py
└─ DeepSeek/OpenAI 兼容 SDK 调用、结构化输出和错误映射

src/planner.py
└─ 用户文字要求 → 合法 DeckSpec

src/pptx_renderer.py
└─ DeckSpec → 原生可编辑 PPTX

src/demo_workflow.py
└─ project_id、隔离工作区、PPTX 渲染和预览编排

src/preview.py
└─ PPTX → PowerPoint COM → PNG；进程隔离、超时和清理
```

UI 只依赖工作流接口；模型 SDK、`python-pptx` 和 COM 分别限制在基础设施模块中。

## 5. DeckSpec 合同

`DeckSpec` 是大纲与 PPTX 渲染的唯一事实源，目标合同包含：

- deck ID、标题、受众、用途、语言和 16:9 画面比例；
- 主题名、字体和受校验的颜色；
- 有顺序的页面列表；
- 每页稳定唯一的 `slide_id`、布局、主要结论、文字内容和语义化视觉结构。

视觉结构只描述“画什么”和元素之间的关系，例如：

- 标题、正文、标签和强调文字；
- 卡片组、色块和分区；
- 节点、步骤和连接关系；
- 页面重点、层级和排列意图。

模型不得输出任意坐标、磁盘路径、图片 URL 或 PowerPoint 内部对象 ID。渲染器根据布局目录决定准确坐标、尺寸、边距、字号和降级规则。

`DeckSpec` 2.0 使用 `layout` 作为 discriminator，由九个独立页面模型组成：`cover`、`section`、`bullets`、`two_column`、`metrics`、`timeline`、`process`、`comparison`、`closing`。当前渲染器仍只实现 `cover`、`bullets`、`closing`；其余布局在下一阶段实现本地绘制。`image_text`、`full_image` 和 `chart` 不得进入 MVP。

`SourceRef`、`AssetRef`、Deck 级 `sources`/`assets` 和页面级 `source_ids`/`asset_ids` 已从数据合同删除。

## 6. DeepSeek 规划边界

```text
用户要求
→ prompts/plan_deck.md
→ OpenAI SDK 调用兼容接口
→ 模型结构化输出
→ 转成普通数据
→ Pydantic 二次校验
→ DeckSpec
```

要求：

- 先设计整份演示文稿的故事线，再规划页面。
- 每页只表达一个主要结论。
- 页数必须符合用户要求。
- 不编造具体数字、事实、引用或数据来源。
- 只使用当前 schema 和布局目录允许的值。
- `slide_id` 稳定且唯一。
- 模型只生成结构，不生成图片，不直接生成 PPTX。
- 超时、网络失败、拒绝、空结果和非法结构映射为清晰领域错误。
- API Key 不进入日志、异常、DeckSpec 或前端状态。

普通测试使用 `FakeProvider` 或 Mock SDK；真实 DeepSeek smoke test必须显式启用。

## 7. 原生版式渲染

渲染器负责把语义化视觉结构映射为 PowerPoint 对象：

- 文本框：标题、正文、标签、页码；
- 自选图形：卡片、背景色块、强调框、节点；
- 连接线：步骤、流程和关系；
- 组合与层级：通过确定性顺序保持视觉结构。

约束：

- 默认 16:9，统一安全边距、字体、颜色和网格。
- 布局函数不得修改输入 `DeckSpec`。
- 每种布局定义字段上限、对象上限和最小字号。
- 内容超限必须在 schema 或渲染前明确拒绝，不能静默缩成不可读文字。
- 不使用用户图片、AI 图片、网络图片、整页截图或图表对象。
- 保存后必须能被 `python-pptx` 重新打开，页数和主要文字与 `DeckSpec` 一致。

## 8. PowerPoint 真实预览

```text
最终 PPTX
→ 主进程启动独立 Python 子进程并设置超时
→ pythoncom.CoInitialize()
→ DispatchEx("PowerPoint.Application")
→ 只读打开 PPTX
→ 每页 Slide.Export(..., "PNG", 1920, 1080)
→ 校验文件数量与尺寸
→ 全部成功后移动到最终预览目录
→ finally 关闭 Presentation、退出 PowerPoint、CoUninitialize()
```

- 只在 Windows 和安装 Microsoft PowerPoint 的环境启用。
- 文件名固定为 `slide_NNN.png`。
- 导出先写临时目录，避免留下半成品。
- 超时或失败时清理本次进程并返回明确错误。
- 普通测试不得启动 PowerPoint；真实测试使用 `integration` 与 `powerpoint` marker。

## 9. 运行目录与状态

```text
workspace/<project_id>/
├─ output/
│  └─ sample_deck.pptx
└─ preview/
   └─ preview-*/png/slide_NNN.png
```

- `project_id` 由程序生成。
- 每次确认创建独立工作区，避免错误任务复用旧文件。
- workspace 只保存运行产物，不实现版本历史、current 指针或回退。
- `workspace/`、`output/`、`.env` 和本地密钥文件不进入 Git。

Streamlit 状态至少包含 `intake`、`planning`、`outline`、`rendering`、`ready` 和 `error`。只有合法 `outline` 可以进入渲染，只有 PPTX 与全部真实 PNG 成功后才能进入 `ready`。

## 10. 必要验证

MVP 不建设独立自动 QA 系统，但每条链路仍执行必要的确定性验证：

1. 模型结果通过当前 `DeckSpec` schema。
2. PPTX 存在、非空、可由 `python-pptx` 重新打开。
3. PPTX 页数和主要文字与 `DeckSpec` 一致。
4. PNG 数量等于 PPTX 页数，尺寸符合配置。
5. PowerPoint 进程在成功、失败和超时后均已退出。

这些检查属于模块正确性和错误处理，不产生 `QAReport`，也不建立哈希或版本绑定。

## 11. 当前排除项

以下能力不属于 MVP，不创建对应模块或任务：

- 文件上传和 Markdown、TXT、PDF、DOCX、PPTX、XLSX 解析；
- MarkItDown、来源追踪和素材/图片注册表；
- 用户图片、图片搜索和 AI 文生图；
- `PatchPlan`、多轮局部修改；
- 版本管理和版本回退；
- 自动 QA 系统和哈希绑定；
- `image_text`、`full_image`、`chart` 布局。

如需重新引入，必须先更新产品、架构和新的执行计划。

## 12. 相关文档

- 产品范围：[PRODUCT.md](PRODUCT.md)
- 测试策略：[TESTING.md](TESTING.md)
- MVP 执行计划：[exec-plans/active/mvp.md](exec-plans/active/mvp.md)
- 历史调研参考：[../AI_PPT_Agent_开源项目与技术方案.md](../AI_PPT_Agent_开源项目与技术方案.md)；不作为当前架构范围。
