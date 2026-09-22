# MVP 执行计划：纯文字生成原生可编辑 PPT

## 状态

- 状态：Active
- 当前阶段：第 10 步 DeckSpec 视觉结构扩展已完成；下一步实现原生形状版式
- 更新时间：2026-09-22
- 产品范围：[../../PRODUCT.md](../../PRODUCT.md)
- 架构约束：[../../ARCHITECTURE.md](../../ARCHITECTURE.md)
- 测试策略：[../../TESTING.md](../../TESTING.md)

## 目标流程

```text
用户文字要求
→ DeepSeek 生成结构化 DeckSpec
→ 用户确认大纲
→ python-pptx 绘制文字、卡片、色块、节点和连接线
→ 可编辑 PPTX
→ PowerPoint COM 导出 PNG
→ 预览与下载
```

## 范围约束

本计划只处理纯文字输入和 PowerPoint 原生对象。以下内容已从 MVP 计划删除：

- 文件上传，以及 Markdown、TXT、PDF、DOCX、PPTX、XLSX 素材解析；
- MarkItDown、`SourceRef` 和 `AssetRef` 的后续使用；
- 用户图片、图片搜索和 AI 文生图；
- `PatchPlan`、多轮局部修改；
- 版本管理和版本回退；
- 自动 QA 系统和 PPTX 哈希/预览版本绑定；
- `image_text`、`full_image` 和 `chart` 布局。

[历史调研方案](../../../AI_PPT_Agent_开源项目与技术方案.md)只作为参考资料，不用于确定当前范围。

## 当前已完成

- [x] 使用 Python 3.12、uv、pytest 和 Ruff 初始化 Windows 工程。
- [x] 将 `DeckSpec` 升级为 2.0；九种布局使用独立模型和 `layout` discriminator，删除来源与素材字段。
- [x] 使用 `python-pptx` 确定性生成三页示例 PPTX，文字和形状可编辑。
- [x] 实现 PowerPoint COM 独立子进程预览，逐页导出 1920×1080 PNG，并处理超时和清理。
- [x] 实现 Streamlit 最小闭环：大纲展示、确认、PPTX 生成、真实预览和下载。
- [x] 实现 OpenAI SDK 适配器、结构化输出、Pydantic 二次校验、FakeProvider 和显式真实 smoke test。
- [x] 将自然语言规划器接入 Streamlit，支持确认、重新规划和放弃候选大纲。
- [x] 使用真实 DeepSeek 兼容配置验证自然语言规划和完整 Streamlit 生成链路。

## 接下来的三个任务

### 10. 扩展 DeckSpec 视觉结构（已完成）

目标：让 DeepSeek 能表达页面的语义化视觉结构，同时保持渲染确定性。

实施步骤：

- [x] 盘点当前 `DeckSpec`、fixture、提示词和渲染器的实际依赖。
- [x] 为九种布局建立独立 Pydantic 模型，并用 `layout` 组成 discriminated union。
- [x] 为列表、指标、时间轴、流程和对比项定义文字与数量限制。
- [x] 保持 `slide_id` 稳定唯一，不向模型开放坐标、页面颜色、图片字段或绘图代码。
- [x] 删除 `SourceRef`、`AssetRef`、Deck 级来源/素材和页面引用字段。
- [x] 将 `sample_deck.json` 更新为九布局合同样例，并保留三布局兼容 fixture 供现有渲染链路使用。
- [x] 更新 Pydantic 单元测试及直接受合同拆分影响的大纲展示。

独立测试：

```powershell
uv run pytest tests/test_models.py -q
uv run ruff check src/models.py tests/test_models.py
```

完成条件：合法视觉 `DeckSpec` 可 JSON 往返；重复 ID、悬空连接、非法颜色、超量内容、图片/图表布局和任意坐标均被拒绝。

### 11. 实现原生形状版式渲染

目标：使用 `python-pptx` 把视觉 `DeckSpec` 绘制成布局稳定、可逐对象编辑的 PPTX。

实施步骤：

- [ ] 提取统一画布、网格、安全边距、字体、颜色和层级原语。
- [ ] 实现文字、卡片、色块、节点和连接线的共享绘制函数。
- [ ] 为受支持布局实现确定性排版和容量处理，不接受模型坐标。
- [ ] 保证所有布局函数不修改原始 `DeckSpec`。
- [ ] 保存后使用 `python-pptx` 重新打开，检查页数、主要文字、形状类型和对象数量。
- [ ] 生成一份覆盖全部受支持布局的示例 PPTX，并用 PowerPoint 导出真实 PNG 进行人工检查。

独立测试：

```powershell
uv run pytest tests/test_renderer.py -q
uv run pytest -m "integration and powerpoint" -q
uv run ruff check src tests
```

完成条件：示例文稿全部页面可打开；文字、卡片、色块、节点和连接线是 PowerPoint 原生对象；无整页图片、图片布局或图表布局；PNG 数量和尺寸正确。

### 12. 将 DeepSeek 对齐视觉 DeckSpec

目标：让真实和 Fake Provider 都能生成新视觉合同，并维持现有 Streamlit 确认后生成流程。

实施步骤：

- [ ] 更新 `prompts/plan_deck.md`，要求先规划故事线，再生成页面结论和受控视觉结构。
- [ ] 明确禁止编造数字和事实、禁止图片、禁止图表、禁止自由坐标。
- [ ] 更新 FakeProvider 和 Mock 响应，使普通测试不访问网络。
- [ ] 调整 planner 与大纲视图，让用户能看懂每页主要内容和视觉结构。
- [ ] 保持“未确认不渲染；重新规划不复用旧候选；失败不覆盖成功产物”。
- [ ] 显式运行真实 DeepSeek smoke test，验证指定页数、合法布局和 Pydantic 二次校验。
- [ ] 完成 Streamlit 端到端人工验收：文字输入 → 大纲 → 确认 → PPTX → PNG → 下载。

独立测试：

```powershell
uv run pytest tests/test_llm.py tests/test_planner.py tests/test_app_workflow.py tests/test_app.py -q
$env:RUN_OPENAI_SMOKE_TESTS='1'
uv run pytest tests/test_openai_smoke.py -m provider -q
uv run pytest -q
uv run ruff check .
```

完成条件：普通测试完全离线；真实 DeepSeek 返回合法视觉 `DeckSpec`；用户确认后可生成可编辑 PPTX、显示全部真实 PNG 并下载，结束后无残留 PowerPoint 进程。

## MVP 退出条件

- [ ] 一条纯文字要求可生成 6～10 页合法视觉 `DeckSpec`。
- [ ] 用户可在生成前确认、重新规划或放弃大纲。
- [ ] PPTX 使用原生文字、卡片、色块、节点和连接线，所有核心对象可编辑。
- [ ] 不需要文件上传、外部素材或图片服务。
- [ ] PowerPoint 实际导出全部 1920×1080 PNG，页数一致且无残留进程。
- [ ] 用户可下载 PowerPoint 能正常打开的 PPTX。
- [ ] `uv run pytest -q` 与 `uv run ruff check .` 通过。

## 决策记录

| 日期 | 决策 | 原因 |
|---|---|---|
| 2026-09-21 | 使用 Windows 原生 Python 单体 | 最快形成可运行闭环 |
| 2026-09-21 | `DeckSpec` 作为大纲和渲染唯一事实源 | 隔离模型输出与确定性绘制 |
| 2026-09-22 | 预览固定使用 PowerPoint COM 独立子进程 | 保证预览来自最终 PPTX，并隔离卡死风险 |
| 2026-09-22 | 普通测试使用 FakeProvider | 避免网络、费用和随机性 |
| 2026-09-22 | MVP 改为纯文字输入和原生形状绘制 | 缩小范围，优先提升视觉结构、可编辑性和本地可运行性 |

## 已知限制

- PowerPoint 预览仅支持 Windows，并要求安装可用的 Microsoft PowerPoint。
- `python-pptx` 不覆盖复杂动画、SmartArt 和全部 Office 特性。
- 中文字体和换行会受目标机器已安装字体影响。
- DeepSeek 兼容端点必须支持项目当前使用的结构化输出调用方式。
- 当前已工作的三种基础布局视觉表现有限，需要任务 9.1 和 9.2 扩展。
