# Repository Instructions

## Mission

构建一个 Windows-first 的 AI PPT Agent：用户只输入文字要求，DeepSeek 生成经过校验的结构化 `DeckSpec`，程序使用 `python-pptx` 绘制原生可编辑 PowerPoint 对象，并通过 Microsoft PowerPoint COM 输出最终 PPTX 的真实 PNG 预览。

## Read Before Changing Anything

实施前只读取与任务相关的权威文档，并始终读取当前执行计划：

- 产品范围与验收：[docs/PRODUCT.md](docs/PRODUCT.md)
- 架构与技术边界：[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- 测试策略与质量门禁：[docs/TESTING.md](docs/TESTING.md)
- 当前 MVP 计划：[docs/exec-plans/active/mvp.md](docs/exec-plans/active/mvp.md)
- 历史调研参考：[AI_PPT_Agent_开源项目与技术方案.md](AI_PPT_Agent_开源项目与技术方案.md)；该文档不代表当前 MVP 范围。

## Long-Term Engineering Rules

- 保持 Windows 本地运行，使用 Python 3.12、`uv` 和 PowerShell。
- 当前 MVP 的唯一输入是用户文字，不实现文件上传、素材解析或用户图片。
- 使用 DeepSeek/OpenAI 兼容接口生成结构化 `DeckSpec`；所有模型输出必须再次通过 Pydantic 校验。
- `DeckSpec` 是大纲和渲染的唯一事实源；deck 与 slide 使用稳定且唯一的 ID。
- 模型描述内容、视觉层级和语义结构，程序拥有坐标、边距、字号下限、容量限制和绘制规则。
- 使用 `python-pptx` 生成文字、卡片、色块、节点、连接线等 PowerPoint 原生对象，不把整页扁平化成图片。
- 最终预览必须由生成后的 PPTX 经 PowerPoint COM 导出，不能使用单独绘制的近似预览。
- 模型提供商必须位于适配器之后，不让 SDK 类型污染领域模型。
- 失败结果不得覆盖或冒充最近一次成功产物；PowerPoint 子进程必须在成功、失败或超时后清理。
- 密钥只保存在被忽略的本地配置中，不得提交或写入日志。
- 当前 MVP 不实现素材解析、图片能力、`PatchPlan`、多轮局部修改、版本管理、版本回退、自动 QA 系统或预览哈希绑定。
- 不引入数据库、队列、Node 服务、自由画布或多 Agent 框架，除非先更新架构和执行计划。

## Change Discipline

- 将任务拆成可独立验证的小步骤，不混入无关重构。
- 每次行为变更都添加或更新测试。
- 先运行最窄相关测试，再执行 [docs/TESTING.md](docs/TESTING.md) 规定的完整门禁。
- 保留用户文件和无关工作区修改。
- 完成任务后在当前执行计划记录结果；测试和验收条件未通过时不得标记完成。

## Documentation Ownership

- 用户结果、范围和验收标准写入 `docs/PRODUCT.md`。
- 组件、合同、依赖和数据流写入 `docs/ARCHITECTURE.md`。
- 测试命令、矩阵和门禁写入 `docs/TESTING.md`。
- 实施顺序、进度和任务决策写入当前执行计划。
- 本文件只保存长期规则和文档链接，不复制详细方案。

## Documentation and Command Style

- 仓库文档使用简洁中文，代码标识符与命令保持英文。
- Windows 操作使用 PowerShell 示例。
- 文档链接使用仓库相对路径。
- 保持 `README.md` 的快速启动命令与实际代码一致。
