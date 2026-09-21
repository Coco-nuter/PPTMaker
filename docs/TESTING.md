# 测试策略

## 1. 测试目标

测试不仅验证 Python 函数是否运行，还要验证最终 PPTX 是否可打开、可编辑、可预览、可修改和可回退。

优先保证：

- 领域合同稳定；
- 未命中的页面不会因修改而漂移；
- 失败不会覆盖上一成功版本；
- 最终文件与预览一致；
- 测试默认不依赖网络、真实 API 密钥或人工判断。

## 2. 测试分层

### 2.1 静态检查

工具：Ruff。

覆盖：

- 格式和常见 Python 错误；
- 未使用导入；
- 不安全或可疑代码模式；
- 可配置的复杂度与命名规则。

计划命令：

```powershell
uv run ruff check .
uv run ruff format --check .
```

### 2.2 单元测试

不访问网络、LibreOffice 或真实文件系统边界之外的资源。

重点：

- Pydantic 模型有效/无效输入；
- ID 生成和稳定性；
- 内容容量规则；
- PatchPlan 操作验证；
- Patch 原子应用和未命中页面不变；
- 主题合并；
- 安全路径解析；
- 项目版本号和状态转换；
- QA issue 分级。

### 2.3 组件测试

使用临时目录和本地 fixtures，允许调用单个真实组件。

重点：

- MarkItDown 解析每种支持格式；
- python-pptx 为每种布局生成有效幻灯片；
- PPTX 可由 `python-pptx` 再次打开；
- PyMuPDF 将 fixture PDF 转成预期页数 PNG；
- storage 写入失败时不移动 current 指针；
- fake LLM provider 返回结构化数据并触发正常工作流。

### 2.4 集成测试

组合多个本地组件，但默认不调用真实模型。

重点：

- 用户请求 + fake provider → DeckSpec → PPTX；
- 上传 fixture → Markdown → DeckSpec；
- DeckSpec → PPTX → PDF → PNG；
- 当前版本 + 修改指令/fake PatchPlan → 新版本；
- 失败候选版本 → 上一成功版本仍可下载；
- 回退 → 生成与目标历史版本一致的结构化内容。

### 2.5 端到端测试

通过 Streamlit 测试工具或浏览器自动化覆盖核心用户旅程：

1. 新建项目。
2. 上传 fixture。
3. 输入自然语言请求。
4. 查看并确认大纲。
5. 生成 PPTX 和预览。
6. 修改指定页面。
7. 下载新版本。
8. 回退旧版本。

端到端测试仍使用 fake provider，保证稳定和低成本。

### 2.6 手工兼容性测试

发布前在真实应用中检查：

- Microsoft PowerPoint 当前支持版本；
- WPS Office 当前支持版本；
- LibreOffice Impress。

检查文件打开、文字编辑、形状编辑、图表编辑、字体、换行、图片裁剪和备注。手工结果记录在发布检查单，不代替自动化测试。

## 3. Fixture 设计

计划目录：

```text
tests/
├─ fixtures/
│  ├─ requests/
│  ├─ materials/
│  │  ├─ sample.pdf
│  │  ├─ sample.docx
│  │  ├─ sample.pptx
│  │  ├─ sample.xlsx
│  │  ├─ sample.md
│  │  └─ sample.png
│  ├─ deck_specs/
│  ├─ patch_plans/
│  ├─ expected/
│  └─ fonts/
├─ unit/
├─ component/
├─ integration/
└─ e2e/
```

Fixture 规则：

- 只使用仓库可合法分发的虚构内容。
- 文件尽量小，避免测试仓库膨胀。
- 数字和结论刻意设计成易于验证，便于检查是否编造。
- 至少包含中文、英文、长标题、超量要点、缺图和无效引用等边界情况。
- 二进制 fixture 必须在旁边提供来源说明或生成脚本说明。
- 不使用真实用户材料作为测试数据。

## 4. 模型调用测试

### 4.1 默认策略

自动化测试使用 `FakeModelProvider`：

- 输入固定 request fixture；
- 返回固定且合法的 DeckSpec/PatchPlan；
- 可配置返回超时、拒绝、非法 JSON、非法字段和空结果；
- 记录调用参数供断言，不发送网络请求。

### 4.2 提供商契约测试

真实提供商测试单独标记，例如 `provider`，默认跳过。运行条件：

- 显式环境变量开启；
- 本机有有效密钥；
- 使用最小请求；
- 不在普通 CI 中自动执行；
- 只验证 schema、错误分类和最小可用性，不比较自由文本措辞。

### 4.3 禁止事项

- 单元/集成测试不能依赖模型的随机自然语言输出。
- 不能把 API 密钥写进 fixture、快照或测试日志。
- 不能把真实调用失败简单重试到测试通过。

## 5. DeckSpec 与 Patch 不变量

必须有自动化测试验证：

- ID 唯一且稳定。
- 页面顺序变化不改变页面 ID。
- 不存在的 slide/source/asset ID 被拒绝。
- 非法布局名和字段被拒绝。
- 删除、移动、插入边界正确。
- 一个 PatchPlan 要么全部成功，要么全部不生效。
- `update_slide` 不能修改其他页面。
- `update_theme` 不改变页面内容。
- 旧版本 JSON 不会被新修改覆盖。
- schema 版本不兼容时给出明确错误。

## 6. 渲染测试

### 6.1 每种布局的最小测试

每个布局至少覆盖：

- 最小合法内容；
- 接近容量上限的内容；
- 缺失可选资源；
- 无效必需字段；
- 中文和英文文本；
- 主题颜色与字体应用。

### 6.2 文件完整性

对每个渲染 fixture：

- 输出文件存在且非空；
- ZIP 包可打开；
- `python-pptx` 能重新加载；
- 幻灯片数量正确；
- 关键原生对象数量和类型符合预期；
- 不包含意外的全页截图替代原生内容。

### 6.3 Golden 策略

不要对整个 PPTX 二进制做字节快照，因为 ZIP 时间戳等元数据可能变化。优先比较：

- 规范化 DeckSpec；
- 幻灯片数量、shape 类型、文本和位置的结构摘要；
- PDF/PNG 的页数、尺寸和有限视觉差异指标；
- 手工批准的小规模基准图片。

视觉快照变化必须说明原因并人工查看差异。

## 7. 预览测试

Windows 集成环境需要安装 LibreOffice。测试内容：

- 正确发现配置的 `soffice.exe`。
- 路径包含空格时可以调用。
- 转换超时会终止并报告错误。
- 输出目录隔离，不读取其他项目文件。
- PPTX、PDF 和 PNG 页数一致。
- PNG 命名顺序稳定。
- 并发转换不会覆盖彼此产物。
- LibreOffice 缺失时给出可操作错误，不伪造预览。

LibreOffice 集成测试可用 `libreoffice` marker；在未安装环境中明确 skip，而不是静默通过。

## 8. 素材与安全测试

至少覆盖：

- 支持格式成功解析。
- 不支持扩展名被拒绝。
- 扩展名与实际内容不符。
- 空文件、损坏文件和超大文件。
- `../`、绝对路径、保留设备名和异常 Unicode 文件名。
- 同名上传不会覆盖。
- 图片纵横比保持。
- 远程 URL 默认禁用。
- `.env`、密钥文件和工作区不会被 Git 跟踪。

## 9. MVP 测试矩阵

| 能力 | 单元 | 组件 | 集成 | E2E | 手工 |
|---|---:|---:|---:|---:|---:|
| DeckSpec 校验 | 必须 |  | 必须 |  |  |
| PatchPlan/版本 | 必须 | 必须 | 必须 | 必须 |  |
| Office/PDF 素材解析 |  | 必须 | 必须 | 必须 |  |
| 9 种页面布局 | 必须 | 必须 | 必须 |  | 必须 |
| PPTX 文件完整性 |  | 必须 | 必须 | 必须 | 必须 |
| LibreOffice 预览 |  | 必须 | 必须 | 必须 | 必须 |
| Streamlit 核心流程 |  |  | 必须 | 必须 | 必须 |
| PowerPoint/WPS 兼容 |  |  |  |  | 必须 |
| 模型提供商 |  | 契约测试 | 可选 | 可选 | 必须 |

## 10. 计划命令

项目初始化后，标准本地门禁计划为：

```powershell
uv run ruff check .
uv run ruff format --check .
uv run pytest -m "not provider"
```

只运行 LibreOffice 集成：

```powershell
uv run pytest -m libreoffice
```

显式运行真实提供商契约测试：

```powershell
$env:RUN_PROVIDER_TESTS = "1"
uv run pytest -m provider
```

这些命令在 M0/M1 创建项目配置后生效；README 必须同步更新。

## 11. 缺陷严重级别

- P0：文件损坏、密钥泄漏、越权文件访问、历史版本丢失。
- P1：核心流程无法完成、未修改页面被改写、预览对应错误版本。
- P2：某一布局失败、可编辑性降低、明显溢出或字体错误。
- P3：轻微视觉偏差、文案提示或非核心体验问题。

P0/P1 阻止任何 MVP 发布；P2 必须有明确降级或已接受记录。

## 12. 完成定义

一个开发任务只有同时满足以下条件才算完成：

- 实现符合产品与架构文档；
- 任务中列出的独立测试通过；
- 新增行为有自动化覆盖；
- 相关静态检查通过；
- 错误路径不会覆盖上一成功版本；
- 用户可见行为或命令变化已更新文档；
- 活跃执行计划中的状态和发现已更新。

## 13. 相关文档

- 产品范围：[PRODUCT.md](PRODUCT.md)
- 架构：[ARCHITECTURE.md](ARCHITECTURE.md)
- MVP 任务：[exec-plans/active/mvp.md](exec-plans/active/mvp.md)

