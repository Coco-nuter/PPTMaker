# PPTMaker 新电脑启动与使用

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
