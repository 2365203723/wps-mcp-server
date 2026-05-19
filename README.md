# wps-mcp-server

`wps-mcp-server` 是一个 Windows 上的 stdio MCP Server，用 Python 通过 WPS Office COM 接口控制本机 WPS。它可以让 Claude Code、Codex 以及其他 MCP Client 调用本地 WPS 来创建、读取、修改和打开 WPS Writer 文档、PDF 和 Excel/WPS 表格。

## 功能

当前提供 11 个 MCP tools：

- `create_document`：创建 WPS Writer/Word 文档并保存到桌面。
- `write_document`：向已有文档追加内容或覆盖内容。
- `read_document`：读取文档文本，最多返回 3000 字。
- `create_pdf`：通过 WPS Writer 创建并导出 PDF。
- `read_pdf`：读取文本型 PDF，优先使用 PyMuPDF/pypdf，最后回退到 WPS。
- `modify_pdf`：追加文本并另存为新 PDF，或用新内容覆盖生成新 PDF。
- `create_spreadsheet`：创建 WPS 表格/Excel 文件，写入表头和数据行。
- `read_spreadsheet`：读取指定工作表和范围，返回 JSON 文本。
- `write_spreadsheet`：从指定单元格开始写入二维数组。
- `modify_spreadsheet`：设置单元格值或清空范围。
- `open_file`：用 WPS 打开 `.docx/.doc/.wps/.xlsx/.xls/.et/.pdf` 文件。

## 系统要求

- Windows 桌面环境。
- 本机已安装 WPS Office。
- WPS COM 自动化可用：
  - Writer: `KWps.Application`
  - Spreadsheets: `KET.Application`
- Python 3.9 或更高版本。
- 必需依赖：`pywin32`。

此项目不支持 macOS、Linux、WSL 或无桌面/无 COM 的远程服务器环境。

## 安装

从 GitHub 安装：

```bash
pip install "git+https://github.com/2365203723/wps-mcp-server.git"
```

如果需要更稳定的 PDF 文本读取能力，安装可选 PDF 依赖：

```bash
pip install "wps-mcp-server[pdf] @ git+https://github.com/2365203723/wps-mcp-server.git"
```

验证命令是否可用：

```bash
mcp-wps
```

直接运行 `mcp-wps` 后没有普通 CLI 输出是正常的；它是 stdio MCP Server，会等待 MCP Client 通过 stdin 发送 JSON-RPC 消息。

## Claude Code 配置

推荐用 Claude Code 命令添加：

```bash
claude mcp add wps -- mcp-wps
```

如果希望添加到用户级配置，并且你的 Claude Code 版本支持 scope：

```bash
claude mcp add --scope user wps -- mcp-wps
```

如果 `mcp-wps` 不在 PATH，可改用 Python 模块方式：

```bash
claude mcp add wps -- python -m wps_server
```

也可以手动创建或修改项目下的 `.claude/mcp.json`：

```json
{
  "mcpServers": {
    "wps": {
      "type": "stdio",
      "command": "mcp-wps",
      "args": [],
      "env": {}
    }
  }
}
```

如果需要使用 Python 启动：

```json
{
  "mcpServers": {
    "wps": {
      "type": "stdio",
      "command": "python",
      "args": ["-m", "wps_server"],
      "env": {}
    }
  }
}
```

配置后重启 Claude Code，确认可以看到 `wps` MCP tools。

## Codex 配置

编辑 Codex 配置文件：

```text
~/.codex/config.toml
```

添加：

```toml
[mcp_servers.wps]
type = "stdio"
command = "mcp-wps"
args = []
startup_timeout_sec = 20
tool_timeout_sec = 120
```

如果 `mcp-wps` 不在 PATH，可使用 Python 启动：

```toml
[mcp_servers.wps]
type = "stdio"
command = "python"
args = ["-m", "wps_server"]
startup_timeout_sec = 20
tool_timeout_sec = 120
```

Windows 上如果 Codex 无法直接启动 Python，也可以使用：

```toml
[mcp_servers.wps]
type = "stdio"
command = "cmd"
args = ["/c", "python", "-m", "wps_server"]
startup_timeout_sec = 20
tool_timeout_sec = 120
```

配置后重启 Codex。

## 本地协议测试

可以用下面的方式测试 `tools/list`，不会打开 WPS：

```bash
python - <<'PY'
import json
import subprocess

p = subprocess.Popen(
    ["mcp-wps"],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True,
    encoding="utf-8",
)

p.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}}, ensure_ascii=False) + "\n")
p.stdin.flush()
print(p.stdout.readline())

p.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}}, ensure_ascii=False) + "\n")
p.stdin.flush()
print(p.stdout.readline())

p.terminate()
PY
```

## 注意事项

- 工具会通过 COM 拉起或连接本机 WPS，因此可能弹出 WPS 窗口。
- 创建类工具默认保存到当前用户桌面。
- PDF 修改不是专业 PDF 编辑；复杂版式不保证保持。
- 扫描版 PDF 通常无法直接提取文字，需要 OCR。
- 文件路径建议使用绝对路径。
- 写入/修改类工具会保存文件，请先确认目标文件路径。

## 常见问题

### `ModuleNotFoundError: No module named 'win32com'`

安装 `pywin32`：

```bash
pip install pywin32
```

### `Invalid class string` 或 WPS COM 对象创建失败

通常是 WPS Office 未安装、未正确注册 COM，或当前环境不是 Windows 桌面环境。请先手动打开一次 WPS，再重试；如果仍失败，尝试修复或重装 WPS Office。

### 直接运行 `mcp-wps` 看起来卡住

这是正常行为。`mcp-wps` 是 stdio MCP Server，会等待 MCP Client 输入 JSON-RPC 消息。

### PDF 读取不完整

PDF 文本提取取决于文件内容。建议安装可选 PDF 依赖：

```bash
pip install PyMuPDF pypdf
```

扫描版 PDF 需要 OCR，本项目暂不提供 OCR。

### 能否在 Linux、macOS 或 WSL 使用？

不能。该项目依赖 Windows COM 和本机 WPS Office。

## 开发验证

```bash
python -m py_compile wps_server.py
python -m pip install .
python -c "import wps_server; print(wps_server.__version__); print(callable(wps_server.main))"
python -m build
```

## License

MIT
