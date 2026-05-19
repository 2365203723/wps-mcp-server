#!/usr/bin/env python3
"""
WPS MCP Server - 通过 COM 接口操控 WPS Office
协议：手动实现 MCP JSON-RPC over stdio（兼容 Python 3.9）
"""
import sys
import json
import os

__version__ = "0.1.0"

# ──────────────────────────────────────────────
#  WPS 操作封装
# ──────────────────────────────────────────────

def get_wps_word():
    """获取或启动 WPS Writer (Word) COM 对象"""
    import win32com.client
    try:
        return win32com.client.GetActiveObject("KWps.Application")
    except Exception:
        return win32com.client.Dispatch("KWps.Application")

def get_wps_excel():
    """获取或启动 WPS Spreadsheets (Excel) COM 对象"""
    import win32com.client
    try:
        return win32com.client.GetActiveObject("KET.Application")
    except Exception:
        return win32com.client.Dispatch("KET.Application")

def export_document_to_pdf(doc, save_path: str):
    try:
        doc.ExportAsFixedFormat(save_path, 17)
    except Exception:
        doc.SaveAs(save_path, FileFormat=17)

def get_spreadsheet_sheet(wb, sheet: str):
    if sheet:
        return wb.Worksheets(sheet)
    return wb.Sheets(1)

def json_safe_value(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)

def wps_create_document(title: str, content: str) -> str:
    app = get_wps_word()
    app.Visible = True
    doc = app.Documents.Add()
    if title:
        rng = doc.Range(0, 0)
        rng.InsertAfter(title + "\n")
        rng.Style = doc.Styles("标题 1")
    if content:
        rng2 = doc.Range()
        rng2.Collapse(0)  # 移到末尾
        rng2.InsertAfter(content)
    save_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{title or 'document'}.docx")
    doc.SaveAs(save_path)
    return f"文档已创建并保存到：{save_path}"

def wps_write_document(file_path: str, content: str, mode: str = "append") -> str:
    import win32com.client
    app = get_wps_word()
    app.Visible = True
    doc = app.Documents.Open(file_path)
    rng = doc.Range()
    if mode == "overwrite":
        rng.Text = ""
        rng = doc.Range()
    rng.Collapse(0)
    rng.InsertAfter(content)
    doc.Save()
    return f"内容已写入文档：{file_path}"

def wps_read_document(file_path: str) -> str:
    app = get_wps_word()
    app.Visible = True
    doc = app.Documents.Open(file_path)
    text = doc.Content.Text
    return text[:3000] + ("..." if len(text) > 3000 else "")

def wps_create_pdf(title: str, content: str) -> str:
    app = get_wps_word()
    app.Visible = True
    doc = app.Documents.Add()
    if title:
        rng = doc.Range(0, 0)
        rng.InsertAfter(title + "\n")
        rng.Style = doc.Styles("标题 1")
    if content:
        rng2 = doc.Range()
        rng2.Collapse(0)
        rng2.InsertAfter(content)
    save_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{title or 'document'}.pdf")
    export_document_to_pdf(doc, save_path)
    return f"PDF 已创建并保存到：{save_path}"

def wps_read_pdf(file_path: str) -> str:
    text = ""
    try:
        import fitz
        with fitz.open(file_path) as pdf:
            text = "\n".join(page.get_text() for page in pdf)
    except Exception:
        try:
            try:
                from pypdf import PdfReader
            except Exception:
                from PyPDF2 import PdfReader
            reader = PdfReader(file_path)
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception:
            app = get_wps_word()
            app.Visible = True
            doc = app.Documents.Open(file_path)
            text = doc.Content.Text
    return text[:3000] + ("..." if len(text) > 3000 else "")

def wps_modify_pdf(file_path: str, content: str, output_path: str = "", mode: str = "append") -> str:
    save_path = output_path or os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        f"{os.path.splitext(os.path.basename(file_path))[0] or 'document'}_{mode}.pdf"
    )
    if os.path.abspath(save_path).lower() == os.path.abspath(file_path).lower():
        raise ValueError("输出路径不能与原 PDF 相同")

    app = get_wps_word()
    app.Visible = True
    if mode == "overwrite":
        doc = app.Documents.Add()
        if content:
            rng = doc.Range(0, 0)
            rng.InsertAfter(content)
    elif mode == "append":
        doc = app.Documents.Open(file_path)
        rng = doc.Range()
        rng.Collapse(0)
        rng.InsertAfter(content)
    else:
        raise ValueError(f"不支持的 PDF 修改模式：{mode}")

    export_document_to_pdf(doc, save_path)
    return f"PDF 已修改并另存为：{save_path}"

def wps_create_spreadsheet(file_name: str, headers: list, rows: list) -> str:
    app = get_wps_excel()
    app.Visible = True
    wb = app.Workbooks.Add()
    ws = wb.Sheets(1)
    # 写表头
    for col, header in enumerate(headers, start=1):
        ws.Cells(1, col).Value = header
    # 写数据行
    for row_idx, row in enumerate(rows, start=2):
        for col_idx, val in enumerate(row, start=1):
            ws.Cells(row_idx, col_idx).Value = val
    save_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{file_name}.xlsx")
    wb.SaveAs(save_path)
    return f"表格已创建并保存到：{save_path}"

def wps_read_spreadsheet(file_path: str, sheet: str = "", range_address: str = "", max_rows: int = 100, max_cols: int = 20) -> str:
    app = get_wps_excel()
    app.Visible = True
    wb = app.Workbooks.Open(file_path)
    ws = get_spreadsheet_sheet(wb, sheet)
    rng = ws.Range(range_address) if range_address else ws.UsedRange
    row_count = min(int(rng.Rows.Count), int(max_rows or 100))
    col_count = min(int(rng.Columns.Count), int(max_cols or 20))
    rows = []
    for row_idx in range(1, row_count + 1):
        row = []
        for col_idx in range(1, col_count + 1):
            row.append(json_safe_value(rng.Cells(row_idx, col_idx).Value))
        rows.append(row)
    return json.dumps({"sheet": ws.Name, "range": range_address or rng.Address, "rows": rows}, ensure_ascii=False)

def wps_write_spreadsheet(file_path: str, data: list, sheet: str = "", start_cell: str = "A1", save: bool = True) -> str:
    app = get_wps_excel()
    app.Visible = True
    wb = app.Workbooks.Open(file_path)
    ws = get_spreadsheet_sheet(wb, sheet)
    start = ws.Range(start_cell or "A1")
    for row_idx, row in enumerate(data):
        for col_idx, value in enumerate(row):
            ws.Cells(start.Row + row_idx, start.Column + col_idx).Value = value
    if save:
        wb.Save()
    return f"表格内容已写入：{file_path}"

def wps_modify_spreadsheet(file_path: str, operations: list, sheet: str = "", save: bool = True) -> str:
    app = get_wps_excel()
    app.Visible = True
    wb = app.Workbooks.Open(file_path)
    ws = get_spreadsheet_sheet(wb, sheet)
    for op in operations:
        action = op.get("action")
        if action == "set":
            ws.Range(op["cell"]).Value = op.get("value", "")
        elif action == "clear":
            ws.Range(op["range"]).ClearContents()
        else:
            raise ValueError(f"不支持的表格操作：{action}")
    if save:
        wb.Save()
    return f"已完成 {len(operations)} 个表格修改操作：{file_path}"

def wps_open_file(file_path: str) -> str:
    import win32com.client
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".docx", ".doc", ".wps"):
        app = get_wps_word()
        app.Visible = True
        app.Documents.Open(file_path)
    elif ext in (".xlsx", ".xls", ".et"):
        app = get_wps_excel()
        app.Visible = True
        app.Workbooks.Open(file_path)
    elif ext == ".pdf":
        app = get_wps_word()
        app.Visible = True
        app.Documents.Open(file_path)
    else:
        return f"不支持的文件格式：{ext}"
    return f"已打开文件：{file_path}"

# ──────────────────────────────────────────────
#  工具定义（MCP Tools Schema）
# ──────────────────────────────────────────────

TOOLS = [
    {
        "name": "create_document",
        "description": "创建一个新的 WPS Word 文档，自动保存到桌面",
        "inputSchema": {
            "type": "object",
            "properties": {
                "title":   {"type": "string", "description": "文档标题（同时作为文件名）"},
                "content": {"type": "string", "description": "文档正文内容"}
            },
            "required": ["title"]
        }
    },
    {
        "name": "write_document",
        "description": "向已有的 WPS Word 文档写入内容",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文档的完整路径"},
                "content":   {"type": "string", "description": "要写入的文本内容"},
                "mode":      {"type": "string", "enum": ["append", "overwrite"], "description": "写入模式：追加或覆盖，默认追加"}
            },
            "required": ["file_path", "content"]
        }
    },
    {
        "name": "read_document",
        "description": "读取 WPS Word 文档的文字内容（最多返回 3000 字）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文档的完整路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "create_pdf",
        "description": "创建一个 PDF 文件，使用 WPS Writer 生成内容并导出到桌面",
        "inputSchema": {
            "type": "object",
            "properties": {
                "title":   {"type": "string", "description": "PDF 标题（同时作为文件名）"},
                "content": {"type": "string", "description": "PDF 正文内容"}
            },
            "required": ["title"]
        }
    },
    {
        "name": "read_pdf",
        "description": "尝试读取文本型 PDF 的文字内容（最多返回 3000 字，依赖本机 WPS 对 PDF 的支持）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "PDF 文件完整路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "modify_pdf",
        "description": "以最小可行方式修改 PDF：追加文本并另存为新 PDF，或用新内容覆盖生成新 PDF；不保证复杂 PDF 原版式保持",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path":   {"type": "string", "description": "原 PDF 文件完整路径"},
                "content":     {"type": "string", "description": "要写入 PDF 的文本内容"},
                "output_path": {"type": "string", "description": "输出 PDF 完整路径；为空则自动保存到桌面"},
                "mode":        {"type": "string", "enum": ["append", "overwrite"], "description": "修改模式：append 尝试追加后另存；overwrite 用新内容生成新 PDF"}
            },
            "required": ["file_path", "content"]
        }
    },
    {
        "name": "create_spreadsheet",
        "description": "创建一个新的 WPS 表格（Excel），写入表头和数据行，自动保存到桌面",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_name": {"type": "string", "description": "文件名（不含扩展名）"},
                "headers":   {"type": "array",  "items": {"type": "string"}, "description": "列标题列表"},
                "rows":      {"type": "array",  "items": {"type": "array"},  "description": "数据行列表，每行是一个数组"}
            },
            "required": ["file_name", "headers", "rows"]
        }
    },
    {
        "name": "read_spreadsheet",
        "description": "读取 WPS 表格/Excel 的指定工作表和范围，返回 JSON 文本",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path":     {"type": "string", "description": "表格文件完整路径"},
                "sheet":         {"type": "string", "description": "工作表名称；为空则读取第一个工作表"},
                "range_address": {"type": "string", "description": "读取范围，例如 A1:D10；为空则读取 UsedRange"},
                "max_rows":      {"type": "integer", "description": "最多读取行数，默认 100"},
                "max_cols":      {"type": "integer", "description": "最多读取列数，默认 20"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "write_spreadsheet",
        "description": "向 WPS 表格/Excel 指定位置写入二维数组数据",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path":  {"type": "string", "description": "表格文件完整路径"},
                "data":       {"type": "array", "items": {"type": "array"}, "description": "要写入的二维数组"},
                "sheet":      {"type": "string", "description": "工作表名称；为空则使用第一个工作表"},
                "start_cell": {"type": "string", "description": "起始单元格，例如 A1，默认 A1"},
                "save":       {"type": "boolean", "description": "写入后是否保存，默认 true"}
            },
            "required": ["file_path", "data"]
        }
    },
    {
        "name": "modify_spreadsheet",
        "description": "修改 WPS 表格/Excel：支持设置单元格值和清空范围",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path":  {"type": "string", "description": "表格文件完整路径"},
                "sheet":      {"type": "string", "description": "工作表名称；为空则使用第一个工作表"},
                "operations": {
                    "type": "array",
                    "description": "操作列表，支持 set 和 clear",
                    "items": {
                        "type": "object",
                        "properties": {
                            "action": {"type": "string", "enum": ["set", "clear"], "description": "操作类型"},
                            "cell":   {"type": "string", "description": "set 操作用的单元格地址，例如 B2"},
                            "range":  {"type": "string", "description": "clear 操作用的范围地址，例如 A1:D10"},
                            "value":  {"description": "set 操作用的新值"}
                        },
                        "required": ["action"]
                    }
                },
                "save": {"type": "boolean", "description": "修改后是否保存，默认 true"}
            },
            "required": ["file_path", "operations"]
        }
    },
    {
        "name": "open_file",
        "description": "用 WPS 打开指定路径的文件（支持 .docx/.doc/.wps/.xlsx/.xls/.et/.pdf）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "要打开的文件完整路径"}
            },
            "required": ["file_path"]
        }
    }
]

# ──────────────────────────────────────────────
#  工具调度
# ──────────────────────────────────────────────

def dispatch_tool(name: str, args: dict) -> str:
    if name == "create_document":
        return wps_create_document(args.get("title", ""), args.get("content", ""))
    if name == "write_document":
        return wps_write_document(args["file_path"], args["content"], args.get("mode", "append"))
    if name == "read_document":
        return wps_read_document(args["file_path"])
    if name == "create_pdf":
        return wps_create_pdf(args.get("title", ""), args.get("content", ""))
    if name == "read_pdf":
        return wps_read_pdf(args["file_path"])
    if name == "modify_pdf":
        return wps_modify_pdf(args["file_path"], args["content"], args.get("output_path", ""), args.get("mode", "append"))
    if name == "create_spreadsheet":
        return wps_create_spreadsheet(args["file_name"], args["headers"], args["rows"])
    if name == "read_spreadsheet":
        return wps_read_spreadsheet(args["file_path"], args.get("sheet", ""), args.get("range_address", ""), args.get("max_rows", 100), args.get("max_cols", 20))
    if name == "write_spreadsheet":
        return wps_write_spreadsheet(args["file_path"], args["data"], args.get("sheet", ""), args.get("start_cell", "A1"), args.get("save", True))
    if name == "modify_spreadsheet":
        return wps_modify_spreadsheet(args["file_path"], args["operations"], args.get("sheet", ""), args.get("save", True))
    if name == "open_file":
        return wps_open_file(args["file_path"])
    return f"未知工具：{name}"

# ──────────────────────────────────────────────
#  MCP JSON-RPC 协议处理
# ──────────────────────────────────────────────

def send(obj: dict):
    line = json.dumps(obj, ensure_ascii=False)
    sys.stdout.write(line + "\n")
    sys.stdout.flush()

def handle(req: dict):
    method = req.get("method", "")
    req_id = req.get("id")

    # 初始化握手
    if method == "initialize":
        return {
            "jsonrpc": "2.0", "id": req_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "serverInfo": {"name": "wps-mcp-server", "version": __version__},
                "capabilities": {"tools": {}}
            }
        }

    # 列出工具
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": TOOLS}}

    # 调用工具
    if method == "tools/call":
        params = req.get("params", {})
        tool_name = params.get("name", "")
        arguments = params.get("arguments", {})
        try:
            result_text = dispatch_tool(tool_name, arguments)
            return {
                "jsonrpc": "2.0", "id": req_id,
                "result": {"content": [{"type": "text", "text": result_text}]}
            }
        except Exception as e:
            return {
                "jsonrpc": "2.0", "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": f"执行失败：{e}"}],
                    "isError": True
                }
            }

    # notifications 不需要回复
    if req_id is None:
        return None

    return {
        "jsonrpc": "2.0", "id": req_id,
        "error": {"code": -32601, "message": f"未知方法：{method}"}
    }

# ──────────────────────────────────────────────
#  主循环
# ──────────────────────────────────────────────

def main():
    # 禁用 Python 的标准输出缓冲，确保每行立即发出
    sys.stdin.reconfigure(encoding="utf-8-sig")
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8")

    for raw_line in sys.stdin:
        raw_line = raw_line.strip()
        if not raw_line:
            continue
        try:
            req = json.loads(raw_line)
        except json.JSONDecodeError as e:
            send({"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": str(e)}})
            continue

        resp = handle(req)
        if resp is not None:
            send(resp)

if __name__ == "__main__":
    main()
