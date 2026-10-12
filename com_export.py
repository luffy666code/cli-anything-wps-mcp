"""Windows-only WPS COM export. No taskkill, no sys.path hacks, no Office install."""
from __future__ import annotations
import os
from pathlib import Path
import sys
from typing import Any


def _rgb(value: str) -> int:
    value = str(value or '#1F2937').strip().lstrip('#')
    if len(value) != 6 or any(c not in '0123456789abcdefABCDEF' for c in value):
        raise ValueError('颜色必须是 #RRGGBB')
    return int(value[0:2], 16) | int(value[2:4], 16) << 8 | int(value[4:6], 16) << 16


def _font(run: Any, size: float | None = None, bold: bool | None = None,
          color: str | None = None, font_name: str | None = None) -> None:
    f = run.Font
    if size is not None: f.Size = float(size)
    if bold is not None: f.Bold = bool(bold)
    if color: f.Color.RGB = _rgb(color)
    if font_name: f.Name = font_name


def _writer_add(doc: Any, item: dict) -> None:
    kind = item['kind']
    rng = doc.Range()
    rng.Collapse(0)  # wdCollapseEnd
    if kind in ('paragraph', 'heading'):
        rng.InsertAfter(item.get('text', ''))
        # InsertAfter expands Range to cover the inserted text; style before
        # InsertParagraphAfter so formatting targets the text, not a blank line.
        try:
            _font(rng, item.get('font_size', 20 if kind == 'heading' else 12),
                  item.get('bold', kind == 'heading'), item.get('color'), item.get('font_name'))
            if kind == 'heading':
                rng.ParagraphFormat.OutlineLevel = min(int(item.get('level', 1)), 9)
            if item.get('alignment'):
                rng.ParagraphFormat.Alignment = {'left': 0, 'center': 1, 'right': 2, 'justify': 3}[item['alignment']]
        except Exception:
            # Some WPS editions do not support all optional style properties.
            pass
        rng.InsertParagraphAfter()
    elif kind == 'table':
        data = item.get('data') or []
        rows, cols = int(item['rows']), int(item['cols'])
        table = doc.Tables.Add(rng, rows, cols)
        for r, row in enumerate(data[:rows]):
            for c, val in enumerate(row[:cols]):
                table.Cell(r + 1, c + 1).Range.Text = str(val)
        doc.Range().InsertParagraphAfter()
    elif kind == 'image':
        p = Path(item['path']).expanduser().resolve()
        if not p.is_file(): raise FileNotFoundError('插入图片不存在: ' + str(p))
        pic = doc.InlineShapes.AddPicture(str(p), False, True, rng)
        if item.get('width'): pic.Width = float(item['width'])
        if item.get('height'): pic.Height = float(item['height'])
    elif kind == 'page_break':
        rng.InsertBreak(7)
    else:
        raise ValueError('未知 Word 内容类型: ' + kind)


def _calc_fill(book: Any, obj: dict) -> None:
    for index, sheet in enumerate(obj['sheets']):
        if index == 0:
            ws = book.Worksheets(1)
        else:
            ws = book.Worksheets.Add()
        ws.Name = sheet['name']
        for address, cell in sheet.get('cells', {}).items():
            rng = ws.Range(address)
            if cell.get('formula'):
                rng.Formula = cell['formula']
            else:
                rng.Value = cell.get('value')
        for item in sheet.get('merges', []):
            ws.Range(item['range']).Merge()


def _add_text(slide: Any, text: str, x: float, y: float, w: float, h: float,
              size: float = 18, color: str = '#1F2937', bold: bool = False,
              font_name: str = 'Microsoft YaHei') -> Any:
    box = slide.Shapes.AddTextbox(1, x, y, w, h)  # msoTextOrientationHorizontal
    box.TextFrame.TextRange.Text = str(text)
    box.TextFrame.TextRange.Font.Size = size
    box.TextFrame.TextRange.Font.Bold = bold
    box.TextFrame.TextRange.Font.Name = font_name
    box.TextFrame.TextRange.Font.Color.RGB = _rgb(color)
    return box


def _impress_fill(pres: Any, obj: dict) -> None:
    width = float(obj.get('slide_width', 960))
    height = float(obj.get('slide_height', 540))
    try:
        pres.PageSetup.SlideWidth = width
        pres.PageSetup.SlideHeight = height
    except Exception:
        pass
    # Always add slides explicitly. WPS Presentations.Add may contain zero slides.
    for i, item in enumerate(obj.get('slides', [])):
        slide = pres.Slides.Add(i + 1, 12)  # ppLayoutBlank
        if item.get('background'):
            # WPS KWPP COM on some Windows editions rejects assigning
            # Slide.FollowMasterGraphics ("Property Add.FollowMasterGraphics can not be set").
            # A full-slide shape inserted before every foreground element gives
            # the same visible background without touching that COM property.
            background = slide.Shapes.AddShape(1, 0, 0, width, height)  # msoShapeRectangle
            background.Fill.ForeColor.RGB = _rgb(item['background'])
            background.Fill.Solid()
            try:
                background.Line.Visible = 0  # msoFalse, optional in WPS
            except Exception:
                pass
        if item.get('title'):
            _add_text(slide, item['title'], 44, 20, width - 88, 65, size=30, bold=True)
        if item.get('content'):
            _add_text(slide, item['content'], 60, 118, width - 120, height - 155, size=19)
        for el in item.get('elements', []):
            k = el['kind']
            x, y, w, h = (float(el[n]) for n in ('x', 'y', 'width', 'height'))
            if k == 'text':
                _add_text(slide, el['text'], x, y, w, h,
                          float(el.get('font_size', 18)), el.get('color', '#1F2937'),
                          bool(el.get('bold', False)))
            elif k == 'shape':
                t = {'rect': 1, 'roundrect': 5, 'ellipse': 9}.get(el.get('shape', 'rect'), 1)
                shape = slide.Shapes.AddShape(t, x, y, w, h)
                shape.Fill.ForeColor.RGB = _rgb(el.get('fill', '#DCE8F8'))
                shape.Fill.Solid()
                if el.get('text'):
                    shape.TextFrame.TextRange.Text = el['text']
            elif k == 'image':
                p = Path(el['path']).expanduser().resolve()
                if not p.is_file(): raise FileNotFoundError('PPT 图片不存在: ' + str(p))
                slide.Shapes.AddPicture(str(p), False, True, x, y, w, h)
            else: raise ValueError('未知 PPT 元素: ' + k)


def _save_office(doc: Any, doc_type: str, ext: str, out: str) -> None:
    if ext == 'pdf':
        if doc_type == 'writer':
            doc.ExportAsFixedFormat(out, 17)
        elif doc_type == 'calc':
            doc.ExportAsFixedFormat(0, out)
        else:
            doc.SaveAs(out, 32)
    elif doc_type == 'writer':
        try: doc.SaveAs2(out, 16)
        except Exception: doc.SaveAs(out)
    elif doc_type == 'calc':
        doc.SaveAs(out, 51)
    else:
        # PPTX constant ppSaveAsOpenXMLPresentation=24 in MS COM model.
        try: doc.SaveAs(out, 24)
        except Exception: doc.SaveAs(out)


def export_project(obj: dict, output_path: str, fmt: str, overwrite: bool = False) -> dict:
    if sys.platform != 'win32':
        raise RuntimeError('本 MCP 仅支持 Windows x86_64 的本机 WPS COM 导出。')
    kind = obj['type']
    ext = fmt.lower().lstrip('.')
    expected = {'writer': {'docx', 'pdf'}, 'calc': {'xlsx', 'pdf'}, 'impress': {'pptx', 'pdf'}}[kind]
    if ext not in expected: raise ValueError(f'{kind} 仅支持输出 {sorted(expected)}')
    path = Path(output_path).expanduser().resolve()
    if path.suffix.lower() != '.' + ext:
        raise ValueError(f'文件扩展名必须为 .{ext}')
    if path.exists() and not overwrite:
        raise FileExistsError('文件已存在；如确认覆盖请设置 overwrite=true：' + str(path))
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        import pythoncom
        import win32com.client
    except ImportError as ex:
        raise RuntimeError('缺少 pywin32。请用运行 MCP 的同一个 Python 执行 python -m pip install pywin32') from ex
    pythoncom.CoInitialize()
    app, doc = None, None
    dedicated = False
    document_closed = False
    progid = {'writer': 'KWPS.Application', 'calc': 'KET.Application', 'impress': 'KWPP.Application'}[kind]
    try:
        try:
            # Avoid changing/closing the user's existing WPS process when possible.
            app = win32com.client.DispatchEx(progid)
            dedicated = True
        except Exception:
            # Certain WPS releases do not expose a separate COM instance.
            # Fallback can attach to a running user instance; never Quit() in that case.
            app = win32com.client.Dispatch(progid)
        if dedicated:
            try: app.Visible = False
            except Exception: pass
            if kind == 'impress':
                # PowerPoint PpAlertLevel: 1 = ppAlertsNone. In particular,
                # do not allow an unsaved automation-only presentation to
                # block the stdio MCP request on a modal Save confirmation.
                try: app.DisplayAlerts = 1
                except Exception: pass
        if kind == 'writer':
            doc = app.Documents.Add()
            for item in obj.get('content', []): _writer_add(doc, item)
        elif kind == 'calc':
            doc = app.Workbooks.Add()
            _calc_fill(doc, obj)
        else:
            doc = app.Presentations.Add()
            _impress_fill(doc, obj)
        _save_office(doc, kind, ext, str(path))
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError('WPS 调用完成，但未生成有效的输出文件，请检查 WPS 版本与 COM 导出能力')
        return {'output': str(path), 'file_size': path.stat().st_size, 'format': ext,
                'type': kind, 'method': 'WPS COM', 'dedicated_instance': dedicated}
    except Exception as ex:
        raise RuntimeError(f'WPS COM 导出失败 [{progid}]：{ex}。请检查同一用户权限、WPS COM 注册及桌面会话。') from ex
    finally:
        if doc is not None:
            try:
                if kind == 'impress':
                    # Presentation.Close() does NOT accept the SaveChanges
                    # argument used by Word Document.Close().
                    # Explicitly mark *our newly created* presentation as
                    # saved so WPS won't ask the user to save it on closing,
                    # even if COM export failed halfway through.
                    try:
                        doc.Saved = True
                    except Exception:
                        pass  # Some WPS builds expose Saved as read-only.
                    doc.Close()
                else:
                    doc.Close(False)
                document_closed = True
            except Exception as ex:
                print(f'[WPS MCP] 关闭自动创建的 {kind} 文档失败: {ex}', file=sys.stderr)
        if app is not None and dedicated:
            if doc is None or document_closed:
                try:
                    app.Quit()
                except Exception as ex:
                    print(f'[WPS MCP] 退出自动创建的 WPS 实例失败: {ex}', file=sys.stderr)
            else:
                # Calling Quit while an unsaved presentation is still open
                # would summon a modal Save/Don't save dialog and hang MCP.
                print('[WPS MCP] 跳过 Quit：自动创建的文档未能关闭，以避免保存弹窗阻塞。', file=sys.stderr)
        pythoncom.CoUninitialize()
