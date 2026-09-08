#!/usr/bin/env python3
"""
大表拆分工具 v1.3 — 按指定列值将大表拆分为多个小表。
每个 Sheet 独立选择拆分依据列，输出文件名=列值，Sheet名=原始Sheet名。
零第三方依赖，仅使用 Python 标准库。独立运行，不依赖合并脚本。
"""

import os, sys, zipfile, io, time, traceback
from xml.etree import ElementTree as ET
import tkinter as tk
from tkinter import messagebox
from collections import OrderedDict

# ── xlsx 命名空间 ──────────────────────────────────────────────
NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ET.register_namespace('', NS_S)
ET.register_namespace('r', NS_R)


# ══════════════════════════════════════════════════════════════════
#  xlsx 工具函数
# ══════════════════════════════════════════════════════════════════

def _tostring_xml(element):
    raw = ET.tostring(element, encoding='unicode')
    if raw.startswith("<?xml version='1.0' encoding='utf-8'?>"):
        raw = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + raw[len("<?xml version='1.0' encoding='utf-8'?>"):]
    elif raw.startswith("<?xml version='1.0' encoding='UTF-8'?>"):
        raw = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + raw[len("<?xml version='1.0' encoding='UTF-8'?>"):]
    return raw.encode('UTF-8')


def _parse_cell_ref(ref):
    """解析单元格引用 'A1' → (col_index, row_num)"""
    import re
    m = re.match(r'^([A-Z]+)(\d+)$', ref)
    if not m:
        return 0, 0
    col_str = m.group(1)
    row_num = int(m.group(2))
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return col - 1, row_num


def _col_letter(col):
    """0-based column index → Excel column letter"""
    result = ''
    col += 1
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result


# ══════════════════════════════════════════════════════════════════
#  文件读取
# ══════════════════════════════════════════════════════════════════

def read_csv(filepath):
    import csv
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = [row for row in reader if row]
    if not rows:
        return {}
    return {'Sheet1': rows}


def read_xlsx(filepath):
    """读取 xlsx 文件 → {sheet_name: [header_row, data_row1, ...]}"""
    z = zipfile.ZipFile(filepath, 'r')
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.parse(z.open('xl/sharedStrings.xml')).getroot()
        for si in root.findall('.//{%s}si' % NS_S):
            texts = []
            for t in si.iter('{%s}t' % NS_S):
                if t.text:
                    texts.append(t.text)
            shared_strings.append(''.join(texts))
    wb_root = ET.parse(z.open('xl/workbook.xml')).getroot()
    sheet_elems = wb_root.findall('.//{%s}sheet' % NS_S)
    sheet_names = [s.get('name', '') for s in sheet_elems]
    result = {}
    for idx, name in enumerate(sheet_names):
        sheet_file = 'xl/worksheets/sheet%d.xml' % (idx + 1)
        if sheet_file not in z.namelist():
            continue
        ws_root = ET.parse(z.open(sheet_file)).getroot()
        rows = ws_root.findall('.//{%s}row' % NS_S)
        sheet_data = []
        for row_el in rows:
            cell_positions = {}
            for c in row_el.findall('{%s}c' % NS_S):
                cell_type = c.get('t', '')
                ref = c.get('r', '')
                v_el = c.find('{%s}v' % NS_S)
                val = v_el.text if v_el is not None else ''
                if cell_type == 's':
                    # 共享字符串索引
                    if val and val.isdigit():
                        idx_s = int(val)
                        if idx_s < len(shared_strings):
                            val = shared_strings[idx_s]
                elif cell_type == 'inlineStr':
                    # 内联字符串：<is><t>文本</t></is>（无 <v>）
                    is_el = c.find('{%s}is' % NS_S)
                    if is_el is not None:
                        val = ''.join(
                            (t.text or '') for t in is_el.iter('{%s}t' % NS_S)
                        )
                elif cell_type == 'str':
                    # 公式计算结果为字符串：<v> 里直接就是文本
                    pass
                elif cell_type == 'b':
                    # 布尔：1=TRUE 0=FALSE
                    val = 'TRUE' if val == '1' else ('FALSE' if val == '0' else val)
                ci, _ = _parse_cell_ref(ref)
                cell_positions[ci] = val
            if not cell_positions:
                continue
            max_col = max(cell_positions.keys())
            row_values = [cell_positions.get(ci, '') for ci in range(max_col + 1)]
            if any(v for v in row_values):
                sheet_data.append(row_values)
        result[name] = sheet_data
    z.close()
    return result


def read_table(filepath):
    if filepath.lower().endswith('.csv'):
        return read_csv(filepath)
    return read_xlsx(filepath)


# ══════════════════════════════════════════════════════════════════
#  xlsx 写入
# ══════════════════════════════════════════════════════════════════

def write_xlsx(filepath, sheets_data):
    shared_strings = []
    ss_map = {}

    def _ss_idx(text):
        s = str(text) if text is not None else ''
        if s not in ss_map:
            ss_map[s] = len(shared_strings)
            shared_strings.append(s)
        return ss_map[s]

    for _, rows in sheets_data.items():
        for row in rows:
            for cell in row:
                if isinstance(cell, str):
                    _ss_idx(cell)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml
        ct_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
        ]
        for i in range(len(sheets_data)):
            ct_parts.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1))
        ct_parts.append('</Types>')
        zf.writestr('[Content_Types].xml', '\n'.join(ct_parts))

        # _rels/.rels
        zf.writestr('_rels/.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
            '</Relationships>')

        # xl/_rels/workbook.xml.rels
        wb_rels_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
        ]
        for i in range(len(sheets_data)):
            wb_rels_parts.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 2, i + 1))
        wb_rels_parts.append('</Relationships>')
        zf.writestr('xl/_rels/workbook.xml.rels', '\n'.join(wb_rels_parts))

        # xl/workbook.xml
        wb = ET.Element('{%s}workbook' % NS_S)
        sheets_el = ET.SubElement(wb, '{%s}sheets' % NS_S)
        for i, name in enumerate(list(sheets_data.keys())):
            ET.SubElement(sheets_el, '{%s}sheet' % NS_S,
                          name=name, sheetId=str(i + 1),
                          **{'{%s}id' % NS_R: 'rId%d' % (i + 2)})
        zf.writestr('xl/workbook.xml', _tostring_xml(wb))

        # xl/sharedStrings.xml
        sst = ET.Element('{%s}sst' % NS_S, count=str(len(shared_strings)), uniqueCount=str(len(shared_strings)))
        for s in shared_strings:
            si = ET.SubElement(sst, '{%s}si' % NS_S)
            t = ET.SubElement(si, '{%s}t' % NS_S)
            t.text = s
        zf.writestr('xl/sharedStrings.xml', _tostring_xml(sst))

        # xl/worksheets/sheetN.xml
        for idx, (sn, rows) in enumerate(sheets_data.items()):
            ws = ET.Element('{%s}worksheet' % NS_S)
            last_col = max(len(r) for r in rows) - 1 if rows else 0
            last_row = len(rows)
            ET.SubElement(ws, '{%s}dimension' % NS_S, ref='A1:%s%d' % (_col_letter(last_col), last_row))
            sv = ET.SubElement(ws, '{%s}sheetViews' % NS_S)
            ET.SubElement(sv, '{%s}sheetView' % NS_S, workbookViewId='0')
            ET.SubElement(ws, '{%s}sheetFormatPr' % NS_S, defaultRowHeight='15')
            sd = ET.SubElement(ws, '{%s}sheetData' % NS_S)
            for row_idx, row in enumerate(rows, 1):
                r_el = ET.SubElement(sd, '{%s}row' % NS_S, r=str(row_idx))
                for col_idx, val in enumerate(row):
                    ref = '%s%d' % (_col_letter(col_idx), row_idx)
                    c = ET.SubElement(r_el, '{%s}c' % NS_S, r=ref)
                    if val is None:
                        val = ''
                    if isinstance(val, (int, float)):
                        v = ET.SubElement(c, '{%s}v' % NS_S)
                        v.text = str(val)
                    elif isinstance(val, bool):
                        c.set('t', 'b')
                        v = ET.SubElement(c, '{%s}v' % NS_S)
                        v.text = '1' if val else '0'
                    else:
                        c.set('t', 's')
                        v = ET.SubElement(c, '{%s}v' % NS_S)
                        v.text = str(ss_map[str(val)])
            zf.writestr('xl/worksheets/sheet%d.xml' % (idx + 1), _tostring_xml(ws))

    with open(filepath, 'wb') as f:
        f.write(buf.getvalue())


# ══════════════════════════════════════════════════════════════════
#  tkinter 工具
# ══════════════════════════════════════════════════════════════════

def _dialog_font():
    import platform
    if platform.system() == 'Windows':
        return ('Microsoft YaHei', 10)
    return ('Sans', 10)


def _show_dialog_root():
    try:
        _root.update()
        _root.lift()
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════
#  日志
# ══════════════════════════════════════════════════════════════════

_log_lines = []

def log(msg):
    ts = time.strftime('%H:%M:%S')
    line = f'[{ts}] {msg}'
    print(line, flush=True)
    _log_lines.append(line)

def write_log(filepath):
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write('\n'.join(_log_lines) + '\n')
    log(f'日志已保存: {filepath}')


# ══════════════════════════════════════════════════════════════════
#  文件名清理
# ══════════════════════════════════════════════════════════════════

def sanitize_filename(name):
    name = str(name).strip()
    for ch in '/\\:*?"<>|[]':
        name = name.replace(ch, '_')
    name = name.strip('. ')
    if len(name) > 31:
        name = name[:31]
    if not name:
        name = '_空值_'
    return name


# ══════════════════════════════════════════════════════════════════
#  文件扫描
# ══════════════════════════════════════════════════════════════════

def _scan_work_dir(work_dir):
    files = []
    for f in os.listdir(work_dir):
        if f.startswith('~$'):
            continue
        low = f.lower()
        if low.endswith('.xlsx') or low.endswith('.csv'):
            files.append(f)
    return sorted(files)


# ══════════════════════════════════════════════════════════════════
#  UI 对话框
# ══════════════════════════════════════════════════════════════════

_root = None


def _select_file_dialog(files):
    """单选大表文件对话框"""
    global _root
    _show_dialog_root()
    result = {'value': None}

    dlg = tk.Toplevel(_root)
    dlg.title('选择大表文件')
    dlg.resizable(False, True)
    dlg.transient(_root)
    dlg.grab_set()
    try:
        dlg.attributes('-topmost', True)
    except Exception:
        pass
    try:
        rx, ry = _root.winfo_x(), _root.winfo_y()
    except Exception:
        rx, ry = 100, 100
    dlg.geometry('+%d+%d' % (rx + 80, ry + 80))

    tk.Label(dlg, text='选择大表文件', font=(_dialog_font()[0], 12, 'bold')
             ).pack(padx=20, pady=15)
    tk.Label(dlg, text='选择需要拆分的大表文件。\n大表的第一行将被识别为表头。',
             font=(_dialog_font()[0], 9), fg='#666'
             ).pack(padx=20, pady=0)

    list_frame = tk.Frame(dlg)
    list_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
    listbox = tk.Listbox(list_frame, font=(_dialog_font()[0], 10),
                         selectmode=tk.SINGLE, height=min(12, len(files)),
                         exportselection=False)
    for f in files:
        listbox.insert(tk.END, f)
    if files:
        listbox.selection_set(0)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll = tk.Scrollbar(list_frame, command=listbox.yview)
    listbox.configure(yscrollcommand=scroll.set)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_ok():
        sel = listbox.curselection()
        if sel:
            result['value'] = files[sel[0]]
            dlg.destroy()

    listbox.bind('<Double-Button-1>', lambda e: _on_ok())

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text='下一步 →', width=14, command=_on_ok,
              bg='#2e86c1', fg='white', font=(_dialog_font()[0], 10, 'bold')
              ).pack()

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


def _configure_split_dialog(filepath):
    """Sheet 拆分配置对话框：展示所有 Sheet，各自勾选 + 各自选拆分列"""
    global _root
    _show_dialog_root()
    result = {'value': None}

    # 读取大表
    try:
        all_sheets = read_table(filepath)
    except Exception as e:
        messagebox.showerror('读取失败', f'无法读取 {os.path.basename(filepath)}:\n{e}')
        return None

    if not all_sheets:
        messagebox.showerror('错误', '文件没有任何 Sheet。')
        return None

    sheet_info = []
    for sn, data in all_sheets.items():
        if data and data[0]:
            cols = [str(h) if h is not None else '' for h in data[0]]
            valid_cols = [(i, c) for i, c in enumerate(cols) if c.strip()]
            has_data = len(data) > 1
            sheet_info.append((sn, valid_cols, has_data))
        else:
            sheet_info.append((sn, [], False))

    if not sheet_info:
        messagebox.showerror('错误', '文件中没有可读的 Sheet。')
        return None

    filename = os.path.basename(filepath)

    # ── 构建 UI ──
    rename_var = tk.BooleanVar(value=False)

    dlg = tk.Toplevel(_root)
    dlg.title('Sheet 拆分配置')
    dlg.resizable(False, True)
    dlg.transient(_root)
    dlg.grab_set()
    try:
        dlg.attributes('-topmost', True)
    except Exception:
        pass
    try:
        rx, ry = _root.winfo_x(), _root.winfo_y()
    except Exception:
        rx, ry = 100, 100
    dlg.geometry('+%d+%d' % (rx + 80, ry + 80))

    tk.Label(dlg, text='Sheet 拆分配置', font=(_dialog_font()[0], 12, 'bold')
             ).pack(padx=20, pady=15)
    tk.Label(dlg, text='大表: %s\n勾选要拆分的 Sheet，选择各自的拆分依据列。\n拆出的文件包含所有勾选的 Sheet。' % filename,
             font=(_dialog_font()[0], 9), fg='#666'
             ).pack(padx=20, pady=0)

    # 滚动区域
    config_canvas = tk.Canvas(dlg, borderwidth=0, highlightthickness=0)
    config_canvas.pack(side=tk.LEFT, padx=20, pady=10, fill=tk.BOTH, expand=True)
    scrollbar = tk.Scrollbar(dlg, orient=tk.VERTICAL, command=config_canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10, padx=0)
    config_canvas.configure(yscrollcommand=scrollbar.set)
    config_inner = tk.Frame(config_canvas)
    config_canvas.create_window((0, 0), window=config_inner, anchor=tk.NW)

    # 表头
    header_row = tk.Frame(config_inner, bg='#2e86c1')
    header_row.pack(fill=tk.X)
    tk.Label(header_row, text='参与', width=5, font=(_dialog_font()[0], 9, 'bold'),
             bg='#2e86c1', fg='white').pack(side=tk.LEFT, padx=2, pady=2)
    tk.Label(header_row, text='Sheet 名称', width=22, anchor=tk.W,
             font=(_dialog_font()[0], 9, 'bold'), bg='#2e86c1', fg='white'
             ).pack(side=tk.LEFT, padx=2, pady=2)
    tk.Label(header_row, text='拆分依据列', width=20,
             font=(_dialog_font()[0], 9, 'bold'), bg='#2e86c1', fg='white'
             ).pack(side=tk.LEFT, padx=2, pady=2)

    check_vars = {}
    col_vars = {}
    select_menus = {}

    def _on_sheet_toggle(sn):
        checked = check_vars[sn].get()
        if sn in select_menus:
            mb, _ = select_menus[sn]
            state = tk.NORMAL if checked else tk.DISABLED
            try:
                mb.configure(state=state)
            except Exception:
                pass

    for i, (sn, valid_cols, has_data) in enumerate(sheet_info):
        bg = '#fff' if i % 2 == 0 else '#f0f4f8'
        row_frame = tk.Frame(config_inner, bg=bg)
        row_frame.pack(fill=tk.X)

        has_cols = len(valid_cols) > 0
        chk_var = tk.BooleanVar(value=has_cols and has_data)
        check_vars[sn] = chk_var

        chk = tk.Checkbutton(row_frame, variable=chk_var, bg=bg, anchor=tk.CENTER,
                             width=3, command=lambda sn=sn: _on_sheet_toggle(sn))
        chk.pack(side=tk.LEFT, padx=2, pady=3)

        tk.Label(row_frame, text=sn, width=22, anchor=tk.W,
                font=(_dialog_font()[0], 10), bg=bg).pack(side=tk.LEFT, padx=2, pady=3)

        if valid_cols:
            col_names = [c for _, c in valid_cols]
            col_var = tk.StringVar(value=col_names[0])
            col_vars[sn] = col_var
            dropdown_frame = tk.Frame(row_frame, bg=bg)
            dropdown_frame.pack(side=tk.LEFT, padx=2, pady=3)
            mb = tk.Menubutton(dropdown_frame, textvariable=col_var,
                               width=18, anchor=tk.W, font=(_dialog_font()[0], 10),
                               relief=tk.RAISED, borderwidth=1, bg='white', indicatoron=True)
            menu = tk.Menu(mb, tearoff=0, font=(_dialog_font()[0], 10))
            for _, cn in valid_cols:
                menu.add_radiobutton(label=cn, variable=col_var, value=cn)
            mb.configure(menu=menu)
            mb.pack()
            select_menus[sn] = (mb, menu)
        else:
            col_vars[sn] = tk.StringVar(value='(无列名)')
            tk.Label(row_frame, text='(无列名)', width=18, anchor=tk.W,
                    font=(_dialog_font()[0], 10), bg=bg, fg='#999'
                    ).pack(side=tk.LEFT, padx=2, pady=3)

    # 初始禁用无列/无数据的 Sheet
    for sn, _, _ in sheet_info:
        if not check_vars[sn].get():
            _on_sheet_toggle(sn)

    # ── Sheet 重命名选项（v1.4：仅单 Sheet 时可用）──
    rename_frame = tk.Frame(dlg)
    rename_frame.pack(pady=6, padx=20, anchor=tk.W)
    rename_chk = tk.Checkbutton(
        rename_frame, variable=rename_var, text='以拆分列值重命名 Sheet（仅单个 Sheet 时可用）',
        font=(_dialog_font()[0], 9), anchor=tk.W
    )
    rename_chk.pack(side=tk.LEFT)

    def _refresh_rename_state():
        selected_count = sum(1 for sn, _, _ in sheet_info if check_vars[sn].get())
        if selected_count == 1:
            rename_chk.configure(state=tk.NORMAL)
        else:
            rename_var.set(False)
            rename_chk.configure(state=tk.DISABLED)

    # 将 rename 状态刷新绑定到 sheet 勾选变化
    _orig_toggle = _on_sheet_toggle
    def _on_sheet_toggle_wrapped(sn):
        _orig_toggle(sn)
        _refresh_rename_state()

    # 重绑 checkbutton 的 command
    for sn, _, _ in sheet_info:
        check_vars[sn].trace_add('write', lambda *a, _sn=sn: _refresh_rename_state())

    _refresh_rename_state()

    # 按钮
    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=10, padx=20)

    def _on_back():
        result['value'] = '__BACK__'
        dlg.destroy()

    def _on_start():
        selected = {}
        for sn, _, _ in sheet_info:
            if check_vars[sn].get():
                selected[sn] = col_vars[sn].get()
        if not selected:
            messagebox.showwarning('提示', '请至少勾选一个 Sheet 进行拆分。')
            return
        result['value'] = (selected, rename_var.get())
        dlg.destroy()

    tk.Button(btn_frame, text='← 上一步', width=12, command=_on_back,
              font=(_dialog_font()[0], 10)).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text='开始拆分', width=14, command=_on_start,
              bg='#2e86c1', fg='white', font=(_dialog_font()[0], 10, 'bold')
              ).pack(side=tk.LEFT, padx=5)

    config_inner.update_idletasks()
    config_canvas.configure(scrollregion=config_canvas.bbox('all'))
    config_canvas.configure(
        width=min(650, config_inner.winfo_reqwidth()),
        height=min(400, config_inner.winfo_reqheight()))

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


# ══════════════════════════════════════════════════════════════════
#  核心拆分逻辑
# ══════════════════════════════════════════════════════════════════

def split_tables(filepath, sheet_configs, rename_sheet=False):
    """按 sheet_configs 拆分大表。

    参数:
      filepath: str
      sheet_configs: {sheet_name: split_col_name}
      rename_sheet: bool（v1.4 预留，改名在 main 写入阶段处理）

    返回:
      {safe_value: {original_sheet_name: [header_row, data_row1, ...]}}
    """
    filename = os.path.basename(filepath)
    log(f'正在读取大表: {filename}')

    try:
        all_sheets = read_table(filepath)
    except Exception as e:
        log(f'  ❌ 读取失败: {e}')
        return {}

    if not all_sheets:
        log(f'  ⚠ 文件无数据')
        return {}

    value_sheets = {}
    sheet_groups = {}
    total_rows = 0
    skipped_empty = 0

    for sheet_name, split_col in sheet_configs.items():
        data = all_sheets.get(sheet_name)
        if not data:
            log(f'  Sheet "{sheet_name}" 不存在于文件中，跳过')
            continue
        if not data[0]:
            log(f'  Sheet "{sheet_name}" 表头为空，跳过')
            continue

        header = [str(h) if h is not None else '' for h in data[0]]
        try:
            split_idx = header.index(split_col)
        except ValueError:
            stripped_header = [h.strip() for h in header]
            try:
                split_idx = stripped_header.index(split_col.strip())
            except ValueError:
                log(f'  Sheet "{sheet_name}" 未找到拆分列 "{split_col}"，跳过')
                continue

        new_header = [str(h) if h is not None else '' for i, h in enumerate(header) if i != split_idx]
        sheet_groups[sheet_name] = {}
        sheet_skipped = 0

        for row in data[1:]:
            if not row or all(v == '' or v is None for v in row):
                continue
            val = row[split_idx] if split_idx < len(row) else ''
            val = str(val).strip() if val is not None else ''
            if not val:
                sheet_skipped += 1
                skipped_empty += 1
                if skipped_empty <= 5:
                    preview = [str(v)[:10] for v in row[:4]]
                    log(f'  Sheet "{sheet_name}" 拆分列值为空，跳过行: {preview}...')
                continue

            safe_val = sanitize_filename(val)
            new_row = [str(row[i]) if i < len(row) and row[i] is not None else ''
                       for i in range(len(row)) if i != split_idx]

            if safe_val not in sheet_groups[sheet_name]:
                sheet_groups[sheet_name][safe_val] = [new_header]
            sheet_groups[sheet_name][safe_val].append(new_row)

            if safe_val not in value_sheets:
                value_sheets[safe_val] = set()
            value_sheets[safe_val].add(sheet_name)

        rows_added = sum(len(grp) - 1 for grp in sheet_groups[sheet_name].values())
        if sheet_skipped > 0:
            log(f'  Sheet "{sheet_name}": {rows_added} 行, 跳过 {sheet_skipped} 空行')
        else:
            log(f'  Sheet "{sheet_name}": {rows_added} 行')
        total_rows += rows_added

    if skipped_empty > 0:
        log(f'共跳过 {skipped_empty} 行（拆分列值为空）')

    # 构建最终输出
    result = {}
    all_selected_sheets = set(sheet_configs.keys())
    for safe_val in value_sheets:
        result[safe_val] = {}
        for sn in all_selected_sheets:
            if sn in sheet_groups and safe_val in sheet_groups[sn]:
                result[safe_val][sn] = sheet_groups[sn][safe_val]
            else:
                # 无匹配行 → 仅保留表头
                data = all_sheets.get(sn)
                if data and data[0]:
                    split_col = sheet_configs.get(sn, '')
                    header = [str(h) if h is not None else '' for h in data[0]]
                    try:
                        si = header.index(split_col)
                    except ValueError:
                        si = None
                    new_header = [str(h) if h is not None else '' for i, h in enumerate(header) if i != si] if si is not None else header
                else:
                    new_header = []
                result[safe_val][sn] = [new_header]

    file_count = len(result)
    sheet_count = len(all_selected_sheets)
    log(f'拆分完成: 共 {total_rows} 行, 分为 {file_count} 个文件 × {sheet_count} Sheet')
    return result


# ══════════════════════════════════════════════════════════════════
#  主流程
# ══════════════════════════════════════════════════════════════════

def main():
    global _root

    _root = tk.Tk()
    _root.withdraw()
    try:
        sw = _root.winfo_screenwidth()
        sh = _root.winfo_screenheight()
    except Exception:
        sw, sh = 200, 200
    _root.geometry('1x1+%d+%d' % (sw // 2, sh // 2))
    _root.deiconify()

    if getattr(sys, 'frozen', False):
        work_dir = os.path.dirname(sys.executable)
    else:
        work_dir = os.getcwd()
    log(f'工作目录: {work_dir}')

    all_files = _scan_work_dir(work_dir)
    if not all_files:
        messagebox.showinfo('提示', '所在目录未找到 Excel/CSV 文件。')
        log('未找到可用文件，退出')
        _root.destroy()
        return
    log(f'扫描到 {len(all_files)} 个文件')

    # 对话框一：选文件
    selected_file = _select_file_dialog(all_files)
    if selected_file is None:
        log('用户取消文件选择，程序退出')
        _root.destroy()
        return
    log(f'选择大表: {selected_file}')
    filepath = os.path.join(work_dir, selected_file)

    # 对话框二：Sheet 配置
    dialog_result = _configure_split_dialog(filepath)
    if dialog_result is None:
        log('用户取消拆分配置，程序退出')
        _root.destroy()
        return

    # 上一步支持
    while dialog_result == '__BACK__':
        log('用户返回上一步')
        selected_file = _select_file_dialog(all_files)
        if selected_file is None:
            log('用户取消文件选择，程序退出')
            _root.destroy()
            return
        log(f'选择大表: {selected_file}')
        filepath = os.path.join(work_dir, selected_file)
        dialog_result = _configure_split_dialog(filepath)
        if dialog_result is None:
            log('用户取消拆分配置，程序退出')
            _root.destroy()
            return

    sheet_configs, rename_sheet = dialog_result
    log(f'拆分配置: {sheet_configs}, 重命名Sheet: {rename_sheet}')

    # 自动递增输出目录
    base_dir = os.path.join(work_dir, '拆分输出')
    output_dir = base_dir
    counter = 2
    while os.path.exists(output_dir):
        output_dir = os.path.join(work_dir, f'拆分输出_{counter}')
        counter += 1
    os.makedirs(output_dir, exist_ok=True)
    log(f'输出目录: {os.path.basename(output_dir)}/')

    # 拆分
    split_result = split_tables(filepath, sheet_configs, rename_sheet=rename_sheet)
    if not split_result:
        messagebox.showinfo('提示', '拆分后没有生成任何文件。\n请检查拆分列的值是否为空。')
        log('无拆分结果，退出')
        write_log(os.path.join(work_dir, '拆分日志.txt'))
        _root.destroy()
        return

    # 写入
    success_count = 0
    for safe_val, sheets_data in split_result.items():
        output_path = os.path.join(output_dir, f'{safe_val}.xlsx')
        try:
            ordered_sheets = OrderedDict()
            if rename_sheet and len(sheet_configs) == 1:
                # v1.4：单 Sheet 时，Sheet 名 = 拆分列值（= 文件名 safe_val）
                only_sn = next(iter(sheet_configs))
                if only_sn in sheets_data and sheets_data[only_sn]:
                    ordered_sheets[safe_val] = sheets_data[only_sn]
            else:
                for sn in sheet_configs:
                    if sn in sheets_data and sheets_data[sn]:
                        ordered_sheets[sn] = sheets_data[sn]
            if ordered_sheets:
                write_xlsx(output_path, ordered_sheets)
                total_rows = sum(len(rows) - 1 for rows in ordered_sheets.values())
                sheet_names = ', '.join(ordered_sheets.keys())
                log(f'  ✅ {safe_val}.xlsx (Sheet: {sheet_names}, {total_rows} 行)')
                success_count += 1
            else:
                log(f'  ⚠ {safe_val}.xlsx 无有效数据，跳过')
        except Exception as e:
            log(f'  ❌ 写入 {safe_val}.xlsx 失败: {e}')

    log(f'共生成 {success_count} 个文件 → {output_dir}')
    write_log(os.path.join(work_dir, '拆分日志.txt'))

    info_lines = ['拆分完成！', '']
    info_lines.append(f'大表: {selected_file}')
    info_lines.append(f'参与 Sheet: {len(sheet_configs)} 个')
    for sn, col in sheet_configs.items():
        info_lines.append(f'  • {sn} → 按「{col}」拆分')
    info_lines.append(f'生成文件: {success_count} 个')
    info_lines.append(f'输出目录: {os.path.basename(output_dir)}/')
    info_lines.append(f'日志文件: 拆分日志.txt')

    messagebox.showinfo('拆分完成', '\n'.join(info_lines))
    _root.destroy()


if __name__ == '__main__':
    try:
        main()
    except SystemExit:
        pass
    except Exception as e:
        err_msg = traceback.format_exc()
        try:
            with open('错误日志_拆分.txt', 'w', encoding='utf-8') as f:
                f.write(err_msg)
            print('\n' + '=' * 50, flush=True)
            print('程序出错，详情已写入 错误日志_拆分.txt', flush=True)
            print('=' * 50, flush=True)
            print(err_msg, flush=True)
        except Exception:
            print(err_msg, flush=True)
        try:
            messagebox.showerror('程序错误', f'程序运行出错：\n{str(e)}\n\n详情已写入 错误日志_拆分.txt')
        except Exception:
            pass
        input('\n按回车键退出...')
    else:
        try:
            input('\n按回车键退出...')
        except Exception:
            pass
