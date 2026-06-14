#!/usr/bin/env python3
"""
Excel 小表并大表工具 — 将多个结构相似的 Excel 小表按列名匹配合并到大表模板中。
零第三方依赖，仅使用 Python 标准库。
"""

import os, sys, zipfile, io, time, traceback, shutil
from xml.etree import ElementTree as ET
import tkinter as tk
from tkinter import messagebox, simpledialog

# ── xlsx 命名空间 ──────────────────────────────────────────────
NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# 注册命名空间避免 ns0: 前缀（同一前缀只能注册一个，优先修复 worksheet）
ET.register_namespace('', NS_S)
ET.register_namespace('r', NS_R)


def _tostring_xml(element):
    """序列化 Element 为 XML 字节串，输出标准 xlsx 兼容的声明头"""
    raw = ET.tostring(element, encoding='unicode')
    # 替换 ElementTree 的单引号声明为标准的双引号+standalone 格式
    if raw.startswith("<?xml version='1.0' encoding='utf-8'?>"):
        raw = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + raw[len("<?xml version='1.0' encoding='utf-8'?>"):]
    elif raw.startswith("<?xml version='1.0' encoding='UTF-8'?>"):
        raw = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + raw[len("<?xml version='1.0' encoding='UTF-8'?>"):]
    return raw.encode('UTF-8')

# ── 日志 ────────────────────────────────────────────────────────
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

def _parse_cell_ref(ref):
    """解析单元格引用 'A1' → (col_index, row_num)。col_index 为 0-based。"""
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


# ══════════════════════════════════════════════════════════════════
#  文件读取（xlsx / csv）
# ══════════════════════════════════════════════════════════════════

def read_csv(filepath):
    """读取 CSV 文件，返回 {'Sheet1': [header_row, data_row1, ...]}"""
    import csv
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = [row for row in reader if row]
    if not rows:
        return {}
    return {'Sheet1': rows}


def read_table(filepath):
    """根据后缀分发读取 xlsx 或 csv"""
    if filepath.lower().endswith('.csv'):
        return read_csv(filepath)
    return read_xlsx(filepath)


def read_xlsx(filepath):
    """读取 xlsx 文件，返回 {sheet_name: [header_row, data_row1, ...]}"""
    z = zipfile.ZipFile(filepath, 'r')

    # 共享字符串表
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.parse(z.open('xl/sharedStrings.xml')).getroot()
        for si in root.findall('.//{%s}si' % NS_S):
            texts = []
            for t in si.iter('{%s}t' % NS_S):
                if t.text:
                    texts.append(t.text)
            shared_strings.append(''.join(texts))

    # 获取 sheet 名称和顺序
    wb_root = ET.parse(z.open('xl/workbook.xml')).getroot()
    sheet_elems = wb_root.findall('.//{%s}sheet' % NS_S)
    sheet_names = [s.get('name', '') for s in sheet_elems]

    # 读取每个 sheet
    result = {}
    for idx, name in enumerate(sheet_names):
        sheet_file = 'xl/worksheets/sheet%d.xml' % (idx + 1)
        if sheet_file not in z.namelist():
            continue

        ws_root = ET.parse(z.open(sheet_file)).getroot()
        rows = ws_root.findall('.//{%s}row' % NS_S)

        sheet_data = []
        for row_el in rows:
            # 按列排序收集单元格，保留位置关系（解决 xlsx XML 中单元格乱序问题）
            cell_positions = {}  # {col_index: value}
            cells = row_el.findall('{%s}c' % NS_S)
            for c in cells:
                cell_type = c.get('t', '')
                ref = c.get('r', '')
                v_el = c.find('{%s}v' % NS_S)
                val = v_el.text if v_el is not None else ''

                if cell_type == 's' and val and val.isdigit():
                    idx_s = int(val)
                    if idx_s < len(shared_strings):
                        val = shared_strings[idx_s]

                ci, _ = _parse_cell_ref(ref)
                cell_positions[ci] = val

            if not cell_positions:
                continue

            # 按列序还原，空位补空字符串
            max_col = max(cell_positions.keys())
            row_values = [cell_positions.get(ci, '') for ci in range(max_col + 1)]

            # 过滤掉全空行
            if any(v for v in row_values):
                sheet_data.append(row_values)

        result[name] = sheet_data

    z.close()
    return result


# ══════════════════════════════════════════════════════════════════
#  xlsx 写入（纯 Python 标准库）
# ══════════════════════════════════════════════════════════════════

def _col_letter(col):
    """0-based column index → Excel column letter (A, B, ..., Z, AA, ...)"""
    result = ''
    col += 1
    while col > 0:
        col -= 1
        result = chr(65 + col % 26) + result
        col //= 26
    return result

def write_xlsx(filepath, sheets_data):
    """
    sheets_data: OrderedDict or dict {sheet_name: [row1, row2, ...]}
    每个 row 是 list of values (str/numbers)。
    """
    from collections import OrderedDict
    shared_strings = []
    ss_map = {}

    def _ss_idx(text):
        s = str(text) if text is not None else ''
        if s not in ss_map:
            ss_map[s] = len(shared_strings)
            shared_strings.append(s)
        return ss_map[s]

    # 预扫描所有字符串
    for _, rows in sheets_data.items():
        for row in rows:
            for cell in row:
                if isinstance(cell, str):
                    _ss_idx(cell)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:

        # [Content_Types].xml (raw XML to avoid ns0 prefix)
        ct_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
        ]
        for i in range(len(sheets_data)):
            ct_parts.append(
                '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
            )
        ct_parts.append('</Types>')
        zf.writestr('[Content_Types].xml', '\n'.join(ct_parts))

        # _rels/.rels (raw XML)
        rels_str = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
            '</Relationships>'
        )
        zf.writestr('_rels/.rels', rels_str)

        # xl/_rels/workbook.xml.rels
        # 用原始 XML 写入，确保 NS_R 使用默认命名空间（Excel 期望无前缀）
        wb_rels_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
        ]
        for i in range(len(sheets_data)):
            wb_rels_parts.append(
                '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 2, i + 1)
            )
        wb_rels_parts.append('</Relationships>')
        zf.writestr('xl/_rels/workbook.xml.rels', '\n'.join(wb_rels_parts))

        # xl/workbook.xml
        wb = ET.Element('{%s}workbook' % NS_S)
        sheets_el = ET.SubElement(wb, '{%s}sheets' % NS_S)
        for i, name in enumerate(list(sheets_data.keys())):
            ET.SubElement(sheets_el, '{%s}sheet' % NS_S,
                          name=name, sheetId=str(i + 1),
                          **{'{%s}id' % NS_R: 'rId%d' % (i + 2)})
        zf.writestr('xl/workbook.xml',
                    _tostring_xml(wb))

        # xl/sharedStrings.xml
        sst = ET.Element('{%s}sst' % NS_S,
                         count=str(len(shared_strings)),
                         uniqueCount=str(len(shared_strings)))
        for s in shared_strings:
            si = ET.SubElement(sst, '{%s}si' % NS_S)
            t = ET.SubElement(si, '{%s}t' % NS_S)
            # Escape XML special chars
            t.text = s
            # Don't preserve spaces by default - let Excel handle it
        zf.writestr('xl/sharedStrings.xml',
                    _tostring_xml(sst))

        # xl/worksheets/sheetN.xml
        for idx, (sn, rows) in enumerate(sheets_data.items()):
            ws = ET.Element('{%s}worksheet' % NS_S)
            # 添加 dimension 元素（格式必须是 A1:X99，不能漏 A1 起点）
            last_col = max(len(r) for r in rows) - 1 if rows else 0
            last_row = len(rows)
            dim_ref = 'A1:%s%d' % (_col_letter(last_col), last_row)
            ET.SubElement(ws, '{%s}dimension' % NS_S, ref=dim_ref)
            # 添加 sheetViews（Excel 要求至少一个 sheetView）
            sv = ET.SubElement(ws, '{%s}sheetViews' % NS_S)
            ET.SubElement(sv, '{%s}sheetView' % NS_S, workbookViewId='0')
            # 格式默认值
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
                        if isinstance(val, int):
                            pass  # default t is fine
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
            zf.writestr('xl/worksheets/sheet%d.xml' % (idx + 1),
                        _tostring_xml(ws))

    with open(filepath, 'wb') as f:
        f.write(buf.getvalue())


# ══════════════════════════════════════════════════════════════════
#  核心合并逻辑
# ══════════════════════════════════════════════════════════════════

# 全局（由 main 设置）
_root = None     # tkinter root
_work_dir = ''
_cancelled_tables = []   # {filename, reason}
_exceptions = []          # {filename, sheet, row_num, col_name, col1_pos, col1_val, col2_pos, col2_val, exc_type, action}
_big_table_snapshot = {}  # {sheet_name: [header_row]}


def _basename_no_ext(filename):
    """去除最后一个 .xlsx 后缀"""
    if filename.lower().endswith('.xlsx'):
        return filename[:-5]
    return os.path.splitext(filename)[0]


def _scan_work_dir(work_dir):
    """扫描工作目录下所有 .xlsx / .csv 文件，排除 ~$ 临时文件"""
    files = []
    for f in os.listdir(work_dir):
        if f.startswith('~$'):
            continue
        low = f.lower()
        if low.endswith('.xlsx') or low.endswith('.csv'):
            files.append(f)
    return sorted(files)


def _template_file_dialog(files):
    """模板文件单选对话框。返回选中的文件名或 None（取消）"""
    _show_dialog_root()
    result = {'value': None}

    dlg = tk.Toplevel(_root)
    dlg.title('选择模板文件')
    dlg.resizable(False, True)
    dlg.transient(_root)
    dlg.grab_set()
    try:
        dlg.attributes('-topmost', True)
    except Exception:
        pass

    tk.Label(dlg, text='选择模板文件', font=(_dialog_font()[0], 12, 'bold')
             ).pack(padx=20, pady=15)

    tk.Label(dlg, text='选择一个 Excel/CSV 文件作为合并模板。\n'
                       '模板仅需包含表头行，不应包含数据。',
             font=(_dialog_font()[0], 9), fg='#666'
             ).pack(padx=20, pady=0)

    # 列表
    list_frame = tk.Frame(dlg)
    list_frame.pack(padx=20, fill=tk.BOTH, expand=True)

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

    tk.Button(dlg, text='确定', width=12, command=_on_ok,
              bg='#2e86c1', fg='white', font=(_dialog_font()[0], 10, 'bold')
              ).pack(pady=10)

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


def _merge_files_dialog(files, template_name):
    """多选待合并文件对话框。返回选中的文件名列表或 None（取消）"""
    _show_dialog_root()
    result = {'value': None}

    dlg = tk.Toplevel(_root)
    dlg.title('选择待合并文件')
    dlg.resizable(False, True)
    dlg.transient(_root)
    dlg.grab_set()
    try:
        dlg.attributes('-topmost', True)
    except Exception:
        pass

    tk.Label(dlg, text='选择待合并文件', font=(_dialog_font()[0], 12, 'bold')
             ).pack(padx=20, pady=15)

    tk.Label(dlg, text='勾选要合并到模板的文件（已排除模板: %s）' % template_name,
             font=(_dialog_font()[0], 9), fg='#666'
             ).pack(padx=20, pady=0)

    # 可滚动的 checkbox 区域
    canvas = tk.Canvas(dlg, borderwidth=0, highlightthickness=0)
    canvas.pack(side=tk.LEFT, padx=20, pady=0, fill=tk.BOTH, expand=True)
    scrollbar = tk.Scrollbar(dlg, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=0, padx=0)
    canvas.configure(yscrollcommand=scrollbar.set)

    inner = tk.Frame(canvas)
    canvas.create_window((0, 0), window=inner, anchor=tk.NW)

    var_map = {}  # {filename: tk.BooleanVar}
    row_map = {}  # {filename: tk.Frame}  for background color

    def _on_toggle(fname, var):
        """勾选/取消时切换背景色"""
        row = row_map.get(fname)
        if row:
            try:
                new_bg = '#d4e6f1' if var.get() else 'white'
                row.configure(bg=new_bg)
                for child in row.winfo_children():
                    child.configure(bg=new_bg)
            except Exception:
                pass

    for f in files:
        var = tk.BooleanVar(value=False)
        var_map[f] = var

        row = tk.Frame(inner, bg='white')
        row.pack(fill=tk.X)
        row_map[f] = row

        cb = tk.Checkbutton(row, text=f, variable=var,
                            font=(_dialog_font()[0], 10),
                            bg='white', anchor=tk.W,
                            command=lambda f=f, v=var: _on_toggle(f, v))
        cb.pack(fill=tk.X, padx=5, pady=2)

    # 全选按钮
    def _select_all():
        for f, var in var_map.items():
            var.set(True)
            _on_toggle(f, var)

    def _deselect_all():
        for f, var in var_map.items():
            var.set(False)
            _on_toggle(f, var)

    btn_bar = tk.Frame(dlg)
    btn_bar.pack(pady=0)

    all_state = {'on': True}
    def _toggle_all():
        if all_state['on']:
            _deselect_all()
            all_state['on'] = False
            btn_all.configure(text='全选')
        else:
            _select_all()
            all_state['on'] = True
            btn_all.configure(text='全不选')

    btn_all = tk.Button(btn_bar, text='全选', width=8, command=_toggle_all)
    btn_all.pack(side=tk.LEFT, padx=5)

    def _on_ok():
        selected = [f for f, var in var_map.items() if var.get()]
        if not selected:
            messagebox.showwarning('提示', '请至少选择一个文件。')
            return
        result['value'] = selected
        dlg.destroy()

    tk.Button(btn_bar, text='确定', width=12, command=_on_ok,
              bg='#2e86c1', fg='white', font=(_dialog_font()[0], 10, 'bold')
              ).pack(side=tk.LEFT, padx=5)

    # 初始全选
    _select_all()

    inner.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox('all'))
    canvas.configure(width=inner.winfo_reqwidth(), height=min(300, inner.winfo_reqheight()))

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


def _read_big_table(filepath):
    """读取大表快照，验证「表名」列存在，返回快照"""
    sheets = read_table(filepath)
    if not sheets:
        raise ValueError('大表没有任何 Sheet')

    # 验证每个 sheet 的第一列表头是否为「表名」
    for sheet_name, data in sheets.items():
        if not data or not data[0]:
            raise ValueError('Sheet「%s」表头为空' % sheet_name)
        header = data[0]
        if not header or header[0] != '表名':
            raise ValueError(
                'Sheet「%s」缺少必填列「表名」，当前第一列为「%s」\n'
                '请检查大表文件，确保第一列为「表名」。' % (sheet_name, header[0] if header else '(空)')
            )

    # 快照：只保留表头行
    snapshot = {name: data[0:1] for name, data in sheets.items()}
    log('大表快照已读取: %s (Sheet: %s)' % (filepath, ', '.join(snapshot.keys())))
    return snapshot


def _sample_vals(col_values, max_n=5):
    """取前 max_n 个值，逗号分隔，空单元格留空。
       返回: "30,,50,60,70" """
    vals = [str(v) if v != '' and v is not None else '' for v in col_values[:max_n]]
    return ','.join(vals)


def _resolve_dup_with_big(small_header, data_rows, dup_col_names,
                          big_header, big_col_groups, filename, sheet_name):
    """
    三步同名列处理：
      Step 1: 小表所有同名列每行值完全一致 → 该值填入大表全部同名列
      Step 2: 小表同名列数量 == 大表同名列数量 → 按列顺序依次对应填入
      Step 3: 小表同名列数量 > 大表同名列数量 → 对大表每一个同名列弹窗选择
    返回 (big_to_small, cancelled)
      big_to_small: {big_col_idx: small_col_idx} 大表列位置到小表列位置的映射
    """
    big_to_small = {}

    for col_name, small_idxs in dup_col_names.items():
        big_idxs = big_col_groups.get(col_name, [])
        n_big = len(big_idxs)
        n_small = len(small_idxs)

        if n_big == 0:
            # 大表没有同名列（不应出现，dup_col_names 已过滤），跳过
            continue

        # 提取小表所有同名列的每列值
        small_col_vals = []  # [(small_idx, values_list)]
        all_same = True
        first_vals = None
        for si in small_idxs:
            vals = [row[si] if si < len(row) else '' for row in data_rows]
            small_col_vals.append((si, vals))
            if first_vals is None:
                first_vals = vals
            elif vals != first_vals:
                all_same = False

        # ═══ Step 1: 值全部一致 → 全部大表列用同一个小表列值 ═══
        if all_same and n_small >= 1:
            for bi in big_idxs:
                big_to_small[bi] = small_idxs[0]
            log('  %s / %s: 同名列「%s」%d列值完全一致，自动填入全部 %d 个大表列' %
                (filename, sheet_name, col_name, n_small, n_big))
            continue

        # ═══ Step 2: 数量相等 → 顺序对应填入 ═══
        if n_small == n_big:
            for i in range(n_small):
                big_to_small[big_idxs[i]] = small_idxs[i]
            log('  %s / %s: 同名列「%s」数量相等(%d列)，按顺序对应填入' %
                (filename, sheet_name, col_name, n_small))
            continue

        # ═══ Step 3: 小表多于大表 → 逐列弹窗 ═══
        # 对每个大表同名列位置，让用户从小表 M 列中选一列填入
        col_choices = [(si, vals, _col_letter(si)) for si, vals in small_col_vals]

        for big_idx_pos, bi in enumerate(big_idxs):
            choice = _dup_col_multi_dialog(
                col_name, big_idx_pos + 1, n_big, n_small, col_choices,
                filename, sheet_name
            )
            if choice == 'cancel_table':
                return None, True
            elif isinstance(choice, int):
                big_to_small[bi] = choice
                log('  %s / %s: 同名列「%s」大表第%d/%d列 → 小表 %s 列' %
                    (filename, sheet_name, col_name,
                     big_idx_pos + 1, n_big, _col_letter(choice)))
            # 'ignore' → 此大表列留空，不记异常（由最终落选判断）

        # 所有弹窗结束后，记录真正落选的小表列
        used_small = set(v for k, v in big_to_small.items() if k in big_idxs)
        for si in small_idxs:
            if si not in used_small:
                col_vals = [row[si] if si < len(row) else '' for row in data_rows]
                sample = _sample_vals(col_vals)
                _exceptions.append({
                    'filename': filename, 'sheet': sheet_name,
                    'row_num': '-', 'col_name': col_name,
                    'col1_pos': _col_letter(si), 'col1_val': sample + ' (共%d行)' % len(data_rows),
                    'col2_pos': '-', 'col2_val': '-',
                    'exc_type': '同名列未选择', 'action': '忽略',
                })

    return big_to_small, False


def _dup_col_multi_dialog(col_name, big_idx_pos, n_big, n_small, col_choices,
                          filename, sheet_name):
    """
    多列选择对话框：从小表 n_small 列中选一列填入大表第 big_idx_pos 个同名列。
    col_choices: [(small_idx, values_list, col_letter), ...]
    返回: int(small_idx) / 'ignore' / 'cancel_table'
    """
    _show_dialog_root()
    result = {'value': 'cancel_table'}
    max_preview = min(len(col_choices[0][1]) if col_choices else 0, 30)

    dlg = tk.Toplevel(_root)
    dlg.title('同名列选择 — %s (大表 %d/%d)' % (col_name, big_idx_pos, n_big))
    dlg.resizable(True, True)
    dlg.transient(_root)
    dlg.grab_set()

    try:
        dlg.attributes('-topmost', True)
    except Exception:
        pass

    # 标题
    header_info = '文件: %s  |  Sheet: %s  |  列名: %s  |  大表第 %d/%d 个' % (
        filename, sheet_name, col_name, big_idx_pos, n_big)
    tk.Label(dlg, text=header_info, font=_dialog_font(), padx=15, pady=10).pack(anchor=tk.W)

    tk.Label(dlg, text='小表中共有 %d 个「%s」列，请选择哪一列填入大表这个位置：' % (n_small, col_name),
             font=(_dialog_font()[0], 10, 'bold'), fg='#1a5276', padx=15).pack(anchor=tk.W)

    # 展示所有列的值对比
    frame = tk.Frame(dlg)
    frame.pack(padx=15, pady=5, fill=tk.BOTH, expand=True)

    # 构建 Text widget 显示对比
    text_widget = tk.Text(frame, font=('Courier', 10), width=80,
                          height=min(18, max_preview + 3), wrap=tk.NONE)
    scroll_y = tk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
    scroll_x = tk.Scrollbar(frame, orient=tk.HORIZONTAL, command=text_widget.xview)
    text_widget.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    text_widget.grid(row=0, column=0, sticky='nsew')
    scroll_y.grid(row=0, column=1, sticky='ns')
    scroll_x.grid(row=1, column=0, sticky='ew')
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    # 表头行
    header_line = '%-5s' % '行号'
    for _, _, letter in col_choices:
        header_line += '  %-16s' % ('%s(%s)' % (col_name, letter))
    text_widget.insert(tk.END, header_line + '\n', 'header')
    text_widget.insert(tk.END, '─' * 80 + '\n', 'sep')

    # 数据行
    for row_i in range(max_preview):
        line = '%-5d' % (row_i + 1)
        for _, vals, _ in col_choices:
            v = vals[row_i] if row_i < len(vals) else ''
            line += '  %-16s' % (v[:15])
        tag = 'diff' if _row_has_diff(row_i, col_choices) else 'same'
        text_widget.insert(tk.END, line + '\n', tag)

    if len(col_choices[0][1]) > max_preview:
        text_widget.insert(tk.END,
                           '（仅显示前 %d 行，共 %d 行）\n' % (max_preview, len(col_choices[0][1])),
                           'info')

    text_widget.tag_config('header', font=('Courier', 10, 'bold'), foreground='#1a5276')
    text_widget.tag_config('sep', foreground='gray')
    text_widget.tag_config('diff', background='#fef9e7', foreground='#7d6608')
    text_widget.tag_config('same', foreground='#2c3e50')
    text_widget.tag_config('info', foreground='gray')
    text_widget.configure(state=tk.DISABLED)

    # 按钮区
    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=10, padx=15)

    def _set_and_close(v):
        result['value'] = v
        dlg.destroy()

    for i, (si, _, letter) in enumerate(col_choices):
        tk.Button(btn_frame, text='填入第%d列 %s(%s)' % (i + 1, col_name, letter),
                  width=18,
                  command=lambda v=si: _set_and_close(v),
                  bg='#d4e6f1').pack(side=tk.LEFT, padx=3)

    tk.Button(btn_frame, text='忽略(留空)', width=12,
              command=lambda: _set_and_close('ignore'),
              bg='#f9e79f').pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text='取消合并该表', width=14,
              command=lambda: _set_and_close('cancel_table')).pack(side=tk.LEFT, padx=3)

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


def _row_has_diff(row_i, col_choices):
    """检查第 row_i 行各列值是否不完全一致"""
    vals = []
    for _, col_vals, _ in col_choices:
        v = col_vals[row_i] if row_i < len(col_vals) else ''
        vals.append(v)
    return len(set(vals)) > 1


def _extra_cols_dialog(filename, sheet_name, header_len, extra_rows):
    """数据列超出表头范围的对话框"""
    _show_dialog_root()
    result = {'value': 'ignore'}

    dlg = tk.Toplevel(_root)
    dlg.title('数据列超出表头')
    dlg.resizable(False, False)
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

    # 构建详情
    extra_preview = []
    for row_num, row_len, extra_vals in extra_rows[:10]:
        vals_str = ', '.join(str(v)[:8] for v in extra_vals[:5])
        extra_preview.append('  第 %d 行: %d 列 (多余值: %s)' % (row_num, row_len, vals_str))
    more = '\n  ... 还有 %d 行' % (len(extra_rows) - 10) if len(extra_rows) > 10 else ''

    detail = (
        '【数据列超出表头范围】\n\n'
        '文件: %s\n'
        'Sheet: %s\n'
        '表头有 %s 列，但以下数据行超出了表头列数：\n\n'
        '%s%s\n\n'
        '超出范围的列没有列名，无法匹配到大表。' % (
            filename, sheet_name, header_len,
            '\n'.join(extra_preview), more,
        )
    )
    tk.Label(dlg, text=detail, justify=tk.LEFT, padx=20, pady=15,
             font=_dialog_font()).pack()

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=0, padx=20)

    def _set_and_close(v):
        result['value'] = v
        dlg.destroy()

    tk.Button(btn_frame, text='忽略多余列（仅保留表头范围的列）', width=28,
              command=lambda: _set_and_close('ignore')).pack(side=tk.LEFT, padx=4)
    tk.Button(btn_frame, text='取消合并该表', width=16,
              command=lambda: _set_and_close('cancel_table')).pack(side=tk.LEFT, padx=4)

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


def _dialog_font():
    """返回跨平台安全的字体配置"""
    import platform
    if platform.system() == 'Windows':
        return ('Microsoft YaHei', 10)
    else:
        return ('Sans', 10)


def _show_dialog_root():
    """确保 root 窗口可见，否则 Windows 上弹窗可能不显示"""
    try:
        _root.update()
        _root.lift()
    except Exception:
        pass


def _no_header_dialog(filename, sheet_name, col_idx):
    """无列名冲突对话框"""
    _show_dialog_root()
    result = {'value': 'ignore'}

    dlg = tk.Toplevel(_root)
    dlg.title('无列名检测')
    dlg.resizable(False, False)
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

    detail = (
        '【无列名检测】\n\n'
        '文件: %s\n'
        'Sheet: %s\n'
        '第 %d 列（%s）无列名但有数据\n\n'
        '请选择处理方式:' % (
            filename, sheet_name, col_idx + 1, _col_letter(col_idx)
        )
    )
    tk.Label(dlg, text=detail, justify=tk.LEFT, padx=20, pady=15,
             font=_dialog_font()).pack()

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=0, padx=20)

    def _set_and_close(v):
        result['value'] = v
        dlg.destroy()

    tk.Button(btn_frame, text='忽略该列（记录到异常表）', width=24,
              command=lambda: _set_and_close('ignore')).pack(side=tk.LEFT, padx=4)
    tk.Button(btn_frame, text='取消合并该表', width=16,
              command=lambda: _set_and_close('cancel_table')).pack(side=tk.LEFT, padx=4)

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


def _discarded_cols_dialog(filename, sheet_name, discarded_cols):
    """小表列名在大表中不存在的对话框"""
    _show_dialog_root()
    result = {'value': 'ignore'}

    dlg = tk.Toplevel(_root)
    dlg.title('列名不匹配 — %s' % filename)
    dlg.resizable(False, False)
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

    cols_str = '\n'.join('  • %s' % c for c in discarded_cols[:20])
    more = '\n  ... 还有 %d 列' % (len(discarded_cols) - 20) if len(discarded_cols) > 20 else ''

    detail = (
        '【列名不匹配】\n\n'
        '文件: %s\n'
        'Sheet: %s\n\n'
        '以下 %d 个列名在大表中不存在，数据将被丢弃：\n\n'
        '%s%s\n\n'
        '可能是大表或小表的列名写错了，请核实。\n'
        '选择「忽略并继续」将丢弃这些列，并记录到异常表。' % (
            filename, sheet_name, len(discarded_cols),
            cols_str, more,
        )
    )
    tk.Label(dlg, text=detail, justify=tk.LEFT, padx=20, pady=15,
             font=_dialog_font()).pack()

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=0, padx=20)

    def _set_and_close(v):
        result['value'] = v
        dlg.destroy()

    tk.Button(btn_frame, text='忽略并继续（记录到异常表）', width=26,
              command=lambda: _set_and_close('ignore')).pack(side=tk.LEFT, padx=4)
    tk.Button(btn_frame, text='取消合并该表', width=16,
              command=lambda: _set_and_close('cancel_table')).pack(side=tk.LEFT, padx=4)

    try:
        dlg.focus_force()
    except Exception:
        pass
    dlg.wait_window()
    return result['value']


def _confirm_overwrite_dialog():
    """合并结果已存在时的确认对话框"""
    _show_dialog_root()
    return messagebox.askyesno(
        '文件已存在',
        '合并结果.xlsx 已存在，是否覆盖？'
    )


def _auto_fill_value(value):
    """自动填入时转换值类型"""
    if value is None or value == '':
        return ''
    # 尝试转为数字
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped or stripped == '-':
            return value
        # 尝试 int
        try:
            return int(stripped)
        except ValueError:
            pass
        # 尝试 float
        try:
            return float(stripped)
        except ValueError:
            pass
    return value


def process_small_table(filepath, filename, big_snapshot):
    """
    处理一个小表，返回 (success, row_data_dict) 或 (False, None) 表示取消。
    row_data_dict: {sheet_name: [row_values, ...]}
    """
    try:
        small_sheets = read_table(filepath)
    except Exception as e:
        log('错误: 读取 %s 失败 - %s' % (filename, e))
        return False, None

    big_sheet_names = set(big_snapshot.keys())
    small_sheet_names = set(small_sheets.keys())

    # Sheet 名称匹配
    matched_sheets = big_sheet_names & small_sheet_names
    if not matched_sheets:
        log('跳过 %s: Sheet 名称与大表不匹配 (小表: %s, 大表: %s)' %
            (filename, ', '.join(small_sheet_names), ', '.join(big_sheet_names)))
        _cancelled_tables.append({'filename': filename, 'reason': 'Sheet名称不匹配'})
        _exceptions.append({
            'filename': filename, 'sheet': '-', 'row_num': '-',
            'col_name': '-', 'col1_pos': '-', 'col1_val': '-',
            'col2_pos': '-', 'col2_val': '-',
            'exc_type': 'Sheet名称不匹配', 'action': '取消合并',
        })
        messagebox.showwarning('Sheet 不匹配',
                               '小表 %s 的 Sheet 名称与大表不一致，已跳过。' % filename)
        return False, None

    # 记录多余/缺少的 Sheet
    extra_sheets = small_sheet_names - big_sheet_names
    missing_sheets = big_sheet_names - small_sheet_names
    if missing_sheets:
        log('  %s 缺少 Sheet: %s，跳过' % (filename, ', '.join(missing_sheets)))
    if extra_sheets:
        log('  %s 多余 Sheet: %s，忽略' % (filename, ', '.join(extra_sheets)))

    # 缓存当前小表的所有结果行
    row_cache = {sn: [] for sn in matched_sheets}

    for sheet_name in matched_sheets:
        big_header = big_snapshot[sheet_name][0]  # 大表表头
        small_data = small_sheets[sheet_name]

        if not small_data or not small_data[0]:
            log('  %s / %s: 无表头，跳过' % (filename, sheet_name))
            continue

        small_header = small_data[0]
        data_rows = small_data[1:]

        if not data_rows:
            log('  %s / %s: 无数据行，跳过' % (filename, sheet_name))
            continue

        # ── 检测超出表头范围的数据列（小表3的那种情况）──
        header_len = len(small_header)
        extra_rows = []
        for row_idx, row in enumerate(data_rows, 2):
            if len(row) > header_len:
                extra = row[header_len:]  # 超出表头的数据
                extra_rows.append((row_idx, len(row), extra))

        if extra_rows:
            choice = _extra_cols_dialog(filename, sheet_name, header_len, extra_rows)
            if choice == 'cancel_table':
                _cancelled_tables.append({'filename': filename, 'reason': '用户取消合并'})
                return False, None
            else:
                # 按额外列位置逐列汇总（而非逐行记录）
                max_extra = max(len(ev) for _, _, ev in extra_rows)
                n_rows = len(extra_rows)
                for col_offset in range(max_extra):
                    col_letter = _col_letter(header_len + col_offset)
                    col_values = []
                    for _, _, extra_vals in extra_rows:
                        col_values.append(
                            str(extra_vals[col_offset])
                            if col_offset < len(extra_vals) else ''
                        )
                    sample = _sample_vals(col_values)
                    _exceptions.append({
                        'filename': filename, 'sheet': sheet_name,
                        'row_num': '全部行',
                        'col_name': '超表头列数',
                        'col1_pos': col_letter,
                        'col1_val': sample + ' (共%d行)' % n_rows,
                        'col2_pos': '-', 'col2_val': '-',
                        'exc_type': '超表头列数', 'action': '忽略多余列',
                    })
                log('  %s / %s: 检测到 %d 行数据超出表头 %d 列（%d 个额外列），已忽略' %
                    (filename, sheet_name, n_rows, header_len, max_extra))

        # ── 检测无列名列（表头里的空列名）──
        no_header_cols = [ci for ci, h in enumerate(small_header) if h == '']
        for ci in no_header_cols:
            choice = _no_header_dialog(filename, sheet_name, ci)
            if choice == 'cancel_table':
                _cancelled_tables.append({'filename': filename, 'reason': '用户取消合并'})
                return False, None
            col_vals = [row[ci] if ci < len(row) else '' for row in data_rows]
            sample = _sample_vals(col_vals)
            _exceptions.append({
                'filename': filename, 'sheet': sheet_name,
                'row_num': '-', 'col_name': '数据异常，无列名',
                'col1_pos': _col_letter(ci),
                'col1_val': sample + ' (共%d行)' % len(data_rows),
                'col2_pos': '-', 'col2_val': '-',
                'exc_type': '无列名', 'action': '忽略',
            })

        # ── 构建列映射 ──
        # 大表列名 → 所有出现位置（保留重复列的全部位置）
        big_col_groups = {}  # {col_name: [col_idx, ...]}
        for bi, name in enumerate(big_header):
            big_col_groups.setdefault(name, []).append(bi)

        small_col_groups = {}  # {col_name: [col_idx, ...]}
        for si, name in enumerate(small_header):
            if name == '':
                continue
            small_col_groups.setdefault(name, []).append(si)

        dup_col_names = {n: idxs for n, idxs in small_col_groups.items() if len(idxs) > 1}

        # 丢弃的列（小表有但大表没有）
        discarded = [n for n in small_header if n != '' and n not in big_col_groups]
        if discarded:
            choice = _discarded_cols_dialog(filename, sheet_name, discarded)
            if choice == 'cancel_table':
                _cancelled_tables.append({'filename': filename, 'reason': '用户取消合并'})
                return False, None
            for col_name in discarded:
                si = small_header.index(col_name)
                col_vals = [row[si] if si < len(row) else '' for row in data_rows]
                sample = _sample_vals(col_vals)
                _exceptions.append({
                    'filename': filename, 'sheet': sheet_name,
                    'row_num': '-', 'col_name': col_name,
                    'col1_pos': _col_letter(si),
                    'col1_val': sample + ' (共%d行)' % len(data_rows),
                    'col2_pos': '-', 'col2_val': '-',
                    'exc_type': '列名不存在于大表', 'action': '忽略并继续',
                })
            log('  %s / %s: 丢弃列: %s' % (filename, sheet_name, ', '.join(discarded)))

        # ── 同名列处理（三步判断：值一致→顺序填入→逐列弹窗）──
        big_to_small = {}  # {big_col_idx: small_col_idx}
        if dup_col_names:
            big_to_small, cancelled = _resolve_dup_with_big(
                small_header, data_rows, dup_col_names,
                big_header, big_col_groups, filename, sheet_name
            )
            if cancelled:
                _cancelled_tables.append({'filename': filename, 'reason': '用户取消合并'})
                return False, None

        # ── 处理每一行（不逐行弹窗）──
        for row_idx, row in enumerate(data_rows, 2):
            out_row = []
            for bi, big_col_name in enumerate(big_header):
                if big_col_name == '表名':
                    out_row.append(_basename_no_ext(filename))
                elif bi in big_to_small:
                    # 通过大表列位置映射到小表列位置
                    si = big_to_small[bi]
                    val = row[si] if si < len(row) else ''
                    out_row.append(_auto_fill_value(val))
                elif big_col_name in small_col_groups:
                    si = small_col_groups[big_col_name][0]
                    val = row[si] if si < len(row) else ''
                    out_row.append(_auto_fill_value(val))
                else:
                    out_row.append('')

            row_cache[sheet_name].append(out_row)

    return True, row_cache


# ══════════════════════════════════════════════════════════════════
#  主流程
# ══════════════════════════════════════════════════════════════════

def main():
    global _root, _work_dir, _big_table_snapshot, _cancelled_tables, _exceptions

    # ── 初始化 tkinter ──
    _root = tk.Tk()
    _root.withdraw()
    try:
        sw = _root.winfo_screenwidth()
        sh = _root.winfo_screenheight()
    except Exception:
        sw, sh = 200, 200
    _root.geometry('1x1+%d+%d' % (sw // 2, sh // 2))
    _root.deiconify()

    # ── 定位工作目录 ──
    if getattr(sys, 'frozen', False):
        _work_dir = os.path.dirname(sys.executable)
    else:
        _work_dir = os.getcwd()
    log('工作目录: %s' % _work_dir)

    # ── 扫描所有可用文件 ──
    all_files = _scan_work_dir(_work_dir)
    if not all_files:
        messagebox.showinfo('提示', '所在目录未找到 Excel/CSV 文件。')
        log('未找到可用文件，退出')
        _root.destroy()
        return
    log('扫描到 %d 个文件: %s' % (len(all_files), ', '.join(all_files)))

    # ── 对话框一：选择模板文件 ──
    template_name = _template_file_dialog(all_files)
    if template_name is None:
        log('用户取消模板选择，程序退出')
        _root.destroy()
        return

    template_path = os.path.join(_work_dir, template_name)

    # ── 读取并校验模板 ──
    while True:
        try:
            _big_table_snapshot = _read_big_table(template_path)
            break
        except ValueError as e:
            messagebox.showerror('模板校验失败', str(e))
            template_name = _template_file_dialog(all_files)
            if template_name is None:
                log('用户取消模板选择，程序退出')
                _root.destroy()
                return
            template_path = os.path.join(_work_dir, template_name)
        except Exception as e:
            messagebox.showerror('读取失败', '无法读取 %s:\n%s' % (template_name, e))
            template_name = _template_file_dialog(all_files)
            if template_name is None:
                log('用户取消模板选择，程序退出')
                _root.destroy()
                return
            template_path = os.path.join(_work_dir, template_name)

    log('模板文件: %s' % template_name)

    # ── 对话框二：选择待合并文件（排除模板）──
    merge_candidates = [f for f in all_files if f != template_name]
    if not merge_candidates:
        messagebox.showinfo('提示', '除了模板外没有其他可合并的文件。')
        log('无可合并文件，退出')
        _root.destroy()
        return

    small_files = _merge_files_dialog(merge_candidates, template_name)
    if small_files is None:
        log('用户取消文件选择，程序退出')
        _root.destroy()
        return

    log('待合并 %d 个文件: %s' % (len(small_files), ', '.join(small_files)))

    # ── 逐一处理待合并文件 ──
    from collections import OrderedDict
    result = OrderedDict()
    for sheet_name, snapshot_data in _big_table_snapshot.items():
        result[sheet_name] = list(snapshot_data)

    success_count = 0
    for idx, small_file in enumerate(small_files, 1):
        log('─' * 40)
        log('[%d/%d] 正在处理: %s' % (idx, len(small_files), small_file))
        small_path = os.path.join(_work_dir, small_file)

        ok, row_cache = process_small_table(small_path, small_file, _big_table_snapshot)
        if not ok:
            log('  → 已跳过')
            continue

        for sheet_name, rows in row_cache.items():
            result[sheet_name].extend(rows)
        success_count += 1
        log('  → 已合并 %d 行' % sum(len(r) for r in row_cache.values()))

    log('─' * 40)
    log('合并完成: %d/%d 个文件成功合并, %d 个被取消' %
        (success_count, len(small_files), len(_cancelled_tables)))

    if success_count == 0 and not _cancelled_tables:
        messagebox.showinfo('提示', '没有数据被合并。')
        write_log(os.path.join(_work_dir, '合并日志.txt'))
        _root.destroy()
        return

    # ── Step 4: 生成结果文件 ──
    output_path = os.path.join(_work_dir, '合并结果.xlsx')

    # 检查是否已存在
    if os.path.exists(output_path):
        if not _confirm_overwrite_dialog():
            log('用户取消覆盖，程序退出')
            _root.destroy()
            return
        os.remove(output_path)
        log('已删除旧的合并结果.xlsx')

    # 构建输出 sheets
    output_sheets = OrderedDict()
    for sheet_name, data_rows in result.items():
        if data_rows:
            output_sheets[sheet_name] = data_rows

    # 异常记录表
    if _exceptions:
        exc_header = ['文件名', 'Sheet', '行号', '列名', '列1位置', '列1值',
                      '列2位置(若同名)', '列2值(若同名)', '异常类型', '用户操作']
        exc_rows = [exc_header]
        for ex in _exceptions:
            exc_rows.append([
                ex.get('filename', ''),
                ex.get('sheet', ''),
                ex.get('row_num', ''),
                ex.get('col_name', ''),
                ex.get('col1_pos', ''),
                ex.get('col1_val', ''),
                ex.get('col2_pos', ''),
                ex.get('col2_val', ''),
                ex.get('exc_type', ''),
                ex.get('action', ''),
            ])
        output_sheets['异常记录'] = exc_rows

    # 取消合并的表名
    if _cancelled_tables:
        cancel_header = ['被取消的表名', '取消原因']
        cancel_rows = [cancel_header]
        for ct in _cancelled_tables:
            cancel_rows.append([ct['filename'], ct['reason']])
        output_sheets['取消合并的表名'] = cancel_rows

    try:
        write_xlsx(output_path, output_sheets)
        log('✅ 已生成: %s' % output_path)
        log('  Sheet 列表: %s' % ', '.join(output_sheets.keys()))
    except Exception as e:
        messagebox.showerror('写入失败',
                             '无法写入合并结果.xlsx:\n%s\n'
                             '请确认文件未被其他程序打开。' % str(e))
        log('错误: 写入合并结果.xlsx 失败 - %s' % e)
        _root.destroy()
        return

    # 写日志
    write_log(os.path.join(_work_dir, '合并日志.txt'))

    messagebox.showinfo('完成', '合并完成！\n\n'
                        '成功合并: %d 个小表\n'
                        '被取消: %d 个小表\n'
                        '结果文件: 合并结果.xlsx\n'
                        '日志文件: 合并日志.txt' % (success_count, len(_cancelled_tables)))

    _root.destroy()


if __name__ == '__main__':
    try:
        main()
    except SystemExit:
        pass
    except Exception as e:
        # 闪退保护：捕获所有异常，写出错误日志并等待用户查看
        err_msg = traceback.format_exc()
        try:
            with open('错误日志.txt', 'w', encoding='utf-8') as f:
                f.write(err_msg)
            print('\n' + '=' * 50, flush=True)
            print('程序出错，详情已写入 错误日志.txt', flush=True)
            print('=' * 50, flush=True)
            print(err_msg, flush=True)
        except Exception:
            print(err_msg, flush=True)
        try:
            messagebox.showerror('程序错误', '程序运行出错：\n%s\n\n详情已写入 错误日志.txt' % str(e))
        except Exception:
            pass
        input('\n按回车键退出...')
    else:
        try:
            input('\n按回车键退出...')
        except Exception:
            pass
