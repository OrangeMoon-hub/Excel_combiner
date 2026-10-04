#!/usr/bin/env python3
"""
大表拆分工具 v1.3 — 按指定列值将大表拆分为多个小表。
每个 Sheet 独立选择拆分依据列，输出文件名=列值，Sheet名=原始Sheet名。
使用 openpyxl 生成有效工作簿并调整拆分后的公式引用。独立运行，不依赖合并脚本。
"""

import os, sys, time, traceback, re
from dataclasses import dataclass, replace
import tkinter as tk
from tkinter import messagebox
from collections import OrderedDict

try:
    import openpyxl
    from openpyxl.formula import Tokenizer
    from openpyxl.utils import column_index_from_string, get_column_letter
except ImportError:
    openpyxl = None
    Tokenizer = None

@dataclass(frozen=True)
class FormulaCell:
    """Formula plus enough source context to relocate it during splitting."""

    formula: str
    origin: str
    sheet_name: str
    cached_value: object = None
    removed_column: int = None


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
    """读取工作簿，同时保留公式、原坐标和最后计算缓存值。"""
    if openpyxl is None:
        raise RuntimeError('拆分含公式的工作簿需要 openpyxl。请先安装：pip install openpyxl')

    formula_wb = openpyxl.load_workbook(filepath, data_only=False, read_only=True)
    value_wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    result = {}
    for ws in formula_wb.worksheets:
        value_ws = value_wb[ws.title]
        sheet_data = []
        for row_number, cells in enumerate(ws.iter_rows(), 1):
            row_values = []
            for column_number, cell in enumerate(cells, 1):
                value = cell.value
                if cell.data_type == 'f':
                    value = FormulaCell(
                        formula=value if str(value).startswith('=') else '=' + str(value),
                        origin=cell.coordinate,
                        sheet_name=ws.title,
                        cached_value=value_ws.cell(row=row_number, column=column_number).value,
                    )
                elif value is None:
                    value = ''
                row_values.append(value)
            if any(v != '' and v is not None for v in row_values):
                sheet_data.append(row_values)
        result[ws.title] = sheet_data
    formula_wb.close()
    value_wb.close()
    return result


def read_table(filepath):
    if filepath.lower().endswith('.csv'):
        return read_csv(filepath)
    return read_xlsx(filepath)


# ══════════════════════════════════════════════════════════════════
#  xlsx 写入
# ══════════════════════════════════════════════════════════════════

_A1_CELL_RE = re.compile(r'^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$')


def _unquote_sheet_name(value):
    if value.startswith("'") and value.endswith("'"):
        return value[1:-1].replace("''", "'")
    return value


def _translate_formula_reference(reference, formula_cell, destination):
    """Translate a local A1 reference after filtering rows and deleting a column.

    Named ranges, structured references and references to another sheet are left
    unchanged. Relative row references follow the formula to its new output row;
    absolute rows remain fixed. Deleting the split column shifts references on
    its right and turns a direct reference to that column into ``#REF!``.
    """
    prefix = ''
    coordinate = reference
    if '!' in reference:
        sheet_part, coordinate = reference.rsplit('!', 1)
        if _unquote_sheet_name(sheet_part).casefold() != formula_cell.sheet_name.casefold():
            return reference
        prefix = sheet_part + '!'

    endpoints = coordinate.split(':')
    if len(endpoints) > 2:
        return reference
    parsed = [_A1_CELL_RE.match(endpoint) for endpoint in endpoints]
    if not all(parsed):
        return reference

    origin_match = _A1_CELL_RE.match(formula_cell.origin)
    destination_match = _A1_CELL_RE.match(destination)
    if origin_match is None or destination_match is None:
        return reference
    row_delta = int(destination_match.group(4)) - int(origin_match.group(4))

    translated = []
    for match in parsed:
        column_absolute, column_letters, row_absolute, row_text = match.groups()
        column_number = column_index_from_string(column_letters)
        if formula_cell.removed_column:
            if column_number == formula_cell.removed_column:
                return '#REF!'
            if column_number > formula_cell.removed_column:
                column_number -= 1
        row_number = int(row_text)
        if not row_absolute:
            row_number += row_delta
            if row_number < 1:
                return '#REF!'
        translated.append('%s%s%s%d' % (
            column_absolute,
            get_column_letter(column_number),
            row_absolute,
            row_number,
        ))
    return prefix + ':'.join(translated)


def _translate_split_formula(formula_cell, destination):
    try:
        tokenizer = Tokenizer(formula_cell.formula)
        pieces = []
        for token in tokenizer.items:
            value = token.value
            if token.type == 'OPERAND' and token.subtype == 'RANGE':
                value = _translate_formula_reference(value, formula_cell, destination)
            pieces.append(value)
        return '=' + ''.join(pieces)
    except Exception as exc:
        log('  警告: 公式 %s 无法安全调整，将使用缓存值: %s' %
            (formula_cell.origin, exc))
        return None


def write_xlsx(filepath, sheets_data):
    """用 openpyxl 生成拆分结果，并按新行列位置调整公式引用。"""
    if openpyxl is None:
        raise RuntimeError('拆分含公式的工作簿需要 openpyxl。请先安装：pip install openpyxl')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    translated_count = 0
    cached_count = 0
    for sheet_name, rows in sheets_data.items():
        ws = wb.create_sheet(title=sheet_name)
        for row_index, row in enumerate(rows, 1):
            for column_index, value in enumerate(row, 1):
                cell = ws.cell(row=row_index, column=column_index)
                if isinstance(value, FormulaCell):
                    formula = _translate_split_formula(value, cell.coordinate)
                    if formula is not None:
                        cell.value = formula
                        translated_count += 1
                    elif value.cached_value is not None:
                        cell.value = value.cached_value
                        cached_count += 1
                    else:
                        cell.value = ''
                else:
                    cell.value = '' if value is None else value
    wb.calculation.calcMode = 'auto'
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(filepath)
    wb.close()
    if translated_count:
        log('  已保护并调整 %d 个公式引用' % translated_count)
    if cached_count:
        log('  %d 个无法调整的公式已改用缓存计算值' % cached_count)


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
    # Windows 设备名即使带扩展名也不能用作普通文件名。
    reserved = {'CON', 'PRN', 'AUX', 'NUL'}
    reserved.update('COM%d' % i for i in range(1, 10))
    reserved.update('LPT%d' % i for i in range(1, 10))
    if name.split('.')[0].upper() in reserved:
        name = '_' + name
    name = name[:31].rstrip('. ')
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

def _plain_cell_value(value):
    if isinstance(value, FormulaCell):
        return value.cached_value if value.cached_value is not None else ''
    return value


def _split_output_value(value, removed_column):
    if isinstance(value, FormulaCell):
        return replace(value, removed_column=removed_column)
    return str(value) if value is not None else ''

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

        header = [str(_plain_cell_value(h)) if _plain_cell_value(h) is not None else ''
                  for h in data[0]]
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
            val = _plain_cell_value(row[split_idx]) if split_idx < len(row) else ''
            val = str(val).strip() if val is not None else ''
            if not val:
                sheet_skipped += 1
                skipped_empty += 1
                if skipped_empty <= 5:
                    preview = [str(v)[:10] for v in row[:4]]
                    log(f'  Sheet "{sheet_name}" 拆分列值为空，跳过行: {preview}...')
                continue

            new_row = [_split_output_value(row[i], split_idx + 1)
                       if i < len(row) else ''
                       for i in range(len(row)) if i != split_idx]

            # 分组使用原始业务值，文件名清理不能改变分组身份。
            if val not in sheet_groups[sheet_name]:
                sheet_groups[sheet_name][val] = [new_header]
            sheet_groups[sheet_name][val].append(new_row)

            if val not in value_sheets:
                value_sheets[val] = set()
            value_sheets[val].add(sheet_name)

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
    all_selected_sheets = list(sheet_configs)
    used_filenames = set()
    for val in value_sheets:
        base_name = sanitize_filename(val)
        safe_val = base_name
        counter = 2
        # 所有 Sheet 共用一次命名分配；兼顾 Windows 大小写不敏感路径。
        while safe_val.casefold() in used_filenames:
            suffix = '_%d' % counter
            safe_val = base_name[:31 - len(suffix)] + suffix
            counter += 1
        used_filenames.add(safe_val.casefold())
        log('  拆分值「%s」→ 文件「%s.xlsx」' % (val, safe_val))
        result[safe_val] = {}
        for sn in all_selected_sheets:
            if sn in sheet_groups and val in sheet_groups[sn]:
                result[safe_val][sn] = sheet_groups[sn][val]
            else:
                # 无匹配行 → 仅保留表头
                data = all_sheets.get(sn)
                if data and data[0]:
                    split_col = sheet_configs.get(sn, '')
                    header = [str(_plain_cell_value(h)) if _plain_cell_value(h) is not None else ''
                              for h in data[0]]
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
