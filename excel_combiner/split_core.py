"""Headless split engine. This module has no GUI or task-global state."""

import csv
import re
import tempfile
from collections import OrderedDict
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Callable, Dict, Mapping, Optional

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter

from .errors import InputError, ProcessingError
from .models import SplitRequest, SplitRunResult


Logger = Optional[Callable[[str], None]]


def _emit(logger: Logger, message: str) -> None:
    if logger is not None:
        logger(message)


@dataclass(frozen=True)
class FormulaCell:
    """Formula plus enough source context to relocate it during splitting."""

    formula: str
    origin: str
    sheet_name: str
    cached_value: object = None
    removed_column: int = None


def read_csv(filepath):
    with open(filepath, 'r', encoding='utf-8-sig') as source:
        rows = [row for row in csv.reader(source) if row]
    return {'Sheet1': rows} if rows else {}


def read_xlsx(filepath):
    formula_wb = openpyxl.load_workbook(filepath, data_only=False, read_only=True)
    value_wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    try:
        result = {}
        for worksheet in formula_wb.worksheets:
            value_ws = value_wb[worksheet.title]
            sheet_data = []
            for row_number, cells in enumerate(worksheet.iter_rows(), 1):
                row_values = []
                for column_number, cell in enumerate(cells, 1):
                    value = cell.value
                    if cell.data_type == 'f':
                        value = FormulaCell(
                            formula=value if str(value).startswith('=') else '=' + str(value),
                            origin=cell.coordinate,
                            sheet_name=worksheet.title,
                            cached_value=value_ws.cell(row=row_number, column=column_number).value,
                        )
                    elif value is None:
                        value = ''
                    row_values.append(value)
                if any(value != '' and value is not None for value in row_values):
                    sheet_data.append(row_values)
            result[worksheet.title] = sheet_data
        return result
    finally:
        formula_wb.close()
        value_wb.close()


def read_table(filepath):
    return read_csv(filepath) if str(filepath).lower().endswith('.csv') else read_xlsx(filepath)


_A1_CELL_RE = re.compile(r'^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$')


def _unquote_sheet_name(value):
    if value.startswith("'") and value.endswith("'"):
        return value[1:-1].replace("''", "'")
    return value


def _translate_formula_reference(reference, formula_cell, destination):
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
        column_absolute, letters, row_absolute, row_text = match.groups()
        column_number = column_index_from_string(letters)
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
            column_absolute, get_column_letter(column_number), row_absolute, row_number))
    return prefix + ':'.join(translated)


def _translate_split_formula(formula_cell, destination, logger=None):
    try:
        pieces = []
        for token in Tokenizer(formula_cell.formula).items:
            value = token.value
            if token.type == 'OPERAND' and token.subtype == 'RANGE':
                value = _translate_formula_reference(value, formula_cell, destination)
            pieces.append(value)
        return '=' + ''.join(pieces)
    except Exception as exc:
        _emit(logger, '  警告: 公式 %s 无法安全调整，将使用缓存值: %s' %
              (formula_cell.origin, exc))
        return None


def write_xlsx(filepath, sheets_data, logger=None):
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    translated_count = 0
    cached_count = 0
    for sheet_name, rows in sheets_data.items():
        worksheet = workbook.create_sheet(title=sheet_name)
        for row_index, row in enumerate(rows, 1):
            for column_index, value in enumerate(row, 1):
                cell = worksheet.cell(row=row_index, column=column_index)
                if isinstance(value, FormulaCell):
                    formula = _translate_split_formula(value, cell.coordinate, logger)
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
    workbook.calculation.calcMode = 'auto'
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(filepath)
    workbook.close()
    if translated_count:
        _emit(logger, '  已保护并调整 %d 个公式引用' % translated_count)
    if cached_count:
        _emit(logger, '  %d 个无法调整的公式已改用缓存计算值' % cached_count)


def sanitize_filename(name):
    name = str(name).strip()
    for character in '/\\:*?"<>|[]':
        name = name.replace(character, '_')
    name = name.strip('. ')
    if len(name) > 31:
        name = name[:31]
    if not name:
        name = '_空值_'
    reserved = {'CON', 'PRN', 'AUX', 'NUL'}
    reserved.update('COM%d' % number for number in range(1, 10))
    reserved.update('LPT%d' % number for number in range(1, 10))
    if name.split('.')[0].upper() in reserved:
        name = '_' + name
    return name[:31].rstrip('. ')


def _plain_cell_value(value):
    if isinstance(value, FormulaCell):
        return value.cached_value if value.cached_value is not None else ''
    return value


def _split_output_value(value, removed_column):
    if isinstance(value, FormulaCell):
        return replace(value, removed_column=removed_column)
    # Keep the source workbook's cell type. Converting numeric operands to text
    # makes formulas such as SUMPRODUCT calculate as zero after splitting.
    return value if value is not None else ''


def split_tables(filepath, sheet_configs, rename_sheet=False, logger=None):
    filename = Path(filepath).name
    _emit(logger, '正在读取大表: %s' % filename)
    try:
        all_sheets = read_table(filepath)
    except Exception as exc:
        raise InputError('读取 %s 失败: %s' % (filename, exc))
    if not all_sheets:
        raise InputError('文件无数据: %s' % filename)

    value_sheets = {}
    sheet_groups = {}
    total_rows = 0
    skipped_empty = 0
    valid_configs = 0
    for sheet_name, split_col in sheet_configs.items():
        data = all_sheets.get(sheet_name)
        if not data:
            raise InputError('Sheet「%s」不存在于文件中' % sheet_name)
        if not data[0]:
            raise InputError('Sheet「%s」表头为空' % sheet_name)
        header = [str(_plain_cell_value(item)) if _plain_cell_value(item) is not None else ''
                  for item in data[0]]
        try:
            split_idx = header.index(split_col)
        except ValueError:
            stripped_header = [item.strip() for item in header]
            try:
                split_idx = stripped_header.index(split_col.strip())
            except ValueError:
                raise InputError('Sheet「%s」未找到拆分列「%s」' % (sheet_name, split_col))
        valid_configs += 1
        new_header = [str(item) if item is not None else ''
                      for index, item in enumerate(header) if index != split_idx]
        sheet_groups[sheet_name] = {}
        sheet_skipped = 0
        for row in data[1:]:
            if not row or all(item == '' or item is None for item in row):
                continue
            value = _plain_cell_value(row[split_idx]) if split_idx < len(row) else ''
            value = str(value).strip() if value is not None else ''
            if not value:
                sheet_skipped += 1
                skipped_empty += 1
                continue
            new_row = [_split_output_value(row[index], split_idx + 1)
                       if index < len(row) else ''
                       for index in range(len(row)) if index != split_idx]
            sheet_groups[sheet_name].setdefault(value, [new_header]).append(new_row)
            value_sheets.setdefault(value, set()).add(sheet_name)
        rows_added = sum(len(group) - 1 for group in sheet_groups[sheet_name].values())
        _emit(logger, '  Sheet「%s」: %d 行%s' % (
            sheet_name, rows_added, '，跳过 %d 个空值' % sheet_skipped if sheet_skipped else ''))
        total_rows += rows_added
    if valid_configs == 0:
        raise InputError('没有可执行的 Sheet 拆分配置')

    result = {}
    selected_sheets = list(sheet_configs)
    used_filenames = set()
    for value in value_sheets:
        base_name = sanitize_filename(value)
        safe_value = base_name
        counter = 2
        while safe_value.casefold() in used_filenames:
            suffix = '_%d' % counter
            safe_value = base_name[:31 - len(suffix)] + suffix
            counter += 1
        used_filenames.add(safe_value.casefold())
        result[safe_value] = {}
        for sheet_name in selected_sheets:
            if value in sheet_groups[sheet_name]:
                result[safe_value][sheet_name] = sheet_groups[sheet_name][value]
            else:
                data = all_sheets[sheet_name]
                header = [str(_plain_cell_value(item)) if _plain_cell_value(item) is not None else ''
                          for item in data[0]]
                split_col = sheet_configs[sheet_name]
                try:
                    split_idx = header.index(split_col)
                except ValueError:
                    split_idx = [item.strip() for item in header].index(split_col.strip())
                result[safe_value][sheet_name] = [[item for index, item in enumerate(header)
                                                   if index != split_idx]]
    _emit(logger, '拆分完成: 共 %d 行，生成 %d 个文件' % (total_rows, len(result)))
    return result


def run_split(request: SplitRequest, logger=None) -> SplitRunResult:
    source = Path(request.source)
    output_dir = Path(request.output_dir)
    if not source.is_file():
        raise InputError('输入文件不存在: %s' % source)
    if not request.sheet_configs:
        raise InputError('至少需要一个 --sheet Sheet名=拆分列 配置')
    if output_dir.exists():
        raise InputError('输出目录已存在，请指定一个新目录: %s' % output_dir)
    try:
        groups = split_tables(source, request.sheet_configs, request.rename_sheet, logger)
        output_dir.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(
                prefix=output_dir.name + '-', dir=str(output_dir.parent)) as temporary_name:
            temporary_dir = Path(temporary_name)
            names = []
            for safe_value, sheets in groups.items():
                output_sheets = sheets
                if request.rename_sheet and len(request.sheet_configs) == 1:
                    only_sheet = next(iter(request.sheet_configs))
                    output_sheets = OrderedDict([(safe_value, sheets[only_sheet])])
                filename = safe_value + '.xlsx'
                write_xlsx(str(temporary_dir / filename), output_sheets, logger)
                names.append(filename)
            temporary_dir.replace(output_dir)
        return SplitRunResult([output_dir / filename for filename in names])
    except InputError:
        raise
    except Exception as exc:
        raise ProcessingError('拆分失败: %s' % exc)
