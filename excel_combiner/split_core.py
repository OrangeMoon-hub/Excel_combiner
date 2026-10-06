"""Headless split engine. This module has no GUI or task-global state."""

import csv
import re
import tempfile
import zipfile
from collections import OrderedDict
from copy import copy, deepcopy
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Callable, Dict, Mapping, Optional

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.utils import (column_index_from_string, coordinate_to_tuple,
                            get_column_letter, range_boundaries)
from openpyxl.worksheet.cell_range import MultiCellRange

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


class SplitGroup(dict):
    """Backward-compatible sheet mapping with the original business group value."""

    def __init__(self, group_value):
        super().__init__()
        self.group_value = group_value


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
        result[safe_value] = SplitGroup(value)
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
    _emit(logger, '拆分分组统计: 共 %d 行，%d 个非空分组' % (total_rows, len(result)))
    return result


def _split_column_index(worksheet, split_column_name):
    headers = [str(cell.value).strip() if cell.value is not None else ''
               for cell in worksheet[1]]
    try:
        return headers.index(split_column_name) + 1
    except ValueError:
        stripped_name = split_column_name.strip()
        try:
            return headers.index(stripped_name) + 1
        except ValueError:
            raise InputError('Sheet「%s」未找到拆分列「%s」' %
                             (worksheet.title, split_column_name))


def _row_group_value(worksheet, row_number, split_column):
    value = worksheet.cell(row=row_number, column=split_column).value
    if value is None:
        return ''
    return str(value).strip()


def _formula_reference_parts(reference, current_sheet):
    sheet_name = current_sheet
    coordinate = reference
    explicit_sheet = False
    if '!' in reference:
        sheet_part, coordinate = reference.rsplit('!', 1)
        if '[' in sheet_part or ']' in sheet_part:
            return None
        sheet_name = _unquote_sheet_name(sheet_part)
        explicit_sheet = True
    endpoints = coordinate.split(':')
    if len(endpoints) > 2:
        return None
    parsed = [_A1_CELL_RE.match(endpoint) for endpoint in endpoints]
    if not all(parsed):
        return None
    return sheet_name, explicit_sheet, parsed


def _check_formula_safety(workbook, sheet_configs):
    """Reject formula dependencies that cannot survive filtering without data leaks."""
    contexts = {}
    for sheet_name, split_column_name in sheet_configs.items():
        if sheet_name not in workbook.sheetnames:
            raise InputError('Sheet「%s」不存在于文件中' % sheet_name)
        worksheet = workbook[sheet_name]
        split_column = _split_column_index(worksheet, split_column_name)
        row_groups = {
            row_number: _row_group_value(worksheet, row_number, split_column)
            for row_number in range(2, worksheet.max_row + 1)
        }
        contexts[sheet_name] = (split_column, row_groups)

    for sheet_name, (split_column, row_groups) in contexts.items():
        worksheet = workbook[sheet_name]
        for row_number in range(2, worksheet.max_row + 1):
            group_value = row_groups[row_number]
            if not group_value:
                continue
            for cell in worksheet[row_number]:
                if cell.data_type != 'f' or not cell.value:
                    continue
                for token in Tokenizer(cell.value).items:
                    if token.type != 'OPERAND' or token.subtype != 'RANGE':
                        continue
                    parts = _formula_reference_parts(token.value, sheet_name)
                    if parts is None:
                        raise InputError(
                            'Sheet「%s」单元格 %s 含无法安全调整的复杂引用「%s」，已停止拆分' %
                            (sheet_name, cell.coordinate, token.value))
                    referenced_sheet, _, parsed = parts
                    if referenced_sheet.casefold() != sheet_name.casefold():
                        raise InputError(
                            'Sheet「%s」单元格 %s 含跨 Sheet 引用「%s」，为避免 #REF! 已停止拆分' %
                            (sheet_name, cell.coordinate, token.value))
                    columns = [column_index_from_string(match.group(2)) for match in parsed]
                    if min(columns) <= split_column <= max(columns):
                        raise InputError(
                            'Sheet「%s」单元格 %s 的公式引用了将被删除的拆分列，已停止拆分' %
                            (sheet_name, cell.coordinate))
                    start_row = int(parsed[0].group(4))
                    end_row = int(parsed[-1].group(4))
                    for referenced_row in range(min(start_row, end_row), max(start_row, end_row) + 1):
                        if referenced_row == 1:
                            continue
                        if row_groups.get(referenced_row) != group_value:
                            raise InputError(
                                'Sheet「%s」单元格 %s 的公式依赖其他分组或将被删除的行 %d，已停止拆分' %
                                (sheet_name, cell.coordinate, referenced_row))
    return contexts


def _check_unsupported_package_features(source):
    """Stop before openpyxl can silently discard controls, shapes, OLE or VBA."""
    if source.suffix.lower() == '.xlsm':
        raise InputError('宏工作簿暂不支持拆分；为避免静默丢失 VBA，已停止处理')
    if source.suffix.lower() != '.xlsx':
        return
    try:
        with zipfile.ZipFile(source) as archive:
            names = archive.namelist()
            lowered = [name.lower() for name in names]
            unsupported_paths = (
                'xl/ctrlprops/', 'xl/controls/', 'xl/embeddings/',
                'xl/richdata/', 'xl/vbaproject.bin',
            )
            if any(any(name.startswith(prefix) for prefix in unsupported_paths)
                   for name in lowered):
                raise InputError('工作簿包含宏按钮、复选框、OLE 或富数据控件；当前版本无法可靠保留，已停止拆分')
            for name in names:
                lower_name = name.lower()
                if not lower_name.endswith(('.xml', '.rels', '.vml')):
                    continue
                data = archive.read(name)
                lower_data = data.lower()
                if (b'/control' in lower_data or b'/oleobject' in lower_data or
                        b'objecttype="button"' in lower_data or
                        b'objecttype="checkbox"' in lower_data or
                        b'<xdr:sp>' in lower_data or b'<xdr:grpsp>' in lower_data):
                    raise InputError('工作簿包含当前版本不能可靠保留的按钮、控件或形状；已停止拆分')
    except zipfile.BadZipFile as exc:
        raise InputError('Excel 文件结构无效: %s' % exc)


def _translated_output_range(reference, split_column, row_mapping):
    min_column, min_row, max_column, max_row = range_boundaries(reference)
    if min_column == max_column == split_column:
        return None
    translated_min_column = min_column - (1 if min_column > split_column else 0)
    translated_max_column = max_column - (1 if max_column > split_column else 0)
    mapped_rows = []
    if min_row <= 1 <= max_row:
        mapped_rows.append(1)
    mapped_rows.extend(destination for source, destination in row_mapping.items()
                       if min_row <= source <= max_row)
    if not mapped_rows:
        return None
    return '%s%d:%s%d' % (
        get_column_letter(translated_min_column), min(mapped_rows),
        get_column_letter(translated_max_column), max(mapped_rows))


def _translate_preserved_formula(formula, sheet_name, destination_sheet_name,
                                  split_column, row_mapping):
    pieces = []
    for token in Tokenizer(formula).items:
        value = token.value
        if token.type == 'OPERAND' and token.subtype == 'RANGE':
            parts = _formula_reference_parts(value, sheet_name)
            if parts is None:
                raise InputError('公式含无法安全调整的复杂引用「%s」' % value)
            referenced_sheet, explicit_sheet, parsed = parts
            if referenced_sheet.casefold() != sheet_name.casefold():
                raise InputError('公式含跨 Sheet 引用「%s」' % value)
            translated = []
            for match in parsed:
                column_absolute, letters, row_absolute, row_text = match.groups()
                column_number = column_index_from_string(letters)
                if column_number == split_column:
                    raise InputError('公式引用了将被删除的拆分列')
                if column_number > split_column:
                    column_number -= 1
                source_row = int(row_text)
                if source_row == 1:
                    destination_row = 1
                elif source_row in row_mapping:
                    destination_row = row_mapping[source_row]
                else:
                    raise InputError('公式引用了其他分组或已删除行 %d' % source_row)
                translated.append('%s%s%s%d' % (
                    column_absolute, get_column_letter(column_number),
                    row_absolute, destination_row))
            value = ':'.join(translated)
            if explicit_sheet:
                escaped = destination_sheet_name.replace("'", "''")
                value = "'%s'!%s" % (escaped, value)
        pieces.append(value)
    return '=' + ''.join(pieces)


def _translate_rule_formula(formula, sheet_name, destination_sheet_name,
                            split_column, row_mapping, rule_name):
    """Move formula-bearing worksheet rules or reject unsafe dependencies."""
    if formula is None or not isinstance(formula, str) or not formula.strip():
        return formula
    stripped = formula.strip()
    if len(stripped) >= 2 and stripped[0] == stripped[-1] == '"':
        return formula
    has_equals = stripped.startswith('=')
    candidate = stripped if has_equals else '=' + stripped
    try:
        translated = _translate_preserved_formula(
            candidate, sheet_name, destination_sheet_name,
            split_column, row_mapping)
    except InputError as exc:
        raise InputError('%s无法安全迁移: %s' % (rule_name, exc))
    return translated if has_equals else translated[1:]


def _translate_multicell_range(value, split_column, row_mapping):
    translated = []
    for cell_range in MultiCellRange(str(value)).ranges:
        new_range = _translated_output_range(str(cell_range), split_column, row_mapping)
        if new_range:
            translated.append(new_range)
    return ' '.join(translated)


def _translate_chart_formula(formula, worksheet_name, destination_sheet_name,
                             split_column, row_mapping):
    parts = _formula_reference_parts(formula, worksheet_name)
    if parts is None:
        raise InputError('图表含无法安全调整的数据引用「%s」' % formula)
    referenced_sheet, _, parsed = parts
    if referenced_sheet.casefold() != worksheet_name.casefold():
        raise InputError('图表含跨 Sheet 数据引用「%s」，已停止拆分' % formula)
    coordinate = ':'.join(match.group(0) for match in parsed)
    translated = _translated_output_range(coordinate, split_column, row_mapping)
    if translated is None:
        return None
    escaped = destination_sheet_name.replace("'", "''")
    return "'%s'!%s" % (escaped, translated)


def _update_chart_references(worksheet, destination_sheet_name, split_column, row_mapping):
    if not row_mapping:
        worksheet._charts = []
        return
    for chart in worksheet._charts:
        for series in chart.series:
            for attribute in ('val', 'cat', 'xVal', 'yVal', 'bubbleSize', 'tx'):
                parent = getattr(series, attribute, None)
                if parent is None:
                    continue
                for ref_attribute in ('numRef', 'strRef'):
                    reference = getattr(parent, ref_attribute, None)
                    if reference is None or not getattr(reference, 'f', None):
                        continue
                    translated = _translate_chart_formula(
                        reference.f, worksheet.title, destination_sheet_name,
                        split_column, row_mapping)
                    if translated is None:
                        worksheet._charts = []
                        return
                    reference.f = translated


def _preserve_dimensions(worksheet, split_column, row_mapping,
                         source_column_dimensions, source_row_dimensions):
    worksheet.column_dimensions.clear()
    for source_index, dimension in source_column_dimensions.items():
        if source_index == split_column:
            continue
        destination_index = source_index - (1 if source_index > split_column else 0)
        destination_letter = get_column_letter(destination_index)
        copied = copy(dimension)
        copied.index = destination_letter
        copied.min = destination_index
        copied.max = destination_index
        worksheet.column_dimensions[destination_letter] = copied
    worksheet.row_dimensions.clear()
    if 1 in source_row_dimensions:
        header = copy(source_row_dimensions[1])
        header.index = 1
        worksheet.row_dimensions[1] = header
    for source_row, destination_row in row_mapping.items():
        if source_row not in source_row_dimensions:
            continue
        copied = copy(source_row_dimensions[source_row])
        copied.index = destination_row
        worksheet.row_dimensions[destination_row] = copied


def _filter_worksheet(worksheet, group_value, split_column, destination_sheet_name):
    keep_rows = [row_number for row_number in range(2, worksheet.max_row + 1)
                 if _row_group_value(worksheet, row_number, split_column) == group_value]
    row_mapping = {source_row: destination_row
                   for destination_row, source_row in enumerate(keep_rows, 2)}
    formulas = []
    for source_row, destination_row in row_mapping.items():
        for source_column in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=source_row, column=source_column)
            if cell.data_type != 'f' or not cell.value:
                continue
            destination_column = source_column - (1 if source_column > split_column else 0)
            formulas.append((destination_row, destination_column, cell.value))

    table_state = {
        name: (worksheet.tables[name].ref,
               list(worksheet.tables[name].tableColumns))
        for name in worksheet.tables
    }
    source_column_dimensions = {
        column_index_from_string(letter): copy(dimension)
        for letter, dimension in worksheet.column_dimensions.items()
    }
    source_row_dimensions = {
        index: copy(dimension) for index, dimension in worksheet.row_dimensions.items()
    }
    source_merges = [str(cell_range) for cell_range in worksheet.merged_cells.ranges]
    freeze_panes = worksheet.freeze_panes
    conditional_rules = list(worksheet.conditional_formatting._cf_rules.items())
    data_validations = list(worksheet.data_validations.dataValidation)
    auto_filter_ref = worksheet.auto_filter.ref

    for row_number in range(worksheet.max_row, 1, -1):
        if row_number not in row_mapping:
            worksheet.delete_rows(row_number)
    worksheet.delete_cols(split_column)

    for destination_row, destination_column, formula in formulas:
        worksheet.cell(destination_row, destination_column).value = _translate_preserved_formula(
            formula, worksheet.title, destination_sheet_name, split_column, row_mapping)

    _preserve_dimensions(
        worksheet, split_column, row_mapping,
        source_column_dimensions, source_row_dimensions)

    worksheet.merged_cells.ranges = set()
    for source_range in source_merges:
        translated = _translated_output_range(source_range, split_column, row_mapping)
        if translated:
            worksheet.merge_cells(translated)

    for name, (source_ref, source_columns) in table_state.items():
        table = worksheet.tables[name]
        if not row_mapping:
            del worksheet.tables[name]
            continue
        translated_ref = _translated_output_range(source_ref, split_column, row_mapping)
        if translated_ref is None:
            del worksheet.tables[name]
            continue
        min_column, _, max_column, _ = range_boundaries(source_ref)
        columns = list(source_columns)
        if min_column <= split_column <= max_column:
            del columns[split_column - min_column]
        for index, table_column in enumerate(columns, 1):
            table_column.id = index
        table.tableColumns = columns
        table.ref = translated_ref
        if table.autoFilter is not None:
            table.autoFilter.ref = translated_ref

    rebuilt_rules = OrderedDict()
    for conditional_formatting, rules in conditional_rules:
        translated = _translate_multicell_range(
            conditional_formatting.sqref, split_column, row_mapping)
        if not translated:
            continue
        copied = deepcopy(conditional_formatting)
        copied.sqref = MultiCellRange(translated)
        copied_rules = []
        for rule in rules:
            copied_rule = deepcopy(rule)
            if copied_rule.formula:
                copied_rule.formula = [
                    _translate_rule_formula(
                        formula, worksheet.title, destination_sheet_name,
                        split_column, row_mapping, '条件格式公式')
                    for formula in copied_rule.formula
                ]
            copied_rules.append(copied_rule)
        rebuilt_rules[copied] = copied_rules
    worksheet.conditional_formatting._cf_rules = rebuilt_rules

    kept_validations = []
    for validation in data_validations:
        translated = _translate_multicell_range(validation.sqref, split_column, row_mapping)
        if translated:
            copied_validation = deepcopy(validation)
            copied_validation.sqref = MultiCellRange(translated)
            copied_validation.formula1 = _translate_rule_formula(
                validation.formula1, worksheet.title, destination_sheet_name,
                split_column, row_mapping, '数据验证公式')
            copied_validation.formula2 = _translate_rule_formula(
                validation.formula2, worksheet.title, destination_sheet_name,
                split_column, row_mapping, '数据验证公式')
            kept_validations.append(copied_validation)
    worksheet.data_validations.dataValidation = kept_validations

    if auto_filter_ref:
        worksheet.auto_filter.ref = _translated_output_range(
            auto_filter_ref, split_column, row_mapping)
    if freeze_panes:
        freeze_coordinate = (freeze_panes.coordinate
                             if hasattr(freeze_panes, 'coordinate') else str(freeze_panes))
        freeze_row, source_freeze_column = coordinate_to_tuple(freeze_coordinate)
        freeze_column = source_freeze_column - (
            1 if source_freeze_column > split_column else 0)
        worksheet.freeze_panes = '%s%d' % (
            get_column_letter(max(1, freeze_column)), freeze_row)

    _update_chart_references(
        worksheet, destination_sheet_name, split_column, row_mapping)


def _write_preserved_xlsx(source, filepath, group_value, sheet_configs,
                          rename_sheet=False, logger=None):
    workbook = openpyxl.load_workbook(source, data_only=False)
    try:
        for worksheet in list(workbook.worksheets):
            if worksheet.title not in sheet_configs:
                workbook.remove(worksheet)
        only_sheet = next(iter(sheet_configs)) if len(sheet_configs) == 1 else None
        for sheet_name, split_column_name in sheet_configs.items():
            worksheet = workbook[sheet_name]
            split_column = _split_column_index(worksheet, split_column_name)
            destination_sheet_name = (sanitize_filename(group_value)
                                      if rename_sheet and only_sheet == sheet_name
                                      else sheet_name)
            _filter_worksheet(
                worksheet, group_value, split_column, destination_sheet_name)
            if destination_sheet_name != worksheet.title:
                worksheet.title = destination_sheet_name
        workbook.calculation.calcMode = 'auto'
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.save(filepath)
    finally:
        workbook.close()
    _emit(logger, '  已保留源单元格样式、日期格式、表格、图表与冻结窗格')


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
        _check_unsupported_package_features(source)
        if source.suffix.lower() == '.xlsx':
            preflight = openpyxl.load_workbook(source, data_only=False, read_only=True)
            try:
                _check_formula_safety(preflight, request.sheet_configs)
            finally:
                preflight.close()
        groups = split_tables(source, request.sheet_configs, request.rename_sheet, logger)
        if not groups:
            raise InputError('拆分列没有可输出的非空分组')
        output_dir.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(
                prefix=output_dir.name + '-', dir=str(output_dir.parent)) as temporary_name:
            temporary_dir = Path(temporary_name)
            names = []
            for safe_value, sheets in groups.items():
                filename = safe_value + '.xlsx'
                if source.suffix.lower() == '.xlsx':
                    _write_preserved_xlsx(
                        source, str(temporary_dir / filename), sheets.group_value,
                        request.sheet_configs, request.rename_sheet, logger)
                else:
                    output_sheets = sheets
                    if request.rename_sheet and len(request.sheet_configs) == 1:
                        only_sheet = next(iter(request.sheet_configs))
                        output_sheets = OrderedDict([(safe_value, sheets[only_sheet])])
                    write_xlsx(str(temporary_dir / filename), output_sheets, logger)
                names.append(filename)
            temporary_dir.replace(output_dir)
        return SplitRunResult([output_dir / filename for filename in names])
    except InputError:
        raise
    except Exception as exc:
        raise ProcessingError('拆分失败: %s' % exc)
