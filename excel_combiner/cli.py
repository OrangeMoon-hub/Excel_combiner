"""Supported command-line interface for repeatable merge and split tasks."""

import argparse
import sys
from pathlib import Path

from . import __version__
from .errors import InputError, ProcessingError
from .merge_core import run_merge
from .models import MergeRequest, SplitRequest
from .split_core import run_split


def _mapping(value, option_name):
    if '=' not in value:
        raise argparse.ArgumentTypeError('%s 必须使用 左侧=右侧 格式' % option_name)
    left, right = value.split('=', 1)
    left = left.strip()
    right = right.strip()
    if not left or not right:
        raise argparse.ArgumentTypeError('%s 两侧都不能为空' % option_name)
    return left, right


def _sheet_mapping(value):
    return _mapping(value, '--sheet-map')


def _split_sheet(value):
    return _mapping(value, '--sheet')


def build_parser():
    parser = argparse.ArgumentParser(
        prog='excel-combiner',
        description='Excel Combiner 1.7：无弹窗执行合并或拆分任务。')
    parser.add_argument('--version', action='version', version='Excel Combiner %s' % __version__)
    commands = parser.add_subparsers(dest='command', required=True)

    merge = commands.add_parser('merge', help='按模板列名合并一个或多个文件')
    merge.add_argument('--template', required=True, type=Path, help='大表模板路径')
    merge.add_argument('--input', required=True, action='append', type=Path,
                       help='输入文件路径；多个文件可重复提供')
    merge.add_argument('--output', required=True, type=Path, help='输出文件路径')
    merge.add_argument('--sheet-map', action='append', type=_sheet_mapping, default=[],
                       metavar='来源Sheet=目标Sheet', help='显式 Sheet 映射；可重复提供')
    merge.add_argument('--add-source-column', action='store_true',
                       help='在模板的“表名”列写入来源文件名')

    split = commands.add_parser('split', help='按指定列将工作簿拆分成多个文件')
    split.add_argument('--source', required=True, type=Path, help='源文件路径')
    split.add_argument('--output-dir', required=True, type=Path, help='输出目录')
    split.add_argument('--sheet', required=True, action='append', type=_split_sheet,
                       metavar='Sheet=拆分列', help='参与拆分的 Sheet 和列；可重复提供')
    split.add_argument('--rename-sheet', action='store_true', help='把输出 Sheet 改为分组值')
    return parser


def main(arguments=None):
    parser = build_parser()
    args = parser.parse_args(arguments)
    try:
        if args.command == 'merge':
            result = run_merge(MergeRequest(
                template=args.template,
                inputs=args.input,
                output=args.output,
                sheet_mapping=dict(args.sheet_map),
                add_source_column=args.add_source_column,
            ), logger=print)
            print('合并完成: %s' % result.output)
            return 0
        result = run_split(SplitRequest(
            source=args.source,
            output_dir=args.output_dir,
            sheet_configs=dict(args.sheet),
            rename_sheet=args.rename_sheet,
        ), logger=print)
        print('拆分完成: %d 个文件，输出目录: %s' %
              (len(result.output_files), args.output_dir))
        return 0
    except InputError as exc:
        print('输入或配置错误: %s' % exc, file=sys.stderr)
        return 3
    except ProcessingError as exc:
        print('处理失败: %s' % exc, file=sys.stderr)
        return 4


if __name__ == '__main__':
    sys.exit(main())
