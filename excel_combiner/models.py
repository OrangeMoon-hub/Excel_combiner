"""Task inputs and results shared by the GUI and CLI adapters."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence


@dataclass(frozen=True)
class MergeRequest:
    template: Path
    inputs: Sequence[Path]
    output: Path
    sheet_mapping: Dict[str, Optional[str]] = field(default_factory=dict)
    add_source_column: bool = False


@dataclass
class MergeRunResult:
    output: Path
    written_sheets: List[str]
    exceptions: List[dict] = field(default_factory=list)
    cancelled_tables: List[dict] = field(default_factory=list)


@dataclass(frozen=True)
class SplitRequest:
    source: Path
    output_dir: Path
    sheet_configs: Dict[str, str]
    rename_sheet: bool = False


@dataclass
class SplitRunResult:
    output_files: List[Path] = field(default_factory=list)
