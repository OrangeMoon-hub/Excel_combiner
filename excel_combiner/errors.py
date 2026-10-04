"""Public errors used by the core and command-line interface."""


class ExcelCombinerError(Exception):
    """Base class for expected application errors."""


class InputError(ExcelCombinerError):
    """An input path or explicit configuration is invalid."""


class DecisionRequired(InputError):
    """The input needs an explicit choice that a headless run cannot guess."""


class ProcessingError(ExcelCombinerError):
    """Processing failed after inputs and configuration were accepted."""
