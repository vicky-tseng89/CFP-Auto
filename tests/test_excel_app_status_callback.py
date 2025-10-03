import sys
import types
from pathlib import Path

import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

def _ensure_stub_module(name, **attributes):
    module = types.ModuleType(name)
    for attr_name, attr_value in attributes.items():
        setattr(module, attr_name, attr_value)
    sys.modules[name] = module
    return module


if 'docx' not in sys.modules:
    class _Document:  # pragma: no cover - trivial stub
        pass

    class _Inches:  # pragma: no cover - trivial stub
        def __init__(self, value):
            self.value = value

    docx_module = _ensure_stub_module('docx', Document=_Document)
    docx_shared = _ensure_stub_module('docx.shared', Inches=_Inches)
    docx_module.shared = docx_shared

if 'docxtpl' not in sys.modules:
    class _DocxTemplate:  # pragma: no cover - trivial stub
        def __init__(self, *_, **__):
            pass

    class _InlineImage:  # pragma: no cover - trivial stub
        def __init__(self, *_, **__):
            pass

    _ensure_stub_module('docxtpl', DocxTemplate=_DocxTemplate, InlineImage=_InlineImage)

if 'matplotlib' not in sys.modules:
    matplotlib_module = _ensure_stub_module('matplotlib')
    _ensure_stub_module('matplotlib.pyplot')
    matplotlib_module.pyplot = sys.modules['matplotlib.pyplot']

if 'numpy' not in sys.modules:
    numpy_module = _ensure_stub_module('numpy')
    numpy_module.array = lambda *args, **kwargs: args[0] if args else None  # pragma: no cover - trivial stub

if 'openpyxl' not in sys.modules:
    class _DummyWorkbook:  # pragma: no cover - trivial stub
        def __init__(self, *_, **__):
            pass

    def _load_workbook(*_, **__):  # pragma: no cover - trivial stub
        return _DummyWorkbook()

    openpyxl_module = _ensure_stub_module('openpyxl', load_workbook=_load_workbook)
    openpyxl_styles = _ensure_stub_module('openpyxl.styles')
    openpyxl_styles_colors = _ensure_stub_module('openpyxl.styles.colors', Color=type('Color', (), {}))
    openpyxl_module.styles = openpyxl_styles
    openpyxl_styles.colors = openpyxl_styles_colors

if 'pandas' not in sys.modules:
    pandas_module = _ensure_stub_module('pandas')
    pandas_module.read_excel = lambda *_, **__: None  # pragma: no cover - trivial stub
    pandas_module.DataFrame = type('DataFrame', (), {})

if 'xlsxwriter' not in sys.modules:
    _ensure_stub_module('xlsxwriter')

if 'win32com' not in sys.modules:
    win32com_module = _ensure_stub_module('win32com')
    def _dispatch_ex(*_, **__):  # pragma: no cover - trivial stub
        raise RuntimeError('DispatchEx is not available in the test stub')

    win32com_client = _ensure_stub_module('win32com.client', DispatchEx=_dispatch_ex)
    win32com_module.client = win32com_client

if 'pythoncom' not in sys.modules:
    def _noop(*_, **__):  # pragma: no cover - trivial stub
        return None

    _ensure_stub_module(
        'pythoncom',
        CoInitialize=_noop,
        CoUninitialize=_noop,
    )

from excel_processing import ExcelApp


def test_status_callback_defaults_to_no_op():
    app = ExcelApp()

    assert app._original_status_callback is None
    assert app._has_status_callback is False
    assert callable(app.status_callback)

    # Should not raise even though no callback was supplied.
    app.status_callback("initial message")
    app._notify_status("follow up message")


def test_status_callback_invoked_when_provided():
    received = []

    def capture(message):
        received.append(message)

    app = ExcelApp(status_callback=capture)
    app._notify_status("hello")

    assert received == ["hello"]
    assert app._has_status_callback is True
    assert app._original_status_callback is capture


def test_notify_status_handles_non_callable_assignment():
    app = ExcelApp()
    app.status_callback = None

    # Reassigning to None should not break status updates.
    app._notify_status("still safe")
