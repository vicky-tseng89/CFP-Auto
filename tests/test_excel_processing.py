import sys
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


class DummyDataFrame:
    def __init__(self, rows):
        self._rows = rows

    class _ILoc:
        def __init__(self, outer):
            self._outer = outer

        def __getitem__(self, idx):
            return list(self._outer._rows[idx])

    @property
    def iloc(self):
        return DummyDataFrame._ILoc(self)

    @property
    def shape(self):
        if not self._rows:
            return (0, 0)
        return (len(self._rows), len(self._rows[0]))

    def replace(self, *args, **kwargs):
        return self

    def fillna(self, *args, **kwargs):
        return self


docx_module = types.ModuleType("docx")
docx_module.Document = object
sys.modules.setdefault("docx", docx_module)

docx_shared_module = types.ModuleType("docx.shared")
docx_shared_module.Inches = lambda *args, **kwargs: None
sys.modules.setdefault("docx.shared", docx_shared_module)

docxtpl_module = types.ModuleType("docxtpl")
docxtpl_module.DocxTemplate = object
docxtpl_module.InlineImage = object
sys.modules.setdefault("docxtpl", docxtpl_module)

openpyxl_module = types.ModuleType("openpyxl")
openpyxl_module.load_workbook = lambda *args, **kwargs: None
sys.modules.setdefault("openpyxl", openpyxl_module)

openpyxl_styles_module = types.ModuleType("openpyxl.styles")
sys.modules.setdefault("openpyxl.styles", openpyxl_styles_module)

openpyxl_colors_module = types.ModuleType("openpyxl.styles.colors")


class _Color:
    def __init__(self, *args, **kwargs):
        pass


openpyxl_colors_module.Color = _Color
sys.modules.setdefault("openpyxl.styles.colors", openpyxl_colors_module)

matplotlib_module = types.ModuleType("matplotlib")
matplotlib_pyplot_module = types.ModuleType("matplotlib.pyplot")
matplotlib_pyplot_module.figure = lambda *args, **kwargs: None
matplotlib_pyplot_module.plot = lambda *args, **kwargs: None
matplotlib_pyplot_module.close = lambda *args, **kwargs: None
matplotlib_module.pyplot = matplotlib_pyplot_module
sys.modules.setdefault("matplotlib", matplotlib_module)
sys.modules.setdefault("matplotlib.pyplot", matplotlib_pyplot_module)

pseudo_numpy = types.ModuleType("numpy")
pseudo_numpy.inf = float("inf")
pseudo_numpy.nan = float("nan")
sys.modules.setdefault("numpy", pseudo_numpy)

pandas_module = types.ModuleType("pandas")
pandas_module.DataFrame = DummyDataFrame
pandas_module.read_excel = lambda *args, **kwargs: DummyDataFrame([])
sys.modules.setdefault("pandas", pandas_module)

xlsxwriter_module = types.ModuleType("xlsxwriter")


class _DummyWorkbook:
    def __init__(self, *args, **kwargs):
        pass

    def add_worksheet(self, *args, **kwargs):
        return object()

    def add_format(self, fmt):
        return fmt

    def close(self):
        pass


xlsxwriter_module.Workbook = _DummyWorkbook
sys.modules.setdefault("xlsxwriter", xlsxwriter_module)

pythoncom_module = types.ModuleType("pythoncom")
pythoncom_module.CoInitialize = lambda *args, **kwargs: None
pythoncom_module.CoUninitialize = lambda *args, **kwargs: None
sys.modules.setdefault("pythoncom", pythoncom_module)

win32com_module = types.ModuleType("win32com")
win32com_client_module = types.ModuleType("win32com.client")


class _DummyDispatch:
    def __getattr__(self, name):
        return lambda *args, **kwargs: None


win32com_client_module.DispatchEx = lambda *args, **kwargs: _DummyDispatch()
win32com_module.client = win32com_client_module
sys.modules.setdefault("win32com", win32com_module)
sys.modules.setdefault("win32com.client", win32com_client_module)


from excel_processing import ExcelApp


def test_process_insert_positions_skips_when_template_has_extra_placeholders(capsys):
    app = ExcelApp()

    template_rows = [
        [{"value": "◎", "format": {}}],
        [{"value": "header", "format": {}}],
        [{"value": "◎", "format": {}}],
        [{"value": "footer", "format": {}}],
        [{"value": "end", "format": {}}],
    ]
    base_insert_positions = [0, 2]
    source_sheet_list = ["SourceA"]
    source_data = {"SourceA": DummyDataFrame([["row-data"]])}

    result_rows = app._process_insert_positions(
        template_rows,
        base_insert_positions,
        source_sheet_list,
        source_data,
        "Raw Material",
    )

    assert len(result_rows) == len(template_rows) + 1
    assert result_rows[3] == ["row-data"]

    captured = capsys.readouterr()
    assert "Raw Material" in captured.out
    assert "略過" in captured.out or "無對應來源資料" in captured.out


def test_process_insert_positions_warns_with_fallback_name(capsys):
    app = ExcelApp()

    template_rows = [
        [{"value": "◎", "format": {}}],
        [{"value": "◎", "format": {}}],
        [{"value": "end", "format": {}}],
    ]
    base_insert_positions = [0, 1]
    source_sheet_list = ["SourceA"]
    source_data = {"SourceA": DummyDataFrame([["row-data"]])}

    result_rows = app._process_insert_positions(
        template_rows,
        base_insert_positions,
        source_sheet_list,
        source_data,
        "Raw Material",
    )

    assert len(result_rows) == len(template_rows) + 1

    captured = capsys.readouterr()
    assert "Raw Material#2" in captured.out
    assert "略過" in captured.out or "無對應來源資料" in captured.out
