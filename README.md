# CFP-Auto

Excel 與 Word 自動化工具，用於產品碳足跡資料轉換、盤查計算與報告書產出。

## 功能
- Transform：將來源 Excel 轉成盤查表單格式（依模板插入資料與保留格式）。
- Process：計算各階段數值並輸出 `result_*.xlsx`、`report_*.xlsx`。
- Process All：可選擇先更新 `INPUT` 工作表，再串接 Transform + Process。
- Report：依廠區模板（竹南 / 竹北 / 越南）產出完整 Word 報告。
- Batch：可一次處理多個 Excel 檔，並輸出批次摘要 CSV。

## 系統需求
- Windows 10/11（必要，程式使用 `pywin32` 與 Excel COM）
- Microsoft Excel（桌面版）
- Python 3.10 以上（建議）

## 安裝
```bash
pip install -r requirements.txt
```

## 主要相依套件
- `pandas`, `openpyxl`, `xlsxwriter`
- `python-docx`, `docxtpl`
- `pywin32`
- `tkcalendar`

## 必要資源檔案
以下模板與資源檔案預設應放在 `resources/` 目錄：
- `resources/PLCI_table_format.xlsx`
- `resources/report_temp.xlsx`
- `resources/智邦-產品碳足跡盤查總報告書_竹南_temp.docx`
- `resources/智邦-產品碳足跡盤查總報告書_竹北_temp.docx`
- `resources/智邦-產品碳足跡盤查總報告書_越南_temp.docx`

程式會優先從 `resources/` 讀取模板；若找不到，才會回頭檢查舊版的程式同層路徑。因此：
- 開發模式：請放在專案根目錄下的 `resources/`
- 打包成 `exe`：請放在 `exe` 同層的 `resources/`
- 舊版相容：根目錄同名檔案仍可作為備援，但不建議再作為正式放置位置

## 輸入 Excel 基本要求
程式流程會使用以下工作表名稱（需一致）：
- `overview`
- `INPUT`
- `simapro10.2.0.0`
- `Raw Material`
- `Manufacturing`
- `Distribution`
- `Usage`
- `Recycling`

`simapro10.2.0.0` 需包含欄位：
- `單位對照`
- `fossil(kg CO2-eq)`
- `biogenic(kg CO2-eq)`
- `land transformation (kg CO2-eq)`
- `unit`

Transform 也會讀取以下來源工作表（名稱需一致）：
- `Raw Material(Direct Material)`
- `Raw Material(Indirect Material)`
- `Raw Material(Direct Transport)`
- `Raw Material(Indirect Transport`（依程式目前字串）
- `Manufacturing(Manufacturing)`
- `Manufacturing(Gas)`
- `Manufacturing(Electricity)`
- `Manufacturing(Transport)`
- `Manufacturing(Waste treatment)`
- `Distribution(Local)`
- `Distribution(Air)`
- `Distribution(Warehouse)`
- `Distribution(Customer)`
- `Recyling(Recyling)`（依程式目前字串）
- `Usage`

## 執行方式
```bash
python GUI_test.py
```

## GUI 使用流程
1. 在下方批次區域加入一個或多個 Excel 檔。
2. 視需求執行 `Transform`、`Process` 或 `Process All`。
3. 需要報告時，在 `Report` 頁籤選擇廠區模板後執行。

## 輸出位置
- `output/result/`：`merged_result_*.xlsx`、`result_*.xlsx`、`report_*.xlsx`
- `output/report/`：`智邦-產品碳足跡盤查總報告書_*.docx`
- `output/charts/`：中間圖表檔
- `output/tmp/`：中間 Word 檔
- `output/batch_summary_*.csv`：批次執行摘要
- `logs/excel_processing.log`：執行紀錄

## 常見問題
- 找不到模板檔：請優先確認上述必要資源檔案都放在 `resources/` 目錄。
- Excel 無法儲存或被鎖定：關閉所有 Excel 視窗後重試。
- 欄位或工作表錯誤：請確認輸入檔工作表名稱與程式要求完全一致。

## 打包（選用）
```bash
pyinstaller GUI_test.spec
```
