# CFP-Auto 碳足跡自動化專案

產品碳足跡盤查 Excel 轉換、碳排計算與 AUP 報告書產生工具。

此工具以 Tkinter GUI 操作，將 Accton 原始 Excel 表單轉成 PLCI 盤查表格式，依 SimaPro 對照表計算各階段碳排，並可套用 Word 範本產出完整產品碳足跡盤查報告書。

目前版本：`1.1.8`

這個 repository 目前同時保存兩條工作線：

1. 現行的產品碳足跡 Excel / Word 自動化程式。
2. 與 AI 討論後產生的多代理資料治理與方法論方案。

為了避免程式、方案文件、輸入樣板、輸出結果和暫存檔混在一起，請以這份 README 作為專案入口，並以 [docs/project_organization_plan.md](docs/project_organization_plan.md) 作為後續整理依據。

## 主要功能

- `轉換格式`：將原始 Excel 的各分類工作表整併到 `PLCI_table_format.xlsx` 模板，輸出 `merged_result_*.xlsx`。
- `處理數據`：計算 Raw Material、Manufacturing、Distribution、Usage、Recycling 各階段碳排，輸出 `result_*.xlsx` 與 `report_*.xlsx`。
- `完整處理`：依序執行 `轉換格式` 與 `處理數據`。
- `完整報告書生成`：讀取已處理盤查表單，套用竹北、竹南或越南 Word 範本，輸出完整報告書。
- 批次匯入：可一次加入多個 Excel 檔案，批次執行轉換、處理或完整流程。
- INPUT 重新整理：可輸入一個或多個產品 F 階機種，自動更新來源檔 `INPUT!B1:B3` 並重新整理 Excel 連線與公式。
- 運輸距離計算：可依起點、終點與運輸方式補算距離與 ton-km，支援本地對照表優先、快取與強制重新計算。
- 進度、取消與錯誤追蹤：長時間作業會顯示進度視窗，錯誤會寫入 `logs/excel_processing.log`。

## 快速入口

- 執行 GUI：`python GUI_test.py`
- 主處理邏輯：`excel_processing.py`
- 運輸距離計算：`transport_distance.py`
- AI 多代理方案：[docs/ai/multiple_agents_carbon_methodology.md](docs/ai/multiple_agents_carbon_methodology.md)
- 專案整理藍圖：[docs/project_organization_plan.md](docs/project_organization_plan.md)
- 舊版自動化副本：[CFP-Auto/](CFP-Auto/)

## 目前目錄定位

| 路徑 | 內容定位 |
| --- | --- |
| `GUI_test.py`, `excel_processing.py`, `transport_distance.py`, `main.py` | 現行可執行程式入口與核心流程。暫時保留在根目錄，避免影響既有執行方式。 |
| `carbon_model/` | 產品碳足跡 canonical data model。 |
| `adapters/` | PLCI、PACT、客戶 Excel 等輸出格式轉換。 |
| `agents/` | AI 多代理角色的本機 deterministic workflow。 |
| `reviews/` | ISO 14067 checklist、change review、PDCA report 等檢核模組。 |
| `tests/` | pytest 測試。 |
| `resources/` | 程式執行需要的樣板、係數表、地點對照表等輸入資源。 |
| `resources/reference/` | 欄位對照、來源說明、資料字典等參考文件。 |
| `docs/` | 專案文件、整理方案、使用手冊與圖片素材。 |
| `docs/ai/` | AI 方案、AI 執行方法論、多代理設計。 |
| `docs/manuals/` | 使用說明書等正式文件。 |
| `docs/assets/` | 文件用圖片，例如系統流程圖。 |
| `output/` | 程式產生的客戶結果、暫存輸出與批次結果。 |
| `報告/` | AUP、查證、平台簡介等正式報告材料。 |
| `之前記錄/` | 歷史資料、舊快照、debug snapshot，不作為現行程式來源。 |
| `CFP-Auto/` | 舊版或獨立副本，用於比對與備查，不是主要工作入口。 |
| `codex_tmp/`, `.tmp_*`, `logs/` | 本機暫存、驗證或執行紀錄。 |

## 新檔案放置原則

- 新的 AI 討論方案、方法論、執行計畫：放在 `docs/ai/`。
- 專案整理規則、架構說明、維護指南：放在 `docs/`。
- 程式執行必要樣板與對照表：放在 `resources/`；純參考表放在 `resources/reference/`。
- 客戶或批次輸出結果：放在 `output/<客戶或任務名稱>/`，臨時輸出放在 `output/tmp/`。
- 正式報告、查證文件、平台簡介：放在 `報告/` 或 `docs/manuals/`。
- 暫存測試、AI scratch、一次性驗證資料：放在 `codex_tmp/` 或 `.tmp_*`，不要放根目錄。
- 舊版程式與歷史快照：放在 `之前記錄/` 或保留於 `CFP-Auto/`，並在文件中標註用途。

根目錄只保留專案入口、設定檔、主要執行檔與目前仍需維持相容的核心模組。

## 執行環境

- Windows 10/11
- Microsoft Excel 桌面版
- Python 3.11 以上
- 建議在可連線網路的環境執行運輸距離計算；若環境無法連線，可關閉 `執行距離計算`

本工具使用 `pywin32` 控制 Excel COM，因此需要安裝 Microsoft Excel，且執行時請避免手動開啟正在處理的檔案。

## 安裝

建立虛擬環境後安裝依賴：

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

若使用 `uv`：

```bash
uv sync
```

## 啟動 GUI

```bash
python GUI_test.py
```

`main.py` 目前只作為簡單入口保留；主要 GUI 邏輯仍在 `GUI_test.py`。

## 必要資源檔

請確認下列檔案存在於 `resources/`：

- `PLCI_table_format.xlsx`
- `report_temp.xlsx`
- `airport_port_land_location_mapping.xlsx`
- `智邦-產品碳足跡盤查總報告書_竹北_temp.docx`
- `智邦-產品碳足跡盤查總報告書_竹南_temp.docx`
- `智邦-產品碳足跡盤查總報告書_越南_temp.docx`

程式會優先從 `resources/` 讀取模板與對照表。打包成 exe 時，請將 `resources/` 與 `VERSION` 放在 exe 同層目錄。

## 輸入 Excel 格式

來源 Excel 建議包含下列工作表：

- `overview`
- `INPUT`
- `simapro10.2.0.0`
- `Raw Material(Direct Material)`
- `Raw Material(Indirect Material)`
- `Raw Material(Direct Transport)`
- `Raw Material(Indirect Transport`
- `Manufacturing(Manufacturing)`
- `Manufacturing(Gas)`
- `Manufacturing(Electricity)`
- `Manufacturing(Transport)`
- `Manufacturing(Waste treatment)`
- `Distribution(Local)`
- `Distribution(Air)`
- `Distribution(Warehouse)`
- `Distribution(Customer)`
- `Usage`
- `Recyling(Recyling)`

`simapro10.2.0.0` 需包含下列欄位：

- `單位對照`
- `unspecified(kg CO2-eq)`
- `fossil(kg CO2-eq)`
- `biogenic(kg CO2-eq)`
- `land transformation (kg CO2-eq)`
- `unit`

報告書生成的來源檔必須是已處理盤查表單，至少包含：

- `overview`
- `Raw Material`
- `Manufacturing`
- `Distribution`
- `Usage`
- `Recycling`

## GUI 使用方式

1. 執行 `python GUI_test.py`。
2. 使用下方 `批次匯入檔案` 區塊加入一個或多個 Excel 檔案。
3. 選擇功能分頁：
   - `轉換格式`：輸入廠區、產品 F 階機種與日期後執行格式轉換。
   - `處理數據`：選擇要計算的碳排階段，可勾選或取消 `執行距離計算`、`使用快取 / 本地對照表優先`、`重新計算所有距離`。
   - `完整處理`：一次完成轉換與計算，並可選擇碳邊界與距離計算選項。
   - `完整報告書生成`：選擇已處理盤查表單與區域後產生 Word 報告書。
4. 若要重新整理來源檔 `INPUT`，請勾選 `啟用重新整理功能`，再輸入產品 F 階機種與起訖日期。
5. 作業中可關閉進度視窗來取消目前批次；後續未處理項目會在批次摘要中標示 skipped。

### 產品 F 階機種規則

- 一行輸入一個機種。
- 啟用重新整理時，每個機種去除空白後必須為 13 個字元。
- 若輸入多個機種，必須先勾選 `啟用重新整理功能`。
- 批次處理時，程式會針對每個來源檔與每個機種建立一個處理工作。

## 輸出位置

程式會自動建立 `output/` 與 `logs/`：

- `output/result/merged_result_*.xlsx`：轉換格式輸出。
- `output/result/result_*.xlsx`：碳排計算後的盤查表單。
- `output/result/batch_summary_*.csv`：多檔或多機種批次摘要。
- `output/report/report_*.xlsx`：各階段彙總用 Excel 報表。
- `output/report/智邦-產品碳足跡盤查總報告書_*.docx`：完整 Word 報告書。
- `output/charts/`：報告書圖表暫存。
- `output/tmp/`：重新整理來源檔與 Word 生成暫存檔。
- `logs/excel_processing.log`：執行紀錄與錯誤訊息。
- `logs/cache/transport_distance_cache.json`：運輸距離與端點查詢快取。

## 運輸距離計算

`處理數據` 與 `完整處理` 預設會執行距離計算：

- Raw Material 會處理第 3、4 個運輸表格。
- Distribution 會處理所有可辨識的運輸表格。
- Road 類型會優先使用 `airport_port_land_location_mapping.xlsx` 的對照資料。
- Air 與 Sea 會依地點查詢與內建路線模型估算距離。
- 若既有 `distance transported (km)` 已有非零值，程式會沿用該距離並更新 ton-km。
- 勾選 `使用快取 / 本地對照表優先` 時，會優先使用本地對照表與 persistent cache。
- 勾選 `重新計算所有距離` 時，會忽略既有距離與可用快取，重新計算可辨識路線。

若來源資料缺少 `starting point`、`end point`、`type of transport` 或 `distance transported (km)` 等必要欄位，該階段可能會失敗。

### 距離計算模組

`transport_distance.py` 提供運輸距離計算核心，可供 GUI 流程、Excel 處理流程或命令列測試共用：

- Road、Driving、Walking、Cycling 會透過 OSRM route service 取得路線距離與幾何線段。
- Air 會以大圓航線估算航空距離，並依設定的分段上限產生路徑節點。
- Sea 會用內建航運 waypoint 網路估算可航行路線，避免只用直線距離低估海運路徑。
- 地點文字查詢會依序嘗試 Nominatim、Photon、ArcGIS 等 geocoding 服務；若直接提供座標，Air 與 Sea 可離線估算。
- 回傳結果包含距離、時間、GeoJSON LineString、分段資訊、查詢 URL 與 metadata，方便寫回 Excel 或除錯。

命令列範例：

```bash
python transport_distance.py --mode driving --from-lat 24.8138 --from-lon 120.9675 --to-lat 25.0330 --to-lon 121.5654
```

## 版本資訊

GUI 顯示的版本依下列優先順序取得：

1. 環境變數 `CFP_AUTO_VERSION`
2. 專案根目錄 `VERSION`
3. `git describe --tags --always --dirty`
4. 預設值 `0.0.0`

目前 `VERSION` 檔內容為 `1.1.8`。

## 測試

```bash
pytest
```

目前測試包含 Excel 報告處理邏輯、GUI 匯出摘要，以及專案內其他實驗性模組的單元測試。

如果測試會產生 Excel、Word 或暫存檔，請確認輸出位置落在 `output/`、`codex_tmp/` 或 `.tmp_*`，不要新增到根目錄。

## 常見問題

### 找不到模板或報告範本

確認 `resources/` 中是否有必要資源檔。錯誤訊息會列出程式實際檢查過的路徑。

### Excel 開啟或儲存失敗

- 關閉正在處理的 Excel 檔案。
- 確認沒有其他 Excel 視窗卡住對話框。
- 可用 GUI 右側的 `Excel X` 按鈕關閉所有 Excel 程序後重試。

### 重新整理 INPUT 失敗

- 確認來源檔存在 `INPUT` 工作表。
- 確認產品 F 階機種去除空白後為 13 碼。
- 若 Excel 連線需要權限或網路，請先在 Excel 中確認可正常重新整理。

### 報告書生成失敗

- 請使用 `處理數據` 或 `完整處理` 產生的 `result_*.xlsx`。
- 確認來源檔含有 `overview`、`Raw Material`、`Manufacturing`、`Distribution`、`Usage`、`Recycling`。
- 查看 `logs/excel_processing.log` 中對應的 `run_id` 與錯誤內容。

### 運輸距離查詢失敗或很慢

- 檢查網路連線。
- 確認起點、終點資料可被地理查詢服務辨識。
- 若只要完成碳排計算，可在 GUI 取消 `執行距離計算`。

## 打包提醒

若使用 PyInstaller 打包，請確保下列項目會被複製到 exe 同層：

- `resources/`
- `VERSION`

打包後程式會以 exe 所在資料夾作為基準路徑，並在該位置建立 `output/`、`logs/` 與必要暫存目錄。

## 本機資料政策

本 workspace 採用 no-egress policy。所有公司資料、客戶檔案、Excel、Word、AI 方案與執行結果都只能在本機與本 repository 內處理，不應透過網路、雲端、Email、Issue tracker 或任何外部服務傳出。
