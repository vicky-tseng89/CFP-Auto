# CFP-Auto 碳足跡自動化專案

這個 repository 目前同時保存兩條工作線：

1. 現行的產品碳足跡 Excel / Word 自動化程式。
2. 與 AI 討論後產生的多代理資料治理與方法論方案。

為了避免程式、方案文件、輸入樣板、輸出結果和暫存檔混在一起，請以這份 README 作為專案入口，並以 [docs/project_organization_plan.md](docs/project_organization_plan.md) 作為後續整理依據。

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

## 安裝

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

也可以使用 `uv`：

```bash
uv sync
```

## 執行

```bash
python GUI_test.py
```

`main.py` 目前只作為簡單入口保留；主要 GUI 邏輯仍在 `GUI_test.py`。

## 測試

```bash
pytest
```

如果測試會產生 Excel、Word 或暫存檔，請確認輸出位置落在 `output/`、`codex_tmp/` 或 `.tmp_*`，不要新增到根目錄。

## 本機資料政策

本 workspace 採用 no-egress policy。所有公司資料、客戶檔案、Excel、Word、AI 方案與執行結果都只能在本機與本 repository 內處理，不應透過網路、雲端、Email、Issue tracker 或任何外部服務傳出。
