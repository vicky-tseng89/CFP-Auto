# CFP-Auto
Carbon Footprint Automation

## 技術分析報告
整體而言， GUI 、 ProgressBarWindow 和 ExcelApp 三個類別形成了本系統的核心結構。GUI 持有
ExcelApp 與 ProgressBarWindow，並負責協調兩者：開始處理時建立進度視窗並將自己與 ExcelApp 聯繫起
來（設定回呼）；處理過程中 ExcelApp 不斷透過回呼要求 ProgressBarWindow 更新UI；處理結束後 GUI 關
閉進度視窗並提示使用者。透過類別方法的呼叫與回呼串連，系統達成了前端即時反映後端進程的效果。在工
程實作上，類別和方法名稱清晰地表達了各自職責，例如 ExcelApp.process_file /
transform_sheet / generate_report 對應不同任務階段，GUI 的 run_process_all 將多個步驟
組合執行 等。這種模組化結構讓每個部分的維護和理解相對獨立，同時又藉由明確的介面（方法與回呼）
形成有機的整體，滿足了應用需求。

## 使用方式
編輯 config.yaml 或用 GUI 介面載入 Excel 檔。

按下「開始處理」按鈕。

在 output/ 資料夾裡取回生成的 Word 報告。

功能清單：
1.  依據 SimaPro 9.3 版格式自動轉換 Excel raw data
2.  自動計算損壞百分比並內嵌到 Word 範本
3.  Tkinter GUI，支援進度條顯示

# Download
```bash
git clone https://github.com/vicky-tseng89/CFP-Auto.git
cd CFP-Auto
pip install -r requirements.txt

