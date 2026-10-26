# RTU 內網產報

網站可用於本機測試；Excel 排程仍需在實際活頁簿上驗證，尚未提供排程註冊與逾時監控。

原始三個 VBA 檔保持不變。排程只使用 `automation/Run-ExcelMacros.ps1`；舊的 VBA 產生工具已移除。腳本不產生、修改或匯入 VBA，你自行撰寫活頁簿中的 macro。

網站以 `RTU_REPORT_ROOT` 環境變數指定報表根目錄。日期目錄採 `yyyy-MM-dd`，內含 `main.html`、`no-reason.html`、`summary.html` 及代表整組已完成的 `complete` 檔案。網站只顯示完成的最近七個日期。

## 啟動網站

在專案根目錄執行：

```powershell
$env:RTU_REPORT_ROOT = Join-Path $PWD 'reports'
dotnet run --project .\RtuWeb --no-launch-profile --urls http://localhost:5080
```

瀏覽 `http://localhost:5080`。在首頁日期卡片選擇要看的報表，上方提供「首頁／主報表／當日彙總／無原因」切換，不提供日期下拉選單。報表在同一分頁的 iframe 顯示，切換種類會保留選定日期；查看其他日期時先返回首頁選取卡片。報表原始 HTML 與其互動不修改。內網測試可將啟動網址改成 `http://0.0.0.0:5080`，並依主機防火牆設定開放連線。

## 建立昨日資料夾並執行 Excel 巨集

```powershell
.\automation\Run-ExcelMacros.ps1 `
    -ReportRoot 'C:\RTU\reports' `
    -WorkbookPath 'C:\RTU\報表工具.xlsm'
```

`ReportRoot` 與 `WorkbookPath` 均需填寫實際值。`ReportRoot` 應與網站的 `RTU_REPORT_ROOT` 一致。腳本內的 `$sourceRTU`、`$sourceOHTM`、`$sourceWendy` 也要改成實際來源資料夾。三個來源各取最後修改時間最新的 CSV／XLSX，忽略 Excel 暫存檔。

腳本依台灣時間建立昨天的 `yyyy-MM-dd` 子資料夾並移除舊 `complete`。依序將三份來源路徑傳給 `Module1.ImportSource1`、`Module1.ImportSource2`、`Module1.ImportSource3`，再將輸出資料夾傳給 `Module1.GenerateReports`。這些是你自行撰寫的 macro，請在腳本內替換成實際名稱。

三份 HTML 全部存在且非空後，儲存並關閉 Excel，再建立 `complete`。之後清理三個來源資料夾（CSV／XLSX 合計各保留最新七份），全部成功後最後通知 C#。只清理舊來源檔，不刪除報表日期資料夾。macro 或報表檢查失敗不建立完成標記、不清理、不通知；來源清理失敗則保留已完成報表，但不發送更新事件。巨集自己儲存的檔案不會被回復。

若本機顯示「禁止執行指令碼」，可用單次處理程序執行，不更改永久設定：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\automation\Run-ExcelMacros.ps1 `
    -ReportRoot 'C:\RTU\reports' `
    -WorkbookPath 'C:\RTU\報表工具.xlsm'
```

巨集需已存在於活頁簿，Excel 必須允許執行。Excel 預設隱藏，手動除錯可在腳本參數加上 `-ShowExcel`。舊的 `-Hidden` 仍可使用，若同時指定 `-ShowExcel` 與 `-Hidden`，以隱藏為準。隱藏主視窗不會消除 macro 的選檔與 `MsgBox`，用於排程的 macro 需自行避免互動。

## macro 成功契約與 Windows 排程

`GenerateReports` 接收 PowerShell 傳入的輸出資料夾，將三份 HTML 完整寫入該位置後返回。不需要回傳值或建立 `success.flag`。PowerShell 確認檔案、儲存並關閉 Excel 後才建立 `complete`。例如完成後目錄應為：

```text
C:\RTU\reports\2026-10-25\
  main.html
  no-reason.html
  summary.html
  complete
```

只有成功產報後才清理來源檔案；報表日期資料夾不清理，網站清單仍只顯示最近七個有完成標記的日期。

Windows 工作排程器的程式設為 `powershell.exe`，引數填寫（路徑換成實際值）：

```text
-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\RTU\automation\Run-ExcelMacros.ps1" -ReportRoot "C:\RTU\reports" -WorkbookPath "C:\RTU\報表工具.xlsm" -NotifyUrl "http://localhost:5080/api/reports/refresh"
```

排程採「只有使用者登入時才執行」並設定不啟動並行執行個體，觸發時間由排程器設定。`-WindowStyle Hidden` 是 `powershell.exe` 的啟動參數，用來隱藏 PowerShell 視窗；Excel 視窗則由腳本控制。正常完成沒有終端訊息，不建立 log，結束碼為 `0`；失敗結束碼為 `1`。

## 錯誤紀錄

只有失敗時才建立腳本旁的 `logs` 資料夾，追加寫入 `logs\error-yyyy-MM-dd.log`。檔名日期以錯誤發生時的台灣日期為準，不一定等於報表日期。每行是一筆 UTF-8 JSON 紀錄，包含 `time`、`stage`、`reportDate`、活頁簿與相關檔案路徑、錯誤訊息、例外及呼叫位置；`cleanupErrors` 另記錄關閉 Excel／釋放物件時的附加錯誤，不覆蓋原始失敗原因。

例如來源路徑錯誤為 `ReadSourceRTU`、macro 執行錯誤為 `ImportRTU`／`GenerateReports`、來源清理錯誤為 `CleanSourceFiles`、通知失敗為 `NotifyWebsite`。錯誤只寫 log，不另輸出至終端；如果 log 本身無法寫入，才將記錄失敗原因、原始錯誤與附加關閉錯誤送到標準錯誤串流。腳本不會跳出自訂訊息視窗。

網站應先持續啟動。`NotifyUrl` 預設為上述本機網址。PowerShell 在清理完成後以 `application/json` POST 傳送 `{"reportDate":"2026-10-25"}`（實際使用執行日的昨天）。C# 只接受 loopback 呼叫，嚴格驗證指定日期資料夾的 `complete` 與三份非空 HTML；成功回傳 `200` 與 `reportDate`，才透過 SSE 推送更新。缺少／錯誤日期回傳 `400`、日期資料夾不存在 `404`、缺旗標或報表缺少／空白 `409`、檔案暫時無法讀取 `503`；失敗不發送事件。不會改驗證其他日期或用舊報表充當本次成功。

瀏覽器透過 `/api/reports/events` 更新日期卡片並重新載入選取報表。連線初次建立及重新連線會同步清單，沒有定時輪詢或資料夾監看。若主機上有反向代理，不可將 refresh 端點對外轉送（代理本機連入也會被視為 loopback）。

若 error log 的步驟為 `NotifyWebsite`，訊息會明確記錄報表及來源清理已成功，只需重試通知，不要重跑 macro。修正問題後可單獨通知，日期須填要刷新的實際報表日期：

```powershell
Invoke-RestMethod -Method Post -Uri 'http://localhost:5080/api/reports/refresh' `
    -ContentType 'application/json' -Body '{"reportDate":"2026-10-25"}'
```

## 手動放入報表

將三份 HTML 複製到日期資料夾並命名為 `main.html`、`no-reason.html`、`summary.html`。確認齊全後再建立空白標記，例如：

```powershell
New-Item -ItemType File -Path 'C:\RTU\reports\2026-10-23\complete'
```

手動測試時，`complete` 沒有副檔名，放好標記後重新整理首頁，或執行上述 POST 通知已開啟的頁面。正式排程會自動建立 `complete` 並通知。若個別 HTML 不存在，切換到該種類時會顯示缺檔提示。

## 複製當日彙總到 Notes

首頁每張日期卡片上的「複製當日彙總」只讀取該日期的 `summary.html`。複製範圍限 `#main-view` 的分類：大項去除末尾數量、所有細項保留數量，僅 `others` 保留其下所有 `.row` 明細。不複製一般細項明細、隱藏的 `.row-detail`，或其他報表分頁。

剪貼簿同時提供 HTML 格式與純文字；HTML 使用報表本身的字型、字級、顏色、粗體、換行與段落樣式。貼到 Notes 文件的富文字欄位，使用一般貼上。若瀏覽器限制自動複製，網站會顯示預覽及再次複製按鈕，也可選取預覽內容按 Ctrl+C。Notes 最終呈現仍需在實際客戶端確認。

## 驗證

`dotnet build RtuWeb` 建置網站。`powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-ExcelRunner.ps1` 使用模擬 Excel 驗證資料夾與呼叫流程，不會執行真實巨集。`python tests/test_web.py` 驗證 HTTP 及瀏覽器互動，使用已安裝的 Chrome／Edge，不另裝套件。實際巨集、產出報表的圖表／按鈕、鎖定畫面及內網連線仍需公司環境驗收。
