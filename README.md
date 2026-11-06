# 屏東飼料廠 製程與成品品質化驗日報（Q-10）線上版

網頁版製程日報：上傳 NIR 匯出 Excel 自動填入、物性試驗（粗細度、含粉率、堅牢度、比重）計算、依 R-20 製程管制標準自動判定、品質指數與本月累計平均／標準差、匯出 Excel。資料存在 Google 試算表。

## 網址
- GitHub Pages：https://dachanqa-report.github.io/pingtung-daily-report/
- Apps Script 網頁版：https://script.google.com/macros/s/AKfycbxioIpmEueosa3NqCq015aaZKePVJVxnxzTrgMBRY1B-QSLflTbUamBGQYFVxcu5JT92Q/exec

## 檔案
- `index.html`：網頁（在 Apps Script 內用 google.script.run，在 GitHub Pages 上透過 Apps Script 網址的 POST API 存取試算表）
- `apps-script/Code.gs`：Apps Script 後端（試算表讀寫、doPost API、選用的 Gemini AI 照片辨識）

## 更新方式
1. 在 Apps Script 編輯器修改程式碼.gs／index.html 後存檔
2. 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）
3. 同步把 index.html 更新到這個 repo，GitHub Pages 才會是新版
