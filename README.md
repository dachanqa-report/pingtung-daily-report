# 屏東飼料廠 製程與成品品質化驗日報（Q-10）線上版

網頁版製程日報：上傳 NIR 匯出 Excel 自動填入、物性試驗（粗細度、含粉率、堅牢度、比重）計算、依 R-20 飼料成份製程管制標準表自動判定、品質指數與本月累計平均／標準差、匯出與原日報相同格式的 Excel（屏東廠、data、Sheet1 三個分頁）。資料存在 Google 試算表。

## 網址
- GitHub Pages：https://dachanqa-report.github.io/pingtung-daily-report/
- Apps Script 網頁版：https://script.google.com/macros/s/AKfycbxioIpmEueosa3NqCq015aaZKePVJVxnxzTrgMBRY1B-QSLflTbUamBGQYFVxcu5JT92Q/exec

## 檔案
- `index.html`：網頁（在 Apps Script 內用 google.script.run，在 GitHub Pages 上透過 Apps Script 網址的 POST API 存取試算表）
- `Code.gs`：Apps Script 後端（日報讀寫、標準表版本、doPost API、選用的 Gemini AI 照片辨識）
- `q10-template.xlsx`：日報空白範本，由「屏東廠製程成品日報1151117.xlsx」清除當日數據而成，含三個分頁：
  - 屏東廠：Q-10 日報（標準用 VLOOKUP 對 Sheet1，累計平均值／標準差指向 data 分頁）
  - data：本月明細，每個品項依當月樣品數自動展開列數
  - Sheet1：飼料成份製程管制標準表（網頁預設標準，115/11/03 版）

## 網頁功能
- 上傳已完成的日報：讀第一頁當天數據；data 分頁的本月其他日期樣品可補進系統；檔案內標準表可一併更新
- 上傳飼料成份製程管制標準表：存到試算表「標準表版本」，之後判定、匯出的 Sheet1 和第一頁右下角日期都用新版

## 更新方式
1. 在 Apps Script 編輯器修改程式碼.gs／index.html 後存檔
2. 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）
3. 同步把 index.html 更新到這個 repo，GitHub Pages 才會是新版
