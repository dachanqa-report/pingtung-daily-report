/**
 * 屏東飼料廠 製程與成品品質化驗日報（Q-10）線上版 — Google Apps Script 後端
 *
 * 試算表裡會有三個工作表：
 *   日報資料：每天一列（日期、更新時間、品質指數、JSON），網頁讀寫用，請勿手動修改 JSON 欄
 *   明細：每天每個品項一列，方便在試算表裡查詢、篩選、做圖
 *   使用說明：欄位說明
 *
 * 選用：AI 照片辨識需要在「專案設定 → 指令碼屬性」新增 GEMINI_API_KEY
 *       （可再加 GEMINI_MODEL 指定模型，預設 gemini-2.5-flash）
 */

const SHEET_DAYS = '日報資料';
const SHEET_ROWS = '明細';
const SS_ID = '1iZHA71p9rnmldWmxJ9L48DyO3Aw069j8Ry2xHQkiyv4'; // 屏東廠製程日報 試算表

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SS_ID);
}

function doGet(e) {
  // ?raw=index 回傳網頁原始碼（給 GitHub 備份用）
  if (e && e.parameter && e.parameter.raw === 'index') {
    return ContentService.createTextOutput(HtmlService.createHtmlOutputFromFile('index').getContent())
      .setMimeType(ContentService.MimeType.TEXT);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('屏東廠製程日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** 第一次使用時在編輯器執行一次：建立工作表與標題列 */
function setup() {
  const ss = ss_();
  const ds = daysSheet_();
  rowsSheet_();
  // 匯入 11/03 原始日報當第一筆資料（只在空白時）
  if (ds.getLastRow() < 2) ds.appendRow(['2026-11-03', new Date(), 24, 25, 24 / 25, SEED_1103]);
  let help = ss.getSheetByName('使用說明');
  if (!help) {
    help = ss.insertSheet('使用說明');
    help.getRange(1, 1, 6, 1).setValues([
      ['屏東廠製程日報線上版'],
      ['網頁按「儲存本日」後，資料寫入「日報資料」（網頁用）和「明細」（查詢用）。'],
      ['同一天重複儲存會覆蓋該日資料。'],
      ['「日報資料」的 JSON 欄請勿手動修改，要改數據請回網頁改。'],
      ['AI 照片辨識：專案設定 → 指令碼屬性 → 新增 GEMINI_API_KEY。'],
      ['網址：部署 → 管理部署作業 → 網頁應用程式網址。'],
    ]);
    help.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    help.setColumnWidth(1, 640);
  }
  const first = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  return 'ok';
}

function daysSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_DAYS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DAYS);
    sh.getRange(1, 1, 1, 6).setValues([['日期', '更新時間', '符合項數', '抽檢項數', '品質指數', 'JSON']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('0.0%');
    sh.setColumnWidth(6, 300);
  }
  return sh;
}

function rowsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_ROWS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ROWS);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

/** 網頁讀取：回傳每天的 JSON 字串 */
function listDays() {
  const sh = daysSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 6, n, 1).getValues().map(r => String(r[0])).filter(Boolean);
}

/** 網頁儲存：覆寫該日資料與明細 */
function saveDay(date, json, header, rows, idx) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日期格式錯誤');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // 日報資料
    const sh = daysSheet_();
    const last = sh.getLastRow();
    const dates = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]) : [];
    const i = dates.indexOf(date);
    const rate = idx && idx.n ? idx.ok / idx.n : '';
    const row = [date, new Date(), idx ? idx.ok : '', idx ? idx.n : '', rate, json];
    if (i >= 0) sh.getRange(i + 2, 1, 1, 6).setValues([row]);
    else sh.appendRow(row);
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).sort({ column: 1, ascending: true });

    // 明細：刪掉該日舊列再寫入
    const rs = rowsSheet_();
    const head = ['日期'].concat(header);
    if (rs.getLastRow() === 0 || rs.getRange(1, 1).getValue() !== '日期') {
      rs.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    }
    const rl = rs.getLastRow();
    if (rl > 1) {
      const col = rs.getRange(2, 1, rl - 1, 1).getDisplayValues();
      for (let r = col.length - 1; r >= 0; r--) {
        if (col[r][0] === date) rs.deleteRow(r + 2);
      }
    }
    const data = rows.map(r => [date].concat(r));
    if (data.length) {
      const width = head.length;
      const fixed = data.map(r => { const x = r.slice(0, width); while (x.length < width) x.push(''); return x; });
      rs.getRange(rs.getLastRow() + 1, 1, fixed.length, width).setValues(fixed);
      const n = rs.getLastRow() - 1;
      rs.getRange(2, 1, n, width).sort({ column: 1, ascending: true });
    }
    return 'ok';
  } finally {
    lock.releaseLock();
  }
}

/** 外部網頁（GitHub Pages）呼叫用的資料 API：POST {fn, args} → {ok, data | error} */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fns = { listDays: listDays, saveDay: saveDay, aiEnabled: aiEnabled, aiExtract: aiExtract };
    if (!fns[req.fn]) throw new Error('不支援的功能：' + req.fn);
    out = { ok: true, data: fns[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** AI 是否已設定 */
function aiEnabled() {
  return !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

/** AI 照片辨識（Gemini） */
function aiExtract(prompt, images) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('尚未設定 GEMINI_API_KEY');
  const model = props.getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  const parts = [{ text: prompt }].concat((images || []).map(im => ({ inline_data: { mime_type: im.mime, data: im.data } })));
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify({ contents: [{ parts: parts }], generationConfig: { responseMimeType: 'application/json', temperature: 0 } }),
      muteHttpExceptions: true,
    }
  );
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');
  if (code !== 200) throw new Error('AI 服務錯誤 ' + code + '：' + ((body.error && body.error.message) || ''));
  const c = body.candidates && body.candidates[0];
  const text = c && c.content && c.content.parts ? c.content.parts.map(p => p.text || '').join('') : '';
  return text || '[]';
}

/** 115/11/03 日報（由 屏東廠製程日報1151103.xlsx 匯入） */
const SEED_1103 = "{\"date\":\"2026-11-03\",\"entries\":{\"n2\":{\"mfg\":\"11.03\",\"qc\":\"C302581\",\"item\":\"112桶\",\"extra\":true,\"M\":11.8,\"CP\":17.3,\"F\":5.3,\"p\":{\"s6\":1.62,\"s610\":29.95,\"s35\":11.41}},\"n3\":{\"mfg\":\"11.03\",\"qc\":\"C302582\",\"item\":\"106桶\",\"extra\":true,\"M\":11.7,\"CP\":16.4,\"F\":4,\"p\":{\"s6\":1.27,\"s610\":28.02,\"s35\":10.56}},\"n4\":{\"mfg\":\"11.03\",\"qc\":\"C302583\",\"item\":\"107桶\",\"extra\":true,\"M\":11.9,\"CP\":16.9,\"F\":5.5,\"p\":{\"s6\":1.59,\"s610\":27.58,\"s35\":9.72,\"bd\":0.64}},\"n5\":{\"mfg\":\"11.03\",\"qc\":\"C302584\",\"item\":\"99桶\",\"extra\":false,\"M\":12.2,\"CP\":19.5,\"F\":8.8,\"p\":{\"fines\":1.95,\"pdi\":97.64}},\"n7\":{\"mfg\":\"11.03\",\"qc\":\"C302585\",\"item\":\"132桶\",\"extra\":false,\"M\":11.6,\"CP\":19.3,\"F\":9.2,\"p\":{\"fines\":1.72,\"pdi\":97.98}},\"n8\":{\"mfg\":\"11.03\",\"qc\":\"C302586\",\"item\":\"154桶\",\"extra\":false,\"M\":12.4,\"CP\":17.2,\"F\":8.5,\"p\":{\"fines\":2.52,\"pdi\":98.12}},\"n9\":{\"mfg\":\"11.03\",\"qc\":\"C302587\",\"item\":\"147桶\",\"extra\":false,\"M\":10.7,\"CP\":21.5,\"F\":6.4,\"p\":{\"sA\":1.95,\"sB\":8.1}},\"n15\":{\"mfg\":\"11.03\",\"qc\":\"C302588\",\"item\":\"156桶\",\"extra\":false,\"M\":11.6,\"CP\":18.2,\"F\":7.98,\"p\":{\"fines\":1.35,\"pdi\":98.41}},\"n1\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n6\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n10\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n11\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n12\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n13\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n14\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n16\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n17\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n18\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n19\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n20\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n21\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n22\":{\"ns\":true,\"nsWhy\":\"未說明原因\"},\"n23\":{\"ns\":true,\"nsWhy\":\"未說明原因\"}},\"extras\":[{\"id\":\"x1\",\"code\":\"11052\",\"mfg\":\"11.03\",\"qc\":\"C302589\",\"item\":\"2號機\",\"extra\":true,\"M\":10.8,\"CP\":17.8,\"F\":8.8,\"p\":{\"fines\":1.8,\"pdi\":97.62}},{\"id\":\"x2\",\"code\":\"11082\",\"mfg\":\"11.03\",\"qc\":\"C302590\",\"item\":\"1號機\",\"extra\":true,\"M\":10.6,\"CP\":18.5,\"F\":8.1,\"p\":{\"fines\":2.72,\"pdi\":98.67}}],\"updated\":\"2026-11-03T17:00:00+08:00\",\"source\":\"由 屏東廠製程日報1151103.xlsx 匯入\"}";
