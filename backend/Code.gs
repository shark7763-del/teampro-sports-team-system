/****************************************************************
 * TeamPro 智慧運動團隊管理系統 — Google Apps Script 後端
 * 用途：讓「選手用自己手機回報」「教練後台看誰沒回報」「家長查閱」
 *
 * 部署方式（只做一次）：
 *  1. 到 https://script.google.com 建立新專案，貼上本檔內容
 *  2. 上方「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *     - 執行身分：我（你自己）
 *     - 具有存取權的使用者：「任何人」
 *  3. 部署後複製 /exec 結尾的網址
 *  4. 回到 TeamPro → 設定 → 線上回報 → 貼上這個網址 → 啟用
 *
 *  資料會自動建立在本專案綁定的試算表（第一次執行會自動建表）。
 *  若要指定試算表，把下方 SHEET_ID 填成你的 Google Sheet ID。
 ****************************************************************/

const SHEET_ID = ''; // 留空＝用容器試算表 / 自動建立；或填入指定 Sheet 的 ID

function getSS(){
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  // 獨立部署時自動建立一個試算表並記住
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SS_ID');
  if (!id){
    const ss = SpreadsheetApp.create('TeamPro 線上回報資料');
    id = ss.getId();
    props.setProperty('SS_ID', id);
  }
  return SpreadsheetApp.openById(id);
}

function sheetOf(name, headers){
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh){ sh = ss.insertSheet(name); sh.appendRow(headers); }
  return sh;
}
const ROSTER_HEADERS  = ['teamCode','teamName','studentName','updatedAt'];
const REPORT_HEADERS  = ['createdAt','teamCode','date','studentName','attendance','body','sleep','fatigue','injury','injuryPart','water','kpiAvg','light','note'];

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 統一日期格式：Sheet 可能把 'YYYY-MM-DD' 自動轉成 Date，讀回來要正規化
function normDate(v){
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

/* ---------- GET：讀取（名單／回報狀況／單一學生）---------- */
function doGet(e){
  try{
    const p = e.parameter || {};
    const action = p.action || '';
    if (action === 'getRoster')  return json(getRoster(p.team));
    if (action === 'getStatus')  return json(getStatus(p.team, p.date));
    if (action === 'getStudent') return json(getStudent(p.team, p.name));
    return json({ ok:true, msg:'TeamPro API is running' });
  }catch(err){ return json({ ok:false, error:String(err) }); }
}

/* ---------- POST：寫入（儲存名單／提交回報）---------- */
function doPost(e){
  try{
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = body.action || '';
    if (action === 'saveRoster')    return json(saveRoster(body));
    if (action === 'submitReport')  return json(submitReport(body));
    return json({ ok:false, msg:'unknown action' });
  }catch(err){ return json({ ok:false, error:String(err) }); }
}

/* ---------- 名單 ---------- */
function saveRoster(b){
  const sh = sheetOf('Roster', ROSTER_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--){
    if (data[i][0] === b.teamCode) sh.deleteRow(i + 1);  // 先清掉舊名單
  }
  (b.students || []).forEach(name => sh.appendRow([b.teamCode, b.teamName || '', name, new Date()]));
  return { ok:true, team:b.teamCode, count:(b.students||[]).length };
}
function getRoster(team){
  const sh = sheetOf('Roster', ROSTER_HEADERS);
  const data = sh.getDataRange().getValues();
  const students = []; let teamName = '';
  for (let i = 1; i < data.length; i++){
    if (data[i][0] === team){ students.push(data[i][2]); teamName = data[i][1]; }
  }
  return { ok:true, team, teamName, students };
}

/* ---------- 回報 ---------- */
function submitReport(b){
  const sh = sheetOf('Reports', REPORT_HEADERS);
  sh.appendRow([ new Date(), b.teamCode, b.date, b.studentName, b.attendance, b.body,
    b.sleep, b.fatigue, b.injury, b.injuryPart, b.water, b.kpiAvg, b.light, b.note ]);
  return { ok:true };
}

/* ---------- 教練後台：今日誰回報、誰沒回報 ---------- */
function getStatus(team, date){
  const roster = getRoster(team).students;
  const sh = sheetOf('Reports', REPORT_HEADERS);
  const data = sh.getDataRange().getValues();
  const reportedMap = {};       // 以「最後一筆」為準
  const reports = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    if (row[1] === team && normDate(row[2]) === String(date)){
      reportedMap[row[3]] = { name:row[3], attendance:row[4], body:row[5],
        sleep:row[6], fatigue:row[7], injury:row[8], water:row[10], kpiAvg:row[11], light:row[12] };
    }
  }
  Object.keys(reportedMap).forEach(k => reports.push(reportedMap[k]));
  const reported = reports.map(r => r.name);
  const notReported = roster.filter(n => reported.indexOf(n) < 0);
  return { ok:true, team, date, roster, reported, notReported, reports };
}

/* ---------- 家長查閱：單一學生最近紀錄 ---------- */
function getStudent(team, name){
  const sh = sheetOf('Reports', REPORT_HEADERS);
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    if (row[1] === team && row[3] === name){
      list.push({ date:normDate(row[2]), attendance:row[4], body:row[5], sleep:row[6],
        fatigue:row[7], injury:row[8], water:row[10], kpiAvg:row[11], light:row[12], note:row[13] });
    }
  }
  list.sort((a,b) => (a.date < b.date ? 1 : -1));   // 新到舊
  return { ok:true, team, name, reports:list.slice(0, 14) };
}
