/**
 * Backend Google Apps Script cho Product Knowledge Quiz.
 * Spreadsheet: https://docs.google.com/spreadsheets/d/1aqyrjbYFu0Efo1So1oVpCdsl8z9nkIdYGGPhnnqckbc/edit
 *
 * Cách dùng:
 * 1) Dán toàn bộ file này vào Apps Script.
 * 2) Chạy setupSheet() MỘT LẦN và cấp quyền.
 * 3) Deploy > Manage deployments > Edit > New version > Web app
 *    Execute as: Me | Who has access: Anyone.
 * 4) Nếu bạn cập nhật deployment hiện tại, URL /exec có thể giữ nguyên.
 */

const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1aqyrjbYFu0Efo1So1oVpCdsl8z9nkIdYGGPhnnqckbc',
  RESULTS_SHEET: 'Results',
  ATTEMPTS_SHEET: 'Attempts',
  QUIZ_MS: 5 * 60 * 1000,
  TOTAL_QUESTIONS: 10,
  POINTS_PER_QUESTION: 10,
  TIMEZONE: 'Asia/Ho_Chi_Minh'
});

// ĐÁP ÁN CHỈ NẰM Ở SERVER, KHÔNG NẰM TRONG index.html.
const ANSWER_KEY = Object.freeze({
  1: ['B'],
  2: ['A','B','C'],
  3: ['B'],
  4: ['C'],
  5: ['B'],
  6: ['B'],
  7: ['C'],
  8: ['A','B','C'],
  9: ['C'],
  10:['A','B','C']
});

function doGet() {
  return json_({
    ok: true,
    service: 'Product Knowledge Quiz API',
    status: 'ready',
    serverNow: Date.now()
  });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').trim();

    if (action === 'start')  return handleStart_(body);
    if (action === 'submit') return handleSubmit_(body);
    if (action === 'status') return handleStatus_(body);
    if (action === 'result') return handleResult_(body);

    return json_({ok:false, message:'Action không hợp lệ.'});
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return json_({ok:false, message:'Lỗi máy chủ: ' + (err && err.message ? err.message : String(err))});
  }
}

/** Chạy thủ công 1 lần sau khi dán code. */
function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  let results = ss.getSheetByName(CONFIG.RESULTS_SHEET);
  if (!results) results = ss.insertSheet(CONFIG.RESULTS_SHEET);
  results.clear();
  const resultHeaders = [
    'Rank','Submitted At','Name','Group','Score','Correct','Total',
    'Duration (sec)','Start At','Reveal At','Attempt ID','Answers JSON','Review Status'
  ];
  results.getRange(1,1,1,resultHeaders.length).setValues([resultHeaders]);
  styleHeader_(results, resultHeaders.length);
  results.setFrozenRows(1);
  results.setColumnWidths(1, resultHeaders.length, 130);
  results.setColumnWidth(3, 190);
  results.setColumnWidth(4, 130);
  results.setColumnWidth(12, 320);
  results.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  results.getRange('I:J').setNumberFormat('dd/MM/yyyy HH:mm:ss');

  let attempts = ss.getSheetByName(CONFIG.ATTEMPTS_SHEET);
  if (!attempts) attempts = ss.insertSheet(CONFIG.ATTEMPTS_SHEET);
  attempts.clear();
  const attemptHeaders = [
    'Attempt ID','Name','Group','Start At (ms)','Reveal At (ms)','Submitted At (ms)',
    'Duration (sec)','Score','Correct','Answers JSON','Submitted','Reviewed At (ms)'
  ];
  attempts.getRange(1,1,1,attemptHeaders.length).setValues([attemptHeaders]);
  styleHeader_(attempts, attemptHeaders.length);
  attempts.setFrozenRows(1);
  try { attempts.hideSheet(); } catch (_) {}

  return 'Setup hoàn tất: đã tạo Results và Attempts.';
}

function handleStart_(body) {
  const name = sanitize_(body.name, 80);
  const group = sanitize_(body.group, 60);
  if (name.length < 2) return json_({ok:false, message:'Tên không hợp lệ.'});
  if (!group) return json_({ok:false, message:'Nhóm không hợp lệ.'});

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    const now = Date.now();
    const revealAt = now + CONFIG.QUIZ_MS;
    const token = Utilities.getUuid();
    const sheet = getSS_().getSheetByName(CONFIG.ATTEMPTS_SHEET);
    sheet.appendRow([token,name,group,now,revealAt,'','','','','',false,'']);
    return json_({ok:true, token, name, group, startAt:now, revealAt, serverNow:Date.now()});
  } finally {
    lock.releaseLock();
  }
}

function handleStatus_(body) {
  ensureSheets_();
  const token = sanitize_(body.token, 100);
  const found = findAttempt_(token);
  if (!found) return json_({ok:false, message:'Không tìm thấy lượt thi.'});

  const v = found.values;
  return json_({
    ok:true,
    name:String(v[1] || ''),
    group:String(v[2] || ''),
    startAt:Number(v[3]),
    revealAt:Number(v[4]),
    submitted:toBool_(v[10]),
    durationSeconds:v[6] === '' ? null : Number(v[6]),
    serverNow:Date.now()
  });
}

function handleSubmit_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSheets_();
    const token = sanitize_(body.token, 100);
    const found = findAttempt_(token);
    if (!found) return json_({ok:false, message:'Không tìm thấy lượt thi.'});

    const a = found.values;
    const startAt = Number(a[3]);
    const revealAt = Number(a[4]);
    const alreadySubmitted = toBool_(a[10]);

    if (alreadySubmitted) {
      return json_({
        ok:true,
        alreadySubmitted:true,
        durationSeconds:Number(a[6]),
        revealAt,
        serverNow:Date.now()
      });
    }

    const answers = normalizeAnswers_(body.answers || {});
    const scored = score_(answers);
    const submittedAt = Date.now();
    const durationSeconds = Math.max(0, Math.min(CONFIG.QUIZ_MS, submittedAt - startAt)) / 1000;
    const durationRounded = Math.round(durationSeconds * 10) / 10;
    const answersJson = JSON.stringify(answers);

    const attemptSheet = found.sheet;
    attemptSheet.getRange(found.row, 6, 1, 6).setValues([[
      submittedAt,
      durationRounded,
      scored.score,
      scored.correctCount,
      answersJson,
      true
    ]]);

    const results = getSS_().getSheetByName(CONFIG.RESULTS_SHEET);
    results.appendRow([
      '',
      new Date(submittedAt),
      String(a[1] || ''),
      String(a[2] || ''),
      scored.score,
      scored.correctCount,
      CONFIG.TOTAL_QUESTIONS,
      durationRounded,
      new Date(startAt),
      new Date(revealAt),
      token,
      answersJson,
      'Locked until timer ends'
    ]);
    sortResults_();

    // Cố ý KHÔNG trả đáp án đúng và cũng không trả review ở đây.
    return json_({
      ok:true,
      submitted:true,
      durationSeconds:durationRounded,
      revealAt,
      serverNow:Date.now()
    });
  } finally {
    lock.releaseLock();
  }
}

function handleResult_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureSheets_();
    const token = sanitize_(body.token, 100);
    const found = findAttempt_(token);
    if (!found) return json_({ok:false, message:'Không tìm thấy lượt thi.'});

    const a = found.values;
    const revealAt = Number(a[4]);
    const now = Date.now();
    if (!toBool_(a[10])) return json_({ok:false, message:'Bài chưa được nộp.'});
    if (now < revealAt) {
      return json_({ok:false, message:'Chưa đến thời điểm mở đáp án.', remainingMs:revealAt-now, serverNow:now});
    }

    const answers = safeParse_(a[9], {});
    const scored = score_(answers);
    const review = Object.keys(ANSWER_KEY).map(id => {
      const qid = Number(id);
      const selected = Array.isArray(answers[qid]) ? answers[qid] : [];
      const correct = ANSWER_KEY[qid].slice();
      return {id:qid, selected, correct, isCorrect:sameSet_(selected, correct)};
    });

    found.sheet.getRange(found.row, 12).setValue(now);
    markReviewed_(token);

    return json_({
      ok:true,
      score:scored.score,
      correctCount:scored.correctCount,
      durationSeconds:Number(a[6]),
      review,
      serverNow:now
    });
  } finally {
    lock.releaseLock();
  }
}

function normalizeAnswers_(raw) {
  const out = {};
  for (let i = 1; i <= CONFIG.TOTAL_QUESTIONS; i++) {
    const value = raw[i] !== undefined ? raw[i] : raw[String(i)];
    let arr = Array.isArray(value) ? value : (value ? [value] : []);
    arr = [...new Set(arr.map(x => String(x).toUpperCase().trim()).filter(x => /^[A-D]$/.test(x)))].sort();
    out[i] = arr;
  }
  return out;
}

function score_(answers) {
  let correctCount = 0;
  for (let i = 1; i <= CONFIG.TOTAL_QUESTIONS; i++) {
    if (sameSet_(answers[i] || [], ANSWER_KEY[i] || [])) correctCount++;
  }
  return {correctCount, score:correctCount * CONFIG.POINTS_PER_QUESTION};
}

function sameSet_(a, b) {
  const x = [...a].map(String).sort();
  const y = [...b].map(String).sort();
  return x.length === y.length && x.every((v,i) => v === y[i]);
}

function sortResults_() {
  const sheet = getSS_().getSheetByName(CONFIG.RESULTS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const cols = 13;
  sheet.getRange(2,1,lastRow-1,cols).sort([
    {column:5, ascending:false}, // Score cao trước
    {column:8, ascending:true},  // Cùng điểm: thời gian ngắn hơn trước
    {column:2, ascending:true}   // Nếu vẫn bằng: nộp sớm hơn trước
  ]);
  const ranks = Array.from({length:lastRow-1}, (_,i) => [i+1]);
  sheet.getRange(2,1,ranks.length,1).setValues(ranks);
}

function markReviewed_(token) {
  const sheet = getSS_().getSheetByName(CONFIG.RESULTS_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const finder = sheet.getRange(2,11,last-1,1).createTextFinder(token).matchEntireCell(true).findNext();
  if (finder) sheet.getRange(finder.getRow(),13).setValue('Review opened');
}

function findAttempt_(token) {
  if (!token) return null;
  const sheet = getSS_().getSheetByName(CONFIG.ATTEMPTS_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const cell = sheet.getRange(2,1,last-1,1).createTextFinder(token).matchEntireCell(true).findNext();
  if (!cell) return null;
  const row = cell.getRow();
  return {sheet, row, values:sheet.getRange(row,1,1,12).getValues()[0]};
}

function ensureSheets_() {
  const ss = getSS_();
  if (!ss.getSheetByName(CONFIG.RESULTS_SHEET) || !ss.getSheetByName(CONFIG.ATTEMPTS_SHEET)) {
    throw new Error('Chưa setup sheet. Hãy chạy hàm setupSheet() một lần trong Apps Script.');
  }
}

function getSS_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function styleHeader_(sheet, columnCount) {
  const r = sheet.getRange(1,1,1,columnCount);
  r.setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
}

function sanitize_(value, maxLen) {
  return String(value || '').replace(/[\r\n\t]+/g,' ').replace(/\s+/g,' ').trim().slice(0,maxLen);
}

function toBool_(v) {
  return v === true || String(v).toLowerCase() === 'true';
}

function safeParse_(value, fallback) {
  try { return JSON.parse(String(value || '')); } catch (_) { return fallback; }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
