/**
 * Bài kiểm tra Trimafort / Grafort - Backend Google Apps Script
 * - Nhận bài làm từ trang web (index.html), tự chấm lại điểm ở phía máy chủ
 * - Ghi vào sheet "KetQua" (dữ liệu gốc)
 * - Sheet "XepHang" được Apps Script tự sắp xếp:
 *   Điểm giảm dần -> Thời gian làm tăng dần -> Nộp sớm hơn
 *
 * Bản này KHÔNG dùng công thức ARRAYFORMULA/SORT trong sheet XepHang
 * để tránh lỗi dấu phân cách công thức theo Locale Google Sheets.
 */

var SPREADSHEET_ID = '1aqyrjbYFu0Efo1So1oVpCdsl8z9nkIdYGGPhnnqckbc';
var RAW_SHEET  = 'KetQua';
var RANK_SHEET = 'XepHang';
var POINTS_PER_QUESTION = 10;

var ANSWER_KEY = {
  1: 'B', 2: 'ABC', 3: 'B', 4: 'C', 5: 'B',
  6: 'B', 7: 'C', 8: 'ABC', 9: 'C', 10: 'ABC'
};

var HEADERS = [
  'Thời điểm nộp', 'Họ và tên', 'Nhóm', 'Điểm', 'Số câu đúng',
  'Thời gian làm (giây)', 'Tự động nộp khi hết giờ', 'Câu sai',
  'Đáp án đã chọn', 'Lần làm', 'Mã lượt làm'
];

var RANK_HEADERS = [
  'Hạng', 'Họ và tên', 'Nhóm', 'Điểm',
  'Số câu đúng', 'Thời gian làm (giây)', 'Thời điểm nộp', 'Câu sai'
];

var COL = {
  TIME: 1, NAME: 2, GROUP: 3, SCORE: 4, CORRECT: 5,
  DURATION: 6, AUTO: 7, WRONG: 8, ANSWERS: 9, ATTEMPT: 10, ID: 11
};

var TIME_FORMAT = 'dd/MM/yyyy HH:mm:ss';

function doGet() {
  return json_({ ok: true, message: 'Quiz endpoint đang hoạt động' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: 'Không có dữ liệu' });
    }

    var d = JSON.parse(e.postData.contents);

    var name = clean_(d.name, 60);
    var group = clean_(d.group, 40);
    var attemptId = String(d.attemptId || '')
      .replace(/[^\w-]/g, '')
      .slice(0, 80);

    if (!name || !group || !attemptId) {
      return json_({
        ok: false,
        error: 'Thiếu họ tên, nhóm hoặc mã lượt làm'
      });
    }

    var result = grade_(d.answers || {});
    var sh = ensureSheets_();

    var attempt = 1;
    var last = sh.getLastRow();

    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();

      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][COL.ID - 1]) === attemptId) {
          return json_({
            ok: true,
            duplicate: true,
            score: rows[i][COL.SCORE - 1]
          });
        }

        if (
          same_(rows[i][COL.NAME - 1], name) &&
          same_(rows[i][COL.GROUP - 1], group)
        ) {
          attempt++;
        }
      }
    }

    var duration = Math.max(
      0,
      Math.min(3600, Math.round(Number(d.durationSec) || 0))
    );

    var answersText = Object.keys(ANSWER_KEY)
      .map(function(q) {
        return 'Q' + q + '=' +
          (normalize_(d.answers && d.answers[q]) || '-');
      })
      .join('; ');

    sh.appendRow([
      new Date(),
      name,
      group,
      result.score,
      result.correct,
      duration,
      d.auto ? 'Có' : 'Không',
      result.wrong.length
        ? 'Câu ' + result.wrong.join(', ')
        : 'Không sai',
      answersText,
      attempt,
      attemptId
    ]);

    refreshRanking_();

    return json_({
      ok: true,
      score: result.score,
      correct: result.correct
    });

  } catch (err) {
    return json_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });

  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (x) {}
    }
  }
}

function grade_(answers) {
  var correct = 0;
  var wrong = [];

  Object.keys(ANSWER_KEY).forEach(function(q) {
    if (normalize_(answers[q]) === ANSWER_KEY[q]) {
      correct++;
    } else {
      wrong.push(q);
    }
  });

  return {
    correct: correct,
    wrong: wrong,
    score: correct * POINTS_PER_QUESTION
  };
}

function normalize_(s) {
  return String(s || '')
    .toUpperCase()
    .replace(/[^A-D]/g, '')
    .split('')
    .sort()
    .join('');
}

function clean_(s, max) {
  s = String(s == null ? '' : s)
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max);

  if (/^[=+\-@]/.test(s)) {
    s = "'" + s;
  }

  return s;
}

function same_(a, b) {
  return String(a).replace(/^'/, '').toLowerCase() ===
    String(b).replace(/^'/, '').toLowerCase();
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  ensureSheets_();
  refreshRanking_();
  return 'Đã sẵn sàng';
}

function ensureSheets_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var raw = ss.getSheetByName(RAW_SHEET);

  if (!raw) {
    var first = ss.getSheets()[0];

    if (first.getLastRow() === 0 && first.getLastColumn() === 0) {
      raw = first;
      raw.setName(RAW_SHEET);
    } else {
      raw = ss.insertSheet(RAW_SHEET);
    }
  }

  if (raw.getRange(1, 1).getValue() !== HEADERS[0]) {
    raw.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight('bold')
      .setBackground('#0b6b52')
      .setFontColor('#ffffff')
      .setWrap(true)
      .setVerticalAlignment('middle');

    raw.setFrozenRows(1);

    raw.getRange(2, COL.TIME, Math.max(1, raw.getMaxRows() - 1), 1)
      .setNumberFormat(TIME_FORMAT);

    raw.getRange(2, COL.ID, Math.max(1, raw.getMaxRows() - 1), 1)
      .setNumberFormat('@');

    var widths = [150, 190, 90, 60, 90, 110, 130, 130, 330, 70, 260];
    widths.forEach(function(w, i) {
      raw.setColumnWidth(i + 1, w);
    });
  }

  var rk = ss.getSheetByName(RANK_SHEET);

  if (!rk) {
    rk = ss.insertSheet(RANK_SHEET);
  }

  rk.getRange(1, 1, 1, RANK_HEADERS.length)
    .setValues([RANK_HEADERS])
    .setFontWeight('bold')
    .setBackground('#073f31')
    .setFontColor('#ffffff')
    .setWrap(true)
    .setVerticalAlignment('middle');

  rk.setFrozenRows(1);

  var w2 = [60, 190, 90, 60, 90, 110, 150, 130];
  w2.forEach(function(w, i) {
    rk.setColumnWidth(i + 1, w);
  });

  return raw;
}

function refreshRanking_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var raw = ss.getSheetByName(RAW_SHEET);
  var rk = ss.getSheetByName(RANK_SHEET);

  if (!raw) return;

  if (!rk) {
    rk = ss.insertSheet(RANK_SHEET);
  }

  rk.getRange(1, 1, 1, RANK_HEADERS.length)
    .setValues([RANK_HEADERS])
    .setFontWeight('bold')
    .setBackground('#073f31')
    .setFontColor('#ffffff')
    .setWrap(true)
    .setVerticalAlignment('middle');

  rk.setFrozenRows(1);

  if (rk.getMaxRows() > 1) {
    rk.getRange(2, 1, rk.getMaxRows() - 1, RANK_HEADERS.length)
      .clearContent();
  }

  var last = raw.getLastRow();
  if (last <= 1) return;

  var values = raw
    .getRange(2, 1, last - 1, HEADERS.length)
    .getValues()
    .filter(function(r) {
      return String(r[COL.NAME - 1]).trim() !== '';
    });

  values.sort(function(a, b) {
    var scoreA = Number(a[COL.SCORE - 1]) || 0;
    var scoreB = Number(b[COL.SCORE - 1]) || 0;

    if (scoreA !== scoreB) return scoreB - scoreA;

    var durationA = Number(a[COL.DURATION - 1]) || 0;
    var durationB = Number(b[COL.DURATION - 1]) || 0;

    if (durationA !== durationB) return durationA - durationB;

    return dateMs_(a[COL.TIME - 1]) - dateMs_(b[COL.TIME - 1]);
  });

  var output = values.map(function(r, i) {
    return [
      i + 1,
      r[COL.NAME - 1],
      r[COL.GROUP - 1],
      r[COL.SCORE - 1],
      r[COL.CORRECT - 1],
      r[COL.DURATION - 1],
      r[COL.TIME - 1],
      r[COL.WRONG - 1]
    ];
  });

  if (output.length) {
    rk.getRange(2, 1, output.length, RANK_HEADERS.length)
      .setValues(output);

    rk.getRange(2, 7, output.length, 1)
      .setNumberFormat(TIME_FORMAT);
  }

  var w2 = [60, 190, 90, 60, 90, 110, 150, 130];
  w2.forEach(function(w, i) {
    rk.setColumnWidth(i + 1, w);
  });
}

function dateMs_(v) {
  if (v instanceof Date) return v.getTime();

  var t = new Date(v).getTime();
  return isNaN(t) ? Number.MAX_SAFE_INTEGER : t;
}
