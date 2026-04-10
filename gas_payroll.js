// ============================================================
// Steak Teppei — Payroll History Google Apps Script
// ============================================================
// セットアップ手順：
// 1. https://sheets.new でスプレッドシートを新規作成
// 2. 拡張機能 → Apps Script を開く
// 3. このファイルの中身を全部コピペして保存
// 4. デプロイ → 新しいデプロイ → ウェブアプリ
//    - 実行ユーザー: 自分
//    - アクセスできるユーザー: 全員
// 5. デプロイ後に表示されるURLをpayroll.htmlのGAS_URLに貼り付ける
// ============================================================

var SHEET_NAME = 'Payroll History';
var DETAIL_SHEET = 'Payroll Detail';

// ── GET リクエスト（履歴一覧取得）──
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === 'list') {
    return jsonResponse(getRecords());
  }
  return jsonResponse({ status: 'ok', message: 'ST Payroll GAS running' });
}

// ── POST リクエスト（保存・削除）──
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'save') {
      saveRecord(body.record);
      return jsonResponse({ status: 'ok' });
    }
    if (action === 'delete') {
      deleteRecord(body.id);
      return jsonResponse({ status: 'ok' });
    }
    return jsonResponse({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── SAVE ──
function saveRecord(rec) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── サマリーシート ──
  var sh = getOrCreateSheet(ss, SHEET_NAME, [
    'ID', 'Saved At', 'From', 'To', 'Pay Date',
    'Total Tips ($)', 'Total Hours (h)', 'Tip Rate ($/h)', 'Employees'
  ]);

  // 既存チェック（同期間の上書き）
  var data = sh.getDataRange().getValues();
  var existRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === rec.from && String(data[i][3]) === rec.to) {
      existRow = i + 1; break;
    }
  }

  var summaryRow = [
    rec.id,
    rec.savedAt || new Date().toISOString(),
    rec.from,
    rec.to,
    rec.payDate,
    rec.tipsTotal,
    rec.totalHours,
    Math.round(rec.tipRate * 10000) / 10000,
    rec.employees ? rec.employees.length : 0
  ];

  if (existRow > 0) {
    sh.getRange(existRow, 1, 1, summaryRow.length).setValues([summaryRow]);
  } else {
    sh.appendRow(summaryRow);
  }

  // ── 詳細シート ──
  var dsh = getOrCreateSheet(ss, DETAIL_SHEET, [
    'Period ID', 'From', 'To', 'Pay Date',
    'Employee', 'Regular Hrs', 'OT Hrs', 'Tip Hours', 'Paycheck Tips ($)'
  ]);

  // 既存の同期間の行を削除
  var ddata = dsh.getDataRange().getValues();
  var toDelete = [];
  for (var i = ddata.length - 1; i >= 1; i--) {
    if (String(ddata[i][1]) === rec.from && String(ddata[i][2]) === rec.to) {
      toDelete.push(i + 1);
    }
  }
  toDelete.forEach(function(r) { dsh.deleteRow(r); });

  // 従業員行を追記
  if (rec.employees) {
    rec.employees.forEach(function(emp) {
      dsh.appendRow([
        rec.id,
        rec.from,
        rec.to,
        rec.payDate,
        emp.display,
        emp.reg,
        emp.ot,
        emp.tipHours,
        emp.tips
      ]);
    });
  }

  // 書式設定
  formatSheets(ss);
}

// ── GET RECORDS ──
function getRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { records: [] };

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { records: [] };

  // ヘッダー行をスキップ、新しい順に返す
  var dsh = ss.getSheetByName(DETAIL_SHEET);
  var detailData = dsh ? dsh.getDataRange().getValues() : [];

  var records = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var id = row[0];

    // 対応する従業員行を取得
    var employees = [];
    for (var j = 1; j < detailData.length; j++) {
      if (String(detailData[j][0]) === String(id)) {
        employees.push({
          display: detailData[j][4],
          reg: detailData[j][5],
          ot: detailData[j][6],
          tipHours: detailData[j][7],
          tips: detailData[j][8]
        });
      }
    }

    records.push({
      id: id,
      savedAt: row[1],
      from: row[2],
      to: row[3],
      payDate: row[4],
      tipsTotal: row[5],
      totalHours: row[6],
      tipRate: row[7],
      employees: employees
    });
  }
  return { records: records };
}

// ── DELETE ──
function deleteRecord(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // サマリーシート
  var sh = ss.getSheetByName(SHEET_NAME);
  if (sh) {
    var data = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(id)) sh.deleteRow(i + 1);
    }
  }

  // 詳細シート
  var dsh = ss.getSheetByName(DETAIL_SHEET);
  if (dsh) {
    var ddata = dsh.getDataRange().getValues();
    for (var j = ddata.length - 1; j >= 1; j--) {
      if (String(ddata[j][0]) === String(id)) dsh.deleteRow(j + 1);
    }
  }
}

// ── HELPERS ──
function getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1e2330')
      .setFontColor('#f97316');
    sh.setFrozenRows(1);
  }
  return sh;
}

function formatSheets(ss) {
  var sh = ss.getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) {
    // Tips列を通貨フォーマット
    sh.getRange(2, 6, sh.getLastRow() - 1, 1).setNumberFormat('$#,##0.00');
    sh.getRange(2, 7, sh.getLastRow() - 1, 1).setNumberFormat('#,##0.00');
    sh.getRange(2, 8, sh.getLastRow() - 1, 1).setNumberFormat('$#,##0.0000');
    sh.autoResizeColumns(1, 9);
  }
  var dsh = ss.getSheetByName(DETAIL_SHEET);
  if (dsh && dsh.getLastRow() > 1) {
    dsh.getRange(2, 9, dsh.getLastRow() - 1, 1).setNumberFormat('$#,##0.00');
    dsh.autoResizeColumns(1, 9);
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
