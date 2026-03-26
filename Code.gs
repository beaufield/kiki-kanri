// ============================================================
// ビューフィールド 貸出管理アプリ — バックエンド
// VERSION: GAS 1.5.1
// 更新日: 2026-03-26
// ============================================================

const VERSION = 'GAS 1.5.1';
const SHEET_ID = '1o12RSbRWmNsiEjVPCb2dIjyw4U4Ntn47-6Lc80E_jvk';
const LINEWORKS_WEBHOOK_URL = 'https://webhook.worksmobile.com/message/bf4bbf8b-e26f-4760-b2f2-5ea20b4cc025';
const ss = SpreadsheetApp.openById(SHEET_ID);

// ─── エントリーポイント（POST） ───────────────────────────────
function doPost(e) {
  try {
    const action = e.parameter.action;
    const data = JSON.parse(e.parameter.data);
    let result;

    if      (action === 'saveDevice')          result = saveDevice(data);
    else if (action === 'saveLoan')            result = saveLoan(data);
    else if (action === 'registerDevice')      result = registerDevice(data);
    else if (action === 'saveLoanTransaction') result = saveLoanTransaction(data);
    else if (action === 'saveSalesRep')        result = saveSalesRep(data);
    else if (action === 'deleteSalesRep')      result = deleteSalesRep(data.id);
    else if (action === 'uploadImage')         result = uploadImage(data);
    else if (action === 'saveMaker')           result = saveMaker(data);
    else if (action === 'deleteMaker')         result = deleteMaker(data.id);
    else if (action === 'issueLabel')          result = issueLabel(data);
    else if (action === 'updatePrintStatus')   result = updatePrintStatus(data);
    else if (action === 'assignLabel')         result = assignLabel(data);
    else result = { error: 'Unknown action: ' + action };

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── エントリーポイント（GET） ───────────────────────────────
function doGet(e) {
  const result = getAllData();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── 全データ取得（起動時に呼ばれる） ────────────────────────
function getAllData() {
  const devices  = getSheet('DeviceMaster');
  const loans    = getSheet('LoanLog');
  const reps     = getSheet('SalesRep');
  const makers   = getSheet('MakerMaster');
  const labels   = getSheet('LabelPool');   // ← LabelPool追加
  return { devices, loans, salesReps: reps, makers, labels };
}

// ─── 汎用シート取得（ヘッダー付き配列をオブジェクト配列に変換） ─
function getSheet(name) {
  const sheet = ss.getSheetByName(name);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ─── 商品登録トランザクション（saveDevice + assignLabel を1回で処理） ─
// フロントエンドから1回のリクエストで完結させ、通信往復を削減する
function registerDevice(data) {
  const savedResult = saveDevice(data.device);
  const savedDevice = savedResult.device;
  if (data.labelId) {
    assignLabel({ labelId: data.labelId, deviceId: savedDevice.id });
    savedDevice.labelId = data.labelId; // labelIdを確実に返す
  }
  return { success: true, device: savedDevice };
}

// ─── 貸出/返却トランザクション（saveDevice + saveLoan を1回で処理） ─
// フロントエンドから1回のリクエストで完結させ、通信往復を削減する
function saveLoanTransaction(data) {
  saveDevice(data.device);
  saveLoan(data.loan);
  return { success: true };
}

// ─── 商品マスタ保存 ──────────────────────────────────────────
function saveDevice(device) {
  const sheet = ss.getSheetByName('DeviceMaster');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (!device.id) {
    device.id = Date.now();
    device.createdAt = new Date().toISOString().split('T')[0];
  }
  device.updatedAt = new Date().toISOString().split('T')[0];

  const existingRow = data.findIndex((row, i) => i > 0 && row[0] == device.id);
  if (existingRow > 0) {
    const rowData = headers.map(h => device[h] !== undefined ? device[h] : '');
    sheet.getRange(existingRow + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    const rowData = headers.map(h => device[h] !== undefined ? device[h] : '');
    sheet.appendRow(rowData);
  }
  return { success: true, device };
}

// ─── 貸出ログ保存 + LINE WORKS即時通知 ───────────────────────
function saveLoan(loan) {
  const sheet = ss.getSheetByName('LoanLog');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!loan.id) loan.id = Date.now();
  const rowData = headers.map(h => loan[h] !== undefined ? loan[h] : '');
  sheet.appendRow(rowData);

  // 貸出登録の通知
  if (loan.type === '貸出') {
    const due = loan.returnDueDate || '未設定';
    const msg = [
      '【貸出登録】',
      '商品ID: ' + (loan.labelId || ''),
      '商品名: ' + (loan.deviceName || ''),
      '貸出先: ' + (loan.loanTo || ''),
      '返却予定日: ' + due,
      '営業担当: ' + (loan.salesRep || ''),
      '操作者: ' + (loan.registeredBy || '')
    ].join('\n');
    sendLineWorksMessage(msg);
  }

  // 返却登録の通知
  if (loan.type === '返却') {
    const msg = [
      '【返却登録】',
      '商品ID: ' + (loan.labelId || ''),
      '商品名: ' + (loan.deviceName || ''),
      '返却日: ' + (loan.date || ''),
      '操作者: ' + (loan.registeredBy || '')
    ].join('\n');
    sendLineWorksMessage(msg);
  }

  return { success: true, loan };
}

// ─── 営業担当マスタ保存 ──────────────────────────────────────
function saveSalesRep(rep) {
  const sheet = ss.getSheetByName('SalesRep');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (!rep.id) rep.id = Date.now();
  const existingRow = data.findIndex((row, i) => i > 0 && row[0] == rep.id);
  if (existingRow > 0) {
    const rowData = headers.map(h => rep[h] !== undefined ? rep[h] : '');
    sheet.getRange(existingRow + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    const rowData = headers.map(h => rep[h] !== undefined ? rep[h] : '');
    sheet.appendRow(rowData);
  }
  return { success: true, rep };
}

// ─── 営業担当削除 ────────────────────────────────────────────
function deleteSalesRep(id) {
  const sheet = ss.getSheetByName('SalesRep');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, i) => i > 0 && row[0] == id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex + 1);
  return { success: true };
}

// ─── 画像アップロード（Google Drive） ───────────────────────
function uploadImage(data) {
  const folder = getOrCreateFolder('貸出管理_画像');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(data.base64),
    data.mimeType,
    data.fileName
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  const imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
  return { success: true, imageUrl };
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// ─── メーカーマスタ保存 ──────────────────────────────────────
function saveMaker(maker) {
  const sheet = ss.getSheetByName('MakerMaster');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (!maker.id) maker.id = Date.now();
  const existingRow = data.findIndex((row, i) => i > 0 && row[0] == maker.id);
  if (existingRow > 0) {
    const rowData = headers.map(h => maker[h] !== undefined ? maker[h] : '');
    sheet.getRange(existingRow + 1, 1, 1, headers.length).setValues([rowData]);
  } else {
    const rowData = headers.map(h => maker[h] !== undefined ? maker[h] : '');
    sheet.appendRow(rowData);
  }
  return { success: true, maker };
}

// ─── メーカー削除 ────────────────────────────────────────────
function deleteMaker(id) {
  const sheet = ss.getSheetByName('MakerMaster');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, i) => i > 0 && row[0] == id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex + 1);
  return { success: true };
}

// ============================================================
// ─── LabelPool（ラベル管理）────────────────────────────────
// ============================================================

// ラベル発行
// data: { count: 発行枚数（数値） }
// LabelPoolシートに新しいラベルを「未印刷」ステータスで追加する
// labelIdは既存の最大番号から連番で採番する（形式: BF-00001）
function issueLabel(data) {
  const sheet = ss.getSheetByName('LabelPool');
  const count = parseInt(data.count, 10) || 50;
  const today = new Date().toISOString().split('T')[0];

  // 既存ラベルの最大番号を取得して連番の開始値を決める
  const existing = sheet.getDataRange().getValues();
  let maxNum = 0;
  if (existing.length > 1) {
    for (let i = 1; i < existing.length; i++) {
      // labelIdはB列（index 1）を想定
      const labelId = existing[i][1] || '';
      const match = String(labelId).match(/BF-(\d+)/);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
  }

  // ヘッダー確認（シートが空の場合はヘッダーを追加）
  const headers = ['id', 'labelId', 'status', 'issuedAt', 'printedAt', 'assignedAt', 'deviceId'];
  if (existing.length === 0) {
    sheet.appendRow(headers);
  }

  // 新規ラベル行を追加
  const newRows = [];
  for (let i = 0; i < count; i++) {
    const num = maxNum + i + 1;
    const labelId = 'BF-' + String(num).padStart(5, '0');
    const id = Date.now() + i; // 一意なID
    newRows.push([id, labelId, '未印刷', today, '', '', '']);
  }

  // まとめて書き込み（1行ずつより高速）
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  return { success: true, issued: count, startLabelId: 'BF-' + String(maxNum + 1).padStart(5, '0') };
}

// 印刷済みにする
// data: { labelIds: ['BF-00001', 'BF-00002', ...] }
// 対象ラベルのstatusを「未印刷」→「印刷済」に更新し、printedAtに本日日付を記録する
function updatePrintStatus(data) {
  const sheet = ss.getSheetByName('LabelPool');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const labelIdCol = headers.indexOf('labelId');
  const statusCol  = headers.indexOf('status');
  const printedCol = headers.indexOf('printedAt');

  const targetIds = new Set(data.labelIds || []);
  const today = new Date().toISOString().split('T')[0];
  let updatedCount = 0;

  for (let i = 1; i < rows.length; i++) {
    if (targetIds.has(rows[i][labelIdCol]) && rows[i][statusCol] === '未印刷') {
      sheet.getRange(i + 1, statusCol + 1).setValue('印刷済');
      sheet.getRange(i + 1, printedCol + 1).setValue(today);
      updatedCount++;
    }
  }

  return { success: true, updated: updatedCount };
}

// ラベル割当済みにする
// data: { labelId: 'BF-00001', deviceId: 機器ID }
// 商品登録成功後に呼ばれ、ラベルのstatusを「印刷済」→「割当済」に更新する
function assignLabel(data) {
  const sheet = ss.getSheetByName('LabelPool');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const labelIdCol  = headers.indexOf('labelId');
  const statusCol   = headers.indexOf('status');
  const assignedCol = headers.indexOf('assignedAt');
  const deviceIdCol = headers.indexOf('deviceId');

  const today = new Date().toISOString().split('T')[0];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][labelIdCol] === data.labelId) {
      sheet.getRange(i + 1, statusCol + 1).setValue('割当済');
      sheet.getRange(i + 1, assignedCol + 1).setValue(today);
      if (data.deviceId) {
        sheet.getRange(i + 1, deviceIdCol + 1).setValue(data.deviceId);
      }
      return { success: true, labelId: data.labelId };
    }
  }

  // 対象ラベルが見つからない場合はエラーを返す
  return { success: false, error: 'ラベルが見つかりません: ' + data.labelId };
}

// ─── LINE WORKS通知 ────────────────────────────────────────
function sendLineWorksMessage(text) {
  try {
    UrlFetchApp.fetch(LINEWORKS_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ body: { text: text } })
    });
  } catch(e) {
    // 通知失敗してもメイン処理は継続する
    console.error('LINE WORKS通知エラー:', e.toString());
  }
}

// ─── 週次レポート（毎週火曜9:00にトリガー登録） ─────────────
// 返却期限未設定の貸出中商品をLINE WORKSに通知する
function sendWeeklyReport() {
  const devices = getSheet('DeviceMaster');
  const targets = devices.filter(function(d) {
    return d.status === '貸出中' && !d.returnDueDate;
  });

  if (targets.length === 0) {
    sendLineWorksMessage('【週次レポート】返却期限未設定の貸出中商品はありません。');
    return;
  }

  let msg = '【週次レポート】返却期限未設定の貸出中商品（' + targets.length + '点）';
  targets.forEach(function(d, i) {
    msg += '\n\n' + (i + 1) + '. [' + (d.labelId || d.id) + '] ' + d.name;
    msg += '\n　貸出先: ' + (d.loanTo || '未設定');
    if (d.salesRep) msg += '\n　営業担当: ' + d.salesRep;
  });

  sendLineWorksMessage(msg);
}

// ─── テスト用関数（GASエディタから手動実行） ────────────────
function testNotify() {
  sendLineWorksMessage('【テスト】LINE WORKS通知の動作確認です。');
}
