const TIMESTAMP_SHEET_NAME = '打刻記録';
const STAY_SHEET_NAME = '現在入室中';

/**
 * メイン画面表示 / API エンドポイント
 */
function doGet(e) {
  const action = e.parameter.action;

  // API モードの処理
  if (action) {
    let result = {};
    try {
      switch (action) {
        case 'getStayGuestList':
          result = getStayGuestList();
          break;
        case 'recordGuestTimestamp':
          const payload = {
            name: e.parameter.name,
            company: e.parameter.company,
            type: e.parameter.type,
            qrValue: e.parameter.qrValue
          };
          result = recordGuestTimestamp(payload);
          break;
        default:
          result = { ok: false, message: 'Invalid action: ' + action };
      }
    } catch (err) {
      result = { ok: false, message: err.message };
    }

    const output = JSON.stringify(result);
    const callback = e.parameter.callback;
    if (callback) {
      // JSONP レスポンス
      return ContentService.createTextOutput(callback + '(' + output + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    // 通常の JSON レスポンス
    return ContentService.createTextOutput(output)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 通常の HTML 表示モード
  const tmpl = HtmlService.createTemplateFromFile('index');
  return tmpl.evaluate()
    .setTitle('株式会社ワークアズライフ|ゲスト入退室記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setFaviconUrl('https://drive.google.com/uc?id=1YkdqM2adcpxtVM-nA8uVGGGPi2WYPkRu&.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * テンプレート内でHTMLファイルをインクルードするためのヘルパー
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 現在入室中のゲスト一覧を取得
 * @return {Array<{name: string, company: string, entryTime: string}>}
 */
function getStayGuestList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STAY_SHEET_NAME);
  if (!sheet) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue; // 名前がない場合はスキップ
    result.push({
      name: String(row[0]),
      company: String(row[1] || ''),
      entryTime: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), 'HH:mm') : ''
    });
  }
  return result;
}

/**
 * ゲストの入退室を記録
 * @param {object} payload
 *   payload = {
 *     name: string,
 *     company: string,
 *     type: 'in' | 'out',
 *     qrValue: string
 *   }
 */
function recordGuestTimestamp(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestampSheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  const staySheet = ss.getSheetByName(STAY_SHEET_NAME);

  if (!timestampSheet || !staySheet) {
    throw new Error('必要なシートが見つかりません（打刻記録 / 現在入室中）');
  }

  const now = new Date();
  const userAgent = Session.getActiveUser().getEmail() || 'guest';

  // 1. 打刻記録シートへの追記
  timestampSheet.appendRow([
    now,
    payload.company || '',
    payload.name || '',
    payload.type === 'in' ? '入室' : '退室',
    payload.qrValue || '',
    userAgent
  ]);

  // 2. 現在入室中シートの更新
  if (payload.type === 'in') {
    // 入室の場合：行を追加
    staySheet.appendRow([
      payload.name,
      payload.company || '',
      now
    ]);
  } else if (payload.type === 'out') {
    // 退室の場合：該当するゲストを削除
    const values = staySheet.getDataRange().getValues();
    for (let i = values.length - 1; i >= 1; i--) {
      // 名前と会社名が一致する最初の行を削除（名前は必須なので名前のみでも良いが、念のため両方）
      if (values[i][0] === payload.name && values[i][1] === (payload.company || '')) {
        staySheet.deleteRow(i + 1);
        break;
      }
    }
  }

  return {
    ok: true,
    timestamp: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    message: payload.type === 'in' ? `ようこそ、${payload.name}様` : `お疲れ様でした、${payload.name}様`
  };
}

/**
 * 現在入室中のリストをリセット（毎日0時実行用）
 */
function resetStayGuestList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staySheet = ss.getSheetByName(STAY_SHEET_NAME);
  if (!staySheet) return;

  const lastRow = staySheet.getLastRow();
  if (lastRow > 1) {
    staySheet.deleteRows(2, lastRow - 1);
  }
  console.log('Stay guest list reset at ' + new Date());
}

/**
 * トリガーの設定（手動または初回実行用）
 */
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'resetStayGuestList') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎日午前0時に実行するトリガーを作成
  ScriptApp.newTrigger('resetStayGuestList')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
}
