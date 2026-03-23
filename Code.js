// ====== 定数 ======
const DB_NAME_RESERVATIONS = 'Reservations';
const DB_NAME_SETTINGS = 'Settings';
// 許可するメールアドレスのドメイン（複数のドメインを許可可能）
const ALLOWED_DOMAIN_SUFFIXES = ['@my.co.jp', '.my.co.jp'];
// アプリのタイトル（ブラウザのタブ等に表示される名前）
const APP_TITLE = '部屋予約システム';

// ====== 部屋ごとの設定 (汎用化) ======
const ROOM_CONFIGS = [
  {
    name: 'room201',
    primary: '#2563eb', // メイン色
    primaryLight: '#eff6ff', // 背景の薄い色
    bgGradient: 'linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%)', // 全体背景
    notice: '定員50名・プロジェクター完備<br>飲食禁止'
  },
  {
    name: 'B5-302',
    primary: '#059669', // メイン色
    primaryLight: '#ecfdf5', // 背景の薄い色
    bgGradient: 'linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%)', // 全体背景
    notice: '定員15名・プロジェクターなし'
  }
];

// ========= ユーティリティ =========
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====== 初期化処理 (DB構築) ======
function initDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Reservationsシート作成
  let resSheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  if (!resSheet) {
    resSheet = ss.insertSheet(DB_NAME_RESERVATIONS);
    resSheet.appendRow(['ID', 'Token', 'Room', 'StartTime', 'EndTime', 'Email', 'Name', 'Title', 'Timestamp', 'Status', 'ConfirmExpiry']);
    // Status: 'Confirmed' (確定) | 'Tentative' (仮予約)
    resSheet.setFrozenRows(1);
    // 日時カラムのフォーマット
    resSheet.getRange("D:E").setNumberFormat("yyyy/MM/dd hh:mm:ss");
  }
  
  // Settingsシート作成
  let setSheet = ss.getSheetByName(DB_NAME_SETTINGS);
  if (!setSheet) {
    setSheet = ss.insertSheet(DB_NAME_SETTINGS);
    setSheet.appendRow(['Key', 'Value', 'Description']);
    setSheet.appendRow(['ADMIN_PASSWORD', 'admin123', '管理者用パスワード (CSVアップロード用)']);
    setSheet.setFrozenRows(1);
  }
}

// ====== 設定取得 ======
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_SETTINGS);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function handleConfirm(token) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>予約の確定</title>
    </head>
    <body style="font-family: sans-serif; text-align: center; padding: 50px; background: #f9f9f9;">
      <div style="background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); max-width: 400px; margin: 0 auto;">
        <h2>予約はまだ確定していません</h2>
        <p style="color: #666; margin-bottom: 30px;">以下のボタンをクリックして、予約を最終確定してください。</p>
        <button id="btn" onclick="doConfirm()" style="padding: 12px 24px; background: #4CAF50; color: white; border: none; border-radius: 5px; font-size: 16px; cursor: pointer; width: 100%; transition: background 0.3s;">予約を確定する</button>
        <div id="msg" style="margin-top: 20px; line-height: 1.6;"></div>
      </div>
      <script>
        function doConfirm() {
          var btn = document.getElementById('btn');
          btn.disabled = true;
          btn.innerText = '通信中...';
          btn.style.background = '#9e9e9e';
          
          google.script.run.withSuccessHandler(function(res) {
            document.getElementById('msg').innerHTML = res.message;
            btn.style.display = 'none';
            document.querySelector('h2').style.display = 'none';
            document.querySelector('p').style.display = 'none';
          }).withFailureHandler(function(err) {
            document.getElementById('msg').innerHTML = '<span style="color:red;">エラー: ' + err.message + '</span>';
            btn.disabled = false;
            btn.innerText = '予約を確定する';
            btn.style.background = '#4CAF50';
          }).executeConfirm('${token}');
        }
      </script>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle("予約の確定").addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ====== API: 実際の予約確定処理 ======
function executeConfirm(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  const appUrl = ScriptApp.getService().getUrl();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === token) {
      const status = data[i][9] || 'Confirmed'; 
      const expiry = data[i][10] ? new Date(data[i][10]) : null;
      const room = data[i][2];
      const roomParam = '?room=' + encodeURIComponent(room);
      
      if (status === 'Confirmed') {
        return { success: false, message: 'すでに予約は確定済みです。<br><br><a href="' + appUrl + roomParam + '" target="_top" style="padding:8px 16px; background:#e0e0e0; color:#333; text-decoration:none; border-radius:5px; display:inline-block;">カレンダーへ戻る</a>' };
      }
      
      if (status === 'Tentative') {
        if (expiry && new Date() > expiry) {
          // 期限切れ
          return { success: false, message: '<span style="color: #E53935; font-weight:bold;">予約の有効期限（15分）が切れました。</span><p>お手数ですが、再度最初から予約をやり直してください。</p><br><a href="' + appUrl + roomParam + '" target="_top" style="padding:8px 16px; background:#e0e0e0; color:#333; text-decoration:none; border-radius:5px; display:inline-block;">カレンダーへ戻る</a>' };
        } else {
          // 確定成功
          sheet.getRange(i + 1, 10).setValue('Confirmed');
          sheet.getRange(i + 1, 11).setValue(''); // Expiryクリア
          return { success: true, message: '<span style="color: #2E7D32; font-size: 1.3em; font-weight:bold;">予約が確定しました！</span><p>ご利用をお待ちしております。</p><br><a href="' + appUrl + roomParam + '" target="_top" style="padding:10px 20px; background:#4CAF50; color:white; text-decoration:none; border-radius:5px; display:inline-block; margin-top:10px;">カレンダーへ戻る</a>' };
        }
      }
    }
  }
  return { success: false, message: '無効なURLです。すでにキャンセルされたか、URLが間違っています。<br><br><a href="' + appUrl + '" target="_top" style="padding:8px 16px; background:#e0e0e0; color:#333; text-decoration:none; border-radius:5px; display:inline-block;">トップページへ戻る</a>' };
}

// ====== ルーティング ======
function doGet(e) {
  // 1. カラーテーマ生成ツール ?tool=colors
  if (e && e.parameter && e.parameter.tool === 'colors') {
    return HtmlService.createHtmlOutputFromFile('color_tool')
      .setTitle('RoomReserv カラーテーマジェネレーター')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 2. 予約確定 ?confirm=token
  if (e.parameter.confirm) {
    return handleConfirm(e.parameter.confirm);
  }

  // CSVテンプレートのダウンロード
  if (e.parameter.csvTemplate) {
    const csvContent = "Room,StartTime,EndTime,Title\nB6-202,2024/04/01 09:00,2024/04/01 10:30,講義A\nB5-302,2024/04/02 13:00,2024/04/02 14:30,プロジェクト会議\n";
    return ContentService.createTextOutput(csvContent)
      .setMimeType(ContentService.MimeType.CSV)
      .downloadAsFile("reservation_template.csv");
  }
  
  // 表示設定判定
  const template = HtmlService.createTemplateFromFile('index');
  template.editToken = (e.parameter && e.parameter.token) ? e.parameter.token : "";
  template.defaultRoom = (e.parameter && e.parameter.room) ? e.parameter.room : "";
  template.appUrl = ScriptApp.getService().getUrl();
  template.roomConfigs = ROOM_CONFIGS; // 汎用設定を渡す
  template.appTitle = APP_TITLE;
  template.allowedDomains = ALLOWED_DOMAIN_SUFFIXES;
  
  return template.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ====== API: 予約データの取得 ======
function getReservations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const rows = data.slice(1);
  const results = [];
  
  const now = new Date();
  rows.forEach(row => {
    const status = row[9] || 'Confirmed';
    const expiry = row[10] ? new Date(row[10]) : null;
    
    // 仮予約かつ期限切れならスキップ（画面に表示しない・ブロックしない）
    if (status === 'Tentative' && expiry && now > expiry) return;
    
    let title = row[7];
    if (status === 'Tentative') {
      title += ' (仮予約)';
    }

    // メールアドレスとトークンは機密情報なので返さない
    results.push({
      id: row[0],
      room: row[2],
      start: row[3] ? new Date(row[3]).toISOString() : '',
      end: row[4] ? new Date(row[4]).toISOString() : '',
      name: row[6],
      title: title,
    });
  });
  return results;
}

// ====== API: トークンによる予約詳細の取得 ======
function getReservationByToken(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === token) {
      return {
        id: data[i][0],
        room: data[i][2],
        start: data[i][3] ? new Date(data[i][3]).toISOString() : '',
        end: data[i][4] ? new Date(data[i][4]).toISOString() : '',
        email: data[i][5],
        name: data[i][6],
        title: data[i][7]
      };
    }
  }
  throw new Error('無効なリンクです。すでにキャンセルされているか、URLが間違っています。');
}

// ====== API: 新規予約の作成 ======
function createReservation(form) {
  const { room, date, startTime, endTime, email, name, title } = form;
  
  if (!email || !(email.endsWith(ALLOWED_DOMAIN_SUFFIXES[0]) || email.endsWith(ALLOWED_DOMAIN_SUFFIXES[1]))) {
    throw new Error(`許可されていないメールアドレスのドメインです。`);
  }
  
  // 日付のパース
  const dateStr = date.replace(/-/g, '/');
  const startObj = new Date(`${dateStr} ${startTime}`);
  const endObj = new Date(`${dateStr} ${endTime}`);
  
  if (isNaN(startObj.getTime()) || isNaN(endObj.getTime())) {
    throw new Error('日時の形式が正しくありません。');
  }
  if (endObj <= startObj) {
    throw new Error('終了時間は開始時間より後に設定してください。');
  }
  
  // 3ヶ月先制限のチェック
  const today = new Date();
  today.setHours(0,0,0,0);
  const maxDate = new Date();
  maxDate.setMonth(today.getMonth() + 3);
  maxDate.setHours(23,59,59,999);
  
  if (startObj < today || startObj > maxDate) {
    throw new Error('予約は本日から3ヶ月先までの期間で指定してください。');
  }
  
  // 重複チェック
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const rRoom = data[i][2];
    const rStart = new Date(data[i][3]);
    const rEnd = new Date(data[i][4]);
    
    if (rRoom === room) {
      const rStatus = data[i][9] || 'Confirmed';
      const rExpiry = data[i][10] ? new Date(data[i][10]) : null;
      
      // 期限切れの仮予約は無視して上書き予約可能とする
      if (rStatus === 'Tentative' && rExpiry && new Date() > rExpiry) {
        continue;
      }

      // 既存の予約と被っているか [s1 < e2 && e1 > s2]
      if (startObj < rEnd && endObj > rStart) {
        throw new Error('指定された時間はすでに他の予約が入っています。');
      }
    }
  }
  
  const id = Utilities.getUuid();
  const token = Utilities.getUuid();
  const timestamp = new Date();
  
  const status = 'Tentative';
  const expiry = new Date(timestamp.getTime() + 15 * 60 * 1000); // 15分後

  sheet.appendRow([
    id,
    token, // 機密の編集および確定用トークン
    room,
    startObj,
    endObj,
    email,
    name,
    title,
    timestamp,
    status,
    expiry
  ]);
  
  // 確定用リンクを含むメール送信
  try {
    const appUrl = ScriptApp.getService().getUrl();
    const confirmUrl = `${appUrl}?confirm=${token}`;
    const editUrl = `${appUrl}?token=${token}`;
    const subject = `【${APP_TITLE}】${room} の仮予約を受け付けました（※まだ確定していません）`;
    const body = `
${name} 様

予約の仮受付を完了しました。
※この時点ではまだ予約は確定していません。

以下のURLを【15分以内】にクリックして、予約を確定させてください。
（15分を過ぎると自動的にキャンセル扱いとなります）

▼ 予約確定用URL
${confirmUrl}

■ 予約内容
場所: ${room}
日時: ${Utilities.formatDate(startObj, "Asia/Tokyo", "yyyy/MM/dd HH:mm")} 〜 ${Utilities.formatDate(endObj, "Asia/Tokyo", "HH:mm")}
目的: ${title}

--------------------------------------------------
▼ 確定後のキャンセル・変更について
予約確定後、キャンセルを行う場合は以下のURLをご利用ください。
${editUrl}
※このURLは予約者本人のみ有効な秘密のリンクです。
`;
    MailApp.sendEmail({ to: email, subject: subject, body: body });
  } catch (e) {
    console.error("メール送信エラー", e);
  }
  
  return { success: true };
}

// ====== API: 予約のキャンセル ======
function cancelReservation(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === token) {
      sheet.deleteRow(i + 1); // 1-indexed (row 1 is header)
      return { success: true };
    }
  }
  throw new Error('無効なトークンです。');
}

// ====== API: 管理者用CSV一括アップロード ======
function uploadCsv(csvText, adminPassword) {
  const settings = getSettings();
  if (settings['ADMIN_PASSWORD'] !== adminPassword) {
    throw new Error('管理者パスワードが間違っています。');
  }
  
  const rows = Utilities.parseCsv(csvText);
  if (rows.length === 0) throw new Error('CSVが空です。');
  
  // 1行目はヘッダーと仮定してスキップ
  const dataRows = rows.slice(1);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  
  let addedCount = 0;
  
  dataRows.forEach(row => {
    if (row.length < 4) return;
    const room = row[0].trim();
    const startStr = row[1].trim();
    const endStr = row[2].trim();
    const title = row[3].trim();
    
    if (!room || !startStr || !endStr) return;
    
    const startObj = new Date(startStr);
    const endObj = new Date(endStr);
    
    if (isNaN(startObj.getTime()) || isNaN(endObj.getTime())) return;
    
    const id = Utilities.getUuid();
    sheet.appendRow([
      id,
      '', // 管理者用は自身によるキャンセルURL無し（スプレッドシートから直接消すか、管理者機能で上書き）
      room,
      startObj,
      endObj,
      'Admin(CSV一括)',
      '管理者',
      title,
      new Date(),
      'Confirmed',
      ''
    ]);
    addedCount++;
  });
  
  return { success: true, count: addedCount };
}

// ====== バッチジョブ: 古い予約の削除 ======
function deleteOldReservations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_NAME_RESERVATIONS);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const threeMonthsAgo = new Date();
  threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
  
  const now = new Date();
  // 行のインデックスがずれないよう、下から上に向かって削除
  for (let i = data.length - 1; i >= 1; i--) {
    const endStr = data[i][4];
    const endObj = new Date(endStr);
    const status = data[i][9] || 'Confirmed';
    const expiry = data[i][10] ? new Date(data[i][10]) : null;

    // 1. 3ヶ月以上前の予約を削除
    if (endObj < threeMonthsAgo) {
      sheet.deleteRow(i + 1);
      continue;
    }
    
    // 2. 有効期限切れの仮予約を削除
    if (status === 'Tentative' && expiry && now > expiry) {
      sheet.deleteRow(i + 1);
    }
  }
}

// 初回設定用（開発者が手動実行してトリガーを作成）
function setupEnvironment() {
  initDb();
  
  // 既存の同名トリガーを削除して再作成
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'deleteOldReservations') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('deleteOldReservations')
    .timeBased()
    .everyDays(1)
    .atHour(2) // 毎日深夜2時に実行
    .create();
}
