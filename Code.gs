/**
 * 在庫監視システム（最小構成版）
 * 条件：B列の在庫数が10未満の行を検出
 */
function checkInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 各シートを取得
  const dataSheet = ss.getSheetByName('データ');
  const configSheet = ss.getSheetByName('設定');
  const logSheet = ss.getSheetByName('ログ');
  
  if (!dataSheet || !configSheet || !logSheet) {
    Logger.log('エラー: 必要なシートが見つかりません');
    return;
  }
  
  // 設定値を読み込み
  const config = getConfig(configSheet);
  Logger.log('設定値: ' + JSON.stringify(config));
  
  // 列番号に変換
  const stockCol = columnToIndex(config.monitorColumn);
  const dateCol = columnToIndex(config.dateColumn);
  const statusCol = columnToIndex(config.statusColumn);
  
  // データ範囲を取得（全列）
  const lastRow = dataSheet.getLastRow();
  const lastCol = dataSheet.getLastColumn();
  
  if (lastRow < config.startRow) {
    Logger.log('データが存在しません');
    return;
  }
  
  const dataRange = dataSheet.getRange(config.startRow, 1, lastRow - config.startRow + 1, lastCol);
  const data = dataRange.getValues();
  
  // アラートを種類別に分類
  const alerts = {
    stock: [],      // 在庫アラート
    deadline: [],   // 納期アラート
    status: [],     // ステータスアラート
    multiple: []    // 複合条件アラート
  };
  
  // 各行をチェック
  data.forEach((row, index) => {
    const rowNumber = index + config.startRow;
    const productName = row[0];
    
    // 各条件のチェック結果
    const checks = {
      stock: false,
      deadline: false,
      status: false
    };
    
    // 1. 在庫チェック
    if (stockCol && typeof row[stockCol - 1] === 'number') {
      const stock = row[stockCol - 1];
      if (stock < config.threshold) {
        checks.stock = true;
      }
    }
    
    // 2. 納期チェック
    if (dateCol && row[dateCol - 1]) {
      const deadline = row[dateCol - 1];
      if (isDateWithinDays(deadline, config.dateDaysThreshold)) {
        checks.deadline = true;
        
        // 残り日数を計算
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const targetDate = new Date(deadline);
        targetDate.setHours(0, 0, 0, 0);
        const daysLeft = Math.ceil((targetDate - today) / (1000 * 60 * 60 * 24));
        
        checks.daysLeft = daysLeft;
      }
    }
    
    // 3. ステータスチェック
    if (statusCol && row[statusCol - 1]) {
      const status = row[statusCol - 1];
      if (status === config.targetStatus) {
        checks.status = true;
      }
    }
    
    // アラートの分類と記録
    const matchCount = [checks.stock, checks.deadline, checks.status].filter(Boolean).length;
    
    if (matchCount === 0) {
      return; // アラート対象外
    }
    
    // メッセージ構築
    let message = `${productName}:`;
    const details = [];
    
    if (checks.stock) {
      details.push(`在庫${row[stockCol - 1]}個`);
    }
    if (checks.deadline) {
      details.push(`納期まで${checks.daysLeft}日`);
    }
    if (checks.status) {
      details.push(`${row[statusCol - 1]}`);
    }
    
    message += ' ' + details.join(', ');
    
    // 複数条件マッチの場合
    if (matchCount >= 2) {
      alerts.multiple.push(message);
      Logger.log(`🚨 【複合】${message}（行${rowNumber}）`);
    } else if (checks.stock) {
      alerts.stock.push(message);
      Logger.log(`⚠️ 【在庫】${message}（行${rowNumber}）`);
    } else if (checks.deadline) {
      alerts.deadline.push(message);
      Logger.log(`📅 【納期】${message}（行${rowNumber}）`);
    } else if (checks.status) {
      alerts.status.push(message);
      Logger.log(`📋 【ステータス】${message}（行${rowNumber}）`);
    }
    
    // 最終チェック時刻を記録
    dataSheet.getRange(rowNumber, 3).setValue(new Date());
  });
  
  // 全アラートを統合
  const allAlerts = [
    ...alerts.multiple,
    ...alerts.stock,
    ...alerts.deadline,
    ...alerts.status
  ];
  
  // ログシートに記録
  writeLog(logSheet, allAlerts);
  
  // 結果サマリー
  const totalCount = allAlerts.length;
  
  if (totalCount > 0) {
    Logger.log('\n=== 検出結果 ===');
    Logger.log(`複合条件: ${alerts.multiple.length}件`);
    Logger.log(`在庫のみ: ${alerts.stock.length}件`);
    Logger.log(`納期のみ: ${alerts.deadline.length}件`);
    Logger.log(`ステータスのみ: ${alerts.status.length}件`);
    Logger.log(`合計: ${totalCount}件`);
    
    // LINE通知
    let lineMessage = '⚠️ アラート通知\n\n';
    
    if (alerts.multiple.length > 0) {
      lineMessage += '🚨【複合条件】\n' + alerts.multiple.join('\n') + '\n\n';
    }
    if (alerts.stock.length > 0) {
      lineMessage += '📦【在庫】\n' + alerts.stock.join('\n') + '\n\n';
    }
    if (alerts.deadline.length > 0) {
      lineMessage += '📅【納期】\n' + alerts.deadline.join('\n') + '\n\n';
    }
    if (alerts.status.length > 0) {
      lineMessage += '📋【ステータス】\n' + alerts.status.join('\n') + '\n\n';
    }
    
    lineMessage += `合計: ${totalCount}件`;
    
    sendLineMessage(lineMessage);
    
  } else {
    Logger.log('アラート対象なし');
  }
}

function getConfig(configSheet){
  //基本設定
  const basicConfig = configSheet.getRange('B2:B5').getValues();
  //日付監視設定
  const dateConfig = configSheet.getRange('B7:B8').getValues();
  //ステータス監視設定
  const statusConfig = configSheet.getRange('B10:B11').getValues();

  return{
    //基本設定
    monitorColumn: basicConfig[0][0],
    threshold: basicConfig[1][0],
    startRow: basicConfig[2][0],
    lineToken: basicConfig[3][0],

    //日付監視設定
    dateColumn: dateConfig[0][0],
    dateDaysThreshold: dateConfig[1][0],

    //ステータス監視設定
    statusColumn: statusConfig[0][0],
    targetStatus: statusConfig[1][0]
  };
}

function writeLog(logSheet, alerts){
  const now = new Date();
  const alertCount = alerts.length;
  const details = alertCount > 0 ? alerts.join(', '):'アラートなし';

  //ログシートの最終行に追加
  logSheet.appendRow([now, alertCount, details]);

  Logger.log(`ログに記録しました：${alertCount}件`);
}

//トリガーを設定するセットアップ関数。この関数を1回だけ手動実行してトリガーを作成する
function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  //時間ベーストリガー
  ScriptApp.newTrigger('checkInventory')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅トリガーを設定しました：1時間ごとに実行');
  Browser.msgBox('✅ 設定完了', 'トリガーを設定しました。\n1時間ごとに自動チェックします。', Browser.Buttons.OK);

  // 補足: 他のトリガーパターン例
  // 毎日9時に実行:
  // ScriptApp.newTrigger('checkInventory')
  //   .timeBased()
  //   .atHour(9)
  //   .everyDays(1)
  //   .create();
}

//トリガーを削除する関数
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log('✅すべてのトリガーを削除しました');
}

//現在設定されているトリガーを確認する関数
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  if(triggers.length === 0){
    Logger.log('設定されているトリガーはありません');
    return;
  }

  Logger.log(`=== 設定中のトリガー一覧(${triggers.length}件)===`);
  triggers.forEach((trigger, index) => {
    Logger.log(`${index + 1}. 関数：${trigger.getHandlerFunction()}`);
    Logger.log(` 種類：${trigger.getEventType()}`);
  });
}

//グローバル変数として最後の実行時刻を保持
let lastEditTime = 0;

//スプレッドシート編集時に自動実行される関数。在庫数が編集されたら即座にチェック
function onEdit(e){
  //重複実行防止：1秒以内の再実行は無視する
  const now = new Date().getTime();
  if(now - lastEditTime < 1000){
    Logger.log('重複実行を防止しました');
    return;
  }
  lastEditTime = now;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('ログ');

  //デバック：onEditが呼ばれたことを記録する
  //logSheet.appendRow([new Date(), 'DEBUG', 'onEdit関数が呼ばれました']);

  //編集されたセルの情報を取得
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const col = range.getColumn();
  const row = range.getRow();

  //デバック情報を記録
  //const debugInfo = `シート：${sheetName}, 行：${row}, 列：${col}`;
  //logSheet.appendRow([new Date(), 'DEBUG', debugInfo]);

  //データシート以外の編集は無視
  if(sheet.getName() !== 'データ'){
    //logSheet.appendRow([new Date(), 'DEBUG', 'データシート以外なのでスキップ']);
    return;
  }

  //在庫数の編集のみ対応
  if(col !== 2){
    //logSheet.appendRow([new Date(), 'DEBUG', 'B列以外なのでスキップ']);
    return;
  }

  Logger.log('在庫数が編集されました。チェックを実行します...');
  //logSheet.appendRow([new Date(), 'DEBUG', 'checkInventoryを実行します']);
  checkInventory();
}

//スプレッドシート起動時にカスタムメニューを追加
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊在庫監視')
    .addItem('🔍今すぐチェック実行','checkInventory')
    .addSeparator()
    .addItem('⚙️トリガー設定','setupTriggers')
    .addItem('🗑️トリガー削除','deleteTriggers')
    .addItem('📋トリガー確認','listTriggers')
    .addToUi();
}

//LINEにメッセージを送信する関数
function sendLineMessage(message){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('設定');

  //Channel Access Tokenを取得する
  const token = configSheet.getRange('B5').getValue();

  if(!token){
    Logger.log('エラー：LINE Channel Access Tokenが設定されていません');
    return false;
  }

  //LINEのエンドポイント（宛先）。
  const url = 'https://api.line.me/v2/bot/message/broadcast';

  //荷物の中身
  const payload = {
    messages: [
      {
        type:'text',
        text:message
      }
    ]
  };

  //APIリクエストのオプション
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),//荷物を通信用の文字データに変換する
    muteHttpExceptions: true //エラーが起きてもプログラムを強制停止させない
  };

  try{
    //送信実行
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if(responseCode === 200){
      Logger.log('✅LINEメッセージ送信成功');
      return true;
    }else{
      Logger.log('❌LINEメッセージ送信失敗：'+responseCode);
      Logger.log(response.getContentText());
      return false;
    }
  }catch(error){
    Logger.log('❌LINE送信エラー：'+error);
    return false;
  }
}

//列A,B,Cを列番号1,2,3に変換する
function columnToIndex(column){
  if(!column) return null;

  column = column.toUpperCase();
  let index = 0;

  for(let i=0; i<column.length;i++){
    index = index * 26 + (column.charCodeAt(i) - 64);
  }
  return index;
}

//日付が指定日数以内かチェック
function isDateWithinDays(dateValue, days){
  if(!dateValue || !(dateValue instanceof Date)) {
    return false;
  }

  const today = new Date();
  today.setHours(0,0,0,0)

  const targetDate = new Date(dateValue);
  targetDate.setHours(0,0,0,0);

  //日数差を計算
  const diffTime = targetDate - today;
  const diffDays = Math.ceil(diffTime / (1000*60*60*24));

  return diffDays >= 0 && diffDays <=days;
}



