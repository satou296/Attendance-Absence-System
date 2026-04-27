function doGet() {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setTitle('QR受付システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * QRコードから読み取ったIDを元にスプレッドシートを更新する
 */
function updateUserStatus(scannedId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('参加者リスト');
  const data = sheet.getDataRange().getValues();
  
  // ヘッダーを除いた2行目から探索
  for (let i = 1; i < data.length; i++) {
    // Connpass IDが一致するか確認 (数値・文字列両対応のため == )
    if (data[i][0] == scannedId) {
      
      // 既に「出席」の場合は警告を返す
      if (data[i][2] === '出席') {
        return {
          success: false,
          message: `既に受付済みです: ${data[i][1]} さん`,
          color: 'orange'
        };
      }
      
      // ステータスと時間を更新
      const row = i + 1;
      const now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
      sheet.getRange(row, 3).setValue('出席');    // C列
      sheet.getRange(row, 4).setValue(now);       // D列
      
      return {
        success: true,
        message: `${data[i][1]} さん、ようこそ！`,
        color: '#28a745'
      };
    }
  }
  
  // IDが見つからなかった場合
  return {
    success: false,
    message: '未登録のユーザーです',
    color: '#dc3545'
  };
}