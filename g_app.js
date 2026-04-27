// GETリクエスト（データ受信）を受け取る関数
function doGet(e) {
  const id = e.parameter.id;
  const result = updateUserStatus(id);
  
  // JSON形式で結果を返す（CORS対策）
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function updateUserStatus(scannedId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('参加者リスト');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == scannedId) {
      if (data[i][2] === '出席') {
        return { success: false, message: `既に受付済み: ${data[i][1]}様`, color: 'orange' };
      }
      const row = i + 1;
      const now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
      sheet.getRange(row, 3).setValue('出席');
      sheet.getRange(row, 4).setValue(now);
      return { success: true, message: `${data[i][1]}様、ようこそ！`, color: '#28a745' };
    }
  }
  return { success: false, message: '未登録のIDです', color: '#dc3545' };
}