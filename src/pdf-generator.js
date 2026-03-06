/**
 * PDF生成ロジック
 */

/**
 * 指定シートをPDFに変換し、Google Driveに保存する
 * @param {string} sheetName - PDFにするシート名
 * @return {Object} { fileId, fileUrl, downloadUrl }
 */
function generatePdf(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('シート「' + sheetName + '」が見つかりません。');
  }

  var ssId = ss.getId();
  var sheetId = sheet.getSheetId();

  // PDF生成パラメータ（A4縦）
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' +
    'exportFormat=pdf' +
    '&format=pdf' +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +        // 幅に合わせる
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +
    '&gridlines=false' +
    '&fzr=false' +
    '&gid=' + sheetId +
    '&top_margin=0.5' +
    '&bottom_margin=0.5' +
    '&left_margin=0.75' +
    '&right_margin=0.75';

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF生成に失敗しました。しばらくしてからお試しください。');
  }

  var blob = response.getBlob().setName(sheetName + '.pdf');

  // 保存先フォルダを取得or作成
  var folder = getOrCreateInvoiceFolder_();
  var file = folder.createFile(blob);

  return {
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    fileName: sheetName + '.pdf'
  };
}

/**
 * 請求書保存用フォルダを取得or作成
 * @return {Folder}
 */
function getOrCreateInvoiceFolder_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssFile = DriveApp.getFileById(ss.getId());
  var parents = ssFile.getParents();
  var parentFolder;

  if (parents.hasNext()) {
    parentFolder = parents.next();
  } else {
    parentFolder = DriveApp.getRootFolder();
  }

  // 「請求書PDF」フォルダを探す
  var folderName = '請求書PDF';
  var folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(folderName);
}

/**
 * 請求書を作成→PDF生成まで一気に行う
 * @param {Object} formData - サイドバーからの入力データ
 * @return {Object} { invoiceData, pdf }
 */
function createInvoice(formData) {
  // 月間制限チェック
  var limitCheck = checkMonthlyLimit();
  if (!limitCheck.allowed) {
    throw new Error('今月の無料枠（' + limitCheck.limit + '枚）を超えました。Basicプラン以上にアップグレードしてください。');
  }

  // 請求書データ構築
  var invoiceData = buildInvoiceData(formData);

  // シートに書き出し
  var sheetResult = writeInvoiceToSheet(invoiceData);

  // シートの変更を確実に反映してからPDF生成
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  // PDF生成
  var pdfResult = generatePdf(sheetResult.sheetName);

  // 使用量カウント
  incrementMonthlyUsage();

  // 履歴保存
  saveInvoiceHistory(invoiceData);

  return {
    invoiceData: invoiceData,
    pdf: pdfResult,
    sheetName: sheetResult.sheetName
  };
}
