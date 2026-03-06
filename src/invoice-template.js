/**
 * 請求書テンプレートをスプレッドシートに生成する
 */

/**
 * 請求書をシートに書き出す
 * @param {Object} invoiceData - buildInvoiceData()の戻り値
 * @return {Object} { sheetName: string, spreadsheetId: string }
 */
function writeInvoiceToSheet(invoiceData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '請求書';

  // 既存の請求書シートがあれば削除して作り直す
  var existing = ss.getSheetByName(sheetName);
  if (existing) {
    ss.deleteSheet(existing);
  }

  var sheet = ss.insertSheet(sheetName);

  // シート設定（A4縦に近いレイアウト）
  sheet.setColumnWidth(1, 40);   // No.
  sheet.setColumnWidth(2, 220);  // 品目
  sheet.setColumnWidth(3, 70);   // 数量
  sheet.setColumnWidth(4, 100);  // 単価
  sheet.setColumnWidth(5, 60);   // 税率
  sheet.setColumnWidth(6, 110);  // 金額
  // 合計幅: 600px ≒ A4幅。不要な列を非表示にする
  for (var c = 7; c <= sheet.getMaxColumns(); c++) {
    sheet.setColumnWidth(c, 1);
  }
  if (sheet.getMaxColumns() > 6) {
    sheet.hideColumns(7, sheet.getMaxColumns() - 6);
  }

  // ========== ヘッダーエリア ==========
  var row = 1;

  // タイトル
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('請 求 書')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  row += 2;

  // 請求書番号・日付（右寄せ）
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('請求書番号: ' + invoiceData.invoiceNumber)
    .setHorizontalAlignment('right').setFontSize(10);
  row++;

  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('発行日: ' + formatDateJP(invoiceData.issueDate))
    .setHorizontalAlignment('right').setFontSize(10);
  row++;

  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('支払期限: ' + formatDateJP(invoiceData.dueDate))
    .setHorizontalAlignment('right').setFontSize(10);
  row += 2;

  // 請求先
  sheet.getRange(row, 1, 1, 3).merge();
  sheet.getRange(row, 1).setValue(invoiceData.client.name + ' 御中')
    .setFontSize(14).setFontWeight('bold');
  row++;

  if (invoiceData.client.address) {
    sheet.getRange(row, 1, 1, 3).merge();
    sheet.getRange(row, 1).setValue(invoiceData.client.address).setFontSize(10);
    row++;
  }
  row++;

  // ご請求金額
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('下記の通りご請求申し上げます。')
    .setFontSize(10);
  row++;

  sheet.getRange(row, 1, 1, 3).merge();
  sheet.getRange(row, 1).setValue('ご請求金額')
    .setFontSize(10).setFontWeight('bold');
  sheet.getRange(row, 4, 1, 3).merge();
  sheet.getRange(row, 4).setValue('¥' + formatCurrency(invoiceData.grandTotal))
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('right');

  // 下線
  sheet.getRange(row, 1, 1, 6).setBorder(null, null, true, null, null, null,
    '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  row += 2;

  // ========== 品目テーブル ==========
  var tableStartRow = row;

  // ヘッダー行
  var headers = ['No.', '品目', '数量', '単価', '税率', '金額'];
  sheet.getRange(row, 1, 1, 6).setValues([headers])
    .setFontWeight('bold').setBackground('#f0f0f0')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  row++;

  // 品目行
  var items = invoiceData.items || [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var amount = (item.quantity || 0) * (item.unitPrice || 0);
    var taxLabel = (item.taxRate === 8) ? '8%' : '10%';

    sheet.getRange(row, 1, 1, 6).setValues([[
      i + 1,
      item.description || '',
      item.quantity || 0,
      item.unitPrice || 0,
      taxLabel,
      amount
    ]]).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

    // 数値フォーマット
    sheet.getRange(row, 3).setNumberFormat('#,##0');
    sheet.getRange(row, 4).setNumberFormat('#,##0');
    sheet.getRange(row, 6).setNumberFormat('#,##0');
    sheet.getRange(row, 5).setHorizontalAlignment('center');

    row++;
  }

  // 空行の追加（最低5行にする）
  var minRows = 5;
  while (items.length < minRows && row < tableStartRow + 1 + minRows) {
    sheet.getRange(row, 1, 1, 6).setValues([['', '', '', '', '', '']])
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    row++;
    minRows--;
  }

  row++;

  // ========== 合計エリア ==========
  var summaryCol = 4; // D列から

  // 小計
  sheet.getRange(row, summaryCol, 1, 2).merge();
  sheet.getRange(row, summaryCol).setValue('小計').setHorizontalAlignment('right');
  sheet.getRange(row, 6).setValue(invoiceData.subtotal).setNumberFormat('#,##0').setHorizontalAlignment('right');
  row++;

  // 10%対象
  if (invoiceData.subtotal10 > 0) {
    sheet.getRange(row, summaryCol, 1, 2).merge();
    sheet.getRange(row, summaryCol).setValue('10%対象計').setHorizontalAlignment('right');
    sheet.getRange(row, 6).setValue(invoiceData.subtotal10).setNumberFormat('#,##0').setHorizontalAlignment('right');
    row++;

    sheet.getRange(row, summaryCol, 1, 2).merge();
    sheet.getRange(row, summaryCol).setValue('消費税（10%）').setHorizontalAlignment('right');
    sheet.getRange(row, 6).setValue(invoiceData.tax10).setNumberFormat('#,##0').setHorizontalAlignment('right');
    row++;
  }

  // 8%対象
  if (invoiceData.subtotal8 > 0) {
    sheet.getRange(row, summaryCol, 1, 2).merge();
    sheet.getRange(row, summaryCol).setValue('8%対象計（軽減税率）').setHorizontalAlignment('right');
    sheet.getRange(row, 6).setValue(invoiceData.subtotal8).setNumberFormat('#,##0').setHorizontalAlignment('right');
    row++;

    sheet.getRange(row, summaryCol, 1, 2).merge();
    sheet.getRange(row, summaryCol).setValue('消費税（8%）').setHorizontalAlignment('right');
    sheet.getRange(row, 6).setValue(invoiceData.tax8).setNumberFormat('#,##0').setHorizontalAlignment('right');
    row++;
  }

  // 源泉徴収
  if (invoiceData.withholdingEnabled && invoiceData.withholding > 0) {
    sheet.getRange(row, summaryCol, 1, 2).merge();
    sheet.getRange(row, summaryCol).setValue('源泉徴収税額').setHorizontalAlignment('right');
    sheet.getRange(row, 6).setValue(-invoiceData.withholding).setNumberFormat('#,##0').setHorizontalAlignment('right');
    row++;
  }

  // 合計
  sheet.getRange(row, summaryCol, 1, 2).merge();
  sheet.getRange(row, summaryCol).setValue('合計').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(row, 6).setValue(invoiceData.grandTotal).setNumberFormat('#,##0')
    .setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(row, summaryCol, 1, 3).setBorder(null, null, true, null, null, null,
    '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  row += 2;

  // ========== 区切り線 ==========
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1, 1, 6).setBorder(null, null, true, null, null, null,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  row++;

  // ========== 振込先・登録番号ボックス ==========
  var bank = invoiceData.bankAccount;
  var hasBank = bank && bank.bankName;
  var hasRegNum = invoiceData.issuer.registrationNumber;

  if (hasBank || hasRegNum) {
    var boxStartRow = row;

    // 振込先ヘッダー
    sheet.getRange(row, 1, 1, 6).merge();
    sheet.getRange(row, 1).setValue('お振込先')
      .setFontWeight('bold').setFontSize(10)
      .setBackground('#f8f8f8').setFontColor('#333333');
    sheet.getRange(row, 1, 1, 6).setBackground('#f8f8f8');
    row++;

    if (hasBank) {
      // 銀行名・支店
      sheet.getRange(row, 1, 1, 3).merge();
      sheet.getRange(row, 1).setValue(bank.bankName + '　' + bank.branchName)
        .setFontSize(11).setFontWeight('bold');
      sheet.getRange(row, 4, 1, 3).merge();
      sheet.getRange(row, 4).setValue(bank.accountType + '　' + bank.accountNumber)
        .setFontSize(11).setHorizontalAlignment('right');
      row++;

      // 口座名義
      sheet.getRange(row, 1, 1, 6).merge();
      sheet.getRange(row, 1).setValue('口座名義：' + bank.accountHolder)
        .setFontSize(10).setFontColor('#555555');
      row++;
    }

    if (hasRegNum) {
      sheet.getRange(row, 1, 1, 6).merge();
      sheet.getRange(row, 1).setValue('適格請求書発行事業者登録番号：' + invoiceData.issuer.registrationNumber)
        .setFontSize(10).setFontColor('#555555');
      row++;
    }

    // ボックスの枠線
    sheet.getRange(boxStartRow, 1, row - boxStartRow, 6)
      .setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

    row++;
  }

  // ========== 備考ボックス ==========
  if (invoiceData.note) {
    var noteStartRow = row;

    sheet.getRange(row, 1, 1, 6).merge();
    sheet.getRange(row, 1).setValue('備考')
      .setFontWeight('bold').setFontSize(10)
      .setBackground('#f8f8f8').setFontColor('#333333');
    sheet.getRange(row, 1, 1, 6).setBackground('#f8f8f8');
    row++;

    sheet.getRange(row, 1, 1, 6).merge();
    sheet.getRange(row, 1).setValue(invoiceData.note)
      .setFontSize(10).setFontColor('#555555')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');
    row++;

    sheet.getRange(noteStartRow, 1, row - noteStartRow, 6)
      .setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

    row++;
  }

  row++;

  // ========== 発行者情報（右寄せ） ==========
  sheet.getRange(row, 4, 1, 3).merge();
  sheet.getRange(row, 4).setValue(invoiceData.issuer.name)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('right');
  row++;

  if (invoiceData.issuer.address) {
    sheet.getRange(row, 4, 1, 3).merge();
    sheet.getRange(row, 4).setValue(invoiceData.issuer.address)
      .setFontSize(9).setHorizontalAlignment('right').setFontColor('#555555');
    row++;
  }

  var contactParts = [];
  if (invoiceData.issuer.phone) contactParts.push('TEL: ' + invoiceData.issuer.phone);
  if (invoiceData.issuer.email) contactParts.push(invoiceData.issuer.email);
  if (contactParts.length > 0) {
    sheet.getRange(row, 4, 1, 3).merge();
    sheet.getRange(row, 4).setValue(contactParts.join('  |  '))
      .setFontSize(9).setHorizontalAlignment('right').setFontColor('#555555');
    row++;
  }

  // 印刷範囲を設定
  sheet.getRange(1, 1, row, 6).setFontFamily('Noto Sans JP');

  return {
    sheetName: sheetName,
    spreadsheetId: ss.getId()
  };
}
