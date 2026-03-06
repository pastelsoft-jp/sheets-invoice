/**
 * 消費税・源泉徴収税額の計算
 */

/**
 * 品目リストから税額を計算する
 * @param {Array} items - [{description, quantity, unitPrice, taxRate}]
 * @return {Object} 計算結果
 */
function calculateTax(items) {
  var subtotal10 = 0;
  var subtotal8 = 0;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var amount = (item.quantity || 0) * (item.unitPrice || 0);
    if (item.taxRate === 8) {
      subtotal8 += amount;
    } else {
      subtotal10 += amount;
    }
  }

  var tax10 = Math.floor(subtotal10 * 0.1);
  var tax8 = Math.floor(subtotal8 * 0.08);
  var subtotal = subtotal10 + subtotal8;
  var totalTax = tax10 + tax8;

  return {
    subtotal: subtotal,
    subtotal10: subtotal10,
    subtotal8: subtotal8,
    tax10: tax10,
    tax8: tax8,
    totalTax: totalTax,
    totalWithTax: subtotal + totalTax
  };
}

/**
 * 源泉徴収税額を計算する
 * 支払金額 <= 100万円: 支払金額 x 10.21%
 * 支払金額 > 100万円: (100万 x 10.21%) + (超過分 x 20.42%)
 * @param {number} amount - 税込支払金額
 * @return {number} 源泉徴収税額（端数切り捨て）
 */
function calculateWithholding(amount) {
  if (!amount || amount <= 0) return 0;

  if (amount <= 1000000) {
    return Math.floor(amount * 0.1021);
  } else {
    return Math.floor(1000000 * 0.1021 + (amount - 1000000) * 0.2042);
  }
}

/**
 * 請求書の全額を計算する
 * @param {Array} items - 品目リスト
 * @param {boolean} withholdingEnabled - 源泉徴収を適用するか
 * @return {Object} 全計算結果
 */
function calculateInvoiceTotal(items, withholdingEnabled) {
  var tax = calculateTax(items);
  var withholding = 0;

  if (withholdingEnabled) {
    withholding = calculateWithholding(tax.totalWithTax);
  }

  return {
    subtotal: tax.subtotal,
    subtotal10: tax.subtotal10,
    subtotal8: tax.subtotal8,
    tax10: tax.tax10,
    tax8: tax.tax8,
    totalTax: tax.totalTax,
    totalWithTax: tax.totalWithTax,
    withholding: withholding,
    grandTotal: tax.totalWithTax - withholding
  };
}

/**
 * 登録番号のバリデーション（T + 13桁数字）
 * @param {string} number - 登録番号
 * @return {boolean}
 */
function validateRegistrationNumber(number) {
  if (!number) return false;
  return /^T\d{13}$/.test(number);
}

/**
 * 金額をフォーマットする
 * @param {number} amount
 * @return {string} 例: "1,234,567"
 */
function formatCurrency(amount) {
  if (amount === null || amount === undefined) return '0';
  return Math.floor(amount).toLocaleString('ja-JP');
}
