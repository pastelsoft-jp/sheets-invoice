/**
 * 請求書データの管理
 * 発行履歴をPropertiesServiceに保存
 */

var HISTORY_KEY = 'invoice_maker_history';
var MAX_HISTORY = 100;

/**
 * 請求書データを構築する
 * @param {Object} formData - サイドバーからの入力データ
 * @return {Object} 構造化された請求書データ
 */
function buildInvoiceData(formData) {
  var settings = loadSettings();
  var items = formData.items || [];

  // 税額計算
  var calc = calculateInvoiceTotal(items, formData.withholdingEnabled);

  var invoiceData = {
    invoiceNumber: formData.invoiceNumber || generateInvoiceNumber(),
    issueDate: formData.issueDate || formatDateISO_(new Date()),
    dueDate: formData.dueDate || calculateDueDate_(settings.defaultDueDays),

    issuer: {
      name: settings.issuer.name,
      registrationNumber: settings.issuer.registrationNumber,
      address: settings.issuer.address,
      phone: settings.issuer.phone,
      email: settings.issuer.email
    },

    client: {
      name: formData.clientName || '',
      address: formData.clientAddress || '',
      contactPerson: formData.clientContactPerson || ''
    },

    items: items,

    subtotal: calc.subtotal,
    subtotal10: calc.subtotal10,
    subtotal8: calc.subtotal8,
    tax10: calc.tax10,
    tax8: calc.tax8,
    totalTax: calc.totalTax,
    totalWithTax: calc.totalWithTax,
    withholdingEnabled: formData.withholdingEnabled || false,
    withholding: calc.withholding,
    grandTotal: calc.grandTotal,

    bankAccount: settings.bankAccount,
    note: formData.note || settings.defaultNote,
    logoUrl: settings.logoUrl,

    createdAt: new Date().toISOString()
  };

  return invoiceData;
}

/**
 * 請求書を発行履歴に保存
 * @param {Object} invoiceData
 */
function saveInvoiceHistory(invoiceData) {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(HISTORY_KEY);
    var history = raw ? JSON.parse(raw) : [];

    // 先頭に追加
    history.unshift({
      invoiceNumber: invoiceData.invoiceNumber,
      clientName: invoiceData.client.name,
      grandTotal: invoiceData.grandTotal,
      issueDate: invoiceData.issueDate,
      createdAt: invoiceData.createdAt
    });

    // 上限を超えたら古いものを削除
    if (history.length > MAX_HISTORY) {
      history = history.slice(0, MAX_HISTORY);
    }

    PropertiesService.getUserProperties().setProperty(HISTORY_KEY, JSON.stringify(history));
  } catch (e) {
    Logger.log('履歴保存エラー: ' + e.message);
  }
}

/**
 * 発行履歴を取得
 * @return {Array}
 */
function getInvoiceHistory() {
  try {
    var plan = checkSubscription();
    if (plan.plan !== 'pro') {
      throw new Error('履歴管理はProプランの機能です。');
    }

    var raw = PropertiesService.getUserProperties().getProperty(HISTORY_KEY);
    if (!raw) return [];
    return JSON.parse(raw);
  } catch (e) {
    if (e.message.indexOf('Proプラン') !== -1) throw e;
    Logger.log('履歴取得エラー: ' + e.message);
    return [];
  }
}

/**
 * ISO形式の日付文字列を返す
 */
function formatDateISO_(date) {
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

/**
 * 支払期限を計算する
 */
function calculateDueDate_(days) {
  var due = new Date();
  due.setDate(due.getDate() + (days || 30));
  return formatDateISO_(due);
}

/**
 * 日付を日本語フォーマットで返す
 */
function formatDateJP(dateStr) {
  if (!dateStr) return '';
  var parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  return parts[0] + '年' + parseInt(parts[1]) + '月' + parseInt(parts[2]) + '日';
}
