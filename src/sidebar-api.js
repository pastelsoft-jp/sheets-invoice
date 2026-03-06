/**
 * サイドバーHTMLから呼ばれるAPI関数群
 * google.script.run でアクセスされる
 */

/**
 * サイドバーからオンボーディング完了
 */
function completeOnboardingFromSidebar() {
  var settings = loadSettings();
  settings.hasSeenOnboarding = true;
  saveSettings(settings);
}

/**
 * サイドバーから事業者情報を保存（オンボーディング）
 */
function saveIssuerFromSidebar(issuerData) {
  try {
    var settings = loadSettings();
    settings.hasSeenOnboarding = true;
    settings.issuer = {
      name: issuerData.name || '',
      registrationNumber: issuerData.registrationNumber || '',
      address: issuerData.address || '',
      phone: issuerData.phone || '',
      email: issuerData.email || ''
    };
    if (issuerData.bankName) {
      settings.bankAccount = {
        bankName: issuerData.bankName || '',
        branchName: issuerData.branchName || '',
        accountType: issuerData.accountType || '普通',
        accountNumber: issuerData.accountNumber || '',
        accountHolder: issuerData.accountHolder || ''
      };
    }
    saveSettings(settings);
  } catch (error) {
    Logger.log('saveIssuerFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから請求書を作成
 */
function createInvoiceFromSidebar(formData) {
  try {
    return JSON.stringify(createInvoice(formData));
  } catch (error) {
    Logger.log('createInvoiceFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから税額をリアルタイム計算
 */
function calculateTaxFromSidebar(items, withholdingEnabled) {
  try {
    return JSON.stringify(calculateInvoiceTotal(items, withholdingEnabled));
  } catch (error) {
    Logger.log('calculateTaxFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから顧客一覧取得
 */
function getCustomersFromSidebar() {
  try {
    return JSON.stringify(getCustomers());
  } catch (error) {
    Logger.log('getCustomersFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから顧客保存
 */
function saveCustomerFromSidebar(customerData) {
  try {
    return JSON.stringify(saveCustomer(customerData));
  } catch (error) {
    Logger.log('saveCustomerFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから顧客削除
 */
function deleteCustomerFromSidebar(customerId) {
  try {
    deleteCustomer(customerId);
  } catch (error) {
    Logger.log('deleteCustomerFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから設定保存
 */
function saveSettingsFromSidebar(uiSettings) {
  try {
    var settings = loadSettings();

    if (uiSettings.issuer) settings.issuer = uiSettings.issuer;
    if (uiSettings.bankAccount) settings.bankAccount = uiSettings.bankAccount;
    if (uiSettings.invoicePrefix !== undefined) settings.invoicePrefix = uiSettings.invoicePrefix;
    if (uiSettings.withholdingEnabled !== undefined) settings.withholdingEnabled = uiSettings.withholdingEnabled;
    if (uiSettings.defaultNote !== undefined) settings.defaultNote = uiSettings.defaultNote;
    if (uiSettings.defaultDueDays !== undefined) settings.defaultDueDays = uiSettings.defaultDueDays;

    saveSettings(settings);
  } catch (error) {
    Logger.log('saveSettingsFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから設定リセット
 */
function resetSettingsFromSidebar() {
  try {
    PropertiesService.getUserProperties().deleteAllProperties();
  } catch (error) {
    Logger.log('resetSettingsFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}

/**
 * サイドバーから発行履歴取得
 */
function getInvoiceHistoryFromSidebar() {
  try {
    return JSON.stringify(getInvoiceHistory());
  } catch (error) {
    Logger.log('getInvoiceHistoryFromSidebar error: ' + error.message);
    throw new Error(error.message);
  }
}
