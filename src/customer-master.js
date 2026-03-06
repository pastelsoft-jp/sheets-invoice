/**
 * 顧客マスタ CRUD
 * PropertiesServiceに保存（最大50件）
 */

var CUSTOMERS_KEY = 'invoice_maker_customers';
var MAX_CUSTOMERS = 50;

/**
 * 全顧客を取得
 * @return {Array} 顧客リスト
 */
function getCustomers() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(CUSTOMERS_KEY);
    if (!raw) return [];
    return JSON.parse(raw);
  } catch (e) {
    Logger.log('顧客一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * 顧客を保存（新規or更新）
 * @param {Object} customerData - {id?, name, address, contactPerson, email, defaultItems}
 * @return {Object} 保存した顧客データ
 */
function saveCustomer(customerData) {
  var plan = checkSubscription();
  if (plan.plan === 'free') {
    throw new Error('顧客マスタはBasicプラン以上の機能です。');
  }

  var customers = getCustomers();

  if (customerData.id) {
    // 更新
    var found = false;
    for (var i = 0; i < customers.length; i++) {
      if (customers[i].id === customerData.id) {
        customers[i] = customerData;
        found = true;
        break;
      }
    }
    if (!found) {
      throw new Error('指定された顧客が見つかりません。');
    }
  } else {
    // 新規
    if (customers.length >= MAX_CUSTOMERS) {
      throw new Error('顧客マスタの上限（' + MAX_CUSTOMERS + '件）に達しています。');
    }
    customerData.id = 'cust_' + new Date().getTime();
    customers.push(customerData);
  }

  try {
    PropertiesService.getUserProperties().setProperty(CUSTOMERS_KEY, JSON.stringify(customers));
  } catch (e) {
    Logger.log('顧客保存エラー: ' + e.message);
    throw new Error('顧客の保存に失敗しました。');
  }

  return customerData;
}

/**
 * 顧客を削除
 * @param {string} customerId
 */
function deleteCustomer(customerId) {
  var customers = getCustomers();
  var filtered = [];
  for (var i = 0; i < customers.length; i++) {
    if (customers[i].id !== customerId) {
      filtered.push(customers[i]);
    }
  }

  PropertiesService.getUserProperties().setProperty(CUSTOMERS_KEY, JSON.stringify(filtered));
}

/**
 * 顧客をIDで取得
 * @param {string} customerId
 * @return {Object|null}
 */
function getCustomerById(customerId) {
  var customers = getCustomers();
  for (var i = 0; i < customers.length; i++) {
    if (customers[i].id === customerId) {
      return customers[i];
    }
  }
  return null;
}
