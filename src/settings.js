/**
 * ユーザー設定の保存・読み込み
 */

var SETTINGS_KEY = 'invoice_maker_settings';
var SUBSCRIPTION_KEY = 'invoice_maker_subscription';
var USAGE_KEY = 'invoice_maker_usage';

/**
 * デフォルト設定
 */
function getDefaultSettings_() {
  return {
    hasSeenOnboarding: false,
    issuer: {
      name: '',
      registrationNumber: '',
      address: '',
      phone: '',
      email: ''
    },
    bankAccount: {
      bankName: '',
      branchName: '',
      accountType: '普通',
      accountNumber: '',
      accountHolder: ''
    },
    invoicePrefix: 'INV',
    nextInvoiceNumber: 1,
    withholdingEnabled: false,
    defaultNote: 'お支払い期限までにお振込みください。',
    defaultDueDays: 30,
    logoUrl: ''
  };
}

/**
 * 設定を読み込む
 */
function loadSettings() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(SETTINGS_KEY);
    if (!raw) {
      return getDefaultSettings_();
    }
    var saved = JSON.parse(raw);
    var defaults = getDefaultSettings_();
    for (var key in defaults) {
      if (!(key in saved)) {
        saved[key] = defaults[key];
      }
    }
    // ネストオブジェクトもマージ
    if (!saved.issuer) saved.issuer = defaults.issuer;
    if (!saved.bankAccount) saved.bankAccount = defaults.bankAccount;
    return saved;
  } catch (e) {
    Logger.log('設定読み込みエラー: ' + e.message);
    return getDefaultSettings_();
  }
}

/**
 * 設定を保存する
 */
function saveSettings(settings) {
  try {
    PropertiesService.getUserProperties().setProperty(SETTINGS_KEY, JSON.stringify(settings));
  } catch (e) {
    Logger.log('設定保存エラー: ' + e.message);
    throw new Error('設定の保存に失敗しました。');
  }
}

/**
 * 次の請求書番号を生成して番号をインクリメント
 */
function generateInvoiceNumber() {
  var settings = loadSettings();
  var prefix = settings.invoicePrefix || 'INV';
  var year = new Date().getFullYear();
  var num = settings.nextInvoiceNumber || 1;
  var paddedNum = ('0000' + num).slice(-4);
  var invoiceNumber = prefix + '-' + year + '-' + paddedNum;

  settings.nextInvoiceNumber = num + 1;
  saveSettings(settings);

  return invoiceNumber;
}

/**
 * サブスクリプション状態を確認
 */
function checkSubscription() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(SUBSCRIPTION_KEY);
    if (!raw) {
      return { active: false, plan: 'free' };
    }
    var sub = JSON.parse(raw);

    if (sub.expiresAt) {
      var expiresAt = new Date(sub.expiresAt);
      if (expiresAt < new Date()) {
        var refreshed = refreshSubscriptionStatus_();
        if (refreshed) return refreshed;
        return { active: false, plan: 'free' };
      }
    }

    return {
      active: sub.active === true,
      plan: sub.plan || (sub.active ? 'basic' : 'free'),
      customerId: sub.customerId || '',
      subscriptionId: sub.subscriptionId || ''
    };
  } catch (e) {
    Logger.log('サブスクリプション確認エラー: ' + e.message);
    return { active: false, plan: 'free' };
  }
}

/**
 * サブスクリプション状態を保存
 */
function saveSubscription(subData) {
  try {
    PropertiesService.getUserProperties().setProperty(SUBSCRIPTION_KEY, JSON.stringify(subData));
  } catch (e) {
    Logger.log('サブスクリプション保存エラー: ' + e.message);
  }
}

/**
 * Stripeからサブスクリプション状態をリフレッシュ
 */
function refreshSubscriptionStatus_() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(SUBSCRIPTION_KEY);
    if (!raw) return null;

    var sub = JSON.parse(raw);
    if (!sub.subscriptionId) return null;

    var stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
    if (!stripeKey) return null;

    var response = UrlFetchApp.fetch(
      'https://api.stripe.com/v1/subscriptions/' + sub.subscriptionId,
      {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + stripeKey },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) return null;

    var stripeSub = JSON.parse(response.getContentText());
    var isActive = stripeSub.status === 'active' || stripeSub.status === 'trialing';

    // プランを判定（price IDから）
    var plan = 'basic';
    if (stripeSub.items && stripeSub.items.data && stripeSub.items.data.length > 0) {
      var priceId = stripeSub.items.data[0].price.id;
      var proPrice = PropertiesService.getScriptProperties().getProperty('STRIPE_PRICE_PRO');
      var proYearPrice = PropertiesService.getScriptProperties().getProperty('STRIPE_PRICE_PRO_YEAR');
      if (priceId === proPrice || priceId === proYearPrice) {
        plan = 'pro';
      }
    }

    var updated = {
      active: isActive,
      plan: isActive ? plan : 'free',
      customerId: stripeSub.customer,
      subscriptionId: stripeSub.id,
      expiresAt: new Date(stripeSub.current_period_end * 1000).toISOString(),
      lastChecked: new Date().toISOString()
    };

    saveSubscription(updated);

    return {
      active: isActive,
      plan: updated.plan,
      customerId: updated.customerId,
      subscriptionId: updated.subscriptionId
    };

  } catch (e) {
    Logger.log('サブスクリプションリフレッシュエラー: ' + e.message);
    return null;
  }
}

/**
 * 月間使用量を取得
 */
function getMonthlyUsage() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(USAGE_KEY);
    if (!raw) return { month: '', count: 0 };
    return JSON.parse(raw);
  } catch (e) {
    return { month: '', count: 0 };
  }
}

/**
 * 月間使用量をインクリメント
 */
function incrementMonthlyUsage() {
  var now = new Date();
  var currentMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
  var usage = getMonthlyUsage();

  if (usage.month !== currentMonth) {
    usage = { month: currentMonth, count: 0 };
  }

  usage.count++;
  PropertiesService.getUserProperties().setProperty(USAGE_KEY, JSON.stringify(usage));
  return usage;
}

/**
 * 月間制限をチェック
 * @return {Object} { allowed: boolean, remaining: number, limit: number, count: number }
 */
function checkMonthlyLimit() {
  var plan = checkSubscription();
  var usage = getMonthlyUsage();

  var now = new Date();
  var currentMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
  var count = (usage.month === currentMonth) ? usage.count : 0;

  var limit;
  if (plan.plan === 'pro' || plan.plan === 'basic') {
    limit = -1; // 無制限
  } else {
    limit = 3;
  }

  return {
    allowed: limit === -1 || count < limit,
    remaining: limit === -1 ? -1 : Math.max(0, limit - count),
    limit: limit,
    count: count,
    plan: plan.plan
  };
}

/**
 * サイドバーから呼ばれる設定取得API
 */
function getSettingsForSidebar() {
  var settings = loadSettings();
  var plan = checkSubscription();
  var usage = checkMonthlyLimit();

  return JSON.stringify({
    settings: settings,
    plan: plan,
    usage: usage
  });
}

/**
 * プラン情報を返す
 */
function getPlanInfo() {
  var plan = checkSubscription();
  var usage = checkMonthlyLimit();
  return JSON.stringify({ plan: plan, usage: usage });
}
