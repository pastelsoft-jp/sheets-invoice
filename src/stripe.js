/**
 * Stripe決済連携（ポーリング方式）
 * Calendar addonと同一パターン
 */

/**
 * Stripe Checkout Sessionを作成
 * @param {string} planId - 'basic' | 'basic_year' | 'pro' | 'pro_year'
 * @return {string} Checkout URL
 */
function createCheckoutSession(planId) {
  var stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!stripeKey) {
    throw new Error('Stripe設定が完了していません。管理者にお問い合わせください。');
  }

  // プランIDからStripe Price IDを取得
  var priceMap = {
    'basic': 'STRIPE_PRICE_BASIC',
    'basic_year': 'STRIPE_PRICE_BASIC_YEAR',
    'pro': 'STRIPE_PRICE_PRO',
    'pro_year': 'STRIPE_PRICE_PRO_YEAR'
  };

  var priceKey = priceMap[planId] || 'STRIPE_PRICE_BASIC';
  var priceId = PropertiesService.getScriptProperties().getProperty(priceKey);

  if (!priceId) {
    throw new Error('料金プランが設定されていません。');
  }

  var email = Session.getActiveUser().getEmail();
  var existingCustomerId = getOrCreateStripeCustomer_(stripeKey, email);

  var payload = {
    'mode': 'subscription',
    'line_items[0][price]': priceId,
    'line_items[0][quantity]': '1',
    'success_url': 'https://pastelsoft-jp.github.io/sheets-invoice/payment-success.html?session_id={CHECKOUT_SESSION_ID}',
    'cancel_url': 'https://pastelsoft-jp.github.io/sheets-invoice/payment-cancel.html',
    'allow_promotion_codes': 'true',
    'locale': 'ja',
    'metadata[addon]': 'sheets-invoice-maker',
    'metadata[email]': email,
    'metadata[plan]': planId
  };

  if (existingCustomerId) {
    payload['customer'] = existingCustomerId;
  } else {
    payload['customer_email'] = email;
  }

  var response = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + stripeKey },
    payload: payload,
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('Stripe Checkout error: ' + response.getContentText());
    throw new Error('決済ページの作成に失敗しました。');
  }

  var session = JSON.parse(response.getContentText());

  // セッションIDを保存（ポーリング用）
  PropertiesService.getUserProperties().setProperty(
    'pending_checkout_session',
    JSON.stringify({
      sessionId: session.id,
      planId: planId,
      createdAt: new Date().toISOString()
    })
  );

  setupCheckoutPollingTrigger_();
  return session.url;
}

/**
 * Stripe顧客を検索
 */
function getOrCreateStripeCustomer_(stripeKey, email) {
  try {
    var response = UrlFetchApp.fetch(
      'https://api.stripe.com/v1/customers?email=' + encodeURIComponent(email) + '&limit=1',
      {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + stripeKey },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      if (data.data && data.data.length > 0) {
        return data.data[0].id;
      }
    }
  } catch (e) {
    Logger.log('Stripe顧客検索エラー: ' + e.message);
  }
  return null;
}

/**
 * Checkout完了ポーリング
 */
function pollCheckoutSession() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty('pending_checkout_session');
    if (!raw) {
      removeCheckoutPollingTrigger_();
      return;
    }

    var pending = JSON.parse(raw);
    var stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
    if (!stripeKey) return;

    // 1時間以上経過したら破棄
    if (new Date() - new Date(pending.createdAt) > 3600000) {
      PropertiesService.getUserProperties().deleteProperty('pending_checkout_session');
      removeCheckoutPollingTrigger_();
      return;
    }

    var response = UrlFetchApp.fetch(
      'https://api.stripe.com/v1/checkout/sessions/' + pending.sessionId,
      {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + stripeKey },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) return;

    var session = JSON.parse(response.getContentText());

    if (session.status === 'complete' && session.subscription) {
      activateSubscription_(stripeKey, session.subscription, session.customer, pending.planId);
      PropertiesService.getUserProperties().deleteProperty('pending_checkout_session');
      removeCheckoutPollingTrigger_();
      return { activated: true, plan: pending.planId };
    }
    return { activated: false, status: session.status, payment: session.payment_status };
  } catch (e) {
    Logger.log('ポーリングエラー: ' + e.message);
    return { activated: false, error: e.message };
  }
}

/**
 * サイドバーから手動で決済完了を確認
 */
function checkPaymentFromSidebar() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty('pending_checkout_session');
    if (!raw) {
      return JSON.stringify({ activated: false, message: '確認中の決済がありません' });
    }

    var pending = JSON.parse(raw);
    var stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
    if (!stripeKey) {
      return JSON.stringify({ activated: false, message: 'Stripe設定がありません' });
    }

    var response = UrlFetchApp.fetch(
      'https://api.stripe.com/v1/checkout/sessions/' + pending.sessionId,
      {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + stripeKey },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) {
      return JSON.stringify({ activated: false, message: 'Stripe通信エラー' });
    }

    var session = JSON.parse(response.getContentText());
    if (session.status === 'complete' && session.subscription) {
      var plan = pending.planId || 'basic';
      if (plan.indexOf('pro') !== -1) { plan = 'pro'; } else { plan = 'basic'; }

      // 30日後をデフォルトの有効期限にする
      var expiresAt = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString();

      saveSubscription({
        active: true,
        plan: plan,
        customerId: session.customer || '',
        subscriptionId: session.subscription,
        expiresAt: expiresAt,
        lastChecked: new Date().toISOString()
      });

      PropertiesService.getUserProperties().deleteProperty('pending_checkout_session');
      removeCheckoutPollingTrigger_();

      var limits = { 'free': 3, 'basic': -1, 'pro': -1 };
      var usageRaw = PropertiesService.getUserProperties().getProperty('invoice_maker_usage');
      var count = 0;
      if (usageRaw) {
        try {
          var u = JSON.parse(usageRaw);
          var ym = new Date().getFullYear() + '-' + ('0' + (new Date().getMonth() + 1)).slice(-2);
          count = (u.month === ym) ? (u.count || 0) : 0;
        } catch(e) {}
      }
      var limit = limits[plan] || 3;

      return JSON.stringify({
        activated: true, planName: plan, planActive: true,
        usageCount: count, usageLimit: limit, usageRemaining: limit === -1 ? -1 : Math.max(0, limit - count)
      });
    }

    return JSON.stringify({ activated: false, message: '決済がまだ完了していません (status: ' + session.status + ', payment: ' + session.payment_status + ')' });
  } catch (e) {
    Logger.log('checkPaymentFromSidebar error: ' + e.message);
    return JSON.stringify({ activated: false, message: e.message });
  }
}

/**
 * サブスクリプションを有効化
 */
function activateSubscription_(stripeKey, subscriptionId, customerId, planId) {
  try {
    var response = UrlFetchApp.fetch(
      'https://api.stripe.com/v1/subscriptions/' + subscriptionId,
      {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + stripeKey },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) return;

    var sub = JSON.parse(response.getContentText());
    var isActive = sub.status === 'active' || sub.status === 'trialing';

    // プランを判定
    var plan = planId || 'basic';
    if (plan.indexOf('pro') !== -1) {
      plan = 'pro';
    } else {
      plan = 'basic';
    }

    saveSubscription({
      active: isActive,
      plan: isActive ? plan : 'free',
      customerId: customerId,
      subscriptionId: subscriptionId,
      expiresAt: new Date(sub.current_period_end * 1000).toISOString(),
      lastChecked: new Date().toISOString()
    });
  } catch (e) {
    Logger.log('サブスクリプション有効化エラー: ' + e.message);
  }
}

/**
 * Stripe Webhook受信（バックアップ）
 */
function handleStripeWebhook(e) {
  try {
    if (!e || !e.postData) {
      return ContentService.createTextOutput('No data').setMimeType(ContentService.MimeType.TEXT);
    }

    var body = JSON.parse(e.postData.contents);
    var eventType = body.type;

    if (body.data && body.data.object && body.data.object.metadata) {
      if (body.data.object.metadata.addon !== 'sheets-invoice-maker') {
        return ContentService.createTextOutput('Not for this addon').setMimeType(ContentService.MimeType.TEXT);
      }
    }

    var stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');

    switch (eventType) {
      case 'checkout.session.completed':
        var session = body.data.object;
        if (session.subscription && session.customer) {
          activateSubscription_(stripeKey, session.subscription, session.customer, session.metadata.plan);
        }
        break;

      case 'customer.subscription.deleted':
      case 'customer.subscription.updated':
        var subscription = body.data.object;
        var isActive = subscription.status === 'active' || subscription.status === 'trialing';
        saveSubscription({
          active: isActive,
          plan: isActive ? (subscription.metadata.plan || 'basic') : 'free',
          customerId: subscription.customer,
          subscriptionId: subscription.id,
          expiresAt: new Date(subscription.current_period_end * 1000).toISOString(),
          lastChecked: new Date().toISOString()
        });
        break;

      case 'invoice.payment_failed':
        Logger.log('決済失敗: ' + JSON.stringify(body.data.object));
        break;
    }

    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (e) {
    Logger.log('Webhook処理エラー: ' + e.message);
    return ContentService.createTextOutput('Error').setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Stripe Customer Portal
 */
function createCustomerPortalSession() {
  var plan = checkSubscription();
  if (!plan.active || !plan.customerId) {
    throw new Error('有効なサブスクリプションがありません。');
  }

  var stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!stripeKey) {
    throw new Error('Stripe設定が完了していません。');
  }

  var response = UrlFetchApp.fetch('https://api.stripe.com/v1/billing_portal/sessions', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + stripeKey },
    payload: {
      'customer': plan.customerId,
      'return_url': 'https://pastelsoft-jp.github.io/sheets-invoice/portal-return.html'
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('カスタマーポータルの作成に失敗しました。');
  }

  return JSON.parse(response.getContentText()).url;
}

/**
 * ポーリングトリガー管理
 */
function setupCheckoutPollingTrigger_() {
  removeCheckoutPollingTrigger_();
  ScriptApp.newTrigger('pollCheckoutSession')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function removeCheckoutPollingTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollCheckoutSession') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 日次サブスクリプションチェック
 */
function setupDailySubscriptionCheck() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailySubscriptionCheck') {
      return;
    }
  }
  ScriptApp.newTrigger('dailySubscriptionCheck')
    .timeBased()
    .everyHours(24)
    .create();
}

function dailySubscriptionCheck() {
  try {
    var plan = checkSubscription();
    if (plan.active && plan.subscriptionId) {
      refreshSubscriptionStatus_();
    }
  } catch (e) {
    Logger.log('日次チェックエラー: ' + e.message);
  }
}
