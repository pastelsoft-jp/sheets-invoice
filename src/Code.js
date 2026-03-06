/**
 * 請求書メーカー for Sheets - インボイス制度対応
 * Google Sheets Add-on
 *
 * エントリーポイント。メニュー、サイドバー表示、カード構築を担当。
 */

// ========== Add-on エントリーポイント ==========

function onHomepage(e) {
  return buildMainCard_();
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('請求書メーカー')
    .addItem('請求書を作成', 'showSidebar')
    .addItem('設定', 'showSettingsSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

// ========== サイドバー ==========

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('請求書メーカー')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSettingsSidebar() {
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('請求書メーカー - 設定')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========== メインカード構築 (Card Service) ==========

function buildMainCard_() {
  var settings = loadSettings();
  var plan = checkSubscription();
  var usage = checkMonthlyLimit();
  var card = CardService.newCardBuilder();

  if (!settings.hasSeenOnboarding) {
    return buildOnboardingCard_();
  }

  card.setHeader(CardService.newCardHeader()
    .setTitle('請求書メーカー')
    .setSubtitle('インボイス制度対応'));

  // ステータスセクション
  var statusSection = CardService.newCardSection()
    .setHeader('状態');

  var planLabel = plan.plan === 'pro' ? 'Pro' : (plan.plan === 'basic' ? 'Basic' : '無料プラン');
  statusSection.addWidget(CardService.newDecoratedText()
    .setTopLabel('プラン')
    .setText(planLabel)
    .setWrapText(true));

  if (plan.plan === 'free') {
    var remainLabel = usage.remaining >= 0 ? '今月あと' + usage.remaining + '枚' : '無制限';
    statusSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('無料枠')
      .setText(remainLabel)
      .setWrapText(true));
  }

  if (settings.issuer.name) {
    statusSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('発行者')
      .setText(settings.issuer.name)
      .setWrapText(true));
  }

  card.addSection(statusSection);

  // アクションセクション
  var actionSection = CardService.newCardSection()
    .setHeader('操作');

  actionSection.addWidget(CardService.newTextButton()
    .setText('サイドバーを開く（請求書作成）')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction().setFunctionName('openSidebarAction')));

  card.addSection(actionSection);

  // アップグレードセクション
  if (plan.plan === 'free') {
    var upgradeSection = CardService.newCardSection()
      .setHeader('アップグレード');

    upgradeSection.addWidget(CardService.newTextParagraph()
      .setText('Basicプラン（¥500/月）で無制限生成・顧客マスタ・一括生成が利用可能'));

    upgradeSection.addWidget(CardService.newTextButton()
      .setText('Basicにアップグレード')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleUpgrade')
        .setParameters({ planId: 'basic' })));

    card.addSection(upgradeSection);
  }

  return card.build();
}

// ========== オンボーディングカード ==========

function buildOnboardingCard_() {
  var card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
    .setTitle('請求書メーカーへようこそ')
    .setSubtitle('インボイス対応の請求書を簡単作成'));

  var section = CardService.newCardSection();

  section.addWidget(CardService.newTextParagraph()
    .setText('このアドオンは、インボイス制度に完全対応した請求書をGoogleスプレッドシートから作成します。\n\n' +
      '<b>無料プラン</b>\n' +
      '・月3枚まで請求書生成\n' +
      '・消費税（10%/8%）自動計算\n' +
      '・源泉徴収税額自動計算\n' +
      '・PDF出力\n\n' +
      '<b>Basic ¥500/月</b>\n' +
      '・無制限生成・顧客マスタ・一括生成\n\n' +
      '<b>Pro ¥1,500/月</b>\n' +
      '・メール送信・履歴管理・見積書/納品書'));

  section.addWidget(CardService.newTextButton()
    .setText('はじめる')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction().setFunctionName('completeOnboarding')));

  card.addSection(section);

  return card.build();
}

function completeOnboarding() {
  var settings = loadSettings();
  settings.hasSeenOnboarding = true;
  saveSettings(settings);

  var nav = CardService.newNavigation().updateCard(buildMainCard_());
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

// ========== アクションハンドラ ==========

function openSidebarAction() {
  showSidebar();
  return buildNotification_('サイドバーを開きました', false);
}

function handleUpgrade(e) {
  try {
    var params = e.parameters || {};
    var planId = params.planId || 'basic';
    var checkoutUrl = createCheckoutSession(planId);
    return CardService.newActionResponseBuilder()
      .setOpenLink(CardService.newOpenLink()
        .setUrl(checkoutUrl)
        .setOpenAs(CardService.OpenAs.OVERLAY))
      .build();
  } catch (error) {
    Logger.log('handleUpgrade error: ' + error.message);
    return buildNotification_('決済ページの作成に失敗しました。', true);
  }
}

// ========== ユーティリティ ==========

function buildNotification_(message, isError) {
  var notification = CardService.newNotification().setText(
    message.length > 200 ? message.substring(0, 197) + '...' : message
  );
  return CardService.newActionResponseBuilder()
    .setNotification(notification)
    .build();
}

// ========== Stripe Webhook受信 ==========

function doPost(e) {
  return handleStripeWebhook(e);
}
