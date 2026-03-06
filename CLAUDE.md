# 請求書メーカー for Sheets -- インボイス制度対応の請求書をGoogleスプレッドシートから生成

## あなたの役割
PM兼エンジニアとして、このGWSアドオンを自律的に開発する。

## 技術スタック
| 項目 | 技術 |
|------|------|
| ランタイム | Google Apps Script (V8 runtime) |
| UI | HTML + CSS + vanilla JS (サイドバー) |
| 決済 | Stripe Checkout（ポーリング方式。外部サーバー不要） |
| 設定保存 | PropertiesService (ユーザー設定) |
| PDF生成 | SpreadsheetApp → getAs(MimeType.PDF) |

## OAuthスコープ
```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/userinfo.email"
  ]
}
```
- spreadsheets: スプレッドシートの読み書き（sensitive）
- drive.file: PDF生成・保存（sensitive — このスクリプトが作成したファイルのみ）
- script.external_request: Stripe API通信
- userinfo.email: ユーザー識別

## GAS固有の制約
| 制約 | Consumer | Workspace |
|------|----------|-----------|
| 実行時間/回 | 6分 | 6分 |
| UrlFetchApp/日 | 20,000回 | 100,000回 |
| PropertiesService | 500KB/ストア | 500KB/ストア |
| トリガー上限 | 20個/ユーザー/スクリプト | 20個 |
| PDF生成 | SpreadsheetApp制限内 | 同左 |

## MVP定義

### Free機能
1. **請求書テンプレート**: インボイス制度完全対応の日本語テンプレート（A4縦）
2. **自動計算**: 小計・消費税（10%/8%軽減税率）・源泉徴収・合計の自動計算
3. **適格請求書要件**: 登録番号(T+13桁)、税率区分、消費税額を自動表示
4. **PDF出力**: 1クリックでPDF生成・ダウンロード
5. **月3枚まで**: 無料版は月3枚の請求書生成制限

### Basic（¥500/月 or ¥4,800/年）
6. **無制限生成**: 請求書の生成枚数制限なし
7. **顧客マスタ**: 取引先情報を保存・呼び出し（最大50件）
8. **請求書番号の自動採番**: 連番管理（プレフィックス設定可能）
9. **一括生成**: 顧客リストから複数請求書を一括生成
10. **テンプレートカスタマイズ**: ロゴ画像、振込先、備考のデフォルト設定

### Pro（¥1,500/月 or ¥14,800/年）
11. **メール送信**: PDF添付で請求書をメール送信（宛先自動入力）
12. **履歴管理**: 発行済み請求書の一覧・再発行
13. **見積書・納品書・領収書**: 請求書以外の帳票テンプレート
14. **Excel/CSV出力**: 請求データの一括エクスポート
15. **複数テンプレート**: デザインテンプレート5種類

## ディレクトリ構成
```
src/
  Code.gs              -- メインロジック（onOpen, showSidebar, include）
  sidebar.html         -- サイドバーUI
  styles.html          -- CSS
  sidebar-api.gs       -- サイドバーAPI関数
  invoice-template.gs  -- 請求書テンプレート生成（Sheetsに書き出し）
  invoice-data.gs      -- 請求データ管理（入力値の構造化）
  pdf-generator.gs     -- PDF生成ロジック
  customer-master.gs   -- 顧客マスタCRUD
  tax-calculator.gs    -- 消費税・源泉徴収計算
  settings.gs          -- ユーザー設定管理
  stripe.gs            -- Stripe決済連携（ポーリング方式）
  appsscript.json      -- GASマニフェスト
```

## 請求書テンプレート仕様

### 必須項目（インボイス制度要件）
1. **発行者の名称**: 事業者名
2. **発行者の登録番号**: T + 13桁の法人番号
3. **取引年月日**: 請求日
4. **取引内容**: 品目・数量・単価
5. **税率ごとに区分した対価の額**: 10%対象小計 / 8%対象小計
6. **適用税率**: 10% / 8%（軽減税率）
7. **税率ごとの消費税額**: 10%分消費税 / 8%分消費税
8. **相手方の名称**: 請求先

### 追加項目
- 請求書番号（自動採番）
- 支払期限
- 振込先口座情報
- 源泉徴収税額（該当する場合）
- 備考欄
- ロゴ画像

### 源泉徴収計算
```
支払金額 ≤ 100万円: 源泉徴収税額 = 支払金額 × 10.21%
支払金額 > 100万円: 源泉徴収税額 = (100万 × 10.21%) + (超過分 × 20.42%)
```
※ 設計報酬、原稿料、講演料等に適用。設定でON/OFF可能。

### 消費税計算
```
// 端数処理は切り捨て（国税庁推奨）
消費税額(10%) = Math.floor(10%対象合計 × 0.1)
消費税額(8%)  = Math.floor(8%対象合計 × 0.08)
```

## Stripe決済パターン（ポーリング方式）
Calendar adonと同一パターン。Checkout Session作成→ポーリングで完了検知→PropertiesServiceに保存。

## 環境変数（ScriptProperties）
```
STRIPE_SECRET_KEY=sk_live_...
STRIPE_PRICE_BASIC=price_...     (¥500/月)
STRIPE_PRICE_PRO=price_...       (¥1,500/月)
STRIPE_PRICE_BASIC_YEAR=price_...  (¥4,800/年)
STRIPE_PRICE_PRO_YEAR=price_...    (¥14,800/年)
```

## コスト試算
| 項目 | 月額 |
|------|------|
| Google Apps Script | ¥0 |
| Stripe手数料 | MRR x 3.6% |
| ドメイン | ¥0（GitHub Pages共用） |
| **合計** | **MRR x 3.6%のみ** |

## API設計（サイドバーから呼び出す関数）
```javascript
// 設定・初期化
function getSettingsForSidebar()
function saveSettingsFromSidebar(settings)
function completeOnboardingFromSidebar()
function resetSettingsFromSidebar()

// 請求書操作
function createInvoice(invoiceData)      // 請求書テンプレートをシートに生成
function generatePdf(sheetName)          // シートをPDF変換
function sendInvoiceEmail(invoiceId, to)  // PDF添付メール送信（Pro）
function getInvoiceHistory()             // 発行履歴一覧

// 顧客マスタ
function getCustomers()                  // 顧客一覧
function saveCustomer(customerData)      // 顧客保存
function deleteCustomer(customerId)      // 顧客削除

// 税計算
function calculateTax(items)             // 消費税計算
function calculateWithholding(amount)    // 源泉徴収計算

// Stripe
function createCheckoutSession(planId)
function createCustomerPortalSession()
function checkSubscription()

// プラン
function getPlanInfo()                   // 現在のプラン・使用量
function checkMonthlyLimit()             // 月間生成枚数チェック
```

## UI設計（サイドバー 300px幅）

### 画面遷移
1. **オンボーディング** → 事業者情報入力（名前、登録番号、振込先）
2. **メイン画面** → 新規作成ボタン + 最近の請求書 + プラン表示
3. **請求書作成** → 取引先選択 → 品目入力 → プレビュー → 生成
4. **設定画面** → 事業者情報 + テンプレート設定 + プラン管理

### 品目入力UI
- 動的行追加（+ボタン）
- 各行: 品目名 / 数量 / 単価 / 税率(10%/8%) / 小計
- リアルタイム合計表示（小計・消費税・源泉徴収・請求額）
