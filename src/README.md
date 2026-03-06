# 請求書メーカー for Sheets

インボイス制度（適格請求書等保存方式）に完全対応した請求書を、Googleスプレッドシートから1クリックで生成するGWSアドオンです。

## 機能

- インボイス制度の必須8項目を自動表示
- 消費税の自動計算（10%標準/8%軽減税率）
- 源泉徴収税額の自動計算（10.21%/20.42%）
- PDF出力（A4縦、Google Driveに保存）
- 顧客マスタ（取引先情報の保存・呼び出し）
- 請求書番号の自動採番

## インストール

### GASエディタから

1. Googleスプレッドシートを開く
2. メニュー > 拡張機能 > Apps Script
3. `src/` 配下の全ファイルをGASエディタにコピー
4. `appsscript.json` を上書き（エディタ > プロジェクトの設定 > マニフェストファイルを表示）

### clasp（推奨）

```bash
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "請求書メーカー"
# src/ の中身を clasp push
```

## ファイル構成

```
src/
  appsscript.json       -- GASマニフェスト
  Code.gs               -- エントリーポイント（onOpen, onHomepage, Card Service）
  sidebar.html          -- サイドバーUI（品目入力、合計表示、タブ）
  styles.html           -- CSS（Material Design風、600行）
  sidebar-api.gs        -- サイドバーAPI（google.script.run 11関数）
  invoice-template.gs   -- 請求書シート生成
  invoice-data.gs       -- 請求データ管理・履歴
  pdf-generator.gs      -- PDF生成・Google Drive保存
  customer-master.gs    -- 顧客マスタCRUD
  tax-calculator.gs     -- 消費税・源泉徴収計算
  settings.gs           -- ユーザー設定・プラン管理
  stripe.gs             -- Stripe決済（ポーリング方式）
```

## Stripe設定

ScriptProperties に以下を設定:

```
STRIPE_SECRET_KEY=sk_live_...
STRIPE_PRICE_BASIC=price_...       (¥500/月)
STRIPE_PRICE_BASIC_YEAR=price_...  (¥4,800/年)
STRIPE_PRICE_PRO=price_...         (¥1,500/月)
STRIPE_PRICE_PRO_YEAR=price_...    (¥14,800/年)
```

## OAuthスコープ

| スコープ | 用途 | レベル |
|----------|------|--------|
| spreadsheets | シート読み書き | sensitive |
| drive.file | PDF保存 | sensitive |
| script.external_request | Stripe API | sensitive |
| userinfo.email | ユーザー識別 | sensitive |

全てsensitive（CASA審査不要）

## 技術仕様

- ランタイム: Google Apps Script (V8)
- UI: HTML + CSS + vanilla JS
- 決済: Stripe Checkout → ポーリング方式
- データ保存: PropertiesService
- PDF: SpreadsheetApp.export → UrlFetchApp
