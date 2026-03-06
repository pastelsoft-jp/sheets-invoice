# プライバシーポリシー / Privacy Policy

**請求書メーカー for Sheets** (以下「本アドオン」)

最終更新日: 2026年3月7日

## 1. 収集するデータ

本アドオンは以下のデータにアクセスします:

- **Googleスプレッドシートのデータ**: 請求書の作成・表示に必要な範囲のみ
- **Google Driveのファイル**: 本アドオンが生成したPDFファイルのみ（drive.file スコープ）
- **メールアドレス**: ユーザー識別およびStripe決済に使用

## 2. データの保存場所

- **請求書データ**: ユーザーのGoogleスプレッドシート内
- **設定データ**: Google Apps Script PropertiesService（ユーザーごとに暗号化保存）
- **PDFファイル**: ユーザーのGoogle Drive内の「請求書PDF」フォルダ
- **決済データ**: Stripe（PCI DSS準拠の決済プロバイダー）

**外部サーバーへのデータ送信はありません。**
請求書データはお使いのGoogleアカウント内にのみ保存されます。

## 3. 外部サービスとの通信

本アドオンは以下の外部サービスと通信します:

- **Stripe** (https://stripe.com): サブスクリプション決済処理のみ。送信データはメールアドレスとプラン情報のみです。

## 4. データの削除

- サイドバーの「設定」>「全設定をリセット」で全データを削除できます
- アドオンをアンインストールすると、PropertiesServiceのデータは自動削除されます
- スプレッドシート内のシートおよびGoogle Drive内のPDFファイルは手動で削除してください

## 5. Google API Services User Data Policy

本アドオンは [Google API Services User Data Policy](https://developers.google.com/terms/api-services-user-data-policy) に準拠しています。取得したGoogle ユーザーデータは上記の目的以外には使用しません。

## 6. 第三者への提供

ユーザーデータを第三者に販売、共有、提供することはありません。ただし、法令に基づく場合はこの限りではありません。

## 7. お問い合わせ

プライバシーに関するお問い合わせ:
- Email: support@kinarikoubou.com
- GitHub: https://github.com/kinarikoubou/sheets-invoice

---

# Privacy Policy (English)

**Invoice Maker for Sheets** ("the Add-on")

Last updated: March 7, 2026

## Data Collection
The Add-on accesses only the data necessary to create invoices: spreadsheet content, Drive files created by this add-on, and user email for identification.

## Data Storage
All data is stored within the user's Google account (Spreadsheet, Drive, PropertiesService). No data is sent to external servers except Stripe for payment processing.

## Data Deletion
Users can delete all stored data via Settings > Reset. Uninstalling the add-on removes PropertiesService data automatically.

## Contact
support@kinarikoubou.com
