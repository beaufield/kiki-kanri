# ビューフィールド貸出管理 (kiki-kanri) プロジェクト

## プロジェクト概要
美容機器の貸出管理を行う単一ファイルSPA。
営業担当がサロンに商品を貸し出す際の登録・返却・管理を行う。

## 技術構成
- **フロントエンド**: `index.html`（単一ファイルSPA / CSS + HTML + Vanilla JS）
- **バックエンド**: Google Apps Script (GAS)
- **DB**: Google Spreadsheet
- **通知**: LINE WORKS Bot Webhook
- **バージョン**: 1.5.0

## 重要なURL・ID
- **GAS API URL**: `https://script.google.com/macros/s/AKfycbzWQEPQ89DwT7n38vtLFf1COx0rz13-qKC7GM1dZBiEBYP5GV5H_3J-Olz_rR2ckr4o/exec`
- **Spreadsheet ID**: `1o12RSbRWmNsiEjVPCb2dIjyw4U4Ntn47-6Lc80E_jvk`
- **LINE WORKS Webhook URL**: `https://webhook.worksmobile.com/message/bf4bbf8b-e26f-4760-b2f2-5ea20b4cc025`
- **GitHub**: `https://github.com/beaufield/kiki-kanri`

## Spreadsheetシート構成
- `DeviceMaster` — 商品マスタ
- `LoanLog` — 貸出ログ
- `SalesRep` — 営業担当マスタ
- `MakerMaster` — メーカーマスタ

## 実装済み機能
- 商品登録・貸出・返却の基本CRUD
- QRコードスキャン（jsQRライブラリ）
- 廃棄機能・メーカー返却機能（日付・操作者記録）
- メーカーマスタ管理（CRUD）
- 商品一覧のメーカーフィルター（タブ＋サブフィルター）
- 画像アップロード（Google Drive、Base64→GAS uploadImageアクション）
- 操作者記録・localStorage保持（lastOperator）
- 貸出期限アラートバナー（3日前警告・期限超過）
- 重複サロン警告トースト
- 検索強化（営業担当名含む）
- CSV出力（デバイス・貸出ログ、全フィールド対応）
- LINE WORKS即時通知（貸出・返却登録時）
- LINE WORKS定期通知（毎週火曜9:00、返却期限未設定商品リスト）
- LabelPool（ラベル管理）: 発行・在庫状況・印刷用CSV・印刷済み更新
- 新規商品登録フォームのLabelIDをLabelPool印刷済ラベルのドロップダウンに変更

## 主要な技術仕様

### GAS通信
- POST: `application/x-www-form-urlencoded`、`action` + `data`パラメータ（CORS preflight回避）
- GASレスポンス形式: `{ status: 'ok', result: { ... } }`
- 画像URL取得: `res.result.imageUrl`

### クライアントルーター
- `currentView`, `viewStack[]`, `viewData`
- `navigate()`, `goBack()`, `resetTo()`

### デバイスステータス
- `'社内'`, `'貸出中'`, `'廃棄'`, `'メーカー返却済'`
- 通常リストは `廃棄` と `メーカー返却済` を除外

### LINE WORKS Webhook ペイロード形式
```json
{ "body": { "text": "メッセージ本文" } }
```

### CSS変数（主要）
- `--bg`, `--card`, `--primary`, `--danger`, `--success`, `--warning`, `--text`, `--border`
- ※ `--bg-card` は未定義。モーダル背景には `#ffffff` を使用

## コーディングルール
- **コードを修正・提示するときは必ず全文差し替え形式で提供する**（部分スニペット不可）
- CSSはインラインで`<style>`タグ内に記述
- JSはインラインで`<script>`タグ内に記述
