# 土地造成リスク診断アプリ

土地の造成リスクを診断し、PDFレポートを提供するWebアプリです。

## 機能

- 住所入力による土地リスク診断
- Stripe決済（クレジットカード）
- SendGridによるPDFレポートメール送信
- 管理者ダッシュボード

## 技術スタック

- Python / Flask
- Stripe（決済）
- SendGrid（メール送信）
- ReportLab（PDF生成）
- SQLite（開発）/ PostgreSQL（本番推奨）

## ローカル開発

```bash
pip install -r requirements.txt
cp .env.example .env   # APIキーを設定
python app.py
```

## 環境変数

| 変数名 | 説明 |
|--------|------|
| SECRET_KEY | Flaskセッションキー |
| STRIPE_SECRET_KEY | Stripe シークレットキー |
| STRIPE_PUBLISHABLE_KEY | Stripe 公開可能キー |
| STRIPE_WEBHOOK_SECRET | Stripe Webhook署名シークレット |
| SENDGRID_API_KEY | SendGrid APIキー |
| FROM_EMAIL | 送信元メールアドレス |
| ADMIN_PASSWORD | 管理者ログインパスワード |
