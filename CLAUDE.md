# Discord-LINE Bot (DevBot)

## 概要
Discord↔LINE連携Bot。注文通知、アトリエスレッド管理、Canva自動処理。

## スタック
Python 3 + discord.py + Flask + aiohttp

## デプロイ
GitHub (i-tategu/discord-line-bot) → Railway自動デプロイ (main push)

## 主要ファイル
| ファイル | 説明 |
|---------|------|
| discord_bot_server.py | メインBot（Discord↔LINE転送、Webhook受信） |
| canva_handler.py | Canva API自動処理 |
| customer_manager.py | 顧客管理 |
| product_register.py | 商品登録 |

## API エンドポイント
- /health — ヘルスチェック
- /api/canva/process — Canva処理 (POST: {"order_id": 123})
- /api/woo-webhook — WooCommerce Webhook受信

## 検証
- `railway logs --num 50`
- `curl -s https://worker-production-eb8a.up.railway.app/health`
