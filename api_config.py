"""
API・サブスクリプション一覧データ
Discord Bot の /api-list コマンド用。

データ元: H:\マイドライブ\app\サブスクリプション\api_usage_descriptions.json
         H:\マイドライブ\app\サブスクリプション\サブスク・API一覧.csv
更新方法: このファイルを編集して git push（Railway 自動デプロイ）
"""

# サブスクリプション（月額・年額）
SUBSCRIPTIONS = [
    {
        "name": "Claude",
        "plan": "$200/月（Max）",
        "url": "https://claude.ai/",
        "note": "20倍使用量・優先アクセス・音声モード",
        "icon": "🤖",
    },
    {
        "name": "Cursor",
        "plan": "$20/月（Pro）",
        "url": "https://www.cursor.com/",
        "note": "無制限タブ補完・エージェント",
        "icon": "💻",
    },
    {
        "name": "Google AI Pro",
        "plan": "$19.99/月（Pro）",
        "url": "https://one.google.com/ai/pro",
        "note": "2TB・Gemini 2.5 Pro・動画生成",
        "icon": "🔷",
    },
    {
        "name": "Fusion 360",
        "plan": "96,800円/年",
        "url": "https://www.autodesk.co.jp/products/fusion-360/overview",
        "note": "CAD・解析・レンダリング（ローカルのみ）",
        "icon": "🔧",
    },
    {
        "name": "ConoHa WING",
        "plan": "約1,089円/月",
        "url": "https://www.conoha.jp/wing/",
        "note": "レンタルサーバー（ベーシック12ヶ月契約）",
        "icon": "🌐",
    },
    {
        "name": "Railway",
        "plan": "従量（Hobby $5/月）",
        "url": "https://railway.com/dashboard",
        "note": "Discord Bot ホスティング（月約$1〜1.5）",
        "icon": "🚂",
    },
]

# API（従量・無料枠）— コスト自動取得対応あり
API_LIST = [
    {
        "name": "OpenAI API",
        "usage_location": "Voicenotes / OpenClaw",
        "function_description": "GPT-5 Nano/Mini/5.2/Codex。音声メモ分析 + OpenClaw マルチLLM",
        "pricing": "従量",
        "docs_url": "https://platform.openai.com/",
        "dashboard_url": "https://platform.openai.com/usage",
        "cost_tracking": "auto",
        "icon": "🤖",
    },
    {
        "name": "Anthropic API",
        "usage_location": "Claude API / OpenClaw",
        "function_description": "Claude Haiku 4.5/Sonnet 4.5/Opus 4.6。OpenClaw マルチLLM",
        "pricing": "従量",
        "docs_url": "https://docs.anthropic.com/",
        "dashboard_url": "https://console.anthropic.com/settings/billing",
        "cost_tracking": "auto",
        "icon": "🤖",
    },
    {
        "name": "Groq API",
        "usage_location": "Voicenotes / OpenClaw（将来）",
        "function_description": "Whisper 音声認識。録音→テキスト変換",
        "pricing": "従量",
        "docs_url": "https://console.groq.com/",
        "dashboard_url": "https://console.groq.com/settings/usage",
        "cost_tracking": "manual",
        "icon": "⚡",
    },
    {
        "name": "Moonshot (Kimi) API",
        "usage_location": "OpenClaw",
        "function_description": "Kimi K2.5。OpenClaw マルチLLM（Tier 2）",
        "pricing": "従量",
        "docs_url": "https://platform.moonshot.cn/",
        "dashboard_url": "https://platform.moonshot.cn/console",
        "cost_tracking": "manual",
        "icon": "🌙",
    },
    {
        "name": "Stripe",
        "usage_location": "EC",
        "function_description": "決済処理（カード・振込）",
        "pricing": "従量（3.6%+¥30）",
        "docs_url": "https://dashboard.stripe.com/",
        "dashboard_url": "https://dashboard.stripe.com/",
        "cost_tracking": "manual",
        "icon": "💳",
    },
    {
        "name": "Supabase",
        "usage_location": "SPEClaud / Voicenotes",
        "function_description": "DB・認証・ストレージ",
        "pricing": "無料枠",
        "docs_url": "https://supabase.com/dashboard",
        "dashboard_url": "https://supabase.com/dashboard",
        "cost_tracking": "free",
        "icon": "🗄️",
    },
    {
        "name": "Google Gemini API",
        "usage_location": "OpenClaw / SPEClaud",
        "function_description": "Gemini 2.5 Flash/Pro。OpenClaw 意図分類 + マルチLLM",
        "pricing": "無料枠/従量",
        "docs_url": "https://aistudio.google.com/",
        "dashboard_url": "https://aistudio.google.com/",
        "cost_tracking": "manual",
        "icon": "🔷",
    },
    {
        "name": "LINE Messaging API",
        "usage_location": "EC（i.tategu）",
        "function_description": "顧客チャット・通知",
        "pricing": "無料枠",
        "docs_url": "https://developers.line.biz/console/",
        "dashboard_url": "https://developers.line.biz/console/",
        "cost_tracking": "free",
        "icon": "💬",
    },
    {
        "name": "Discord API",
        "usage_location": "EC / 共通",
        "function_description": "Bot運用・開発ログ投稿",
        "pricing": "無料",
        "docs_url": "https://discord.com/developers/applications",
        "dashboard_url": "https://discord.com/developers/applications",
        "cost_tracking": "free",
        "icon": "💬",
    },
]
