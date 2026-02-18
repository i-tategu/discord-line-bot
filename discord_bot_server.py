"""
Discord Bot v2 - サーバー版（Canva自動化統合）
- Discord → LINE 転送
- 顧客ステータス管理
- 一覧表示・自動更新
- WooCommerce Webhook → Canva自動化
"""
import os
import re
import json
import asyncio
import requests
import threading
import hmac
import hashlib
import base64
from flask import Flask, request, jsonify
from dotenv import load_dotenv
import discord
from discord.ext import commands
from discord import app_commands

from customer_manager import (
    CustomerStatus, STATUS_CONFIG,
    add_customer, add_order_customer, update_customer_status, get_customer,
    get_customer_by_channel, get_customer_by_order,
    get_status_summary, get_all_customers_grouped, load_customers, save_customers,
    get_linked_users_by_order, update_linked_customer_statuses
)

# 商品登録モジュール
try:
    import product_register
    PRODUCT_REGISTER_ENABLED = True
except ImportError as e:
    PRODUCT_REGISTER_ENABLED = False
    print(f"[WARN] Product register not available: {e}")

# Canva自動化ハンドラー
try:
    from canva_handler import process_order as canva_process_order, get_current_tokens
    CANVA_ENABLED = True
except ImportError as e:
    CANVA_ENABLED = False
    print(f"[WARN] Canva handler not available: {e}")
    def get_current_tokens():
        return os.environ.get("CANVA_ACCESS_TOKEN"), os.environ.get("CANVA_REFRESH_TOKEN")

# API一覧・コスト取得モジュール
try:
    from api_manager import register_api_commands, APICostView
    API_MANAGER_ENABLED = True
except ImportError as e:
    API_MANAGER_ENABLED = False
    print(f"[WARN] API Manager not available: {e}")

load_dotenv()

# 環境変数（全て遅延読み込み - Railway Railpack対策）
# os.environ.get() を使用（os.getenv検出を回避）
def get_line_token():
    return os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

def get_discord_token():
    return os.environ.get("DISCORD_BOT_TOKEN")

def get_guild_id():
    return os.environ.get("DISCORD_GUILD_ID")

def get_category_active():
    return os.environ.get("DISCORD_CATEGORY_ACTIVE")

def get_category_shipped():
    return os.environ.get("DISCORD_CATEGORY_SHIPPED")

def get_overview_channel():
    return os.environ.get("DISCORD_OVERVIEW_CHANNEL")

def get_forum_completed():
    return os.environ.get("DISCORD_FORUM_COMPLETED")

def get_forum_line():
    return os.environ.get("DISCORD_FORUM_LINE", "1463460598493745225")

def get_forum_atelier():
    return os.environ.get("DISCORD_FORUM_ATELIER", "1472857095031488524")

def get_atelier_webhook_url():
    return os.environ.get("ATELIER_WEBHOOK_URL", "https://i-tategu-shop.com/wp-json/i-tategu/v1/atelier/webhook")

def get_atelier_webhook_secret():
    return os.environ.get("ATELIER_WEBHOOK_SECRET", "")

def get_atelier_inquiry_webhook_url():
    return os.environ.get("ATELIER_INQUIRY_WEBHOOK_URL", "https://i-tategu-shop.com/wp-json/i-tategu/v1/atelier/inquiry/webhook")

def get_canva_access_token():
    access, _ = get_current_tokens()
    return access or os.environ.get("CANVA_ACCESS_TOKEN")

def get_canva_refresh_token():
    _, refresh = get_current_tokens()
    return refresh or os.environ.get("CANVA_REFRESH_TOKEN")

def get_canva_webhook_url():
    return os.environ.get("DISCORD_WEBHOOK_URL")

def get_wc_url():
    return os.environ.get("WC_URL")

def get_wc_consumer_key():
    return os.environ.get("WC_CONSUMER_KEY")

def get_wc_consumer_secret():
    return os.environ.get("WC_CONSUMER_SECRET")

def get_woo_webhook_secret():
    return os.environ.get("WOO_WEBHOOK_SECRET", "")

def get_instagram_page_token():
    return os.environ.get("INSTAGRAM_PAGE_TOKEN", "")

def get_instagram_app_secret():
    return os.environ.get("INSTAGRAM_APP_SECRET", "")

# スレッドマップファイル
THREAD_MAP_FILE = os.path.join(os.path.dirname(__file__), "thread_map.json")
INSTAGRAM_THREAD_MAP_FILE = os.path.join(os.path.dirname(__file__), "instagram_thread_map.json")

# Flask API
api = Flask(__name__)
api.secret_key = os.environ.get("FLASK_SECRET_KEY", os.urandom(24).hex())

# 商品登録ルート登録
if PRODUCT_REGISTER_ENABLED:
    product_register.register_routes(api)
    print("[OK] Product register routes enabled")

# Discord Bot設定
intents = discord.Intents.default()
intents.message_content = True
intents.guilds = True
bot = commands.Bot(command_prefix="!", intents=intents)

# グローバル変数
overview_message_id = None


def send_line_message(user_id, messages):
    """LINEにメッセージ送信"""
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {get_line_token()}"
    }
    data = {"to": user_id, "messages": messages}
    response = requests.post(url, headers=headers, json=data)
    return response.status_code == 200


def get_line_user_id_from_channel(channel):
    """チャンネルのトピックからLINE User IDを取得"""
    if not channel.topic:
        return None
    match = re.search(r'LINE User ID:\s*(\S+)', channel.topic)
    if match:
        return match.group(1)
    return None


def load_thread_map():
    """スレッドマップを読み込み"""
    if os.path.exists(THREAD_MAP_FILE):
        with open(THREAD_MAP_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def get_line_user_id_from_thread(thread_id):
    """スレッドIDからLINE User IDを取得"""
    thread_map = load_thread_map()
    for line_user_id, data in thread_map.items():
        if str(data.get('thread_id')) == str(thread_id):
            return line_user_id
    return None


def get_all_line_users_from_thread(thread_id):
    """スレッドIDから全LINE User IDと表示名を取得（複数ユーザー対応）"""
    thread_map = load_thread_map()
    users = []
    for line_user_id, data in thread_map.items():
        if str(data.get('thread_id')) == str(thread_id):
            users.append({
                'line_user_id': line_user_id,
                'display_name': data.get('display_name', '不明')
            })
    return users


def load_instagram_thread_map():
    """Instagramスレッドマップを読み込み"""
    if os.path.exists(INSTAGRAM_THREAD_MAP_FILE):
        with open(INSTAGRAM_THREAD_MAP_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def get_instagram_user_id_from_thread(thread_id):
    """スレッドIDからInstagram User IDを取得"""
    ig_map = load_instagram_thread_map()
    for ig_user_id, data in ig_map.items():
        if str(data.get('thread_id')) == str(thread_id):
            return ig_user_id
    return None


def get_platform_from_thread(thread_id):
    """スレッドIDからプラットフォームを判定（'line', 'instagram', None）"""
    # LINE thread_map をチェック
    line_map = load_thread_map()
    for _, data in line_map.items():
        if str(data.get('thread_id')) == str(thread_id):
            return 'line'

    # Instagram thread_map をチェック
    ig_map = load_instagram_thread_map()
    for _, data in ig_map.items():
        if str(data.get('thread_id')) == str(thread_id):
            return 'instagram'

    return None


def send_instagram_message(user_id, text):
    """Instagram DM でテキストメッセージを送信"""
    token = get_instagram_page_token()
    if not token:
        print("[IG] No INSTAGRAM_PAGE_TOKEN configured")
        return False

    url = "https://graph.instagram.com/v18.0/me/messages"
    data = {
        "recipient": {"id": user_id},
        "message": {"text": text}
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    try:
        response = requests.post(url, json=data, headers=headers, timeout=10)
        if response.status_code == 200:
            print(f"[IG] Message sent to {user_id}")
            return True
        else:
            print(f"[IG] Send failed: {response.status_code} {response.text}")
            return False
    except Exception as e:
        print(f"[IG] Send error: {e}")
        return False


def send_instagram_image(user_id, image_url):
    """Instagram DM で画像を送信"""
    token = get_instagram_page_token()
    if not token:
        print("[IG] No INSTAGRAM_PAGE_TOKEN configured")
        return False

    url = "https://graph.instagram.com/v18.0/me/messages"
    data = {
        "recipient": {"id": user_id},
        "message": {
            "attachment": {
                "type": "image",
                "payload": {"url": image_url}
            }
        }
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    try:
        response = requests.post(url, json=data, headers=headers, timeout=10)
        if response.status_code == 200:
            print(f"[IG] Image sent to {user_id}")
            return True
        else:
            print(f"[IG] Image send failed: {response.status_code} {response.text}")
            return False
    except Exception as e:
        print(f"[IG] Image send error: {e}")
        return False


# テンプレート（DATA_DIRに保存版があればそちらを優先）
_TEMPLATES_BUNDLED = os.path.join(os.path.dirname(__file__), "line_templates.json")
_DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(__file__))
_TEMPLATES_SAVED = os.path.join(_DATA_DIR, "line_templates.json")

# テンプレートボタンメッセージID追跡（スレッドID → メッセージID）
_template_button_msg_ids = {}
_posting_buttons_lock = set()  # 再投稿ループ防止


def load_templates():
    """LINEテンプレートを読み込み（DATA_DIR優先）"""
    path = _TEMPLATES_SAVED if os.path.exists(_TEMPLATES_SAVED) else _TEMPLATES_BUNDLED
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get("templates", [])
    return []


def save_templates(templates):
    """テンプレートをDATA_DIRに保存"""
    with open(_TEMPLATES_SAVED, 'w', encoding='utf-8') as f:
        json.dump({"templates": templates}, f, ensure_ascii=False, indent=2)


def get_thread_customer_info(thread):
    """フォーラムスレッドから顧客情報を取得"""
    name_match = re.search(r'[\U0001F7E0\U0001F7E1\U0001F535\U0001F7E2\u2705\U0001F4E6\U0001F389\U0001F490\U0001F64F\U0001F4AC\U0001F3A8]\s*(?:#\d+\s+)?(.+?)\s*様', thread.name)
    customer_name = name_match.group(1) if name_match else "お客様"

    order_match = re.search(r'#(\d+)', thread.name)
    order_id = order_match.group(1) if order_match else None

    return customer_name, order_id


async def find_line_user_id_in_thread(thread):
    """スレッドからLINE User IDを検索"""
    line_user_id = get_line_user_id_from_thread(thread.id)
    if line_user_id:
        return line_user_id

    async for msg in thread.history(limit=5, oldest_first=True):
        if msg.content:
            match = re.search(r'LINE User ID:\s*`?([A-Za-z0-9]+)`?', msg.content)
            if match:
                return match.group(1)
        for embed in msg.embeds:
            embed_text = (embed.description or "")
            for field in embed.fields:
                embed_text += f" {field.name} {field.value}"
            match = re.search(r'LINE User ID:\s*`?([A-Za-z0-9]+)`?', embed_text)
            if match:
                return match.group(1)

    return None


async def create_status_embed():
    """ステータス一覧のEmbed作成"""
    summary = get_status_summary()

    embeds = []

    header = discord.Embed(
        title="📊 顧客ステータス一覧",
        description="各ステータスの顧客数と詳細",
        color=0x5865F2
    )
    header.set_footer(text="名前をクリックでチャンネルへジャンプ")
    embeds.append(header)

    for status in CustomerStatus:
        data = summary[status.value]
        config = STATUS_CONFIG[status]

        embed = discord.Embed(
            title=f"{config['emoji']} {config['label']} ({data['count']}件)",
            color=config['color']
        )

        if data['customers']:
            customer_links = []
            for c in data['customers']:
                channel_id = c.get('discord_channel_id')
                name = c.get('display_name', '不明')
                # 注文番号を取得
                order_num = ""
                if c.get('orders'):
                    latest_order = c['orders'][-1]
                    order_num = f"#{latest_order.get('order_id', '')} "
                if channel_id:
                    customer_links.append(f"• {order_num}<#{channel_id}> {name}様")
                else:
                    customer_links.append(f"• {order_num}{name}様")

            # Embed文字数制限(4096)対策: 超える場合は複数Embedに分割
            chunk = []
            chunk_len = 0
            for line in customer_links:
                if chunk_len + len(line) + 1 > 4000 and chunk:
                    embed.description = "\n".join(chunk)
                    embeds.append(embed)
                    embed = discord.Embed(
                        title=f"{config['emoji']} {config['label']} (続き)",
                        color=config['color']
                    )
                    chunk = []
                    chunk_len = 0
                chunk.append(line)
                chunk_len += len(line) + 1
            embed.description = "\n".join(chunk) if chunk else "_該当なし_"
        else:
            embed.description = "_該当なし_"

        embeds.append(embed)

    return embeds


async def update_overview_channel():
    """一覧チャンネルを更新"""
    global overview_message_id

    if not get_overview_channel():
        return

    guild = bot.get_guild(int(get_guild_id()))
    if not guild:
        return

    channel = guild.get_channel(int(get_overview_channel()))
    if not channel:
        return

    embeds = await create_status_embed()

    try:
        if overview_message_id:
            try:
                message = await channel.fetch_message(overview_message_id)
                await message.edit(embeds=embeds)
                return
            except discord.NotFound:
                pass

        async for msg in channel.history(limit=10):
            if msg.author == bot.user:
                await msg.delete()

        message = await channel.send(embeds=embeds)
        overview_message_id = message.id

    except Exception as e:
        print(f"[ERROR] Failed to update overview: {e}")


async def update_atelier_thread_status(order_id, new_status: CustomerStatus):
    """アトリエフォーラムスレッドのステータス絵文字・タグを更新"""
    if not get_forum_atelier():
        return

    guild = bot.get_guild(int(get_guild_id()))
    if not guild:
        return

    forum = guild.get_channel(int(get_forum_atelier()))
    if not forum or not isinstance(forum, discord.ForumChannel):
        return

    config = STATUS_CONFIG[new_status]
    target_prefix = f"#{order_id} "

    # アクティブスレッドから該当注文を検索
    for thread in forum.threads:
        if target_prefix in thread.name:
            try:
                # スレッド名の絵文字更新
                new_name = re.sub(
                    r'^[\U0001F7E0\U0001F7E1\U0001F535\U0001F7E2\u2705\U0001F4E6\U0001F389\U0001F490\U0001F64F]\s*',
                    '',
                    thread.name
                )
                new_name = f"{config['emoji']} {new_name}"
                kwargs = {'name': new_name}

                # フォーラムタグ更新
                target_tag = None
                for tag in forum.available_tags:
                    if config['label'] in tag.name or config['emoji'] in (getattr(tag, 'emoji', None) or ''):
                        target_tag = tag
                        break
                if target_tag:
                    kwargs['applied_tags'] = [target_tag]

                await thread.edit(**kwargs)
                print(f"[Atelier] Updated thread: {new_name}")
            except Exception as e:
                print(f"[Atelier] Thread update failed: {e}")
            return

    # アーカイブスレッドも検索
    try:
        async for thread in forum.archived_threads(limit=50):
            if target_prefix in thread.name:
                try:
                    await thread.edit(archived=False)
                    new_name = re.sub(
                        r'^[\U0001F7E0\U0001F7E1\U0001F535\U0001F7E2\u2705\U0001F4E6\U0001F389\U0001F490\U0001F64F]\s*',
                        '',
                        thread.name
                    )
                    new_name = f"{config['emoji']} {new_name}"
                    await thread.edit(name=new_name)
                    print(f"[Atelier] Updated archived thread: {new_name}")
                except Exception as e:
                    print(f"[Atelier] Archived thread update failed: {e}")
                return
    except Exception as e:
        print(f"[Atelier] Archived thread search failed: {e}")


async def move_channel_to_category(channel, category_id):
    """チャンネルを別カテゴリに移動"""
    if not category_id:
        return False

    guild = channel.guild
    category = guild.get_channel(int(category_id))

    if category and isinstance(category, discord.CategoryChannel):
        try:
            await channel.edit(category=category)
            return True
        except Exception as e:
            print(f"[ERROR] Failed to move channel: {e}")

    return False


async def archive_channel_to_forum(channel, customer_name=None):
    """チャンネルをフォーラムにアーカイブ"""
    if not get_forum_completed():
        print("[WARN] get_forum_completed() not set")
        return False

    guild = channel.guild
    forum = guild.get_channel(int(get_forum_completed()))

    if not forum or not isinstance(forum, discord.ForumChannel):
        print(f"[ERROR] Forum channel not found: {get_forum_completed()}")
        return False

    try:
        if not customer_name:
            customer_name = channel.name.replace("line-", "")

        from datetime import datetime
        year_month = datetime.now().strftime("%Y-%m")

        thread_title = f"[{year_month}] {customer_name} 様"

        messages = []
        async for msg in channel.history(limit=500, oldest_first=True):
            messages.append(msg)

        if not messages:
            thread, _ = await forum.create_thread(
                name=thread_title,
                content=f"📋 {customer_name} 様のやり取り履歴\n（メッセージなし）"
            )
        else:
            first_content = f"📋 **{customer_name} 様** のやり取り履歴\n"
            first_content += f"━━━━━━━━━━━━━━━━━━━━━━\n"
            first_content += f"📅 期間: {messages[0].created_at.strftime('%Y-%m-%d')} 〜 {messages[-1].created_at.strftime('%Y-%m-%d')}\n"
            first_content += f"💬 メッセージ数: {len(messages)}件\n"
            first_content += f"━━━━━━━━━━━━━━━━━━━━━━"

            applied_tags = []
            for tag in forum.available_tags:
                if year_month in tag.name:
                    applied_tags.append(tag)
                    break

            thread, _ = await forum.create_thread(
                name=thread_title,
                content=first_content,
                applied_tags=applied_tags[:5] if applied_tags else None
            )

            for msg in messages:
                timestamp = msg.created_at.strftime("%m/%d %H:%M")
                author = msg.author.display_name

                content_parts = []
                if msg.content:
                    content_parts.append(msg.content)

                for att in msg.attachments:
                    content_parts.append(att.url)

                if content_parts:
                    formatted = f"**[{timestamp}] {author}**\n" + "\n".join(content_parts)
                    if len(formatted) > 2000:
                        formatted = formatted[:1997] + "..."
                    await thread.send(formatted)

        print(f"[OK] Archived to forum: {thread_title}")

        await channel.delete(reason="フォーラムにアーカイブ済み")
        print(f"[OK] Deleted original channel: {channel.name}")

        return True

    except Exception as e:
        print(f"[ERROR] Failed to archive: {e}")
        import traceback
        traceback.print_exc()
        return False


# ================== Discord Events ==================

@bot.event
async def on_ready():
    print("=" * 50)
    print("Discord Bot v2 - Server Edition")
    print("=" * 50)
    print(f"[OK] Logged in as: {bot.user}")
    print(f"[OK] Application ID: {bot.application_id}")

    # Persistent Viewを登録（Bot再起動後もボタンが動作）
    bot.add_view(TemplatePersistentView())

    # API一覧・コスト取得の Persistent View とコマンド登録
    if API_MANAGER_ENABLED:
        bot.add_view(APICostView())
        register_api_commands(bot)
        print("[OK] API Manager commands registered")

    try:
        guild = discord.Object(id=int(get_guild_id()))
        bot.tree.copy_global_to(guild=guild)
        await bot.tree.sync(guild=guild)
        print("[OK] Slash commands synced to EC guild")
    except Exception as e:
        print(f"[WARN] Failed to sync commands to EC guild: {e}")

    # 追加サーバーにもコマンドを同期（環境変数があれば）
    for env_key, label in [
        ("DEV_LOG_GUILD_ID", "dev log"),
        ("OPENCLAW_GUILD_ID", "OpenClaw"),
    ]:
        extra_guild_id = os.environ.get(env_key)
        if extra_guild_id:
            try:
                extra_guild = discord.Object(id=int(extra_guild_id))
                bot.tree.copy_global_to(guild=extra_guild)
                await bot.tree.sync(guild=extra_guild)
                print(f"[OK] Slash commands synced to {label} guild")
            except Exception as e:
                print(f"[WARN] Failed to sync commands to {label} guild: {e}")

    await update_overview_channel()
    print("[OK] Overview channel updated")
    print("=" * 50)


async def handle_atelier_message(message):
    """#atelier フォーラムのメッセージをWordPress webhook に転送（注文 & 問い合わせ対応）"""
    thread_name = message.channel.name
    secret = get_atelier_webhook_secret()

    # 💬 プレフィックス → 問い合わせスレッド
    is_inquiry = thread_name.startswith('💬')

    if is_inquiry:
        # 問い合わせ: スレッド名から inquiry_id を取得（例: "💬 #1 石橋伯昂 様"）
        id_match = re.search(r'#(\d+)', thread_name)
        if not id_match:
            print(f"[Atelier Inquiry] Could not extract inquiry ID from thread: {thread_name}")
            return
        inquiry_id = id_match.group(1)
        webhook_url = get_atelier_inquiry_webhook_url()
    else:
        # 注文: スレッド名から order_id を取得（例: "🎨 #1865 はるか 様"）
        order_match = re.search(r'#(\d+)', thread_name)
        if not order_match:
            print(f"[Atelier] Could not extract order ID from thread: {thread_name}")
            return
        order_id = order_match.group(1)
        webhook_url = get_atelier_webhook_url()

    if not webhook_url or not secret:
        print("[Atelier] Webhook URL or secret not configured")
        return

    # テキストメッセージ
    text = message.content if message.content and not message.content.startswith("!") else ""

    # 画像URL（最初の画像添付のみ）
    image_url = ""
    for attachment in message.attachments:
        if attachment.content_type and attachment.content_type.startswith("image/"):
            image_url = attachment.url
            break

    if not text and not image_url:
        return

    if is_inquiry:
        payload = {
            "inquiry_id": int(inquiry_id),
            "message": text,
            "image_url": image_url,
        }
    else:
        payload = {
            "order_id": int(order_id),
            "message": text,
            "image_url": image_url,
        }

    try:
        resp = requests.post(webhook_url, json=payload, headers={
            "X-Atelier-Secret": secret,
            "Content-Type": "application/json",
        }, timeout=10)

        if resp.status_code == 200:
            await message.add_reaction("✅")
            label = f"inquiry={inquiry_id}" if is_inquiry else f"order={order_id}"
            print(f"[Atelier] Forwarded to WP: {label}")
        else:
            await message.add_reaction("❌")
            print(f"[Atelier] WP webhook failed: {resp.status_code} {resp.text}")
    except Exception as e:
        await message.add_reaction("❌")
        print(f"[Atelier] Webhook error: {e}")


@bot.event
async def on_raw_reaction_add(payload):
    """👀リアクションで既読マーク（#atelierフォーラムのみ）"""
    if str(payload.emoji) != '👀':
        return
    if payload.user_id == bot.user.id:
        return

    # アトリエフォーラムのスレッドか確認
    forum_id = get_forum_atelier()
    if not forum_id:
        return

    channel = bot.get_channel(payload.channel_id)
    if not isinstance(channel, discord.Thread):
        return
    if str(channel.parent_id) != str(forum_id):
        return

    thread_name = channel.name
    secret = get_atelier_webhook_secret()
    if not secret:
        return

    is_inquiry = thread_name.startswith('💬')
    id_match = re.search(r'#(\d+)', thread_name)
    if not id_match:
        return

    target_id = id_match.group(1)

    if is_inquiry:
        webhook_url = get_atelier_inquiry_webhook_url()
        payload_data = {"inquiry_id": int(target_id), "mark_read": True}
    else:
        webhook_url = get_atelier_webhook_url()
        payload_data = {"order_id": int(target_id), "mark_read": True}

    try:
        resp = requests.post(webhook_url, json=payload_data, headers={
            "X-Atelier-Secret": secret,
            "Content-Type": "application/json",
        }, timeout=10)
        if resp.status_code == 200:
            print(f"[Atelier] Marked as read via 👀: {'inquiry' if is_inquiry else 'order'}={target_id}")
        else:
            print(f"[Atelier] Mark read failed: {resp.status_code}")
    except Exception as e:
        print(f"[Atelier] Mark read error: {e}")


@bot.event
async def on_error(event, *args, **kwargs):
    """エラーログ"""
    import traceback
    print(f"[ERROR] Event: {event}")
    traceback.print_exc()


@bot.tree.error
async def on_app_command_error(interaction: discord.Interaction, error):
    """アプリコマンドエラー"""
    print(f"[ERROR] App command error: {error}")
    import traceback
    traceback.print_exc()


@bot.event
async def on_message(message):
    """Discordメッセージを監視してLINEに転送 + テンプレートボタン再投稿"""
    print(f"[MSG] channel={message.channel.name if hasattr(message.channel, 'name') else 'DM'}, author={message.author}, bot={message.author.bot}")

    # LINE対応/アトリエ フォーラムスレッド内のメッセージ → テンプレートボタン再投稿
    if isinstance(message.channel, discord.Thread):
        parent_id = str(message.channel.parent_id)
        is_line_forum = get_forum_line() and parent_id == str(get_forum_line())
        is_atelier_forum = get_forum_atelier() and parent_id == str(get_forum_atelier())
        if is_line_forum or is_atelier_forum:
            thread_key = str(message.channel.id)
            # 自分が投稿したボタンメッセージは無視（ループ防止）
            if message.id != _template_button_msg_ids.get(thread_key):
                # 送信記録Embed（📤）も無視
                is_sent_record = False
                for embed in message.embeds:
                    if embed.author and embed.author.name and "📤" in embed.author.name:
                        is_sent_record = True
                        break
                if not is_sent_record:
                    await post_template_buttons(message.channel)

    if message.author == bot.user:
        return

    if message.author.bot:
        return

    await bot.process_commands(message)

    # ── #atelier フォーラムスレッド → WordPress webhook 転送 ──
    if isinstance(message.channel, discord.Thread) and get_forum_atelier():
        if str(message.channel.parent_id) == str(get_forum_atelier()):
            await handle_atelier_message(message)
            return  # LINE転送は不要

    # ── #LINE対応 フォーラムスレッド → LINE / Instagram 転送 ──
    if not (isinstance(message.channel, discord.Thread) and
            message.channel.parent_id == int(get_forum_line())):
        # フォーラムスレッド外 → 通常チャンネルからの転送（トピックにLINE User IDがあれば転送）
        line_user_id = None
        if hasattr(message.channel, 'topic'):
            line_user_id = get_line_user_id_from_channel(message.channel)
        if not line_user_id:
            return
        # 通常チャンネル → LINE 送信（従来互換）
        if message.content and not message.content.startswith("!"):
            success = send_line_message(line_user_id, [{"type": "text", "text": message.content}])
            if success:
                await message.add_reaction("✅")
            else:
                await message.add_reaction("❌")
        for attachment in message.attachments:
            if attachment.content_type and attachment.content_type.startswith("image/"):
                send_line_message(line_user_id, [{
                    "type": "image",
                    "originalContentUrl": attachment.url,
                    "previewImageUrl": attachment.url
                }])
        return

    # ── フォーラムスレッド内: プラットフォーム判定 ──
    thread_id = message.channel.id
    platform = get_platform_from_thread(thread_id)
    print(f"[DEBUG] Thread {thread_id}: platform={platform}")

    # ── Instagram スレッドの場合 ──
    if platform == 'instagram':
        ig_user_id = get_instagram_user_id_from_thread(thread_id)
        if not ig_user_id:
            print(f"[DEBUG] No Instagram User ID found for thread: {thread_id}")
            return

        # テキスト送信
        if message.content and not message.content.startswith("!"):
            success = send_instagram_message(ig_user_id, message.content)
            if success:
                await message.add_reaction("✅")
            else:
                await message.add_reaction("❌")

        # 画像送信
        for attachment in message.attachments:
            if attachment.content_type and attachment.content_type.startswith("image/"):
                success = send_instagram_image(ig_user_id, attachment.url)
                if success:
                    await message.add_reaction("🖼️")
        return

    # ── LINE スレッドの場合（従来ロジック）──
    line_user_id = get_line_user_id_from_thread(thread_id)
    if not line_user_id:
        starter = message.channel.starter_message
        if starter:
            match = re.search(r'LINE User ID:\s*`?([A-Za-z0-9]+)`?', starter.content)
            if match:
                line_user_id = match.group(1)

        if not line_user_id:
            async for msg in message.channel.history(limit=5, oldest_first=True):
                if msg.content:
                    match = re.search(r'LINE User ID:\s*`?([A-Za-z0-9]+)`?', msg.content)
                    if match:
                        line_user_id = match.group(1)
                        print(f"[DEBUG] Found LINE User ID in content: {line_user_id}")
                        break

                for embed in msg.embeds:
                    embed_text = ""
                    if embed.description:
                        embed_text += embed.description
                    for field in embed.fields:
                        embed_text += f" {field.name} {field.value}"

                    match = re.search(r'LINE User ID:\s*`?([A-Za-z0-9]+)`?', embed_text)
                    if match:
                        line_user_id = match.group(1)
                        print(f"[DEBUG] Found LINE User ID in embed: {line_user_id}")
                        break

                if line_user_id:
                    break

    if not line_user_id:
        print(f"[DEBUG] No LINE User ID found for channel: {message.channel.name}")
        return

    print(f"[DEBUG] LINE User ID found: {line_user_id}")

    # 複数LINEユーザー対応（夫婦連携）
    all_line_users = get_all_line_users_from_thread(thread_id)
    if len(all_line_users) > 1:
        has_content = message.content and not message.content.startswith("!")
        attachment_data = [
            {'url': att.url, 'content_type': att.content_type}
            for att in message.attachments
            if att.content_type and att.content_type.startswith("image/")
        ]
        if has_content or attachment_data:
            view = ReplyTargetView(all_line_users, message.content if has_content else "", attachment_data)
            names = " / ".join(u['display_name'] for u in all_line_users)
            await message.reply(f"📨 送信先を選択してください（{names}）", view=view, mention_author=False)
        return

    # テキストメッセージ送信（単一ユーザー）
    if message.content and not message.content.startswith("!"):
        success = send_line_message(line_user_id, [
            {"type": "text", "text": message.content}
        ])
        if success:
            await message.add_reaction("✅")
        else:
            await message.add_reaction("❌")

    # 画像送信
    for attachment in message.attachments:
        if attachment.content_type and attachment.content_type.startswith("image/"):
            success = send_line_message(line_user_id, [
                {
                    "type": "image",
                    "originalContentUrl": attachment.url,
                    "previewImageUrl": attachment.url
                }
            ])
            if success:
                await message.add_reaction("🖼️")


# ================== Button Interactions ==================

@bot.event
async def on_interaction(interaction: discord.Interaction):
    """ボタンクリック処理"""
    if interaction.type != discord.InteractionType.component:
        return

    custom_id = interaction.data.get("custom_id", "")

    # B2用コピーボタン
    if custom_id.startswith("b2_copy_"):
        order_id = custom_id.replace("b2_copy_", "")
        await handle_b2_copy(interaction, order_id)

    # B2自動入力ボタン（キューセット）
    elif custom_id.startswith("b2_autofill_"):
        order_id = custom_id.replace("b2_autofill_", "")
        await handle_b2_autofill(interaction, order_id)

    # 発送完了ボタン
    elif custom_id.startswith("shipped_"):
        order_id = custom_id.replace("shipped_", "")
        await handle_shipped(interaction, order_id)


async def handle_b2_autofill(interaction: discord.Interaction, order_id: str):
    """B2自動入力キューをセット（Tampermonkeyがポーリングで検出）"""
    await interaction.response.defer(ephemeral=True)

    wc_url = get_wc_url()
    if not wc_url:
        await interaction.followup.send("WC_URL設定がありません", ephemeral=True)
        return

    try:
        url = f"{wc_url}/wp-json/i-tategu/v1/b2-queue"
        shipping_token = os.environ.get("SHIPPING_API_TOKEN", "itg_ship_2026")
        response = requests.post(
            url,
            json={"order_id": order_id},
            headers={
                "X-Shipping-Token": shipping_token,
                "Content-Type": "application/json",
            }
        )

        if response.status_code == 200:
            data = response.json()
            if data.get("success"):
                await interaction.followup.send(
                    f"✅ 注文 #{order_id} をB2自動入力キューにセットしました\n"
                    f"B2クラウドのかんたん発行画面を開いていれば、2秒以内に自動入力されます。",
                    ephemeral=True
                )
            else:
                await interaction.followup.send(f"キュー設定失敗: {data}", ephemeral=True)
        else:
            await interaction.followup.send(f"API呼び出し失敗: {response.status_code}", ephemeral=True)

    except Exception as e:
        await interaction.followup.send(f"エラー: {e}", ephemeral=True)


async def handle_b2_copy(interaction: discord.Interaction, order_id: str):
    """B2クラウド用データを表示"""
    await interaction.response.defer(ephemeral=True)

    # WooCommerceから注文情報取得
    wc_url = get_wc_url()
    wc_key = get_wc_consumer_key()
    wc_secret = get_wc_consumer_secret()

    if not all([wc_url, wc_key, wc_secret]):
        await interaction.followup.send("WooCommerce設定がありません", ephemeral=True)
        return

    try:
        url = f"{wc_url}/wp-json/wc/v3/orders/{order_id}"
        response = requests.get(url, auth=(wc_key, wc_secret))
        if response.status_code != 200:
            await interaction.followup.send(f"注文取得失敗: {response.status_code}", ephemeral=True)
            return

        order = response.json()
        billing = order.get('billing', {})
        shipping = order.get('shipping', {})

        # 発送先情報
        postcode = shipping.get('postcode') or billing.get('postcode', '')
        state = shipping.get('state') or billing.get('state', '')
        city = shipping.get('city') or billing.get('city', '')
        address1 = shipping.get('address_1') or billing.get('address_1', '')
        address2 = shipping.get('address_2') or billing.get('address_2', '')

        # 都道府県コード変換
        JP_STATES = {
            'JP01': '北海道', 'JP02': '青森県', 'JP03': '岩手県', 'JP04': '宮城県',
            'JP05': '秋田県', 'JP06': '山形県', 'JP07': '福島県', 'JP08': '茨城県',
            'JP09': '栃木県', 'JP10': '群馬県', 'JP11': '埼玉県', 'JP12': '千葉県',
            'JP13': '東京都', 'JP14': '神奈川県', 'JP15': '新潟県', 'JP16': '富山県',
            'JP17': '石川県', 'JP18': '福井県', 'JP19': '山梨県', 'JP20': '長野県',
            'JP21': '岐阜県', 'JP22': '静岡県', 'JP23': '愛知県', 'JP24': '三重県',
            'JP25': '滋賀県', 'JP26': '京都府', 'JP27': '大阪府', 'JP28': '兵庫県',
            'JP29': '奈良県', 'JP30': '和歌山県', 'JP31': '鳥取県', 'JP32': '島根県',
            'JP33': '岡山県', 'JP34': '広島県', 'JP35': '山口県', 'JP36': '徳島県',
            'JP37': '香川県', 'JP38': '愛媛県', 'JP39': '高知県', 'JP40': '福岡県',
            'JP41': '佐賀県', 'JP42': '長崎県', 'JP43': '熊本県', 'JP44': '大分県',
            'JP45': '宮崎県', 'JP46': '鹿児島県', 'JP47': '沖縄県'
        }
        state_name = JP_STATES.get(state, state)

        full_address = f"{city}{address1}"
        if address2:
            full_address += f" {address2}"

        customer_name = f"{billing.get('last_name', '')} {billing.get('first_name', '')}"
        customer_phone = billing.get('phone', '')

        # 商品名
        products = [item.get('name', '') for item in order.get('line_items', [])]
        product_name = products[0] if products else "一枚板結婚証明書"

        # B2クラウド用フォーマット（コピペ用）
        b2_data = f"""```
【B2クラウド入力用】
━━━━━━━━━━━━━━━━━━━━
郵便番号: {postcode}
都道府県: {state_name}
市区町村: {city}
番地: {address1}
建物名等: {address2 or ""}
━━━━━━━━━━━━━━━━━━━━
届け先名: {customer_name}
電話番号: {customer_phone}
━━━━━━━━━━━━━━━━━━━━
品名: {product_name}
個数: 1
━━━━━━━━━━━━━━━━━━━━
```"""

        await interaction.followup.send(b2_data, ephemeral=True)

    except Exception as e:
        await interaction.followup.send(f"エラー: {e}", ephemeral=True)


async def handle_shipped(interaction: discord.Interaction, order_id: str):
    """発送完了処理"""
    await interaction.response.defer()

    # WooCommerceのステータス更新
    wc_url = get_wc_url()
    wc_key = get_wc_consumer_key()
    wc_secret = get_wc_consumer_secret()

    if not all([wc_url, wc_key, wc_secret]):
        await interaction.followup.send("WooCommerce設定がありません")
        return

    try:
        url = f"{wc_url}/wp-json/wc/v3/orders/{order_id}"
        response = requests.put(url, auth=(wc_key, wc_secret), json={"status": "completed"})

        if response.status_code == 200:
            # メッセージを更新（ボタン無効化 + 色変更）
            message = interaction.message
            embed = message.embeds[0].to_dict() if message.embeds else {}
            embed["title"] = embed.get("title", "").replace("🟡 未発送", "✅ 発送済み")
            embed["color"] = 0x2ECC71  # 緑

            # ボタンを無効化
            disabled_components = [
                {
                    "type": 1,
                    "components": [
                        {"type": 2, "style": 2, "label": "📋 B2用コピー", "custom_id": f"b2_copy_{order_id}", "disabled": True},
                        {"type": 2, "style": 2, "label": "✅ 発送完了", "custom_id": f"shipped_{order_id}", "disabled": True},
                    ]
                }
            ]

            await message.edit(embed=discord.Embed.from_dict(embed), components=disabled_components)
            await interaction.followup.send(f"✅ 注文 #{order_id} を発送済みに更新しました")
        else:
            await interaction.followup.send(f"ステータス更新失敗: {response.status_code}")

    except Exception as e:
        await interaction.followup.send(f"エラー: {e}")


# ================== Slash Commands ==================

@bot.tree.command(name="status", description="顧客のステータスを変更")
@app_commands.describe(new_status="新しいステータス")
@app_commands.choices(new_status=[
    app_commands.Choice(name="🟡 購入済み", value="purchased"),
    app_commands.Choice(name="🔵 デザイン確定", value="design-confirmed"),
    app_commands.Choice(name="🟢 制作完了", value="produced"),
    app_commands.Choice(name="📦 発送済み", value="shipped"),
])
async def change_status(interaction: discord.Interaction, new_status: str):
    """ステータス変更コマンド"""
    channel = interaction.channel

    if not channel.name.startswith("line-"):
        await interaction.response.send_message("このコマンドはLINE顧客チャンネルでのみ使用できます", ephemeral=True)
        return

    line_user_id = get_line_user_id_from_channel(channel)
    if not line_user_id:
        await interaction.response.send_message("LINE User IDが見つかりません", ephemeral=True)
        return

    try:
        status = CustomerStatus(new_status)
    except ValueError:
        await interaction.response.send_message("無効なステータスです", ephemeral=True)
        return

    customer = get_customer(line_user_id)
    if not customer:
        add_customer(line_user_id, channel.name.replace("line-", ""), str(channel.id))

    update_customer_status(line_user_id, status)

    config = STATUS_CONFIG[status]
    await interaction.response.send_message(
        f"{config['emoji']} ステータスを **{config['label']}** に変更しました"
    )

    if status == CustomerStatus.SHIPPED and get_forum_completed():
        await channel.send("📦 完了一覧にアーカイブ中...")
        customer = get_customer(line_user_id)
        customer_name = customer.get('display_name') if customer else None
        await archive_channel_to_forum(channel, customer_name)
    elif status != CustomerStatus.SHIPPED and get_category_active():
        await move_channel_to_category(channel, get_category_active())

    # アトリエフォーラムスレッドも連動更新
    customer = get_customer(line_user_id)
    if customer and customer.get('orders'):
        for order in customer['orders']:
            await update_atelier_thread_status(order['order_id'], status)

    await update_overview_channel()


@bot.tree.command(name="atelier-url", description="アトリエURLを表示")
@app_commands.describe(order_id="注文番号")
async def atelier_url(interaction: discord.Interaction, order_id: int):
    """指定注文のアトリエURLを生成して表示"""
    await interaction.response.defer(ephemeral=True)

    wc_url = get_wc_url()
    wc_key = get_wc_consumer_key()
    wc_secret = get_wc_consumer_secret()

    if not all([wc_url, wc_key, wc_secret]):
        await interaction.followup.send("WooCommerce設定がありません", ephemeral=True)
        return

    try:
        url = f"{wc_url}/wp-json/wc/v3/orders/{order_id}"
        response = requests.get(url, auth=(wc_key, wc_secret))
        if response.status_code != 200:
            await interaction.followup.send(f"注文 #{order_id} が見つかりません (HTTP {response.status_code})", ephemeral=True)
            return

        order = response.json()
        meta = {m['key']: m['value'] for m in order.get('meta_data', [])}
        atelier_token = meta.get('_atelier_token')

        if not atelier_token:
            await interaction.followup.send(
                f"注文 #{order_id} にアトリエトークンがありません\n"
                f"ステータス: {order.get('status', '不明')}\n"
                f"※ トークンは processing/on-hold 時に自動生成されます",
                ephemeral=True
            )
            return

        atelier_url_str = f"{wc_url}/atelier/?order={order_id}&token={atelier_token}"
        billing = order.get('billing', {})
        customer_name = f"{billing.get('last_name', '')} {billing.get('first_name', '')}".strip()

        embed = discord.Embed(
            title=f"🎨 注文 #{order_id} のアトリエURL",
            color=0xc5a96a
        )
        embed.add_field(name="お客様", value=customer_name or "不明", inline=True)
        embed.add_field(name="ステータス", value=order.get('status', '不明'), inline=True)
        embed.add_field(name="アトリエURL", value=atelier_url_str, inline=False)
        embed.set_footer(text="このURLをインスタDM等でお客様にお送りください")

        await interaction.followup.send(embed=embed, ephemeral=True)

    except Exception as e:
        await interaction.followup.send(f"エラー: {e}", ephemeral=True)


@bot.tree.command(name="overview", description="顧客一覧を更新")
async def refresh_overview(interaction: discord.Interaction):
    """一覧更新コマンド"""
    await interaction.response.defer(ephemeral=True)
    await update_overview_channel()
    await interaction.followup.send("一覧を更新しました", ephemeral=True)


@bot.tree.command(name="register", description="このチャンネルの顧客を登録")
async def register_customer(interaction: discord.Interaction):
    """顧客登録コマンド"""
    channel = interaction.channel

    if not channel.name.startswith("line-"):
        await interaction.response.send_message("このコマンドはLINE顧客チャンネルでのみ使用できます", ephemeral=True)
        return

    line_user_id = get_line_user_id_from_channel(channel)
    if not line_user_id:
        await interaction.response.send_message("LINE User IDが見つかりません", ephemeral=True)
        return

    display_name = channel.name.replace("line-", "")
    add_customer(line_user_id, display_name, str(channel.id))

    await interaction.response.send_message(f"✅ {display_name}様を顧客リストに登録しました")
    await update_overview_channel()


# ================== Template System ==================

class ReplyTargetView(discord.ui.View):
    """複数LINE宛先がある場合の送信先選択UI"""
    def __init__(self, line_users, message_content, attachments=None):
        super().__init__(timeout=120)
        self.line_users = line_users
        self.message_content = message_content
        self.attachments = attachments or []

        options = []
        for user in line_users:
            options.append(discord.SelectOption(
                label=f"{user['display_name']}だけ",
                value=user['line_user_id'],
                description=f"{user['display_name']}様のみに送信"
            ))
        options.append(discord.SelectOption(
            label="両方に送信",
            value="__all__",
            description="全員に送信",
            default=True
        ))

        select = ReplyTargetSelect(options, line_users, message_content, attachments)
        self.add_item(select)


class ReplyTargetSelect(discord.ui.Select):
    """送信先選択セレクトメニュー"""
    def __init__(self, options, line_users, message_content, attachments):
        super().__init__(placeholder="送信先を選択...", options=options)
        self.line_users = line_users
        self.message_content = message_content
        self.attachments = attachments

    async def callback(self, interaction: discord.Interaction):
        selected = self.values[0]

        if selected == "__all__":
            targets = self.line_users
        else:
            targets = [u for u in self.line_users if u['line_user_id'] == selected]

        results = []
        for user in targets:
            uid = user['line_user_id']
            name = user['display_name']

            if self.message_content:
                success = send_line_message(uid, [{"type": "text", "text": self.message_content}])
                results.append(f"{'✅' if success else '❌'} {name}")

            for att in self.attachments:
                if att.get('content_type', '').startswith("image/"):
                    send_line_message(uid, [{
                        "type": "image",
                        "originalContentUrl": att['url'],
                        "previewImageUrl": att['url']
                    }])

        target_names = ", ".join(u['display_name'] for u in targets)
        await interaction.response.edit_message(
            content=f"✅ {target_names}様に送信しました",
            view=None
        )


class TemplateEditModal(discord.ui.Modal):
    """テンプレート編集モーダル（複数ユーザー対応 / Instagram対応）"""
    def __init__(self, template, customer_name, order_id, line_user_ids, platform='line', inquiry_id=None):
        self.template = template
        self.customer_name = customer_name
        self.order_id = order_id
        self.inquiry_id = inquiry_id
        self.line_user_ids = line_user_ids  # [{'line_user_id': ..., 'display_name': ...}]
        self.platform = platform  # 'line', 'instagram', 'atelier', or 'atelier_inquiry'

        title = template["label"]
        if template.get("status_action"):
            try:
                sl = STATUS_CONFIG[CustomerStatus(template["status_action"])]["label"]
                title += f" → {sl}"
            except ValueError:
                pass
        super().__init__(title=title[:45])

        prefilled = template["text"].replace("{name}", customer_name)
        self.message_input = discord.ui.TextInput(
            label="メッセージ内容（編集可能）",
            style=discord.TextStyle.long,
            default=prefilled,
            max_length=2000,
            required=True,
        )
        self.add_item(self.message_input)

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)

        message_text = self.message_input.value
        results = []

        # 1. メッセージ送信（プラットフォーム別）
        all_success = True
        sent_names = []
        platform_labels = {'line': 'LINE', 'instagram': 'Instagram', 'atelier': 'アトリエ', 'atelier_inquiry': 'お問い合わせ'}
        platform_label = platform_labels.get(self.platform, self.platform)

        if self.platform == 'atelier_inquiry':
            # 問い合わせ: inquiry webhook で送信
            webhook_url = get_atelier_inquiry_webhook_url()
            secret = get_atelier_webhook_secret()
            if webhook_url and secret and self.inquiry_id:
                try:
                    resp = requests.post(webhook_url, json={
                        "inquiry_id": int(self.inquiry_id),
                        "message": message_text,
                        "image_url": "",
                    }, headers={
                        "X-Atelier-Secret": secret,
                        "Content-Type": "application/json",
                    }, timeout=10)
                    if resp.status_code == 200:
                        all_success = True
                        sent_names.append(self.customer_name or "顧客")
                    else:
                        all_success = False
                        print(f"[Atelier Inquiry Template] Webhook failed: {resp.status_code} {resp.text}")
                except Exception as e:
                    all_success = False
                    print(f"[Atelier Inquiry Template] Webhook error: {e}")
            else:
                all_success = False
        elif self.platform == 'atelier':
            # アトリエ注文: WordPress webhook で送信
            webhook_url = get_atelier_webhook_url()
            secret = get_atelier_webhook_secret()
            if webhook_url and secret and self.order_id:
                try:
                    resp = requests.post(webhook_url, json={
                        "order_id": int(self.order_id),
                        "message": message_text,
                        "image_url": "",
                    }, headers={
                        "X-Atelier-Secret": secret,
                        "Content-Type": "application/json",
                    }, timeout=10)
                    if resp.status_code == 200:
                        all_success = True
                        sent_names.append(self.customer_name or "顧客")
                    else:
                        all_success = False
                        print(f"[Atelier Template] Webhook failed: {resp.status_code} {resp.text}")
                except Exception as e:
                    all_success = False
                    print(f"[Atelier Template] Webhook error: {e}")
            else:
                all_success = False
        else:
            for user in self.line_user_ids:
                if self.platform == 'instagram':
                    success = send_instagram_message(user['line_user_id'], message_text)
                else:
                    success = send_line_message(user['line_user_id'], [
                        {"type": "text", "text": message_text}
                    ])
                if success:
                    sent_names.append(user['display_name'])
                else:
                    all_success = False

        if not sent_names and not (self.platform == 'atelier' and all_success):
            await interaction.followup.send(f"❌ {platform_label}送信に失敗しました", ephemeral=True)
            return

        if len(self.line_user_ids) > 1:
            results.append(f"✅ {platform_label}送信完了（{', '.join(sent_names)}）")
        else:
            results.append(f"✅ {platform_label}送信完了")

        # 2. WooCommerceステータス更新（一度だけ）
        status_action = self.template.get("status_action")
        if status_action and self.order_id:
            wc_url = get_wc_url()
            wc_key = get_wc_consumer_key()
            wc_secret = get_wc_consumer_secret()

            if all([wc_url, wc_key, wc_secret]):
                try:
                    url = f"{wc_url}/wp-json/wc/v3/orders/{self.order_id}"
                    resp = requests.put(url, auth=(wc_key, wc_secret), json={"status": status_action})
                    if resp.status_code == 200:
                        results.append(f"✅ WooCommerce → {status_action}")
                    else:
                        results.append(f"⚠️ WooCommerce更新失敗 ({resp.status_code})")
                except Exception as e:
                    results.append(f"⚠️ WooCommerceエラー: {e}")

        # 3. customer_managerステータス更新（全ユーザー）
        if status_action:
            try:
                new_status = CustomerStatus(status_action)
                for user in self.line_user_ids:
                    update_customer_status(user['line_user_id'], new_status, self.order_id)
                results.append("✅ 顧客ステータス更新")
            except ValueError:
                pass

        # 4. フォーラムスレッドの名前更新（絵文字変更）
        if status_action:
            try:
                new_status = CustomerStatus(status_action)
                config = STATUS_CONFIG[new_status]
                thread = interaction.channel
                new_name = re.sub(
                    r'^[\U0001F7E0\U0001F7E1\U0001F535\U0001F7E2\u2705\U0001F4E6\U0001F389\U0001F490\U0001F64F]\s*',
                    f"{config['emoji']} ",
                    thread.name
                )
                if new_name != thread.name:
                    await thread.edit(name=new_name)
                    results.append("✅ スレッド名更新")
            except Exception as e:
                print(f"[WARN] Thread name update failed: {e}")

        # 5. フォーラムタグ更新
        if status_action:
            try:
                thread = interaction.channel
                forum = thread.parent
                if forum and isinstance(forum, discord.ForumChannel):
                    new_status = CustomerStatus(status_action)
                    config = STATUS_CONFIG[new_status]
                    target_tag = None
                    for tag in forum.available_tags:
                        if config['label'] in tag.name or config['emoji'] in (getattr(tag, 'emoji', None) or ''):
                            target_tag = tag
                            break

                    if target_tag:
                        await thread.edit(applied_tags=[target_tag])
                        results.append(f"✅ タグ更新: {target_tag.name}")
            except Exception as e:
                print(f"[WARN] Tag update failed: {e}")

        # 6. アトリエフォーラムスレッド連動更新
        if status_action and self.order_id:
            try:
                new_status = CustomerStatus(status_action)
                await update_atelier_thread_status(self.order_id, new_status)
                results.append("✅ アトリエスレッド更新")
            except Exception as e:
                print(f"[WARN] Atelier thread update failed: {e}")

        # 7. スレッドに送信記録を投稿
        from datetime import datetime
        thread = interaction.channel
        sent_embed = discord.Embed(
            description=message_text,
            color=0x06C755
        )
        sent_embed.set_author(name=f"📤 {self.template['label']}")
        footer_platforms = {'line': 'LINE送信済み', 'instagram': 'Instagram送信済み', 'atelier': 'アトリエ送信済み', 'atelier_inquiry': 'お問い合わせ送信済み'}
        footer_platform = footer_platforms.get(self.platform, f'{self.platform}送信済み')
        if len(self.line_user_ids) > 1:
            names = ", ".join(u['display_name'] for u in self.line_user_ids)
            sent_embed.set_footer(text=f"{footer_platform} ({names}) • {datetime.now().strftime('%m/%d %H:%M')}")
        else:
            sent_embed.set_footer(text=f"{footer_platform} • {datetime.now().strftime('%m/%d %H:%M')}")
        await thread.send(embed=sent_embed)

        # 8. 顧客一覧を更新
        await update_overview_channel()

        # 9. テンプレートボタンを再投稿（常に下部に表示）
        await post_template_buttons(thread)

        # 結果報告
        await interaction.followup.send("\n".join(results), ephemeral=True)


class TemplatePersistentView(discord.ui.View):
    """テンプレートボタン常設ビュー（Bot再起動後も動作）"""
    def __init__(self):
        super().__init__(timeout=None)

    async def _handle_button(self, interaction: discord.Interaction, template_id: str):
        """ボタン押下時の共通処理（複数ユーザー対応 / Instagram対応）"""
        templates = load_templates()
        template = next((t for t in templates if t["id"] == template_id), None)
        if not template:
            await interaction.response.send_message("テンプレートが見つかりません", ephemeral=True)
            return

        thread = interaction.channel
        if not isinstance(thread, discord.Thread):
            await interaction.response.send_message("スレッド内で使用してください", ephemeral=True)
            return

        # プラットフォーム判定
        is_atelier = get_forum_atelier() and str(thread.parent_id) == str(get_forum_atelier())

        if is_atelier:
            # アトリエスレッド（問い合わせ or 注文）
            is_inquiry_thread = thread.name.startswith('💬')
            if is_inquiry_thread:
                platform = 'atelier_inquiry'
                customer_name, _ = get_thread_customer_info(thread)
                inq_match = re.search(r'#(\d+)', thread.name)
                inquiry_id = inq_match.group(1) if inq_match else None
            else:
                platform = 'atelier'
                inquiry_id = None
            customer_name, order_id = get_thread_customer_info(thread)
            all_users = [{'line_user_id': '', 'display_name': customer_name}]
        elif get_platform_from_thread(thread.id) == 'instagram':
            # Instagram スレッド
            platform = 'instagram'
            ig_user_id = get_instagram_user_id_from_thread(thread.id)
            if not ig_user_id:
                await interaction.response.send_message("❌ Instagram User IDが見つかりません", ephemeral=True)
                return
            ig_map = load_instagram_thread_map()
            ig_data = ig_map.get(ig_user_id, {})
            customer_name, order_id = get_thread_customer_info(thread)
            all_users = [{'line_user_id': ig_user_id, 'display_name': ig_data.get('display_name', customer_name)}]
        else:
            # LINE スレッド（従来ロジック）
            platform = 'line'
            all_users = get_all_line_users_from_thread(thread.id)
            if not all_users:
                line_user_id = await find_line_user_id_in_thread(thread)
                if not line_user_id:
                    await interaction.response.send_message("❌ LINE User IDが見つかりません", ephemeral=True)
                    return
                customer_name, order_id = get_thread_customer_info(thread)
                all_users = [{'line_user_id': line_user_id, 'display_name': customer_name}]

        # 顧客情報取得
        customer_name, order_id = get_thread_customer_info(thread)
        if not order_id and platform == 'line':
            customer = get_customer(all_users[0]['line_user_id'])
            if customer and customer.get("orders"):
                order_id = str(customer["orders"][-1].get("order_id", ""))

        # モーダル表示（プラットフォーム情報付き）
        inq_id = locals().get('inquiry_id')
        modal = TemplateEditModal(template, customer_name, order_id, all_users, platform=platform, inquiry_id=inq_id)
        await interaction.response.send_modal(modal)

    @discord.ui.button(label="① あいさつ", style=discord.ButtonStyle.secondary, custom_id="tpl_greeting", emoji="👋", row=0)
    async def btn_greeting(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._handle_button(interaction, "greeting")

    @discord.ui.button(label="② デザイン確認", style=discord.ButtonStyle.secondary, custom_id="tpl_design_check", emoji="🎨", row=0)
    async def btn_design_check(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._handle_button(interaction, "design_check")

    @discord.ui.button(label="③ 確定", style=discord.ButtonStyle.primary, custom_id="tpl_design_confirmed", emoji="✅", row=0)
    async def btn_design_confirmed(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._handle_button(interaction, "design_confirmed")

    @discord.ui.button(label="④ 制作完了", style=discord.ButtonStyle.primary, custom_id="tpl_production_done", emoji="🎉", row=0)
    async def btn_production_done(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._handle_button(interaction, "production_done")

    @discord.ui.button(label="⑤ 発送完了", style=discord.ButtonStyle.success, custom_id="tpl_shipped", emoji="📦", row=1)
    async def btn_shipped(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._handle_button(interaction, "shipped")

    @discord.ui.button(label="⑥ お礼①", style=discord.ButtonStyle.secondary, custom_id="tpl_thanks_1", emoji="🙏", row=1)
    async def btn_thanks_1(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._handle_button(interaction, "thanks_1")

    @discord.ui.button(label="⑦ お礼②", style=discord.ButtonStyle.secondary, custom_id="tpl_thanks_2", emoji="💐", row=1)
    async def btn_thanks_2(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._handle_button(interaction, "thanks_2")

    @discord.ui.button(label="テンプレ編集", style=discord.ButtonStyle.secondary, custom_id="tpl_manage", emoji="✏️", row=2)
    async def btn_manage(self, interaction: discord.Interaction, button: discord.ui.Button):
        """テンプレート管理メニュー"""
        templates = load_templates()
        options = []
        for t in templates:
            options.append(discord.SelectOption(
                label=f"{t['emoji']} {t['label']}",
                value=t["id"],
                description="編集"
            ))
        options.append(discord.SelectOption(
            label="＋ 新規テンプレート追加",
            value="__new__",
            emoji="➕"
        ))

        view = discord.ui.View(timeout=120)
        select = TemplateManageSelect(options)
        view.add_item(select)
        await interaction.response.send_message("編集するテンプレートを選択:", view=view, ephemeral=True)


class TemplateManageSelect(discord.ui.Select):
    """テンプレート管理用セレクトメニュー"""
    def __init__(self, options):
        super().__init__(placeholder="テンプレートを選択...", options=options)

    async def callback(self, interaction: discord.Interaction):
        selected = self.values[0]

        if selected == "__new__":
            modal = TemplateManageModal(template_id=None, label="", text="", is_new=True)
            await interaction.response.send_modal(modal)
        else:
            templates = load_templates()
            template = next((t for t in templates if t["id"] == selected), None)
            if not template:
                await interaction.response.send_message("テンプレートが見つかりません", ephemeral=True)
                return
            modal = TemplateManageModal(
                template_id=template["id"],
                label=template["label"],
                text=template["text"],
                is_new=False
            )
            await interaction.response.send_modal(modal)


class TemplateManageModal(discord.ui.Modal):
    """テンプレート編集・追加モーダル"""
    def __init__(self, template_id, label, text, is_new=False):
        self.template_id = template_id
        self.is_new = is_new
        super().__init__(title="テンプレート追加" if is_new else "テンプレート編集")

        self.label_input = discord.ui.TextInput(
            label="テンプレート名",
            style=discord.TextStyle.short,
            default=label,
            placeholder="例: ① 初回あいさつ",
            max_length=50,
            required=True,
        )
        self.add_item(self.label_input)

        self.text_input = discord.ui.TextInput(
            label="メッセージ本文（{name}で顧客名に置換）",
            style=discord.TextStyle.long,
            default=text,
            placeholder="{name}様\n\nメッセージ内容...",
            max_length=2000,
            required=True,
        )
        self.add_item(self.text_input)

    async def on_submit(self, interaction: discord.Interaction):
        templates = load_templates()

        if self.is_new:
            new_id = f"custom_{len(templates) + 1}"
            templates.append({
                "id": new_id,
                "label": self.label_input.value,
                "emoji": "💬",
                "status_action": None,
                "text": self.text_input.value,
            })
            save_templates(templates)
            await interaction.response.send_message(
                f"✅ テンプレート「{self.label_input.value}」を追加しました",
                ephemeral=True
            )
        else:
            for t in templates:
                if t["id"] == self.template_id:
                    t["label"] = self.label_input.value
                    t["text"] = self.text_input.value
                    break
            save_templates(templates)
            await interaction.response.send_message(
                f"✅ テンプレート「{self.label_input.value}」を更新しました",
                ephemeral=True
            )


async def post_template_buttons(thread):
    """テンプレートボタンをスレッドに投稿（前回のを削除して常に最下部に表示）"""
    thread_key = str(thread.id)

    # ループ防止
    if thread_key in _posting_buttons_lock:
        return
    _posting_buttons_lock.add(thread_key)

    try:
        # 前回のボタンメッセージを削除
        old_msg_id = _template_button_msg_ids.get(thread_key)
        if old_msg_id:
            try:
                old_msg = await thread.fetch_message(old_msg_id)
                await old_msg.delete()
            except Exception:
                pass

        view = TemplatePersistentView()
        msg = await thread.send(view=view)
        _template_button_msg_ids[thread_key] = msg.id
    finally:
        _posting_buttons_lock.discard(thread_key)


@bot.tree.command(name="template", description="LINEテンプレートボタンを表示")
async def send_template(interaction: discord.Interaction):
    """テンプレートボタン投稿コマンド"""
    channel = interaction.channel

    if not isinstance(channel, discord.Thread):
        await interaction.response.send_message(
            "このコマンドは #LINE対応 または #atelier フォーラムのスレッド内で使用してください",
            ephemeral=True
        )
        return

    parent_id = str(channel.parent_id)
    allowed_forums = [str(get_forum_line())]
    if get_forum_atelier():
        allowed_forums.append(str(get_forum_atelier()))
    if parent_id not in allowed_forums:
        await interaction.response.send_message(
            "このコマンドは #LINE対応 または #atelier フォーラムのスレッド内で使用してください",
            ephemeral=True
        )
        return

    await interaction.response.defer(ephemeral=True)
    await post_template_buttons(channel)
    await interaction.followup.send("✅ テンプレートボタンを表示しました", ephemeral=True)


# ================== API Endpoints ==================

@api.route("/api/status", methods=["POST"])
def api_update_status():
    """ステータス更新API"""
    data = request.json

    order_id = data.get("order_id")
    new_status = data.get("status")

    if not order_id or not new_status:
        return jsonify({"error": "order_id and status required"}), 400

    try:
        status = CustomerStatus(new_status)
    except ValueError:
        return jsonify({"error": "Invalid status"}), 400

    line_user_id, customer = get_customer_by_order(order_id)

    if not line_user_id:
        return jsonify({"error": "Customer not found for order"}), 404

    update_customer_status(line_user_id, status, order_id)

    asyncio.run_coroutine_threadsafe(update_overview_channel(), bot.loop)

    if status == CustomerStatus.SHIPPED and get_category_shipped():
        channel_id = customer.get("discord_channel_id")
        if channel_id:
            async def move_channel():
                guild = bot.get_guild(int(get_guild_id()))
                if guild:
                    channel = guild.get_channel(int(channel_id))
                    if channel:
                        await move_channel_to_category(channel, get_category_shipped())
            asyncio.run_coroutine_threadsafe(move_channel(), bot.loop)

    return jsonify({"success": True, "status": new_status})


@api.route("/api/customer", methods=["POST"])
def api_add_customer():
    """顧客追加API"""
    data = request.json

    line_user_id = data.get("line_user_id")
    display_name = data.get("display_name")
    discord_channel_id = data.get("discord_channel_id")
    order_id = data.get("order_id")
    order_info = data.get("order_info", {})

    if not line_user_id:
        return jsonify({"error": "line_user_id required"}), 400

    customer = add_customer(line_user_id, display_name, discord_channel_id, order_id, order_info)

    asyncio.run_coroutine_threadsafe(update_overview_channel(), bot.loop)

    return jsonify({"success": True, "customer": customer})


@api.route("/api/customer/delete", methods=["POST"])
def api_delete_customer():
    """顧客削除API"""
    data = request.json
    customer_key = data.get("customer_key")
    if not customer_key:
        return jsonify({"error": "customer_key required"}), 400

    customers = load_customers()
    if customer_key in customers:
        deleted = customers.pop(customer_key)
        save_customers(customers)
        asyncio.run_coroutine_threadsafe(update_overview_channel(), bot.loop)
        return jsonify({"success": True, "deleted": deleted.get("display_name", "")})
    return jsonify({"error": "customer not found"}), 404


@api.route("/api/overview", methods=["GET"])
def api_get_overview():
    """一覧取得API"""
    return jsonify(get_status_summary())


@api.route("/health", methods=["GET"])
def health_check():
    """ヘルスチェック（Railway用）"""
    return jsonify({"status": "ok", "canva_enabled": CANVA_ENABLED})


def verify_woo_webhook_signature(payload, signature, secret):
    """WooCommerce Webhook署名をHMAC-SHA256で検証"""
    if not secret or not signature:
        return False
    expected = base64.b64encode(
        hmac.new(secret.encode('utf-8'), payload, hashlib.sha256).digest()
    ).decode('utf-8')
    return hmac.compare_digest(expected, signature)


@api.route("/api/woo-webhook", methods=["GET", "POST"])
def woo_webhook():
    """WooCommerce Webhook受信 → Canva自動化"""
    # GETリクエスト = WooCommerceのPingテスト
    if request.method == "GET":
        return jsonify({"status": "ok", "message": "Webhook endpoint ready"})

    # Webhook検証（オプション）
    webhook_source = request.headers.get("X-WC-Webhook-Source", "")
    webhook_topic = request.headers.get("X-WC-Webhook-Topic", "")

    # Pingテスト検出（トピックがない、またはボディが空/webhook_idのみ）
    raw_payload = request.get_data()
    data = request.get_json(force=True, silent=True) or {}
    if not data or data.get("webhook_id") and not data.get("id"):
        print(f"[Webhook] Ping test received from {webhook_source}")
        return jsonify({"status": "ok", "message": "Webhook ping successful"})

    if not CANVA_ENABLED:
        return jsonify({"error": "Canva handler not available"}), 503

    # Webhook署名検証（設定されている場合）
    webhook_secret = get_woo_webhook_secret()
    if webhook_secret:
        signature = request.headers.get("X-WC-Webhook-Signature", "")
        if not verify_woo_webhook_signature(raw_payload, signature, webhook_secret):
            print(f"[Webhook] Invalid signature from {webhook_source}")
            return jsonify({"error": "Invalid signature"}), 401

    order_id = data.get("id")
    if not order_id:
        return jsonify({"error": "No order_id"}), 400

    # 注文ステータスをチェック（支払い完了後のみ処理）
    order_status = data.get("status", "")
    print(f"[Webhook] Received order #{order_id} (status: {order_status}) from {webhook_source}")

    # 顧客一覧に追加（入金確認済みのみ）
    if order_status in ("pending", "failed", "cancelled"):
        print(f"[Webhook] Skipping customer add: status={order_status}")
        # pending等はCanva処理もスキップ
        return jsonify({"status": "skipped", "reason": f"order status: {order_status}"})

    try:
        billing = data.get("billing", {})
        customer_name = f"{billing.get('last_name', '')} {billing.get('first_name', '')}".strip()
        email = billing.get("email", "")
        order_info = {
            "total": data.get("total", ""),
            "status": order_status,
            "product": data.get("line_items", [{}])[0].get("name", "") if data.get("line_items") else "",
        }
        add_order_customer(order_id, customer_name, email, order_info)
        print(f"[Webhook] Customer added/updated: {customer_name} ({email})")
        # 顧客一覧を更新
        if bot.loop:
            asyncio.run_coroutine_threadsafe(update_overview_channel(), bot.loop)
    except Exception as e:
        print(f"[WARN] Failed to add customer: {e}")

    # processing（入金確認後）のみ処理 ※ 2026-01-31: designing → processing に変更
    if order_status != "processing":
        print(f"[Webhook] Skipping order #{order_id} - status '{order_status}' not ready for Canva")
        return jsonify({"status": "skipped", "reason": f"Order status '{order_status}' not ready"})

    # 必要な設定が揃っているか確認
    if not all([get_canva_access_token(), get_canva_refresh_token(), get_wc_url(), get_wc_consumer_key(), get_wc_consumer_secret()]):
        print("[ERROR] Missing Canva or WooCommerce configuration")
        return jsonify({"error": "Missing configuration"}), 500

    # 非同期でCanva処理を実行（Webhookレスポンスを待たせない）
    def process_async():
        try:
            config = {
                'wc_url': get_wc_url(),
                'wc_key': get_wc_consumer_key(),
                'wc_secret': get_wc_consumer_secret(),
                'canva_access_token': get_canva_access_token(),
                'canva_refresh_token': get_canva_refresh_token(),
                'discord_webhook': get_canva_webhook_url(),
                'discord_bot_token': get_discord_token(),
            }
            canva_process_order(order_id, config)
        except Exception as e:
            print(f"[ERROR] Canva processing failed: {e}")
            import traceback
            traceback.print_exc()

    thread = threading.Thread(target=process_async, daemon=True)
    thread.start()

    return jsonify({"success": True, "message": f"Processing order #{order_id}"})


@api.route("/api/canva/process", methods=["POST"])
def api_canva_process():
    """手動でCanva処理をトリガー（デバッグ用）"""
    if not CANVA_ENABLED:
        return jsonify({"error": "Canva handler not available"}), 503

    data = request.json
    order_id = data.get("order_id")

    if not order_id:
        return jsonify({"error": "order_id required"}), 400

    if not all([get_canva_access_token(), get_canva_refresh_token(), get_wc_url(), get_wc_consumer_key(), get_wc_consumer_secret()]):
        return jsonify({"error": "Missing configuration"}), 500

    # 同期で処理
    try:
        config = {
            'wc_url': get_wc_url(),
            'wc_key': get_wc_consumer_key(),
            'wc_secret': get_wc_consumer_secret(),
            'canva_access_token': get_canva_access_token(),
            'canva_refresh_token': get_canva_refresh_token(),
            'discord_webhook': get_canva_webhook_url(),
            'discord_bot_token': get_discord_token(),
        }
        result = canva_process_order(order_id, config)
        return jsonify({"success": result, "order_id": order_id})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@api.route("/api/canva/debug-process", methods=["POST"])
def api_canva_debug_process():
    """詳細デバッグ付きCanva処理"""
    from canva_handler import get_order_from_woocommerce, parse_order_data

    data = request.json
    order_id = data.get("order_id")
    debug = {"order_id": order_id, "steps": []}

    # Step 1: Config check
    config = {
        'wc_url': get_wc_url(),
        'wc_key': get_wc_consumer_key(),
        'wc_secret': get_wc_consumer_secret(),
    }
    debug["steps"].append({"step": "config", "wc_url": config['wc_url'], "wc_key_set": bool(config['wc_key']), "wc_secret_set": bool(config['wc_secret'])})

    # Step 2: Get order
    order = get_order_from_woocommerce(order_id, config['wc_url'], config['wc_key'], config['wc_secret'])
    if not order:
        debug["steps"].append({"step": "get_order", "success": False, "error": "Order not found"})
        return jsonify(debug)
    debug["steps"].append({"step": "get_order", "success": True, "order_status": order.get('status')})

    # Step 3: Check if already processed
    meta = {m['key']: m['value'] for m in order.get('meta_data', [])}
    if meta.get('canva_automation_done'):
        debug["steps"].append({"step": "check_processed", "already_done": True})
        return jsonify(debug)
    debug["steps"].append({"step": "check_processed", "already_done": False})

    # Step 4: Parse order data
    order_data = parse_order_data(order)
    debug["steps"].append({
        "step": "parse_order",
        "product_name": order_data.get('product_name'),
        "board_name": order_data.get('board_name'),
        "board_number": order_data.get('board_number'),
        "groom": order_data.get('sim_data', {}).get('groomName'),
        "bride": order_data.get('sim_data', {}).get('brideName'),
    })

    if not order_data['board_name']:
        debug["steps"].append({"step": "board_check", "success": False, "error": "No board name"})
        return jsonify(debug)

    debug["steps"].append({"step": "board_check", "success": True})
    debug["ready_for_canva"] = True

    return jsonify(debug)


@api.route("/api/canva/debug-import", methods=["POST"])
def api_canva_debug_import():
    """Canvaインポートを直接テスト"""
    from canva_handler import (
        get_order_from_woocommerce, parse_order_data, create_pptx,
        import_to_canva, refresh_canva_token
    )
    import tempfile

    data = request.json
    order_id = data.get("order_id")
    debug = {"order_id": order_id, "steps": []}

    config = {
        'wc_url': get_wc_url(),
        'wc_key': get_wc_consumer_key(),
        'wc_secret': get_wc_consumer_secret(),
    }

    # Get order
    order = get_order_from_woocommerce(order_id, config['wc_url'], config['wc_key'], config['wc_secret'])
    if not order:
        debug["error"] = "Order not found"
        return jsonify(debug)

    order_data = parse_order_data(order)
    debug["steps"].append({"step": "parse", "board": order_data.get('board_name')})

    # Create PowerPoint
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            pptx_path = create_pptx(order_data, temp_dir)
            debug["steps"].append({"step": "pptx_created", "path": pptx_path})

            # Check file size
            import os as os_module
            file_size = os_module.path.getsize(pptx_path)
            debug["steps"].append({"step": "pptx_size", "bytes": file_size})

            # Get fresh token
            access_token = get_canva_access_token()
            refresh_token = get_canva_refresh_token()

            # Try import
            canva_title = f"Test_{order_id}"
            design, error_info = import_to_canva(pptx_path, canva_title, access_token, refresh_token)

            if design:
                debug["steps"].append({"step": "import_success", "design_id": design.get('id')})
                debug["success"] = True
            else:
                debug["steps"].append({"step": "import_failed", "error": error_info})
                debug["success"] = False

    except Exception as e:
        import traceback
        debug["error"] = str(e)
        debug["traceback"] = traceback.format_exc()

    return jsonify(debug)


@api.route("/api/canva/debug-token", methods=["GET"])
def api_canva_debug_token():
    """Canvaトークンリフレッシュをテスト（デバッグ用）"""
    import base64
    import requests as req

    # 環境変数の状態
    client_id = os.environ.get("CANVA_CLIENT_ID", "OC-AZvUVtxGhbOD")
    client_secret = os.environ.get("CANVA_CLIENT_SECRET", "")
    refresh_token = get_canva_refresh_token()
    access_token = get_canva_access_token()

    debug_info = {
        "client_id": client_id,
        "client_secret_set": bool(client_secret),
        "client_secret_len": len(client_secret) if client_secret else 0,
        "client_secret_preview": client_secret[:10] + "..." if client_secret and len(client_secret) > 10 else client_secret,
        "refresh_token_set": bool(refresh_token),
        "refresh_token_len": len(refresh_token) if refresh_token else 0,
        "access_token_set": bool(access_token),
        "access_token_len": len(access_token) if access_token else 0,
    }

    # トークンリフレッシュを試行
    if client_id and client_secret and refresh_token:
        url = 'https://api.canva.com/rest/v1/oauth/token'
        credentials = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()

        try:
            response = req.post(url, data={
                'grant_type': 'refresh_token',
                'refresh_token': refresh_token,
            }, headers={
                'Authorization': f'Basic {credentials}',
                'Content-Type': 'application/x-www-form-urlencoded',
            })

            debug_info["refresh_status_code"] = response.status_code
            debug_info["refresh_response"] = response.text[:500] if response.text else ""

            if response.status_code == 200:
                debug_info["refresh_success"] = True
            else:
                debug_info["refresh_success"] = False
        except Exception as e:
            debug_info["refresh_error"] = str(e)
    else:
        debug_info["refresh_skipped"] = "Missing required tokens"

    return jsonify(debug_info)


@api.route("/api/canva/update-tokens", methods=["POST"])
def api_canva_update_tokens():
    """Canvaトークンを手動で更新（ファイル永続化）"""
    data = request.json
    access_token = data.get("access_token")
    refresh_token = data.get("refresh_token")

    updated = {}
    if access_token:
        os.environ['CANVA_ACCESS_TOKEN'] = access_token
        updated['access_token'] = f"Set ({len(access_token)} chars)"
    if refresh_token:
        os.environ['CANVA_REFRESH_TOKEN'] = refresh_token
        updated['refresh_token'] = f"Set ({len(refresh_token)} chars)"

    # ファイルにも保存（再起動後も維持）
    if access_token and refresh_token:
        try:
            from canva_handler import save_tokens_to_file
            save_tokens_to_file(access_token, refresh_token)
            updated['file_saved'] = True
        except Exception as e:
            updated['file_save_error'] = str(e)

    return jsonify({"success": True, "updated": updated})


@api.route("/api/canva/current-tokens", methods=["GET"])
def api_canva_current_tokens():
    """現在のCanvaトークン情報を取得（先頭50文字のみ）"""
    access = get_canva_access_token()
    refresh = get_canva_refresh_token()
    return jsonify({
        "access_token_preview": access[:50] + "..." if access and len(access) > 50 else access,
        "access_token_len": len(access) if access else 0,
        "refresh_token_preview": refresh[:50] + "..." if refresh and len(refresh) > 50 else refresh,
        "refresh_token_len": len(refresh) if refresh else 0,
    })


# OAuth認証用の状態保持
_oauth_state = {}

@api.route("/api/canva/oauth/start", methods=["GET"])
def api_canva_oauth_start():
    """Canva OAuth認証URLを生成"""
    import secrets
    import hashlib

    client_id = os.environ.get("CANVA_CLIENT_ID", "OC-AZvUVtxGhbOD")
    redirect_uri = "https://i-tategu-shop.com/canva-callback/"
    scopes = "design:content:read design:content:write design:meta:read design:permission:read design:permission:write asset:read asset:write folder:read folder:write"

    # PKCE生成
    code_verifier = secrets.token_urlsafe(64)[:128]
    code_challenge = base64.urlsafe_b64encode(
        hashlib.sha256(code_verifier.encode()).digest()
    ).decode().rstrip('=')

    # 状態保存（5分間有効）
    import time
    _oauth_state['code_verifier'] = code_verifier
    _oauth_state['expires'] = time.time() + 300

    params = f"response_type=code&client_id={client_id}&redirect_uri={redirect_uri}&scope={scopes}&code_challenge={code_challenge}&code_challenge_method=S256"
    auth_url = f"https://www.canva.com/api/oauth/authorize?{params}"

    return jsonify({
        "auth_url": auth_url,
        "instructions": "このURLをブラウザで開き、認証後にリダイレクトされたURLの ?code=XXX をコピーして /api/canva/oauth/callback に送信"
    })


@api.route("/api/canva/oauth/callback", methods=["POST"])
def api_canva_oauth_callback():
    """OAuth認証コードをトークンに交換"""
    import time

    data = request.json
    code = data.get("code")

    if not code:
        return jsonify({"error": "code required"}), 400

    # 状態確認
    if not _oauth_state.get('code_verifier') or time.time() > _oauth_state.get('expires', 0):
        return jsonify({"error": "OAuth session expired. Call /api/canva/oauth/start first"}), 400

    code_verifier = _oauth_state['code_verifier']
    client_id = os.environ.get("CANVA_CLIENT_ID", "OC-AZvUVtxGhbOD")
    client_secret = os.environ.get("CANVA_CLIENT_SECRET", "")
    redirect_uri = "https://i-tategu-shop.com/canva-callback/"

    # トークン交換
    auth = base64.b64encode(f'{client_id}:{client_secret}'.encode()).decode()
    response = requests.post(
        'https://api.canva.com/rest/v1/oauth/token',
        headers={
            'Authorization': f'Basic {auth}',
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        data={
            'grant_type': 'authorization_code',
            'code': code,
            'code_verifier': code_verifier,
            'redirect_uri': redirect_uri
        }
    )

    if response.status_code != 200:
        return jsonify({"error": "Token exchange failed", "details": response.text}), 400

    tokens = response.json()
    access_token = tokens.get('access_token')
    refresh_token = tokens.get('refresh_token')

    if access_token and refresh_token:
        # 環境変数更新
        os.environ['CANVA_ACCESS_TOKEN'] = access_token
        os.environ['CANVA_REFRESH_TOKEN'] = refresh_token

        # ファイル保存
        try:
            from canva_handler import save_tokens_to_file
            save_tokens_to_file(access_token, refresh_token)
        except Exception as e:
            print(f"[WARN] Failed to save tokens to file: {e}")

        # 状態クリア
        _oauth_state.clear()

        return jsonify({
            "success": True,
            "access_token_preview": access_token[:50] + "...",
            "refresh_token_preview": refresh_token[:50] + "...",
            "expires_in": tokens.get('expires_in')
        })

    return jsonify({"error": "No tokens in response", "response": tokens}), 400


def run_api():
    """API サーバー起動"""
    port = int(os.getenv("PORT", 5701))
    api.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)


# ================== Main ==================

if __name__ == "__main__":
    if not get_discord_token():
        print("[ERROR] get_discord_token() not set")
        exit(1)

    api_thread = threading.Thread(target=run_api, daemon=True)
    api_thread.start()
    print(f"[OK] API Server started")

    print("[...] Starting Discord Bot...")
    bot.run(get_discord_token())
