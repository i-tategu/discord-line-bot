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
    get_status_summary, get_all_customers_grouped, load_customers
)

# Canva自動化ハンドラー
try:
    from canva_handler import process_order as canva_process_order, get_current_tokens
    CANVA_ENABLED = True
except ImportError as e:
    CANVA_ENABLED = False
    print(f"[WARN] Canva handler not available: {e}")
    def get_current_tokens():
        return os.environ.get("CANVA_ACCESS_TOKEN"), os.environ.get("CANVA_REFRESH_TOKEN")

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

# スレッドマップファイル
THREAD_MAP_FILE = os.path.join(os.path.dirname(__file__), "thread_map.json")

# Flask API
api = Flask(__name__)

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
        if status == CustomerStatus.SHIPPED:
            continue

        data = summary[status]
        config = STATUS_CONFIG[status]

        embed = discord.Embed(
            title=f"{config['emoji']} {config['label']} ({data['count']}件)",
            color=config['color']
        )

        if data['customers']:
            customer_links = []
            for c in data['customers'][:10]:
                channel_id = c.get('discord_channel_id')
                name = c.get('display_name', '不明')
                if channel_id:
                    customer_links.append(f"• <#{channel_id}> {name}様")
                else:
                    customer_links.append(f"• {name}様")

            embed.description = "\n".join(customer_links)

            if len(data['customers']) > 10:
                embed.description += f"\n... 他 {len(data['customers']) - 10}件"
        else:
            embed.description = "_該当なし_"

        embeds.append(embed)

    shipped_data = summary[CustomerStatus.SHIPPED]
    shipped_config = STATUS_CONFIG[CustomerStatus.SHIPPED]
    shipped_embed = discord.Embed(
        title=f"{shipped_config['emoji']} {shipped_config['label']} ({shipped_data['count']}件)",
        color=shipped_config['color']
    )
    if shipped_data['customers']:
        names = [c.get('display_name', '不明') + "様" for c in shipped_data['customers'][:5]]
        shipped_embed.description = "、".join(names)
        if len(shipped_data['customers']) > 5:
            shipped_embed.description += f" 他{len(shipped_data['customers']) - 5}件"
    else:
        shipped_embed.description = "_該当なし_"
    embeds.append(shipped_embed)

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

    try:
        guild = discord.Object(id=int(get_guild_id()))
        bot.tree.copy_global_to(guild=guild)
        await bot.tree.sync(guild=guild)
        print("[OK] Slash commands synced")
    except Exception as e:
        print(f"[WARN] Failed to sync commands: {e}")

    await update_overview_channel()
    print("[OK] Overview channel updated")
    print("=" * 50)


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
    """Discordメッセージを監視してLINEに転送"""
    print(f"[MSG] channel={message.channel.name if hasattr(message.channel, 'name') else 'DM'}, author={message.author}, bot={message.author.bot}")

    if message.author == bot.user:
        return

    if message.author.bot:
        return

    await bot.process_commands(message)

    line_user_id = None

    # フォーラムスレッドからの転送
    if isinstance(message.channel, discord.Thread):
        print(f"[DEBUG] Thread detected: parent_id={message.channel.parent_id}, get_forum_line()={get_forum_line()}")
        if message.channel.parent_id == int(get_forum_line()):
            line_user_id = get_line_user_id_from_thread(message.channel.id)
            if not line_user_id:
                starter = message.channel.starter_message
                if starter:
                    # バッククォートあり・なし両方に対応
                    match = re.search(r'LINE User ID:\s*`?([A-Za-z0-9]+)`?', starter.content)
                    if match:
                        line_user_id = match.group(1)

                # starter_messageがキャッシュされていない場合、履歴から取得
                if not line_user_id:
                    async for msg in message.channel.history(limit=5, oldest_first=True):
                        # メッセージ本文から検索
                        if msg.content:
                            match = re.search(r'LINE User ID:\s*`?([A-Za-z0-9]+)`?', msg.content)
                            if match:
                                line_user_id = match.group(1)
                                print(f"[DEBUG] Found LINE User ID in content: {line_user_id}")
                                break

                        # Embed（埋め込み）から検索
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

    # 通常チャンネルからの転送（トピックにLINE User IDがあれば転送）
    if not line_user_id and hasattr(message.channel, 'topic'):
        line_user_id = get_line_user_id_from_channel(message.channel)

    if not line_user_id:
        print(f"[DEBUG] No LINE User ID found for channel: {message.channel.name}")
        return

    print(f"[DEBUG] LINE User ID found: {line_user_id}")
    # テキストメッセージ送信
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


# ================== Slash Commands ==================

@bot.tree.command(name="status", description="顧客のステータスを変更")
@app_commands.describe(new_status="新しいステータス")
@app_commands.choices(new_status=[
    app_commands.Choice(name="🟡 購入済み", value="purchased"),
    app_commands.Choice(name="🔵 デザイン確定", value="design"),
    app_commands.Choice(name="🟢 制作完了", value="production"),
    app_commands.Choice(name="✅ 発送済み", value="shipped"),
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

    await update_overview_channel()


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

    # 顧客一覧に追加（全ステータスで追加、重複は自動スキップ）
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

    # designing（デザイン打ち合わせ中 = 支払い確認後）のみ処理
    if order_status != "designing":
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
