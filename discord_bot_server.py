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
from flask import Flask, request, jsonify
from dotenv import load_dotenv
import discord
from discord.ext import commands
from discord import app_commands

from customer_manager import (
    CustomerStatus, STATUS_CONFIG,
    add_customer, update_customer_status, get_customer,
    get_customer_by_channel, get_customer_by_order,
    get_status_summary, get_all_customers_grouped, load_customers
)

# Canva自動化ハンドラー
try:
    from canva_handler import process_order as canva_process_order
    CANVA_ENABLED = True
except ImportError as e:
    CANVA_ENABLED = False
    print(f"[WARN] Canva handler not available: {e}")

load_dotenv()

# 環境変数（全て遅延読み込み - Railway Railpack対策）
# os.environ.get() を使用（os.getenv検出を回避）
def get_line_token():
    return os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

def get_discord_token():
    return os.environ.get("get_discord_token()")

def get_guild_id():
    return os.environ.get("get_guild_id()")

def get_category_active():
    return os.environ.get("DISCORD_get_category_active()")

def get_category_shipped():
    return os.environ.get("DISCORD_get_category_shipped()")

def get_overview_channel():
    return os.environ.get("DISCORD_OVERVIEW_CHANNEL")

def get_forum_completed():
    return os.environ.get("DISCORD_FORUM_COMPLETED")

def get_forum_line():
    return os.environ.get("DISCORD_FORUM_LINE", "1463460598493745225")

def get_canva_access_token():
    return os.environ.get("CANVA_ACCESS_TOKEN")

def get_canva_refresh_token():
    return os.environ.get("CANVA_REFRESH_TOKEN")

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


@api.route("/api/woo-webhook", methods=["POST"])
def woo_webhook():
    """WooCommerce Webhook受信 → Canva自動化"""
    if not CANVA_ENABLED:
        return jsonify({"error": "Canva handler not available"}), 503

    # Webhook検証（オプション）
    webhook_source = request.headers.get("X-WC-Webhook-Source", "")
    if WOO_WEBHOOK_SECRET:
        signature = request.headers.get("X-WC-Webhook-Signature", "")
        # TODO: HMACで検証（セキュリティ強化用）

    data = request.json
    if not data:
        return jsonify({"error": "No data"}), 400

    order_id = data.get("id")
    if not order_id:
        return jsonify({"error": "No order_id"}), 400

    print(f"[Webhook] Received order #{order_id} from {webhook_source}")

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
        }
        result = canva_process_order(order_id, config)
        return jsonify({"success": result, "order_id": order_id})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


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
