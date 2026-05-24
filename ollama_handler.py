"""
Discord Bot ↔ Mac Studio ローカル LLM 連携

- Open WebUI 経由 (Ollama API を /ollama/api/chat でプロキシ) を Tailscale Funnel 越しに叩く
- スラッシュコマンド /ai と Bot メンション両方に対応
- 添付画像があれば vision モデルへ自動切替
- Allowlist で利用可能ユーザーを制限

Required env vars:
  OPENWEBUI_URL                  例: https://macstudio.tail9b6868.ts.net
  OPENWEBUI_API_KEY              Open WebUI で発行した sk-xxxx
  OLLAMA_ALLOWED_USER_IDS        カンマ区切りの Discord ユーザー ID (空なら全員許可)
  OLLAMA_DEFAULT_MODEL           既定モデル (デフォルト: qwen3.6:35b)
  OLLAMA_VISION_MODEL            画像入力時のモデル (デフォルト: qwen3-vl:32b)
  OLLAMA_AUTO_RESPOND_CHANNELS   カンマ区切りのチャンネル ID — メンション不要で自動応答
"""
import os
import base64
import logging
import aiohttp
import discord
from discord import app_commands
from discord.ext import commands

log = logging.getLogger(__name__)

DISCORD_MSG_LIMIT = 2000
REQUEST_TIMEOUT = 300  # 推論時間が長いモデル向け

_OPENWEBUI_URL = os.environ.get("OPENWEBUI_URL", "").rstrip("/")
_API_KEY = os.environ.get("OPENWEBUI_API_KEY", "")
_DEFAULT_MODEL = os.environ.get("OLLAMA_DEFAULT_MODEL", "qwen3.6:35b")
_VISION_MODEL = os.environ.get("OLLAMA_VISION_MODEL", "qwen3-vl:32b")

_ALLOWED_USER_IDS: set[int] = {
    int(x.strip())
    for x in os.environ.get("OLLAMA_ALLOWED_USER_IDS", "").split(",")
    if x.strip().isdigit()
}

# このチャンネル内では @メンション無しの全メッセージに自動応答する
_AUTO_RESPOND_CHANNELS: set[int] = {
    int(x.strip())
    for x in os.environ.get("OLLAMA_AUTO_RESPOND_CHANNELS", "").split(",")
    if x.strip().isdigit()
}


def _is_allowed(user_id: int) -> bool:
    return not _ALLOWED_USER_IDS or user_id in _ALLOWED_USER_IDS


async def _fetch_image_b64(url: str) -> str | None:
    """Ollama API は base64 文字列のみ (data: プレフィクス不要) を期待。"""
    try:
        async with aiohttp.ClientSession() as s:
            async with s.get(url, timeout=aiohttp.ClientTimeout(total=30)) as r:
                if r.status != 200:
                    return None
                data = await r.read()
        return base64.b64encode(data).decode()
    except Exception as e:
        log.warning(f"[ollama] image fetch failed: {e}")
        return None


async def _call_chat(messages: list[dict], model: str) -> str:
    """Open WebUI 経由で Ollama の /api/chat を叩く。"""
    payload = {"model": model, "messages": messages, "stream": False}
    headers = {
        "Authorization": f"Bearer {_API_KEY}",
        "Content-Type": "application/json",
    }
    url = f"{_OPENWEBUI_URL}/ollama/api/chat"
    async with aiohttp.ClientSession() as s:
        async with s.post(url, json=payload, headers=headers,
                          timeout=aiohttp.ClientTimeout(total=REQUEST_TIMEOUT)) as r:
            text = await r.text()
            if r.status != 200:
                return f"⚠️ Ollama error ({r.status}):\n```\n{text[:1500]}\n```"
            import json as _json
            data = _json.loads(text)
            try:
                return data["message"]["content"]
            except (KeyError, TypeError):
                return f"⚠️ 予期しない応答形式:\n```\n{text[:1500]}\n```"


def _split_for_discord(text: str) -> list[str]:
    if len(text) <= DISCORD_MSG_LIMIT:
        return [text]
    chunks: list[str] = []
    remaining = text
    while remaining:
        if len(remaining) <= DISCORD_MSG_LIMIT:
            chunks.append(remaining)
            break
        cut = remaining.rfind("\n", 0, DISCORD_MSG_LIMIT)
        if cut < DISCORD_MSG_LIMIT // 2:
            cut = DISCORD_MSG_LIMIT
        chunks.append(remaining[:cut])
        remaining = remaining[cut:].lstrip("\n")
    return chunks


async def _build_messages_and_model(
    prompt: str, attachments: list[discord.Attachment]
) -> tuple[list[dict], str]:
    images_b64: list[str] = []
    for att in attachments:
        if att.content_type and att.content_type.startswith("image/"):
            b64 = await _fetch_image_b64(att.url)
            if b64:
                images_b64.append(b64)

    if images_b64:
        message = {
            "role": "user",
            "content": prompt or "この画像について説明してください。",
            "images": images_b64,
        }
        return [message], _VISION_MODEL

    return [{"role": "user", "content": prompt}], _DEFAULT_MODEL


async def _respond_to_prompt(
    prompt: str, attachments: list[discord.Attachment], model_override: str | None = None
) -> tuple[str, str]:
    """Returns (model_used, reply_text)."""
    messages, default_model = await _build_messages_and_model(prompt, attachments)
    model = model_override or default_model
    reply = await _call_chat(messages, model)
    return model, reply


# /ai で選択可能なモデル一覧。新モデル追加時はここに足す。
_MODEL_CHOICES = [
    app_commands.Choice(name="qwen3.6:35b (汎用 / 高速)", value="qwen3.6:35b"),
    app_commands.Choice(name="llama4:scout (大規模)", value="llama4:scout"),
    app_commands.Choice(name="qwen3-vl:32b (画像対応 vision)", value="qwen3-vl:32b"),
]


def register_ollama_commands(bot: commands.Bot) -> None:
    if not _OPENWEBUI_URL or not _API_KEY:
        log.warning("[ollama] OPENWEBUI_URL or OPENWEBUI_API_KEY not set — skipped registering")
        return

    @bot.tree.command(name="ai", description="Mac Studio のローカル LLM に質問")
    @app_commands.describe(
        prompt="質問内容 (画像はメンションで送る方が楽です)",
        model="使用するモデル (未指定なら既定: qwen3.6:35b)",
    )
    @app_commands.choices(model=_MODEL_CHOICES)
    async def ai_command(
        interaction: discord.Interaction,
        prompt: str,
        model: app_commands.Choice[str] | None = None,
    ):
        if not _is_allowed(interaction.user.id):
            await interaction.response.send_message("⛔ 利用許可されていません", ephemeral=True)
            return
        await interaction.response.defer(thinking=True)
        try:
            model_used, reply = await _respond_to_prompt(
                prompt, [], model_override=model.value if model else None
            )
            chunks = _split_for_discord(f"**[{model_used}]**\n{reply}")
            await interaction.followup.send(chunks[0])
            for c in chunks[1:]:
                await interaction.followup.send(c)
        except Exception as e:
            log.exception("[ollama] /ai failed")
            await interaction.followup.send(f"⚠️ エラー: {e}")

    async def on_mention(message: discord.Message):
        if message.author.bot:
            return
        if not _is_allowed(message.author.id):
            return

        mentioned = bot.user is not None and bot.user in message.mentions
        auto_channel = message.channel.id in _AUTO_RESPOND_CHANNELS
        if not (mentioned or auto_channel):
            return

        prompt = message.content
        if mentioned and bot.user is not None:
            for token in (f"<@{bot.user.id}>", f"<@!{bot.user.id}>"):
                prompt = prompt.replace(token, "")
        prompt = prompt.strip()

        if not prompt and not message.attachments:
            return

        async with message.channel.typing():
            try:
                model, reply = await _respond_to_prompt(prompt, list(message.attachments))
                chunks = _split_for_discord(f"**[{model}]**\n{reply}")
                first_ref = message
                for c in chunks:
                    await message.channel.send(c, reference=first_ref if first_ref else None)
                    first_ref = None
            except Exception as e:
                log.exception("[ollama] auto reply failed")
                await message.channel.send(f"⚠️ エラー: {e}", reference=message)

    bot.add_listener(on_mention, "on_message")
    log.info(
        f"[ollama] commands registered (slash /ai + mention; "
        f"auto-respond channels: {len(_AUTO_RESPOND_CHANNELS)})"
    )
