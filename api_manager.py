"""
API一覧 + コスト取得ボタン — Discord スラッシュコマンド＆ビュー

/api-list  : API・サブスクリプション一覧を Embed で表示
ボタン     : 「今日のコスト」「今月のコスト」で OpenAI / Anthropic 使用量を取得
"""
import discord
from discord import app_commands
from datetime import datetime, timezone, timedelta

from api_config import SUBSCRIPTIONS, API_LIST
from api_cost_fetcher import fetch_all_costs

# 日本時間 (UTC+9)
JST = timezone(timedelta(hours=9))


def _build_api_list_embed() -> discord.Embed:
    """API・サブスクリプション一覧の Embed を作成"""
    embed = discord.Embed(
        title="📋 API・サブスクリプション一覧",
        description="契約中のサービス一覧。下のボタンでコストを取得できます。",
        color=0x5865F2,
    )

    # サブスクリプション
    sub_lines = []
    for s in SUBSCRIPTIONS:
        name_link = f"[{s['name']}]({s['url']})"
        sub_lines.append(f"{s['icon']} **{name_link}** — {s['plan']}\n　　{s['note']}")
    embed.add_field(
        name="💳 サブスクリプション（月額・年額）",
        value="\n".join(sub_lines) if sub_lines else "なし",
        inline=False,
    )

    # API（自動取得）
    auto_lines = []
    for a in API_LIST:
        if a["cost_tracking"] == "auto":
            name_link = f"[{a['name']}]({a['docs_url']})"
            auto_lines.append(
                f"{a['icon']} **{name_link}** — {a['usage_location']}\n"
                f"　　{a['function_description']}（{a['pricing']}）"
            )
    if auto_lines:
        embed.add_field(
            name="🔄 API（コスト自動取得）",
            value="\n".join(auto_lines),
            inline=False,
        )

    # API（手動確認・無料）
    manual_lines = []
    for a in API_LIST:
        if a["cost_tracking"] in ("manual", "free"):
            name_link = f"[{a['name']}]({a['docs_url']})"
            label = "無料" if a["cost_tracking"] == "free" else "管理画面で確認"
            manual_lines.append(
                f"{a['icon']} **{name_link}** — {a['usage_location']}（{label}）"
            )
    if manual_lines:
        embed.add_field(
            name="📋 API（手動確認・無料）",
            value="\n".join(manual_lines),
            inline=False,
        )

    embed.set_footer(text="ボタンを押すと OpenAI / Anthropic / OpenClaw のコストを取得します")
    return embed


def _build_cost_embed(results: dict) -> discord.Embed:
    """コスト取得結果の Embed を作成"""
    period = results["period"]
    period_label = "今日" if period == "today" else "今月"
    now_jst = datetime.now(JST)

    if period == "today":
        period_range = now_jst.strftime("%Y-%m-%d")
    else:
        first_day = now_jst.replace(day=1).strftime("%Y-%m-%d")
        today = now_jst.strftime("%Y-%m-%d")
        period_range = f"{first_day} 〜 {today}"

    embed = discord.Embed(
        title=f"💰 API使用コスト（{period_label}）",
        description=f"📅 期間: {period_range}",
        color=0x2ECC71,
    )

    # --- 自動取得結果（Billing API） ---
    total = 0.0
    billing_lines = []

    # OpenAI
    openai = results.get("openai", {})
    if openai.get("error"):
        billing_lines.append(f"🤖 **OpenAI API**: ❌ {openai['error']}")
    else:
        cost = openai.get("cost", 0) or 0
        total += cost
        billing_lines.append(f"🤖 **OpenAI API**: ${cost:.4f}")

    # Anthropic
    anthropic = results.get("anthropic", {})
    if anthropic.get("error"):
        billing_lines.append(f"🤖 **Anthropic API**: ❌ {anthropic['error']}")
    else:
        cost = anthropic.get("cost", 0) or 0
        total += cost
        billing_lines.append(f"🤖 **Anthropic API**: ${cost:.4f}")

    embed.add_field(
        name="🔄 Billing API",
        value="\n".join(billing_lines),
        inline=False,
    )

    # --- OpenClaw 経由（トークン推定） ---
    openclaw = results.get("openclaw", {})
    openclaw_lines = []

    PROVIDER_DISPLAY = {
        "google": ("🔷", "Google Gemini"),
        "moonshot": ("🌙", "Moonshot (Kimi)"),
        "groq": ("⚡", "Groq"),
        "claude-cli": ("🤖", "Claude CLI (サブスク)"),
    }

    if openclaw.get("error"):
        openclaw_lines.append(f"❌ {openclaw['error']}")
    else:
        providers = openclaw.get("providers", {})
        for key, (icon, name) in PROVIDER_DISPLAY.items():
            if key in providers:
                cost = providers[key].get("cost", 0)
                calls = providers[key].get("calls", 0)
                total += cost
                openclaw_lines.append(f"{icon} **{name}**: ${cost:.4f} ({calls}回)")
        if not openclaw_lines:
            openclaw_lines.append("データなし（OpenClaw 未稼働 or 使用ゼロ）")

    embed.add_field(
        name="📡 OpenClaw（トークン推定）",
        value="\n".join(openclaw_lines),
        inline=False,
    )

    # --- 合計 ---
    embed.add_field(
        name="💰 合計",
        value=f"**${total:.4f}**",
        inline=False,
    )

    # 手動確認が必要なサービス
    manual_lines = []
    for a in API_LIST:
        if a["cost_tracking"] == "manual":
            manual_lines.append(
                f"{a['icon']} [{a['name']}]({a['dashboard_url']}) — 管理画面で確認"
            )
    if manual_lines:
        embed.add_field(
            name="📋 手動確認",
            value="\n".join(manual_lines),
            inline=False,
        )

    fetched_at = now_jst.strftime("%Y-%m-%d %H:%M:%S JST")
    embed.set_footer(text=f"取得時刻: {fetched_at} | ※ Anthropic: 前日確定分 / OpenClaw: トークン推定")
    return embed


class APICostView(discord.ui.View):
    """コスト取得ボタン（Persistent — Bot 再起動後も動作）"""

    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="📊 今日のコスト",
        style=discord.ButtonStyle.primary,
        custom_id="api_cost_today",
        row=0,
    )
    async def today_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._fetch_and_reply(interaction, "today")

    @discord.ui.button(
        label="📈 今月のコスト",
        style=discord.ButtonStyle.secondary,
        custom_id="api_cost_month",
        row=0,
    )
    async def month_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        await self._fetch_and_reply(interaction, "month")

    async def _fetch_and_reply(self, interaction: discord.Interaction, period: str):
        """コストを取得して返信"""
        await interaction.response.defer()

        try:
            results = await fetch_all_costs(period, bot=interaction.client)
            embed = _build_cost_embed(results)
            await interaction.followup.send(embed=embed)
        except Exception as e:
            await interaction.followup.send(f"❌ コスト取得中にエラーが発生しました: {e}")


def register_api_commands(bot):
    """
    /api-list スラッシュコマンドを Bot に登録。
    on_ready() から呼び出す。
    """

    @bot.tree.command(name="api-list", description="API・サブスクリプション一覧を表示")
    async def api_list(interaction: discord.Interaction):
        embed = _build_api_list_embed()
        view = APICostView()
        await interaction.response.send_message(embed=embed, view=view)
