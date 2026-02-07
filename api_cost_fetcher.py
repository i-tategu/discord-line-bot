"""
API コスト取得モジュール
OpenAI Costs API / Anthropic Cost Report API / OpenClaw Discord スナップショットを非同期で取得。

OpenAI:    GET https://api.openai.com/v1/organization/costs
Anthropic: GET https://api.anthropic.com/v1/organizations/cost_report
OpenClaw:  Discord #apiコスト チャンネルの JSON スナップショット（1時間ごと投稿）

注意: Anthropic API は完了した日のバケットのみ返す（当日分は翌日確定）。
      ending_at を未来日にするとエラーになるため、省略して API にデフォルト（現在時刻）を使わせる。
"""
import os
import re
import json
import asyncio
from datetime import datetime, timezone, timedelta

import aiohttp

# タイムアウト（秒）
REQUEST_TIMEOUT = 15


def _get_anthropic_start(period: str) -> str:
    """
    Anthropic API 用の starting_at を返す。
    ending_at は省略（API が自動で現在時刻を使用）。

    period: "today" → 昨日 0:00 UTC（当日バケットは未完了のため取得不可）
            "month" → 今月1日 0:00 UTC
    """
    now = datetime.now(timezone.utc)
    if period == "today":
        # 当日のバケットは未完了でエラーになるため、昨日から取得
        start = now.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)
    else:  # month
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    return start.strftime("%Y-%m-%dT%H:%M:%SZ")


def _get_period_unix(period: str) -> tuple[int, int]:
    """
    期間を Unix タイムスタンプ（秒）で返す。
    OpenAI Costs API 用。
    bucket_width=1d に対応するため、end は翌日0:00にする。
    """
    now = datetime.now(timezone.utc)
    tomorrow = now.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
    if period == "today":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    else:  # month
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    return int(start.timestamp()), int(tomorrow.timestamp())


async def fetch_openai_cost(period: str = "today") -> dict:
    """
    OpenAI Costs API でコストを取得。

    Returns:
        {"cost": float|None, "error": str|None, "details": str}
    """
    admin_key = os.environ.get("OPENAI_ADMIN_API_KEY")
    if not admin_key:
        return {"cost": None, "error": "環境変数 OPENAI_ADMIN_API_KEY 未設定"}

    start_time, end_time = _get_period_unix(period)

    url = "https://api.openai.com/v1/organization/costs"
    headers = {
        "Authorization": f"Bearer {admin_key}",
        "Content-Type": "application/json",
    }
    params = {
        "start_time": start_time,
        "end_time": end_time,
        "bucket_width": "1d",
    }

    try:
        timeout = aiohttp.ClientTimeout(total=REQUEST_TIMEOUT)
        async with aiohttp.ClientSession(timeout=timeout) as session:
            async with session.get(url, headers=headers, params=params) as resp:
                if resp.status == 401:
                    return {"cost": None, "error": "API認証エラー（Admin API Key を確認）"}
                if resp.status == 403:
                    return {"cost": None, "error": "権限不足（Admin API Key が必要）"}
                if resp.status == 429:
                    return {"cost": None, "error": "レート制限。数分後に再試行"}
                if resp.status != 200:
                    text = await resp.text()
                    return {"cost": None, "error": f"HTTPエラー {resp.status}: {text[:200]}"}

                data = await resp.json()

                # data.data[] に日別バケットがあり、results[].amount.value がコスト
                total_cost = 0.0
                for bucket in data.get("data", []):
                    for result in bucket.get("results", []):
                        amount = result.get("amount", {})
                        value = amount.get("value", 0)
                        if value:
                            total_cost += float(value)

                return {"cost": total_cost, "error": None}

    except asyncio.TimeoutError:
        return {"cost": None, "error": "タイムアウト（15秒）"}
    except aiohttp.ClientError as e:
        return {"cost": None, "error": f"接続エラー: {e}"}
    except Exception as e:
        return {"cost": None, "error": f"予期しないエラー: {e}"}


async def fetch_anthropic_cost(period: str = "today") -> dict:
    """
    Anthropic Cost Report API でコストを取得。

    注意:
    - amount はセント単位の10進文字列（"123.45" → $1.2345）
    - API は完了した日のバケットのみ返す（当日分は翌日確定）
    - ending_at を未来日にするとエラーになるため省略

    Returns:
        {"cost": float|None, "error": str|None}
    """
    admin_key = os.environ.get("ANTHROPIC_ADMIN_API_KEY")
    if not admin_key:
        return {"cost": None, "error": "環境変数 ANTHROPIC_ADMIN_API_KEY 未設定"}

    start_iso = _get_anthropic_start(period)

    headers = {
        "x-api-key": admin_key,
        "anthropic-version": "2023-06-01",
    }

    try:
        timeout = aiohttp.ClientTimeout(total=REQUEST_TIMEOUT)
        async with aiohttp.ClientSession(timeout=timeout) as session:
            cost_url = "https://api.anthropic.com/v1/organizations/cost_report"
            total_cost_cents = 0.0
            page = None

            while True:
                req_params = {
                    "starting_at": start_iso,
                    "bucket_width": "1d",
                }
                if page:
                    req_params["page"] = page

                async with session.get(cost_url, headers=headers, params=req_params) as resp:
                    if resp.status == 401:
                        return {"cost": None, "error": "API認証エラー（Admin API Key を確認）"}
                    if resp.status == 403:
                        return {"cost": None, "error": "権限不足（Admin API Key が必要）"}
                    if resp.status == 429:
                        return {"cost": None, "error": "レート制限。数分後に再試行"}
                    if resp.status != 200:
                        text = await resp.text()
                        return {"cost": None, "error": f"HTTPエラー {resp.status}: {text[:200]}"}

                    data = await resp.json()

                    for bucket in data.get("data", []):
                        for result in bucket.get("results", []):
                            amount_str = result.get("amount", "0")
                            try:
                                total_cost_cents += float(amount_str)
                            except (ValueError, TypeError):
                                pass

                    if data.get("has_more") and data.get("next_page"):
                        page = data["next_page"]
                    else:
                        break

            # セント → ドルに変換
            total_cost_dollars = total_cost_cents / 100.0
            return {"cost": total_cost_dollars, "error": None}

    except asyncio.TimeoutError:
        return {"cost": None, "error": "タイムアウト（15秒）"}
    except aiohttp.ClientError as e:
        return {"cost": None, "error": f"接続エラー: {e}"}
    except Exception as e:
        return {"cost": None, "error": f"予期しないエラー: {e}"}


def _parse_cost_snapshot(content: str) -> dict | None:
    """Discord メッセージから OpenClaw コストJSONスナップショットをパース"""
    match = re.search(r"```json\s*\n(.*?)\n```", content, re.DOTALL)
    if not match:
        return None
    try:
        data = json.loads(match.group(1))
        if data.get("type") == "OPENCLAW_COST_SNAPSHOT":
            return data
    except (json.JSONDecodeError, AttributeError):
        pass
    return None


async def fetch_openclaw_costs(bot, period: str = "today") -> dict:
    """
    OpenClaw Bot が Discord #apiコスト チャンネルに投稿した
    コストJSONスナップショットを読み取る。

    OpenClaw は Google/Moonshot/Groq 等のトークンベースコストを
    1時間ごとに JSON 形式で投稿している。

    Returns:
        {"providers": {name: {"cost": float, "calls": int}}, "total": float, "error": str|None}
    """
    guild_id = os.environ.get("OPENCLAW_GUILD_ID")
    channel_id = os.environ.get("OPENCLAW_APICOST_CHANNEL_ID")

    if not guild_id or not channel_id:
        return {"providers": {}, "total": 0, "error": "OPENCLAW 環境変数未設定"}

    try:
        guild = bot.get_guild(int(guild_id))
        if not guild:
            return {"providers": {}, "total": 0, "error": "OpenClaw サーバー未接続"}

        channel = guild.get_channel(int(channel_id))
        if not channel:
            return {"providers": {}, "total": 0, "error": "apiコスト チャンネル未発見"}

        now = datetime.now(timezone.utc)
        today_str = now.strftime("%Y-%m-%d")

        if period == "today":
            # 今日の最新スナップショットを取得
            async for msg in channel.history(limit=50):
                snapshot = _parse_cost_snapshot(msg.content)
                if snapshot and snapshot.get("date") == today_str:
                    return {
                        "providers": snapshot.get("providers", {}),
                        "total": snapshot.get("total", 0),
                        "error": None,
                    }
            return {"providers": {}, "total": 0, "error": None}

        else:  # month
            # 今月の各日の最新スナップショットを集約
            month_prefix = now.strftime("%Y-%m")
            daily_snapshots: dict[str, dict] = {}

            async for msg in channel.history(limit=500):
                snapshot = _parse_cost_snapshot(msg.content)
                if snapshot and snapshot.get("date", "").startswith(month_prefix):
                    date = snapshot["date"]
                    if date not in daily_snapshots:
                        daily_snapshots[date] = snapshot

            total_providers: dict[str, dict] = {}
            grand_total = 0.0
            for snap in daily_snapshots.values():
                for provider, info in snap.get("providers", {}).items():
                    if provider not in total_providers:
                        total_providers[provider] = {"cost": 0.0, "calls": 0}
                    total_providers[provider]["cost"] += info.get("cost", 0)
                    total_providers[provider]["calls"] += info.get("calls", 0)
                grand_total += snap.get("total", 0)

            return {"providers": total_providers, "total": grand_total, "error": None}

    except Exception as e:
        return {"providers": {}, "total": 0, "error": f"OpenClaw取得エラー: {e}"}


async def fetch_all_costs(period: str = "today", bot=None) -> dict:
    """
    全ての自動取得可能な API のコストを並列取得。

    Args:
        period: "today" | "month"
        bot: Discord Bot インスタンス（OpenClaw コスト取得用）

    Returns:
        {
            "openai": {"cost": float|None, "error": str|None},
            "anthropic": {"cost": float|None, "error": str|None},
            "openclaw": {"providers": {...}, "total": float, "error": str|None},
            "period": "today" | "month",
            "fetched_at": "2026-02-06T14:23:45+00:00",
        }
    """
    tasks = [fetch_openai_cost(period), fetch_anthropic_cost(period)]
    if bot:
        tasks.append(fetch_openclaw_costs(bot, period))

    results_list = await asyncio.gather(*tasks, return_exceptions=True)

    openai_result = results_list[0]
    anthropic_result = results_list[1]
    openclaw_result = results_list[2] if len(results_list) > 2 else {"providers": {}, "total": 0, "error": "Bot未接続"}

    # gather が例外を返した場合の処理
    if isinstance(openai_result, Exception):
        openai_result = {"cost": None, "error": str(openai_result)}
    if isinstance(anthropic_result, Exception):
        anthropic_result = {"cost": None, "error": str(anthropic_result)}
    if isinstance(openclaw_result, Exception):
        openclaw_result = {"providers": {}, "total": 0, "error": str(openclaw_result)}

    return {
        "openai": openai_result,
        "anthropic": anthropic_result,
        "openclaw": openclaw_result,
        "period": period,
        "fetched_at": datetime.now(timezone.utc).isoformat(),
    }
