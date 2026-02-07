"""
API コスト取得モジュール
OpenAI Costs API と Anthropic Cost Report API を非同期で呼び出す。

OpenAI:    GET https://api.openai.com/v1/organization/costs
Anthropic: GET https://api.anthropic.com/v1/organizations/cost_report

注意: Anthropic API は完了した日のバケットのみ返す（当日分は翌日確定）。
      ending_at を未来日にするとエラーになるため、省略して API にデフォルト（現在時刻）を使わせる。
"""
import os
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


async def fetch_all_costs(period: str = "today") -> dict:
    """
    全ての自動取得可能な API のコストを並列取得。

    Args:
        period: "today" | "month"

    Returns:
        {
            "openai": {"cost": float|None, "error": str|None},
            "anthropic": {"cost": float|None, "error": str|None},
            "period": "today" | "month",
            "fetched_at": "2026-02-06T14:23:45+00:00",
        }
    """
    openai_result, anthropic_result = await asyncio.gather(
        fetch_openai_cost(period),
        fetch_anthropic_cost(period),
        return_exceptions=True,
    )

    # gather が例外を返した場合の処理
    if isinstance(openai_result, Exception):
        openai_result = {"cost": None, "error": str(openai_result)}
    if isinstance(anthropic_result, Exception):
        anthropic_result = {"cost": None, "error": str(anthropic_result)}

    return {
        "openai": openai_result,
        "anthropic": anthropic_result,
        "period": period,
        "fetched_at": datetime.now(timezone.utc).isoformat(),
    }
