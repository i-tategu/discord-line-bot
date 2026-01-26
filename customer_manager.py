"""
顧客ステータス管理モジュール
"""
import os
import json
from datetime import datetime
from enum import Enum

# ステータス定義
class CustomerStatus(Enum):
    PURCHASED = "purchased"        # 購入済み
    DESIGN_CONFIRMED = "design"    # デザイン確定
    PRODUCTION_DONE = "production" # 制作完了
    SHIPPED = "shipped"            # 発送済み

# ステータス表示設定
STATUS_CONFIG = {
    CustomerStatus.PURCHASED: {
        "label": "購入済み",
        "emoji": "🟡",
        "color": 0xFFD700,  # ゴールド
    },
    CustomerStatus.DESIGN_CONFIRMED: {
        "label": "デザイン確定",
        "emoji": "🔵",
        "color": 0x3498DB,  # ブルー
    },
    CustomerStatus.PRODUCTION_DONE: {
        "label": "制作完了",
        "emoji": "🟢",
        "color": 0x2ECC71,  # グリーン
    },
    CustomerStatus.SHIPPED: {
        "label": "発送済み",
        "emoji": "✅",
        "color": 0x95A5A6,  # グレー
    },
}

# データファイル
DATA_FILE = os.path.join(os.path.dirname(__file__), "customers.json")


def load_customers():
    """顧客データ読み込み"""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_customers(data):
    """顧客データ保存"""
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add_customer(line_user_id, display_name, discord_channel_id, order_id=None, order_info=None):
    """顧客追加"""
    customers = load_customers()

    if line_user_id not in customers:
        customers[line_user_id] = {
            "display_name": display_name,
            "discord_channel_id": discord_channel_id,
            "status": CustomerStatus.PURCHASED.value,
            "orders": [],
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
        }

    # 注文追加
    if order_id:
        order_data = {
            "order_id": order_id,
            "status": CustomerStatus.PURCHASED.value,
            "info": order_info or {},
            "created_at": datetime.now().isoformat(),
        }
        customers[line_user_id]["orders"].append(order_data)
        customers[line_user_id]["updated_at"] = datetime.now().isoformat()

    save_customers(customers)
    return customers[line_user_id]


def update_customer_status(line_user_id, new_status: CustomerStatus, order_id=None):
    """顧客ステータス更新"""
    customers = load_customers()

    if line_user_id not in customers:
        return None

    customers[line_user_id]["status"] = new_status.value
    customers[line_user_id]["updated_at"] = datetime.now().isoformat()

    # 特定の注文のステータスを更新
    if order_id:
        for order in customers[line_user_id]["orders"]:
            if str(order["order_id"]) == str(order_id):
                order["status"] = new_status.value
                break

    save_customers(customers)
    return customers[line_user_id]


def get_customer(line_user_id):
    """顧客情報取得"""
    customers = load_customers()
    return customers.get(line_user_id)


def get_customer_by_channel(discord_channel_id):
    """チャンネルIDから顧客情報取得"""
    customers = load_customers()
    for line_user_id, data in customers.items():
        if str(data.get("discord_channel_id")) == str(discord_channel_id):
            return line_user_id, data
    return None, None


def get_customer_by_order(order_id):
    """注文IDから顧客情報取得"""
    customers = load_customers()
    for line_user_id, data in customers.items():
        for order in data.get("orders", []):
            if str(order["order_id"]) == str(order_id):
                return line_user_id, data
    return None, None


def get_customers_by_status(status: CustomerStatus):
    """ステータス別顧客リスト取得"""
    customers = load_customers()
    result = []
    for line_user_id, data in customers.items():
        if data.get("status") == status.value:
            result.append({
                "line_user_id": line_user_id,
                **data
            })
    return result


def get_all_customers_grouped():
    """全顧客をステータス別にグループ化"""
    customers = load_customers()
    grouped = {status: [] for status in CustomerStatus}

    for line_user_id, data in customers.items():
        status_str = data.get("status", CustomerStatus.PURCHASED.value)
        try:
            status = CustomerStatus(status_str)
        except ValueError:
            status = CustomerStatus.PURCHASED

        grouped[status].append({
            "line_user_id": line_user_id,
            **data
        })

    return grouped


def get_status_summary():
    """ステータス別サマリー取得"""
    grouped = get_all_customers_grouped()
    summary = {}
    for status, customers in grouped.items():
        config = STATUS_CONFIG[status]
        # JSONシリアライズのためenumのvalueを使用
        summary[status.value] = {
            "count": len(customers),
            "customers": customers,
            "label": config["label"],
            "emoji": config["emoji"],
            "color": config["color"],
        }
    return summary
