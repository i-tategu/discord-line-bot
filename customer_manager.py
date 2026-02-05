"""
顧客ステータス管理モジュール
"""
import os
import json
from datetime import datetime
from enum import Enum

# ステータス定義（WooCommerceカスタムステータスと一致）
class CustomerStatus(Enum):
    PURCHASED = "purchased"            # 購入済み
    DESIGN_CONFIRMED = "design-confirmed"  # デザイン確定
    PRODUCED = "produced"              # 制作完了
    SHIPPED = "shipped"                # 発送済み

# 旧ステータス値のマイグレーション
STATUS_MIGRATION = {
    "design": "design-confirmed",
    "production": "produced",
    "designing": "purchased",
    "completed": "shipped",
}

# ステータス表示設定
STATUS_CONFIG = {
    CustomerStatus.PURCHASED: {
        "label": "購入済み",
        "emoji": "🟡",
        "color": 0xFFD700,
    },
    CustomerStatus.DESIGN_CONFIRMED: {
        "label": "デザイン確定",
        "emoji": "🔵",
        "color": 0x3498DB,
    },
    CustomerStatus.PRODUCED: {
        "label": "制作完了",
        "emoji": "🟢",
        "color": 0x2ECC71,
    },
    CustomerStatus.SHIPPED: {
        "label": "発送済み",
        "emoji": "📦",
        "color": 0x9B59B6,
    },
}

# データファイル（Railway Volume対応）
# 環境変数 DATA_DIR が設定されていればそちらを使用（永続化）
# 未設定の場合はローカル開発用にカレントディレクトリを使用
DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(__file__))
DATA_FILE = os.path.join(DATA_DIR, "customers.json")

# ディレクトリが存在しない場合は作成
if DATA_DIR and not os.path.exists(DATA_DIR):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except Exception:
        pass  # 作成失敗時はそのまま続行


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


def add_order_customer(order_id, customer_name, email, order_info=None):
    """WooCommerce注文から顧客追加（LINE未連携）"""
    customers = load_customers()

    # emailをキーとして使用（LINE連携前の仮ID）
    customer_key = f"wc_{email}" if email else f"order_{order_id}"

    # 既存顧客を確認（同じメールアドレス）
    existing_key = None
    for key, data in customers.items():
        if data.get("email") == email and email:
            existing_key = key
            break

    if existing_key:
        customer_key = existing_key
    elif customer_key not in customers:
        customers[customer_key] = {
            "display_name": customer_name,
            "email": email,
            "discord_channel_id": None,  # LINE連携時に設定
            "line_user_id": None,  # LINE連携時に設定
            "status": CustomerStatus.PURCHASED.value,
            "orders": [],
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
        }

    # 注文追加（重複チェック）
    order_exists = any(
        str(o["order_id"]) == str(order_id)
        for o in customers[customer_key].get("orders", [])
    )
    if not order_exists:
        order_data = {
            "order_id": order_id,
            "status": CustomerStatus.PURCHASED.value,
            "info": order_info or {},
            "created_at": datetime.now().isoformat(),
        }
        customers[customer_key]["orders"].append(order_data)
        customers[customer_key]["updated_at"] = datetime.now().isoformat()

    save_customers(customers)
    return customers[customer_key]


def link_line_to_customer(email, line_user_id, discord_channel_id):
    """LINE連携: emailでWooCommerce顧客を検索してLINEアカウントを紐付け"""
    customers = load_customers()

    # emailで既存顧客を検索
    for key, data in customers.items():
        if data.get("email") == email:
            # LINE情報を更新
            data["line_user_id"] = line_user_id
            data["discord_channel_id"] = discord_channel_id
            data["updated_at"] = datetime.now().isoformat()
            save_customers(customers)
            return key, data

    return None, None


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
        # 旧ステータス値をマイグレーション
        status_str = STATUS_MIGRATION.get(status_str, status_str)
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
