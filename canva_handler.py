# -*- coding: utf-8 -*-
"""
Canva自動化ハンドラー（Railway版）
WooCommerce Webhookを受信してCanvaデザインを自動作成
"""
import os
import re
import json
import base64
import requests
import time
import tempfile
from io import BytesIO
from datetime import datetime

# python-pptx
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Pillow
from PIL import Image

# reportlab for PDF creation
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# トークン永続化ファイルパス（Railway Volume対応）
# DATA_DIR環境変数で永続化ディレクトリを指定
DATA_DIR = os.environ.get("DATA_DIR", "/tmp")
TOKEN_FILE_PATH = os.path.join(DATA_DIR, "canva_tokens.json")

# ディレクトリが存在しない場合は作成
if not os.path.exists(DATA_DIR):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except Exception:
        pass

def save_tokens_to_file(access_token, refresh_token):
    """トークンをファイルに保存（再起動後も維持）"""
    try:
        tokens = {
            'access_token': access_token,
            'refresh_token': refresh_token,
            'updated_at': datetime.now().isoformat()
        }
        with open(TOKEN_FILE_PATH, 'w') as f:
            json.dump(tokens, f)
        print(f"[Token] Saved to {TOKEN_FILE_PATH}")
        return True
    except Exception as e:
        print(f"[Token] Failed to save: {e}")
        return False

def load_tokens_from_file():
    """ファイルからトークンを読み込み（環境変数より優先）"""
    try:
        if os.path.exists(TOKEN_FILE_PATH):
            with open(TOKEN_FILE_PATH, 'r') as f:
                tokens = json.load(f)
            print(f"[Token] Loaded from file (updated: {tokens.get('updated_at', 'unknown')})")
            return tokens.get('access_token'), tokens.get('refresh_token')
    except Exception as e:
        print(f"[Token] Failed to load from file: {e}")
    return None, None

def get_current_tokens():
    """現在有効なトークンを取得（ファイル優先、なければ環境変数）"""
    file_access, file_refresh = load_tokens_from_file()
    if file_access and file_refresh:
        return file_access, file_refresh
    # ファイルになければ環境変数から
    return os.getenv("CANVA_ACCESS_TOKEN", ""), os.getenv("CANVA_REFRESH_TOKEN", "")

# 設定（環境変数から取得 - 遅延読み込み）
def get_canva_client_id():
    return os.getenv("CANVA_CLIENT_ID", "OC-AZvUVtxGhbOD")

def get_canva_client_secret():
    return os.getenv("CANVA_CLIENT_SECRET", "")

# サーバー上のcutout画像URL
CUTOUT_BASE_URL = "https://i-tategu-shop.com/wp-content/themes/i-tategu/assets/images/cutouts"
TREE_IMAGES_URL = "https://i-tategu-shop.com/wp-content/themes/i-tategu/assets/images"

# スライドサイズ（1:1正方形、シミュレーター座標を変換）
SLIDE_WIDTH_PX = 1000
SLIDE_HEIGHT_PX = 1000  # 1:1 ratio
EMU_PER_PX = 914400 / 96

# シミュレーターは4:3 (500x375)、PPTXは1:1 (1000x1000)
# Y座標を変換: 4:3コンテンツを1:1の中央に配置
# 垂直オフセット = (1000 - 750) / 2 = 125px
SIMULATOR_ASPECT_HEIGHT = 750  # 1000幅での4:3相当の高さ
Y_OFFSET = (SLIDE_HEIGHT_PX - SIMULATOR_ASPECT_HEIGHT) / 2  # = 125px

# フォントマッピング
FONT_MAP = {
    'Alex Brush': 'Alex Brush',
    'Great Vibes': 'Great Vibes',
    'Pinyon Script': 'Pinyon Script',
    'Sacramento': 'Sacramento',
    'Dancing Script': 'Dancing Script',
    'Parisienne': 'Parisienne',
    'Allura': 'Allura',
    'Satisfy': 'Satisfy',
    'Rouge Script': 'Rouge Script',
    # 日本語フォント（人気）
    'Shippori Mincho': 'Shippori Mincho',
    'Zen Old Mincho': 'Zen Old Mincho',
    'Klee One': 'Klee One',
    'Noto Serif JP': 'Noto Serif JP',
    'Zen Maru Gothic': 'Zen Maru Gothic',
    'Sawarabi Mincho': 'Sawarabi Mincho',
    'Noto Sans JP': 'Noto Sans JP',
    # 日本語フォント（個性派）
    'Yomogi': 'Yomogi',
    'Kaisei Decol': 'Kaisei Decol',
    'Reggae One': 'Reggae One',
}

# フォント表示名マッピング（シミュレーターと同じスタイル表記、英語版）
FONT_DISPLAY_MAP = {
    'Sacramento': 'Holiday style',
    'Pinyon Script': 'Eyesome style',
    'Satisfy': 'Mistrully style',
    'Rouge Script': 'Amsterdam style',
}

# テンプレート
TEMPLATES = {
    'holy': """I hereby certify
that this man and woman
were united in holy matrimony,
in the name of the Father,
of the Son and of the Holy Spirit.""",
    'happy': """We are happy to be married
before Holy God and many witnesses.
We promise that in this new life
we will have joy and happiness.""",
    'promise': """We promise to be husband and wife
before these witnesses here present.
We swear to admire, love and help each other
and to create a peaceful happy family."""
}

TITLES = {
    'wedding': 'Wedding Certificate',
    'marriage': 'Certificate of Marriage'
}

BACKGROUND_MAP = {
    'product': '_1_cutout.png',
    'product_back': '_2_cutout.png',
    'noclear': '_3_cutout.png',
    'noclear_back': '_4_cutout.png',
}


def px_to_emu(px):
    return int(px * EMU_PER_PX)


def sim_y_to_pptx_y(sim_y_pct):
    """シミュレーターY座標(%)をPPTX Y座標(px)に変換

    シミュレーター: 4:3 (500x375)
    PPTX: 1:1 (1000x1000)
    4:3コンテンツを1:1の中央に配置
    """
    return Y_OFFSET + (sim_y_pct / 100) * SIMULATOR_ASPECT_HEIGHT


def detect_board_shape(cutout_img):
    """板の形状と境界を検出（シミュレーターと同じロジック）

    Returns:
        dict: {
            'type': 'portrait' | 'landscape' | 'square',
            'aspectRatio': float,
            'stretchFactor': float (portrait時のみ有効),
            'bounds': {'minX': 0-1, 'maxX': 0-1, 'minY': 0-1, 'maxY': 0-1}  # 正規化された境界
        }
    """
    default_bounds = {'minX': 0, 'maxX': 1, 'minY': 0, 'maxY': 1}

    if cutout_img is None:
        return {'type': 'square', 'aspectRatio': 1.0, 'stretchFactor': 0, 'bounds': default_bounds}

    # 透明部分を除いた実際の板サイズを検出
    if cutout_img.mode != 'RGBA':
        cutout_img = cutout_img.convert('RGBA')

    img_w, img_h = cutout_img.size
    alpha = cutout_img.split()[3]
    alpha_data = alpha.load()

    # 非透明ピクセルの境界を検出
    min_x, max_x, min_y, max_y = img_w, 0, img_h, 0
    for py in range(img_h):
        for px in range(img_w):
            if alpha_data[px, py] > 10:
                min_x = min(min_x, px)
                max_x = max(max_x, px)
                min_y = min(min_y, py)
                max_y = max(max_y, py)

    if max_x <= min_x or max_y <= min_y:
        return {'type': 'square', 'aspectRatio': 1.0, 'stretchFactor': 0, 'bounds': default_bounds}

    # 正規化された境界（0-1の範囲）
    bounds = {
        'minX': min_x / img_w,
        'maxX': max_x / img_w,
        'minY': min_y / img_h,
        'maxY': max_y / img_h
    }

    # 実際の板サイズ
    board_w = max_x - min_x
    board_h = max_y - min_y
    aspect_ratio = board_w / board_h

    # 形状判定（シミュレーターと同じ許容範囲0.15）
    if aspect_ratio > 1.15:
        shape_type = 'landscape'  # 横長
    elif aspect_ratio < 0.85:
        shape_type = 'portrait'   # 縦長
    else:
        shape_type = 'square'     # 正方形

    # 縦長の場合のstretchFactor（縦横比から1を引いた値、最大3）
    stretch_factor = 0
    if shape_type == 'portrait':
        stretch_factor = min((1 / aspect_ratio) - 1, 3)

    print(f"[Shape] Detected: {shape_type}, aspectRatio={aspect_ratio:.2f}, stretchFactor={stretch_factor:.2f}")
    print(f"[Shape] Bounds: X={bounds['minX']:.2f}-{bounds['maxX']:.2f}, Y={bounds['minY']:.2f}-{bounds['maxY']:.2f}")

    return {
        'type': shape_type,
        'aspectRatio': aspect_ratio,
        'stretchFactor': stretch_factor,
        'bounds': bounds
    }


def get_portrait_layout_adjustments(stretch_factor):
    """縦長板用のレイアウト調整値を取得（シミュレーター準拠）"""
    return {
        'titleY': 6,
        'titleSize': 90,
        'bodyY': 15,
        'bodySize': 80,
        'dateY': 50 + 8 * stretch_factor,
        'dateSize': 85,
        'nameY': 62 + 10 * stretch_factor,
        'nameSize': 90,
        'treeX': 0.5,
        'treeY': 0.82,
        'treeSize': 70,
    }


def get_landscape_layout_adjustments():
    """横長板用のレイアウト調整値を取得"""
    return {
        'titleY': 20,
        'titleSize': 95,
        'bodyY': 30,
        'bodySize': 90,
        'dateY': 55,
        'dateSize': 85,
        'nameY': 68,
        'nameSize': 90,
        'treeX': 0.85,
        'treeY': 0.5,
        'treeSize': 60,
    }


def format_date(date_str, format_type):
    try:
        for fmt in ["%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"]:
            try:
                d = datetime.strptime(date_str, fmt)
                break
            except:
                continue
        else:
            return date_str
    except:
        return date_str

    months_full = ['January', 'February', 'March', 'April', 'May', 'June',
                   'July', 'August', 'September', 'October', 'November', 'December']
    months_short = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

    def ordinal(n):
        s = ['th', 'st', 'nd', 'rd']
        v = n % 100
        return str(n) + (s[(v - 20) % 10] if v > 20 else s[v] if v < 4 else s[0])

    if format_type == 'western':
        return f"{d.year}.{d.month:02d}.{d.day:02d}"
    elif format_type == 'us_long':
        return f"{months_full[d.month-1]} {ordinal(d.day)}, {d.year}"
    elif format_type == 'us_short':
        return f"{months_short[d.month-1]} {ordinal(d.day)}, {d.year}"
    elif format_type == 'uk_long':
        return f"{ordinal(d.day)} {months_full[d.month-1]} {d.year}"
    elif format_type == 'uk_short':
        return f"{ordinal(d.day)} {months_short[d.month-1]} {d.year}"
    else:
        return date_str


def add_text_box(slide, text, x_px, y_px, font_name, font_size_pt, center=True, color_rgb=(42, 24, 16)):
    width_px = max(len(text) * font_size_pt * 0.8, 100)
    height_px = font_size_pt * 2

    left = px_to_emu(x_px - width_px / 2) if center else px_to_emu(x_px)
    top = px_to_emu(y_px - height_px / 2)
    width = px_to_emu(width_px)
    height = px_to_emu(height_px)

    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = False

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if center else PP_ALIGN.LEFT

    run = p.add_run()
    run.text = text
    run.font.name = FONT_MAP.get(font_name, font_name)
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = RGBColor(*color_rgb)

    return textbox


def add_multiline_text_box(slide, text, x_px, y_px, font_name, font_size_pt, line_height=1.4, color_rgb=(42, 24, 16)):
    lines = text.strip().split('\n')
    width_px = max(len(line) * font_size_pt * 0.7 for line in lines)
    height_px = len(lines) * font_size_pt * line_height * 1.5

    left = px_to_emu(x_px - width_px / 2)
    top = px_to_emu(y_px)
    width = px_to_emu(width_px)
    height = px_to_emu(height_px)

    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = False

    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(0)
        p.space_after = Pt(font_size_pt * (line_height - 1))

        run = p.add_run()
        run.text = line.strip()
        run.font.name = FONT_MAP.get(font_name, font_name)
        run.font.size = Pt(font_size_pt)
        run.font.color.rgb = RGBColor(*color_rgb)

    return textbox


def extract_board_info(product_name):
    """商品名から板情報を抽出"""
    name = ''
    number = '01'
    size = ''

    # パターン1: "【一点物】タモ 一枚板 結婚証明書 300x300mm"
    match = re.search(r'【一点物】\s*(.+?)\s+一枚板.*?(\d+)x(\d+)', product_name)
    if match:
        name = match.group(1).strip()
        size = f"{match.group(2)}_{match.group(3)}"
        return {'name': name, 'number': number, 'size': size}

    # パターン2: "ケヤキ_01_400_600"
    match = re.match(r'^([^_]+)_(\d+)(?:_(\d+)_(\d+))?', product_name)
    if match:
        name = match.group(1)
        number = match.group(2)
        if match.group(3) and match.group(4):
            size = f"{match.group(3)}_{match.group(4)}"
        return {'name': name, 'number': number.zfill(2), 'size': size}

    # パターン3: "ケヤキ No.01"
    match = re.search(r'^(.+?)\s*No\.?(\d+)', product_name)
    if match:
        name = match.group(1).strip()
        number = match.group(2)
        return {'name': name, 'number': number.zfill(2), 'size': size}

    return {'name': name, 'number': number, 'size': size}


def find_cutout_url(board_name, board_number, board_size, background='product'):
    """cutout画像のURLを構築"""
    suffix = BACKGROUND_MAP.get(background, '_1_cutout.png')

    # サイズありの場合
    if board_size:
        filename = f"{board_name}_{board_number}_{board_size}{suffix}"
    else:
        filename = f"{board_name}_{board_number}_300_300{suffix}"  # デフォルトサイズ

    return f"{CUTOUT_BASE_URL}/{filename}"


def download_image(url, temp_dir, max_size=800, preserve_transparency=False):
    """URLから画像をダウンロードして圧縮保存

    Args:
        preserve_transparency: Trueの場合、PNG形式で透明度を保持
    """
    try:
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            # 画像を開いてリサイズ・圧縮
            img = Image.open(BytesIO(response.content))

            # リサイズ（最大サイズを超える場合）
            if max(img.size) > max_size:
                ratio = max_size / max(img.size)
                new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
                img = img.resize(new_size, Image.LANCZOS)

            if preserve_transparency:
                # PNG形式で透明度を保持
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                filename = os.path.splitext(os.path.basename(url))[0] + '.png'
                filepath = os.path.join(temp_dir, filename)
                img.save(filepath, 'PNG', optimize=True)
            else:
                # JPEG形式（透明度を白背景に変換）
                if img.mode == 'RGBA':
                    white_bg = Image.new('RGB', img.size, (255, 255, 255))
                    white_bg.paste(img, mask=img.split()[3])
                    img = white_bg
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                filename = os.path.splitext(os.path.basename(url))[0] + '.jpg'
                filepath = os.path.join(temp_dir, filename)
                img.save(filepath, 'JPEG', quality=80, optimize=True)

            print(f"[IMG] {'PNG' if preserve_transparency else 'JPEG'}: {os.path.getsize(filepath) / 1024:.1f}KB")
            return filepath
    except Exception as e:
        print(f"[WARN] Failed to download image: {url} - {e}")
    return None


def refresh_canva_token(refresh_token):
    """Canvaトークンをリフレッシュし、環境変数に保存"""
    url = 'https://api.canva.com/rest/v1/oauth/token'
    client_id = get_canva_client_id()
    client_secret = get_canva_client_secret()

    print(f"[Canva Token] Client ID: {client_id}")
    print(f"[Canva Token] Client Secret: {'SET' if client_secret else 'EMPTY'} (len={len(client_secret) if client_secret else 0})")

    credentials = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()

    response = requests.post(url, data={
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
    }, headers={
        'Authorization': f'Basic {credentials}',
        'Content-Type': 'application/x-www-form-urlencoded',
    })

    print(f"[Canva Token] Refresh status: {response.status_code}")

    if response.status_code == 200:
        tokens = response.json()
        new_access = tokens.get('access_token')
        new_refresh = tokens.get('refresh_token', refresh_token)

        # 環境変数を更新（プロセス内）
        os.environ['CANVA_ACCESS_TOKEN'] = new_access
        os.environ['CANVA_REFRESH_TOKEN'] = new_refresh

        # ファイルにも保存（再起動後も維持）
        save_tokens_to_file(new_access, new_refresh)

        print(f"[Canva Token] Refresh successful! Tokens updated in memory and file.")
        print(f"[Canva Token] New refresh token (first 50 chars): {new_refresh[:50]}...")

        return {
            'access_token': new_access,
            'refresh_token': new_refresh
        }

    print(f"[Canva Token] Refresh failed: {response.text[:500]}")
    return None


def import_to_canva(file_path, title, access_token, refresh_token, retry=False):
    """CanvaにファイルをインポートPDF/PPTX対応（エラー情報も返す）"""
    url = "https://api.canva.com/rest/v1/imports"
    title_base64 = base64.b64encode(title.encode("utf-8")).decode("utf-8")

    # ファイル形式を検出してMIMEタイプを設定
    if file_path.endswith('.pdf'):
        mime_type = "application/pdf"
        print(f"[Canva Import] Using PDF format")
    else:
        mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        print(f"[Canva Import] Using PPTX format")

    metadata = {
        "title_base64": title_base64,
        "mime_type": mime_type
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream",
        "Import-Metadata": json.dumps(metadata)
    }

    with open(file_path, "rb") as f:
        response = requests.post(url, headers=headers, data=f)

    print(f"[Canva Import] Status: {response.status_code}")

    # 401エラー時はトークンリフレッシュ
    if response.status_code == 401 and not retry:
        print("[INFO] Token expired, refreshing...")
        new_tokens = refresh_canva_token(refresh_token)
        if new_tokens:
            return import_to_canva(file_path, title, new_tokens['access_token'], new_tokens['refresh_token'], retry=True)
        return None, {"error": "Token refresh failed"}

    if response.status_code != 200:
        error_msg = f"HTTP {response.status_code}: {response.text[:500]}"
        print(f"[ERROR] Canva import failed: {error_msg}")
        return None, {"error": error_msg}

    job = response.json().get("job", {})
    job_id = job.get("id")
    print(f"[Canva Import] Job ID: {job_id}")

    # リフレッシュした場合は新しいトークンを使用
    check_token = access_token

    # ジョブ完了を待機
    for i in range(15):
        time.sleep(2)
        check_url = f"https://api.canva.com/rest/v1/imports/{job_id}"
        check_resp = requests.get(check_url, headers={"Authorization": f"Bearer {check_token}"})

        print(f"[Canva Import] Check {i+1}: {check_resp.status_code}")

        if check_resp.status_code == 200:
            job_data = check_resp.json().get("job", {})
            status = job_data.get("status")
            print(f"[Canva Import] Status: {status}")

            if status == "success":
                designs = job_data.get("result", {}).get("designs", [])
                if designs:
                    return designs[0], None
            elif status == "failed":
                error = job_data.get("error", {})
                error_msg = error.get('message', 'Unknown error')
                print(f"[ERROR] {error_msg}")
                return None, {"error": error_msg, "details": error}
        elif check_resp.status_code == 401:
            # チェック中にトークンが切れた場合、リフレッシュ
            new_tokens = refresh_canva_token(refresh_token)
            if new_tokens:
                check_token = new_tokens['access_token']

    print("[ERROR] Timeout waiting for Canva import")
    return None, {"error": "Timeout"}


def get_order_from_woocommerce(order_id, wc_url, wc_key, wc_secret):
    """WooCommerceから注文を取得"""
    url = f"{wc_url}/wp-json/wc/v3/orders/{order_id}"
    print(f"[WC API] Fetching: {url}")

    try:
        response = requests.get(url, auth=(wc_key, wc_secret))
        print(f"[WC API] Status: {response.status_code}")

        if response.status_code == 200:
            return response.json()
        else:
            print(f"[WC API] Error response: {response.text[:500]}")
            return None
    except Exception as e:
        print(f"[WC API] Exception: {e}")
        return None


def parse_order_data(order):
    """注文データを解析"""
    meta = {m['key']: m['value'] for m in order.get('meta_data', [])}

    # シミュレーションデータ
    sim_data_raw = meta.get('_simulation_data', '{}')
    try:
        sim_data = json.loads(sim_data_raw) if sim_data_raw else {}
    except:
        sim_data = {}

    # シミュレーション画像（透かしなしファイルURLを優先）
    sim_image_url = meta.get('_simulation_image_url', '')
    sim_image = sim_image_url if sim_image_url else meta.get('_simulation_image', '')

    # 刻印情報（フォールバック）
    if not sim_data.get('groomName'):
        sim_data['groomName'] = meta.get('_engraving_name1', '')
    if not sim_data.get('brideName'):
        sim_data['brideName'] = meta.get('_engraving_name2', '')

    # フォント情報（フォールバック）
    if not sim_data.get('baseFont'):
        font_style = meta.get('_font_style', '')
        if font_style:
            sim_data['baseFont'] = font_style

    # 商品情報
    product_name = ''
    for item in order.get('line_items', []):
        product_name = item.get('name', '')
        break

    board_info = extract_board_info(product_name)

    # 日付
    wedding_date = meta.get('_engraving_date', '')
    if not wedding_date:
        wedding_date = sim_data.get('weddingDate', '')

    return {
        'order_id': order['id'],
        'sim_data': sim_data,
        'sim_image': sim_image,
        'product_name': product_name,
        'board_name': board_info['name'],
        'board_number': board_info['number'],
        'board_size': board_info['size'],
        'wedding_date': wedding_date,
    }


def create_pptx(order_data, temp_dir):
    """PowerPointを作成（6ページ構成）"""
    print(f"[Canva] Creating PowerPoint for order #{order_data['order_id']}...")

    sim_data = order_data['sim_data']
    sim_image = order_data['sim_image']
    groom = sim_data.get('groomName', '')
    bride = sim_data.get('brideName', '')

    prs = Presentation()
    prs.slide_width = px_to_emu(SLIDE_WIDTH_PX)
    prs.slide_height = px_to_emu(SLIDE_HEIGHT_PX)
    blank_layout = prs.slide_layouts[6]

    # cutout画像のURLを取得
    background = sim_data.get('background', 'product')
    cutout_urls = {
        'product': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'product'),
        'product_back': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'product_back'),
        'noclear': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'noclear'),
        'noclear_back': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'noclear_back'),
    }

    # ========== 1ページ目: シミュレーション画像 + 注文情報 ==========
    slide1 = prs.slides.add_slide(blank_layout)

    if sim_image:
        try:
            temp_img_path = os.path.join(temp_dir, f"sim_{order_data['order_id']}.jpg")

            # 画像データを取得
            if sim_image.startswith('http'):
                response = requests.get(sim_image, timeout=30)
                response.raise_for_status()
                img = Image.open(BytesIO(response.content))
            else:
                if ',' in sim_image:
                    sim_image = sim_image.split(',')[1]
                img_data = base64.b64decode(sim_image)
                img = Image.open(BytesIO(img_data))

            # 圧縮: RGBA→RGB変換、リサイズ
            if img.mode == 'RGBA':
                white_bg = Image.new('RGB', img.size, (255, 255, 255))
                white_bg.paste(img, mask=img.split()[3])
                img = white_bg
            elif img.mode != 'RGB':
                img = img.convert('RGB')

            # 最大800pxにリサイズ
            max_size = 800
            if max(img.size) > max_size:
                ratio = max_size / max(img.size)
                new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
                img = img.resize(new_size, Image.LANCZOS)

            # JPEGで保存
            img.save(temp_img_path, 'JPEG', quality=80, optimize=True)
            print(f"[IMG] Sim image: {os.path.getsize(temp_img_path) / 1024:.1f}KB")

            img = Image.open(temp_img_path)
            img_width, img_height = img.size
            available_height = SLIDE_HEIGHT_PX * 0.75
            scale = min(SLIDE_WIDTH_PX * 0.95 / img_width, available_height / img_height)
            draw_width = img_width * scale
            draw_height = img_height * scale
            x = (SLIDE_WIDTH_PX - draw_width) / 2
            y = 20

            slide1.shapes.add_picture(
                temp_img_path,
                px_to_emu(x), px_to_emu(y),
                px_to_emu(draw_width), px_to_emu(draw_height)
            )
        except Exception as e:
            print(f"[WARN] Simulation image error: {e}")

    # 下部に注文情報
    info_y = SLIDE_HEIGHT_PX - 160
    add_text_box(slide1, f"注文 #{order_data['order_id']} - {order_data['board_name']} No.{order_data['board_number']}",
                 SLIDE_WIDTH_PX/2, info_y, 'Meiryo UI', 16, center=True, color_rgb=(40, 40, 40))

    info_y += 30
    add_text_box(slide1, f"新郎: {groom}　　新婦: {bride}",
                 SLIDE_WIDTH_PX/2, info_y, 'Meiryo UI', 13, center=True, color_rgb=(60, 60, 60))

    info_y += 26
    add_text_box(slide1, f"挙式日: {order_data['wedding_date']}",
                 SLIDE_WIDTH_PX/2, info_y, 'Meiryo UI', 13, center=True, color_rgb=(60, 60, 60))

    info_y += 26
    base_font_display = sim_data.get('baseFont', 'Alex Brush')
    template_names = {'holy': '教会式①', 'happy': '教会式②', 'promise': '人前式', 'custom': 'カスタム'}
    template_info = template_names.get(sim_data.get('template', 'holy'), '教会式①')

    # 各要素のフォントをチェック（ベースと異なる場合のみ表示）
    font_differences = []
    title_font = sim_data.get('titleFont', '')
    body_font = sim_data.get('bodyFont', '')
    date_font = sim_data.get('dateFont', '')
    name_font = sim_data.get('nameFont', '')

    if title_font and title_font != base_font_display:
        font_differences.append(f"タイトル:{title_font}")
    if body_font and body_font != base_font_display:
        font_differences.append(f"本文:{body_font}")
    if date_font and date_font != base_font_display:
        font_differences.append(f"日付:{date_font}")
    if name_font and name_font != base_font_display:
        font_differences.append(f"名前:{name_font}")

    # フォント表示文字列を構築
    if font_differences:
        font_detail = f"フォント: {base_font_display}（{' / '.join(font_differences)}）"
    else:
        font_detail = f"フォント: {base_font_display}"

    add_text_box(slide1, f"{font_detail}　　本文: {template_info}",
                 SLIDE_WIDTH_PX/2, info_y, 'Meiryo UI', 11, center=True, color_rgb=(100, 100, 100))

    # ========== 2ページ目: 背景cutout + テキスト ==========
    slide2 = prs.slides.add_slide(blank_layout)

    base_font = sim_data.get('baseFont', 'Alex Brush')
    if base_font not in FONT_MAP:
        base_font = 'Alex Brush'

    def get_element_font(element):
        font_key = element + 'Font'
        font = sim_data.get(font_key) or base_font
        if font not in FONT_MAP:
            font = base_font
        return font

    FONT_SCALE = SLIDE_WIDTH_PX / 500
    board_size_pct = sim_data.get('boardSize', 130) / 100
    # テキスト色: 'burn'=茶色, 'white'=白
    text_color_name = sim_data.get('textColor', 'burn')
    text_color = (255, 255, 255) if text_color_name == 'white' else (42, 24, 16)

    # 背景画像をダウンロード（透明度を保持）
    cutout_path = download_image(cutout_urls.get(background, cutout_urls['product']), temp_dir, preserve_transparency=True)

    # 板形状検出（ダウンロード後に実行）
    board_shape = {'type': 'square', 'aspectRatio': 1.0, 'stretchFactor': 0, 'bounds': {'minX': 0, 'maxX': 1, 'minY': 0, 'maxY': 1}}
    layout_adj = {}  # レイアウト調整値

    # 板の実際の位置とサイズ（テキスト配置用）
    actual_board_left = 0
    actual_board_top = 0
    actual_board_width = SLIDE_WIDTH_PX
    actual_board_height = SLIDE_HEIGHT_PX

    if cutout_path and os.path.exists(cutout_path):
        try:
            bg_img = Image.open(cutout_path)
            bg_width, bg_height = bg_img.size

            # 板形状検出（境界情報を含む）
            board_shape = detect_board_shape(bg_img)
            bounds = board_shape.get('bounds', {'minX': 0, 'maxX': 1, 'minY': 0, 'maxY': 1})

            # 形状に応じたレイアウト調整を取得
            if board_shape['type'] == 'portrait':
                layout_adj = get_portrait_layout_adjustments(board_shape['stretchFactor'])
                print(f"[Layout] Portrait adjustments applied: titleY={layout_adj['titleY']}, nameY={layout_adj['nameY']}")
            elif board_shape['type'] == 'landscape':
                layout_adj = get_landscape_layout_adjustments()
                print(f"[Layout] Landscape adjustments applied")

            base_scale = 0.95
            img_ratio = bg_width / bg_height
            max_width = SLIDE_WIDTH_PX * base_scale
            # 4:3コンテンツエリア内に収める
            max_height = SIMULATOR_ASPECT_HEIGHT * base_scale

            if bg_width / max_width > bg_height / max_height:
                base_width = max_width
                base_height = max_width / img_ratio
            else:
                base_height = max_height
                base_width = max_height * img_ratio

            draw_width = base_width * board_size_pct
            draw_height = base_height * board_size_pct
            # boardX, boardY は中心位置（0.5 = 中央）- 4:3エリア内で計算
            board_x_pct = sim_data.get('boardX', 0.5)
            board_y_pct = sim_data.get('boardY', 0.5)
            img_x = SLIDE_WIDTH_PX * board_x_pct - draw_width / 2
            # Y座標: 4:3→1:1変換（board_y_pctは0-1の範囲）
            img_y = Y_OFFSET + board_y_pct * SIMULATOR_ASPECT_HEIGHT - draw_height / 2

            # 板の実際の位置とサイズを計算（透過部分を除く）
            actual_board_left = img_x + draw_width * bounds['minX']
            actual_board_top = img_y + draw_height * bounds['minY']
            actual_board_width = draw_width * (bounds['maxX'] - bounds['minX'])
            actual_board_height = draw_height * (bounds['maxY'] - bounds['minY'])

            print(f"[Board] Image: ({img_x:.0f}, {img_y:.0f}) {draw_width:.0f}x{draw_height:.0f}")
            print(f"[Board] Actual: ({actual_board_left:.0f}, {actual_board_top:.0f}) {actual_board_width:.0f}x{actual_board_height:.0f}")

            slide2.shapes.add_picture(
                cutout_path,
                px_to_emu(img_x), px_to_emu(img_y),
                px_to_emu(draw_width), px_to_emu(draw_height)
            )
        except Exception as e:
            print(f"[WARN] Background image error: {e}")

    # レイアウト調整を適用（sim_dataの値がデフォルトの場合のみ上書き）
    def get_adjusted_value(key, default, is_size=False):
        """sim_dataの値またはレイアウト調整値を取得"""
        sim_val = sim_data.get(key)
        # sim_dataに明示的に設定されている場合はそれを使用
        if sim_val is not None and sim_val != default:
            return sim_val
        # レイアウト調整がある場合はそれを使用
        if key in layout_adj:
            return layout_adj[key]
        return default

    # ========== テキスト配置（板の境界を基準） ==========
    # テキスト位置は板の実際の境界（透過部分を除く）を基準に計算
    # titleX=50, titleY=22 → 板の左から50%、上から22%の位置

    # タイトル
    title_font = get_element_font('title')
    title_key = sim_data.get('title', 'wedding')
    title_text = sim_data.get('customTitle', '') if title_key == 'custom' else TITLES.get(title_key, 'Wedding Certificate')
    title_x = actual_board_left + actual_board_width * (sim_data.get('titleX', 50) / 100)
    title_y_pct = get_adjusted_value('titleY', 22)
    title_y = actual_board_top + actual_board_height * (title_y_pct / 100)
    title_size_pct = get_adjusted_value('titleSize', 100)
    title_size = 24 * (title_size_pct / 100) * FONT_SCALE
    title_box = add_text_box(slide2, title_text, title_x, title_y, title_font, title_size, center=True, color_rgb=text_color)
    title_bottom = title_y + title_size

    # 本文
    body_font = get_element_font('body')
    template_key = sim_data.get('template', 'holy')
    body_text = sim_data.get('customText', '') if template_key == 'custom' else TEMPLATES.get(template_key, '')
    body_x = actual_board_left + actual_board_width * (sim_data.get('bodyX', 50) / 100)
    body_y_pct = get_adjusted_value('bodyY', 32)
    body_y_base = actual_board_top + actual_board_height * (body_y_pct / 100)
    min_gap = 30 * FONT_SCALE
    body_y = max(body_y_base, title_bottom + min_gap)
    body_size_pct = get_adjusted_value('bodySize', 115)
    body_size = 11 * (body_size_pct / 100) * FONT_SCALE
    body_line_height = sim_data.get('bodyLineHeight', 1.4)
    if body_text:
        add_multiline_text_box(slide2, body_text, body_x, body_y, body_font, body_size,
                               line_height=body_line_height, color_rgb=text_color)

    # 日付
    date_font = get_element_font('date')
    date_format_key = sim_data.get('dateFormat', 'western')
    formatted_date = sim_data.get('customDate', '') if date_format_key == 'custom' else format_date(order_data['wedding_date'], date_format_key)
    date_x = actual_board_left + actual_board_width * (sim_data.get('dateX', 50) / 100)
    date_y_pct = get_adjusted_value('dateY', 60)
    date_y = actual_board_top + actual_board_height * (date_y_pct / 100)
    date_size_pct = get_adjusted_value('dateSize', 85)
    date_size = 18 * (date_size_pct / 100) * FONT_SCALE
    if formatted_date:
        add_text_box(slide2, formatted_date, date_x, date_y, date_font, date_size, center=True, color_rgb=text_color)

    # 名前
    name_font = get_element_font('name')
    name_x_pct = sim_data.get('nameX', 50) / 100
    name_y_pct = get_adjusted_value('nameY', 74)
    name_size_pct = get_adjusted_value('nameSize', 90)
    name_size = 32 * (name_size_pct / 100) * FONT_SCALE
    name_center_x = actual_board_left + actual_board_width * name_x_pct
    name_y = actual_board_top + actual_board_height * (name_y_pct / 100)

    groom_width_approx = len(groom) * name_size * 0.6
    amp_width_approx = name_size * 2
    bride_width_approx = len(bride) * name_size * 0.6
    total_width_approx = groom_width_approx + amp_width_approx + bride_width_approx

    if groom:
        groom_x = name_center_x - total_width_approx / 2 + groom_width_approx / 2
        add_text_box(slide2, groom, groom_x, name_y, name_font, name_size, center=True, color_rgb=text_color)

    amp_x = name_center_x
    add_text_box(slide2, "&", amp_x, name_y, name_font, name_size, center=True, color_rgb=text_color)

    if bride:
        bride_x = name_center_x + total_width_approx / 2 - bride_width_approx / 2
        add_text_box(slide2, bride, bride_x, name_y, name_font, name_size, center=True, color_rgb=text_color)

    # ツリー
    if sim_data.get('showTree', False):
        tree_type = sim_data.get('treeType', 'simple')
        # 板形状に応じたツリー位置調整
        default_tree_x = layout_adj.get('treeX', 0.75) if layout_adj else 0.75
        default_tree_y = layout_adj.get('treeY', 0.65) if layout_adj else 0.65
        default_tree_size = layout_adj.get('treeSize', 80) if layout_adj else 80

        # sim_dataに明示的な値がある場合はそちらを優先
        tree_x_pct = sim_data.get('treeX') if sim_data.get('treeX') is not None else default_tree_x
        tree_y_pct = sim_data.get('treeY') if sim_data.get('treeY') is not None else default_tree_y
        tree_size_pct = (sim_data.get('treeSize') if sim_data.get('treeSize') is not None else default_tree_size) / 100

        tree_url = f"{TREE_IMAGES_URL}/tree-{tree_type}.png"
        print(f"[TREE] Downloading: {tree_url} (shape: {board_shape['type']}, pos: {tree_x_pct:.2f}, {tree_y_pct:.2f})")

        try:
            # 直接ダウンロードしてPNG形式を明示的に保持
            response = requests.get(tree_url, timeout=30)
            if response.status_code == 200:
                tree_img = Image.open(BytesIO(response.content))
                print(f"[TREE] Original: {tree_img.size}, mode={tree_img.mode}")

                # リサイズ
                max_size = 800
                if max(tree_img.size) > max_size:
                    ratio = max_size / max(tree_img.size)
                    new_size = (int(tree_img.size[0] * ratio), int(tree_img.size[1] * ratio))
                    tree_img = tree_img.resize(new_size, Image.LANCZOS)

                # RGBA確保
                if tree_img.mode != 'RGBA':
                    tree_img = tree_img.convert('RGBA')
                print(f"[TREE] After resize: {tree_img.size}, mode={tree_img.mode}")

                # 透明度チェック
                alpha = tree_img.split()[3]
                alpha_extrema = alpha.getextrema()
                print(f"[TREE] Alpha range: {alpha_extrema}")

                # BytesIOにPNGとして保存
                tree_buffer = BytesIO()
                tree_img.save(tree_buffer, 'PNG', optimize=False)
                tree_buffer.seek(0)
                print(f"[TREE] PNG buffer size: {len(tree_buffer.getvalue()) / 1024:.1f}KB")

                tree_width, tree_height = tree_img.size
                # シミュレーターと同じ計算: img.width * (treeSize/100) * 0.08
                # 元のツリー画像サイズは3000x3000、シミュレーターのキャンバスは500px幅
                TREE_ORIGINAL_SIZE = 3000  # ツリー画像の元サイズ
                TREE_SCALE_FACTOR = 0.08   # シミュレーターの係数
                SIMULATOR_WIDTH = 500      # シミュレーターのキャンバス幅
                draw_tree_width = TREE_ORIGINAL_SIZE * tree_size_pct * TREE_SCALE_FACTOR * (SLIDE_WIDTH_PX / SIMULATOR_WIDTH)
                draw_tree_height = draw_tree_width * (tree_height / tree_width)
                # ツリー位置も板基準（テキストと同じ）
                # tree_x_pct, tree_y_pct は板の相対位置（0-1）
                tree_x = actual_board_left + actual_board_width * tree_x_pct - draw_tree_width / 2
                tree_y = actual_board_top + actual_board_height * tree_y_pct - draw_tree_height / 2
                print(f"[TREE] Size: {draw_tree_width:.0f}x{draw_tree_height:.0f}px at ({tree_x:.0f}, {tree_y:.0f}) (board-relative)")

                # BytesIOから直接追加
                slide2.shapes.add_picture(
                    tree_buffer,
                    px_to_emu(tree_x), px_to_emu(tree_y),
                    px_to_emu(draw_tree_width), px_to_emu(draw_tree_height)
                )
                print(f"[TREE] Added to slide at ({tree_x:.0f}, {tree_y:.0f})")
        except Exception as e:
            print(f"[WARN] Tree image error: {e}")
            import traceback
            traceback.print_exc()

    # ========== 3-6ページ目: cutout画像 ==========
    cutout_labels = [
        ('product', '水引表'),
        ('product_back', '水引裏'),
        ('noclear', '無塗装表'),
        ('noclear_back', '無塗装裏'),
    ]

    for key, label in cutout_labels:
        slide = prs.slides.add_slide(blank_layout)
        cutout_path = download_image(cutout_urls[key], temp_dir, preserve_transparency=True)

        if cutout_path and os.path.exists(cutout_path):
            try:
                img = Image.open(cutout_path)
                img_width, img_height = img.size
                max_size = min(SLIDE_WIDTH_PX, SLIDE_HEIGHT_PX) * 0.9
                scale = min(max_size / img_width, max_size / img_height)
                draw_width = img_width * scale
                draw_height = img_height * scale
                x = (SLIDE_WIDTH_PX - draw_width) / 2
                y = (SLIDE_HEIGHT_PX - draw_height) / 2

                slide.shapes.add_picture(
                    cutout_path,
                    px_to_emu(x), px_to_emu(y),
                    px_to_emu(draw_width), px_to_emu(draw_height)
                )
            except Exception as e:
                print(f"[WARN] {label} image error: {e}")
                add_text_box(slide, f"No Image: {label}", SLIDE_WIDTH_PX/2, SLIDE_HEIGHT_PX/2,
                             'Meiryo UI', 24, center=True, color_rgb=(150, 150, 150))
        else:
            add_text_box(slide, f"No Image: {label}", SLIDE_WIDTH_PX/2, SLIDE_HEIGHT_PX/2,
                         'Meiryo UI', 24, center=True, color_rgb=(150, 150, 150))

    # 保存
    output_path = os.path.join(temp_dir, f"order_{order_data['order_id']}.pptx")
    prs.save(output_path)
    print(f"[Canva] PowerPoint created: {output_path}")
    return output_path


def create_pdf(order_data, temp_dir):
    """PDFを作成（透明度対応版）"""
    print(f"[Canva] Creating PDF for order #{order_data['order_id']}...")

    sim_data = order_data['sim_data']
    sim_image = order_data['sim_image']
    groom = sim_data.get('groomName', '')
    bride = sim_data.get('brideName', '')

    # PDFページサイズ（ポイント単位、1000x1000px相当）
    PAGE_SIZE = (1000, 1000)

    output_path = os.path.join(temp_dir, f"order_{order_data['order_id']}.pdf")
    c = pdf_canvas.Canvas(output_path, pagesize=PAGE_SIZE)

    # cutout画像のURLを取得
    background = sim_data.get('background', 'product')
    cutout_urls = {
        'product': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'product'),
        'product_back': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'product_back'),
        'noclear': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'noclear'),
        'noclear_back': find_cutout_url(order_data['board_name'], order_data['board_number'], order_data['board_size'], 'noclear_back'),
    }

    base_font = sim_data.get('baseFont', 'Alex Brush')
    # テキスト色: 'burn'=茶色, 'white'=白
    text_color_name = sim_data.get('textColor', 'burn')
    text_color_hex = (1.0, 1.0, 1.0) if text_color_name == 'white' else (42/255, 24/255, 16/255)

    # ========== 1ページ目: シミュレーション画像 + 注文情報 ==========
    if sim_image:
        try:
            if sim_image.startswith('http'):
                response = requests.get(sim_image, timeout=30)
                img = Image.open(BytesIO(response.content))
            else:
                if ',' in sim_image:
                    sim_image = sim_image.split(',')[1]
                img_data = base64.b64decode(sim_image)
                img = Image.open(BytesIO(img_data))

            # 一時保存
            sim_path = os.path.join(temp_dir, 'sim_image.png')
            img.save(sim_path, 'PNG')

            # PPTXと同じロジックで大きく配置
            img_width, img_height = img.size
            available_height = PAGE_SIZE[1] * 0.80  # 80%の高さを使用
            available_width = PAGE_SIZE[0] * 0.95   # 95%の幅を使用
            scale = min(available_width / img_width, available_height / img_height)
            draw_width = img_width * scale
            draw_height = img_height * scale

            # 中央上寄せ配置（PDFは左下原点）
            x = (PAGE_SIZE[0] - draw_width) / 2
            y = PAGE_SIZE[1] - draw_height - 20  # 上から20px下

            c.drawImage(sim_path, x, y, width=draw_width, height=draw_height)
        except Exception as e:
            print(f"[WARN] Sim image error: {e}")

    # 注文情報テキスト
    c.setFillColorRGB(0.16, 0.16, 0.16)
    c.setFont("Helvetica", 14)
    info_y = 150
    c.drawCentredString(PAGE_SIZE[0]/2, info_y, f"Order #{order_data['order_id']} - {order_data['board_name']} No.{order_data['board_number']}")
    c.setFont("Helvetica", 11)
    c.drawCentredString(PAGE_SIZE[0]/2, info_y - 25, f"Groom: {groom}  Bride: {bride}")
    c.drawCentredString(PAGE_SIZE[0]/2, info_y - 50, f"Date: {order_data['wedding_date']}")

    # フォント詳細表示（個別フォント設定がある場合）
    font_display_name = FONT_DISPLAY_MAP.get(base_font, base_font)
    font_differences = []
    title_font_pdf = sim_data.get('titleFont')
    body_font_pdf = sim_data.get('bodyFont')
    date_font_pdf = sim_data.get('dateFont')
    name_font_pdf = sim_data.get('nameFont')

    if title_font_pdf and title_font_pdf != base_font:
        font_differences.append(f"Title:{title_font_pdf}")
    if body_font_pdf and body_font_pdf != base_font:
        font_differences.append(f"Body:{body_font_pdf}")
    if date_font_pdf and date_font_pdf != base_font:
        font_differences.append(f"Date:{date_font_pdf}")
    if name_font_pdf and name_font_pdf != base_font:
        font_differences.append(f"Name:{name_font_pdf}")

    if font_differences:
        font_detail = f"Font: {font_display_name} ({' / '.join(font_differences)})"
    else:
        font_detail = f"Font: {font_display_name}"
    c.drawCentredString(PAGE_SIZE[0]/2, info_y - 75, font_detail)

    c.showPage()

    # ========== 2ページ目: 背景cutout + テキスト + ツリー ==========
    # PPTXと同じスケール係数を使用
    FONT_SCALE = PAGE_SIZE[0] / 500  # = 2.0
    board_size_pct = sim_data.get('boardSize', 130) / 100

    # 板の境界情報（デフォルト値、画像解析後に更新）
    board_bounds = {'minX': 0, 'maxX': 1, 'minY': 0, 'maxY': 1}
    board_center_offset_x = 0
    board_center_offset_y = 0
    draw_w = 0
    draw_h = 0

    # 背景cutout画像
    cutout_url = cutout_urls.get(background, cutout_urls['product'])
    try:
        response = requests.get(cutout_url, timeout=30)
        if response.status_code == 200:
            cutout_img = Image.open(BytesIO(response.content))
            if cutout_img.mode != 'RGBA':
                cutout_img = cutout_img.convert('RGBA')

            # 板の実際の境界を検出（透過部分を除く）
            img_w, img_h = cutout_img.size
            alpha = cutout_img.split()[3]
            alpha_data = alpha.load()

            # 非透明ピクセルの境界を検出
            min_x, max_x, min_y, max_y = img_w, 0, img_h, 0
            for py in range(img_h):
                for px in range(img_w):
                    if alpha_data[px, py] > 10:  # 閾値10以上を板として認識
                        min_x = min(min_x, px)
                        max_x = max(max_x, px)
                        min_y = min(min_y, py)
                        max_y = max(max_y, py)

            # 正規化（0-1の範囲に）
            if max_x > min_x and max_y > min_y:
                board_bounds = {
                    'minX': min_x / img_w,
                    'maxX': max_x / img_w,
                    'minY': min_y / img_h,
                    'maxY': max_y / img_h
                }
                print(f"[PDF] Board bounds: X={board_bounds['minX']:.2f}-{board_bounds['maxX']:.2f}, Y={board_bounds['minY']:.2f}-{board_bounds['maxY']:.2f}")

            cutout_path = os.path.join(temp_dir, 'cutout_bg.png')
            cutout_img.save(cutout_path, 'PNG')

            # シミュレーターと同じロジックで描画サイズを計算
            base_scale = 0.95
            img_ratio = img_w / img_h
            max_width = PAGE_SIZE[0] * base_scale
            max_height = PAGE_SIZE[1] * base_scale

            # アスペクト比を維持してフィット
            if img_w / max_width > img_h / max_height:
                base_width = max_width
                base_height = max_width / img_ratio
            else:
                base_height = max_height
                base_width = max_height * img_ratio

            # boardSizeを適用
            draw_w = base_width * board_size_pct
            draw_h = base_height * board_size_pct
            img_x = (PAGE_SIZE[0] - draw_w) / 2
            img_y = (PAGE_SIZE[1] - draw_h) / 2

            # 実際の板の中心オフセットを計算（画像中心からのずれ）
            board_center_offset_x = ((board_bounds['minX'] + board_bounds['maxX']) / 2 - 0.5) * draw_w
            board_center_offset_y = ((board_bounds['minY'] + board_bounds['maxY']) / 2 - 0.5) * draw_h
            print(f"[PDF] Board center offset: ({board_center_offset_x:.1f}, {board_center_offset_y:.1f})")

            c.drawImage(cutout_path, img_x, img_y, width=draw_w, height=draw_h, mask='auto')
            print(f"[PDF] Cutout: draw={draw_w:.0f}x{draw_h:.0f}, boardSize={board_size_pct*100:.0f}%")
    except Exception as e:
        print(f"[WARN] Cutout error: {e}")

    # 実際の板のサイズ（透過部分を除く）
    actual_board_w = draw_w * (board_bounds['maxX'] - board_bounds['minX'])
    actual_board_h = draw_h * (board_bounds['maxY'] - board_bounds['minY'])

    # 板の中心位置（ページ上での実際の位置）
    board_center_x = PAGE_SIZE[0] / 2 + board_center_offset_x
    board_center_y = PAGE_SIZE[1] / 2 - board_center_offset_y  # Y軸反転

    # テキスト要素（板の実際の境界を基準に配置）
    c.setFillColorRGB(*text_color_hex)

    # タイトル - 板の境界を基準に配置
    title_key = sim_data.get('title', 'wedding')
    title_text = sim_data.get('customTitle', '') if title_key == 'custom' else TITLES.get(title_key, 'Wedding Certificate')
    # titleY=22 は板の上から22%の位置
    title_x = board_center_x + actual_board_w * ((sim_data.get('titleX', 50) - 50) / 100)
    title_y_pct = sim_data.get('titleY', 22) / 100
    title_y = board_center_y + actual_board_h / 2 - actual_board_h * title_y_pct
    title_size = 24 * (sim_data.get('titleSize', 100) / 100) * FONT_SCALE
    c.setFont("Helvetica", title_size)
    c.drawCentredString(title_x, title_y, title_text)

    # 本文
    template_key = sim_data.get('template', 'holy')
    body_text = sim_data.get('customText', '') if template_key == 'custom' else TEMPLATES.get(template_key, '')
    body_x = board_center_x + actual_board_w * ((sim_data.get('bodyX', 50) - 50) / 100)
    body_y_pct = sim_data.get('bodyY', 32) / 100
    body_y_base = board_center_y + actual_board_h / 2 - actual_board_h * body_y_pct
    body_size = 11 * (sim_data.get('bodySize', 115) / 100) * FONT_SCALE

    # 重なり防止: タイトル下端から最低30px確保（PDF座標系ではY軸上向き）
    title_bottom = title_y - title_size  # タイトルの下端
    min_gap = 30 * FONT_SCALE  # 最小間隔
    body_y = min(body_y_base, title_bottom - min_gap)  # 本文はタイトル下に配置
    body_line_height = sim_data.get('bodyLineHeight', 1.4)
    c.setFont("Helvetica", body_size)
    for i, line in enumerate(body_text.split('\n')):
        c.drawCentredString(body_x, body_y - i * body_size * body_line_height, line.strip())

    # 日付
    date_format_key = sim_data.get('dateFormat', 'western')
    formatted_date = sim_data.get('customDate', '') if date_format_key == 'custom' else format_date(order_data['wedding_date'], date_format_key)
    date_x = board_center_x + actual_board_w * ((sim_data.get('dateX', 50) - 50) / 100)
    date_y_pct = sim_data.get('dateY', 60) / 100
    date_y = board_center_y + actual_board_h / 2 - actual_board_h * date_y_pct
    date_size = 18 * (sim_data.get('dateSize', 85) / 100) * FONT_SCALE
    c.setFont("Helvetica", date_size)
    c.drawCentredString(date_x, date_y, formatted_date)

    # 名前
    name_x = board_center_x + actual_board_w * ((sim_data.get('nameX', 50) - 50) / 100)
    name_y_pct = sim_data.get('nameY', 74) / 100
    name_y = board_center_y + actual_board_h / 2 - actual_board_h * name_y_pct
    name_size = 32 * (sim_data.get('nameSize', 90) / 100) * FONT_SCALE
    c.setFont("Helvetica", name_size)
    name_text = f"{groom}  &  {bride}"
    c.drawCentredString(name_x, name_y, name_text)

    # ツリー画像（透明度保持）- 板サイズを基準に計算
    if sim_data.get('showTree', False):
        tree_type = sim_data.get('treeType', 'simple')
        tree_x_pct = sim_data.get('treeX', 0.75)
        tree_y_pct = sim_data.get('treeY', 0.65)
        tree_size_pct = sim_data.get('treeSize', 80) / 100

        tree_url = f"{TREE_IMAGES_URL}/tree-{tree_type}.png"
        print(f"[PDF] Downloading tree: {tree_url}")
        print(f"[PDF] Tree params: x={tree_x_pct}, y={tree_y_pct}, size={tree_size_pct*100}%")
        print(f"[PDF] Board size: {actual_board_w:.0f}x{actual_board_h:.0f}")

        try:
            response = requests.get(tree_url, timeout=30)
            if response.status_code == 200:
                tree_img = Image.open(BytesIO(response.content))
                print(f"[PDF] Tree original: {tree_img.size}, mode={tree_img.mode}")

                # RGBA確保
                if tree_img.mode != 'RGBA':
                    tree_img = tree_img.convert('RGBA')

                # アルファチェック
                alpha = tree_img.split()[3]
                print(f"[PDF] Tree alpha range: {alpha.getextrema()}")

                tree_path = os.path.join(temp_dir, 'tree.png')
                tree_img.save(tree_path, 'PNG')

                # ===== 板サイズ基準のツリーサイズ計算 =====
                # 視覚的調整: 0.35=小, 0.7=中, 1.0=適正, 1.2=大
                tree_base_ratio = 1.0  # 視覚的に確認済み
                draw_w = actual_board_w * tree_size_pct * tree_base_ratio

                # アスペクト比を維持
                tree_w, tree_h = tree_img.size
                aspect_ratio = tree_h / tree_w
                draw_h = draw_w * aspect_ratio

                print(f"[PDF] Tree draw size: {draw_w:.0f}x{draw_h:.0f} ({draw_w/actual_board_w*100:.1f}% of board)")

                # ===== 板を基準にした位置計算 =====
                # treeX, treeY は板内での相対位置（0-1）
                # 板の左上を(0,0)、右下を(1,1)とする
                # 板の実際の左端・上端を計算
                board_left = board_center_x - actual_board_w / 2
                board_top = board_center_y + actual_board_h / 2  # PDF座標系（Y軸上向き）

                # ツリー中心位置（板内座標）
                tree_center_x = board_left + actual_board_w * tree_x_pct
                tree_center_y = board_top - actual_board_h * tree_y_pct  # Y軸反転

                # 左下座標に変換
                x = tree_center_x - draw_w / 2
                y = tree_center_y - draw_h / 2

                print(f"[PDF] Tree position: center=({tree_center_x:.0f}, {tree_center_y:.0f}), corner=({x:.0f}, {y:.0f})")

                # mask='auto' でPNG透明度を自動適用
                c.drawImage(tree_path, x, y, width=draw_w, height=draw_h, mask='auto')
                print(f"[PDF] Tree: size={tree_size_pct*100:.0f}%, draw={draw_w:.0f}x{draw_h:.0f}, pos=({x:.0f}, {y:.0f})")
        except Exception as e:
            print(f"[WARN] Tree error: {e}")
            import traceback
            traceback.print_exc()

    c.showPage()

    # ========== 3-6ページ目: cutout画像 ==========
    cutout_labels = [
        ('product', 'Front Clear'),
        ('product_back', 'Back Clear'),
        ('noclear', 'Front No-Clear'),
        ('noclear_back', 'Back No-Clear'),
    ]

    for key, label in cutout_labels:
        try:
            response = requests.get(cutout_urls[key], timeout=30)
            if response.status_code == 200:
                img = Image.open(BytesIO(response.content))
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')

                max_size = 800
                if max(img.size) > max_size:
                    ratio = max_size / max(img.size)
                    new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
                    img = img.resize(new_size, Image.LANCZOS)

                img_path = os.path.join(temp_dir, f'{key}.png')
                img.save(img_path, 'PNG')

                img_w, img_h = img.size
                max_draw = PAGE_SIZE[0] * 0.9
                scale = min(max_draw / img_w, max_draw / img_h)
                draw_w = img_w * scale
                draw_h = img_h * scale
                x = (PAGE_SIZE[0] - draw_w) / 2
                y = (PAGE_SIZE[1] - draw_h) / 2

                c.drawImage(img_path, x, y, width=draw_w, height=draw_h, mask='auto')
        except Exception as e:
            print(f"[WARN] {label} error: {e}")
            c.setFont("Helvetica", 20)
            c.drawCentredString(PAGE_SIZE[0]/2, PAGE_SIZE[1]/2, f"No Image: {label}")

        c.showPage()

    c.save()
    print(f"[Canva] PDF created: {output_path}")
    return output_path


def send_discord_notification(order_data, design, webhook_url, order=None):
    """Discord通知送信（新規注文 + Canvaリンク統合版）"""
    if not webhook_url:
        return False

    edit_url = design.get("urls", {}).get("edit_url", "")
    groom = order_data['sim_data'].get('groomName', '')
    bride = order_data['sim_data'].get('brideName', '')

    # 注文情報を取得（orderオブジェクトがある場合）
    customer_name = ""
    order_total = ""
    payment_method = ""
    customer_phone = ""
    customer_email = ""

    if order:
        billing = order.get('billing', {})
        customer_name = f"{billing.get('last_name', '')} {billing.get('first_name', '')}"
        order_total = order.get('total', '0')
        payment_method = order.get('payment_method_title', '')
        customer_phone = billing.get('phone', '')
        customer_email = billing.get('email', '')

    embed = {
        "title": f"🛒 新規注文 #{order_data['order_id']}",
        "color": 0x06C755,  # LINE緑
        "fields": [
            {"name": "👤 お客様", "value": customer_name or f"{groom} & {bride}", "inline": True},
            {"name": "💰 金額", "value": f"¥{int(float(order_total)):,}" if order_total else "N/A", "inline": True},
            {"name": "💳 支払方法", "value": payment_method or "N/A", "inline": True},
            {"name": "📦 商品", "value": f"{order_data['board_name']} No.{order_data['board_number']}", "inline": False},
            {"name": "📅 挙式日", "value": order_data['wedding_date'], "inline": False},
            {"name": "📞 連絡先", "value": f"TEL: {customer_phone}\nEmail: {customer_email}" if customer_phone else "N/A", "inline": False},
            {"name": "🎨 Canva", "value": f"[デザインを編集する]({edit_url})", "inline": False},
        ],
        "footer": {"text": "i.tategu 自動化システム（Railway）"},
    }

    # 商品画像があればサムネイルに設定
    if order:
        for item in order.get('line_items', []):
            image_url = item.get('image', {}).get('src', '')
            if image_url:
                embed['thumbnail'] = {'url': image_url}
                break

    response = requests.post(
        webhook_url,
        json={"embeds": [embed]},
        headers={"Content-Type": "application/json"}
    )

    return response.status_code == 204


# 発送管理チャンネルID
DISCORD_SHIPPING_CHANNEL_ID = "1463452139312644240"

def send_shipping_notification(order_data, order, bot_token):
    """発送管理チャンネルに住所情報を投稿"""
    if not order or not bot_token:
        print("[Shipping] Missing order or bot_token")
        return False

    billing = order.get('billing', {})
    shipping = order.get('shipping', {})

    # 発送先情報（shipping優先、なければbilling）
    postcode = shipping.get('postcode') or billing.get('postcode', '')
    state = shipping.get('state') or billing.get('state', '')
    city = shipping.get('city') or billing.get('city', '')
    address1 = shipping.get('address_1') or billing.get('address_1', '')
    address2 = shipping.get('address_2') or billing.get('address_2', '')

    # 都道府県コード変換（簡易版）
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

    full_address = f"{state_name}{city}{address1}"
    if address2:
        full_address += f" {address2}"

    customer_name = f"{billing.get('last_name', '')} {billing.get('first_name', '')}"
    customer_phone = billing.get('phone', '')
    order_total = order.get('total', '0')
    payment_method = order.get('payment_method_title', '')

    # 商品名
    products = []
    for item in order.get('line_items', []):
        products.append(item.get('name', ''))
    product_names = ', '.join(products) if products else order_data.get('board_name', '')

    embed = {
        "title": f"🟡 未発送 | #{order_data['order_id']} {customer_name} 様",
        "color": 0xFFD700,  # 黄色
        "fields": [
            {"name": "📞 電話", "value": customer_phone or "N/A", "inline": True},
            {"name": "📦 商品", "value": product_names, "inline": True},
            {"name": "💰 金額", "value": f"¥{int(float(order_total)):,} / {payment_method}", "inline": True},
            {"name": "〒 住所", "value": f"{postcode} {full_address}" if postcode else full_address, "inline": False},
        ],
    }

    # Discord Bot APIで送信
    url = f"https://discord.com/api/v10/channels/{DISCORD_SHIPPING_CHANNEL_ID}/messages"
    headers = {
        "Authorization": f"Bot {bot_token}",
        "Content-Type": "application/json"
    }

    try:
        response = requests.post(url, json={"embeds": [embed]}, headers=headers)
        if response.status_code in [200, 201]:
            print(f"[Shipping] Notification sent for order #{order_data['order_id']}")
            return True
        else:
            print(f"[Shipping] Failed: {response.status_code} - {response.text}")
            return False
    except Exception as e:
        print(f"[Shipping] Error: {e}")
        return False


def clear_processing_lock(order_id, wc_url, wc_key, wc_secret):
    """処理中ロックを解除（失敗時用）"""
    url = f"{wc_url}/wp-json/wc/v3/orders/{order_id}?consumer_key={wc_key}&consumer_secret={wc_secret}"
    try:
        requests.put(url, json={"meta_data": [{"key": "canva_processing", "value": ""}]})
        print(f"[Canva] Lock released for order #{order_id}")
    except Exception as e:
        print(f"[WARN] Failed to release lock: {e}")


def send_discord_error_notification(order_id, error_message, webhook_url):
    """Discord エラー通知送信"""
    if not webhook_url:
        return False

    embed = {
        "title": f"⚠️ Canva処理エラー #{order_id}",
        "color": 15158332,  # 赤色
        "fields": [
            {"name": "エラー内容", "value": str(error_message)[:500]},
        ],
        "footer": {"text": "i.tategu Canva自動化（Railway）"},
    }

    try:
        response = requests.post(
            webhook_url,
            json={"embeds": [embed]},
            headers={"Content-Type": "application/json"}
        )
        return response.status_code == 204
    except:
        return False


def mark_order_processed(order_id, design_url, wc_url, wc_key, wc_secret):
    """注文を処理済みにマーク + ステータスを「デザイン打ち合わせ中」に変更"""
    # WooCommerce REST APIはクエリパラメータで認証
    url = f"{wc_url}/wp-json/wc/v3/orders/{order_id}?consumer_key={wc_key}&consumer_secret={wc_secret}"

    data = {
        "status": "designing",  # デザイン打ち合わせ中 (short slug for WP 20-char limit)
        "meta_data": [
            {"key": "canva_automation_done", "value": "1"},
            {"key": "canva_processing", "value": ""},  # ロック解除
            {"key": "canva_design_url", "value": design_url},
        ]
    }

    try:
        response = requests.put(url, json=data)
        print(f"[WC Update] Status: {response.status_code}")
        if response.status_code != 200:
            print(f"[WC Update] Error: {response.text[:500]}")
            return False
        print(f"[WC Update] Order #{order_id} marked as processed, status → designing")
        return True
    except Exception as e:
        print(f"[WC Update] Exception: {e}")
        return False


def process_order(order_id, config):
    """
    注文を処理してCanvaデザインを作成

    config: {
        'wc_url': WooCommerce URL,
        'wc_key': Consumer Key,
        'wc_secret': Consumer Secret,
        'canva_access_token': Canva Access Token,
        'canva_refresh_token': Canva Refresh Token,
        'discord_webhook': Discord Webhook URL,
    }
    """
    print(f"\n{'='*50}")
    print(f"[Canva] Processing order #{order_id}")
    print(f"{'='*50}")

    # 注文取得
    order = get_order_from_woocommerce(order_id, config['wc_url'], config['wc_key'], config['wc_secret'])
    if not order:
        print(f"[ERROR] Order not found: {order_id}")
        return False

    # 既に処理済み or 処理中かチェック
    meta = {m['key']: m['value'] for m in order.get('meta_data', [])}
    if meta.get('canva_automation_done') or meta.get('canva_processing'):
        print(f"[SKIP] Already processed or in progress: {order_id}")
        return False

    # 即座に処理中フラグを立てる（重複防止ロック）
    lock_acquired = False
    try:
        lock_url = f"{config['wc_url']}/wp-json/wc/v3/orders/{order_id}?consumer_key={config['wc_key']}&consumer_secret={config['wc_secret']}"
        requests.put(lock_url, json={"meta_data": [{"key": "canva_processing", "value": "1"}]})
        lock_acquired = True
        print(f"[Canva] Lock acquired for order #{order_id}")
    except Exception as e:
        print(f"[WARN] Lock failed: {e}")

    # ロック取得後の処理（失敗時は必ずロック解除）
    success = False
    error_message = None

    try:
        # 注文データ解析
        order_data = parse_order_data(order)

        if not order_data['board_name']:
            error_message = "No board info"
            print(f"[SKIP] No board info: {order_id}")
            return False

        print(f"[Canva] Product: {order_data['board_name']} No.{order_data['board_number']}")
        print(f"[Canva] Names: {order_data['sim_data'].get('groomName', '')} & {order_data['sim_data'].get('brideName', '')}")

        # 一時ディレクトリ作成
        with tempfile.TemporaryDirectory() as temp_dir:
            # PPTX作成（Canva互換性・フォント対応が安定）
            pptx_path = create_pptx(order_data, temp_dir)

            # Canvaタイトル
            groom = order_data['sim_data'].get('groomName', '')
            bride = order_data['sim_data'].get('brideName', '')
            canva_title = f"注文{order_id} {order_data['board_name']} No.{order_data['board_number']} {groom}＆{bride} {order_data['wedding_date']}"

            # Canvaインポート
            print(f"[Canva] Importing PPTX to Canva...")
            design, error_info = import_to_canva(
                pptx_path, canva_title,
                config['canva_access_token'],
                config['canva_refresh_token']
            )

            if not design:
                error_message = f"Canva import failed: {error_info}"
                print(f"[ERROR] {error_message}")
                return False

            design_id = design.get('id')
            print(f"[Canva] Design ID: {design_id}")

            # Discord通知（注文情報+Canvaリンク統合版）
            print(f"[Canva] Sending Discord notification...")
            send_discord_notification(order_data, design, config['discord_webhook'], order)
            print(f"[Canva] Discord notification sent")

            # 発送管理チャンネルへ住所情報通知
            bot_token = config.get('discord_bot_token', '')
            if bot_token:
                print(f"[Canva] Sending shipping notification...")
                send_shipping_notification(order_data, order, bot_token)
            else:
                print(f"[WARN] No bot token, skipping shipping notification")

            # 処理済みマーク（ここでロックも解除される）
            design_url = design.get('urls', {}).get('edit_url', '')
            mark_order_processed(order_id, design_url, config['wc_url'], config['wc_key'], config['wc_secret'])
            print(f"[Canva] Order marked as processed")
            success = True

        print(f"[Canva] Order #{order_id} completed!")
        return True

    except Exception as e:
        error_message = str(e)
        print(f"[ERROR] Processing failed: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 失敗時はロック解除 & エラー通知
        if lock_acquired and not success:
            clear_processing_lock(order_id, config['wc_url'], config['wc_key'], config['wc_secret'])
            if error_message and config.get('discord_webhook'):
                send_discord_error_notification(order_id, error_message, config['discord_webhook'])
