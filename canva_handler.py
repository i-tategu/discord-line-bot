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

# 設定（環境変数から取得 - 遅延読み込み）
def get_canva_client_id():
    return os.getenv("CANVA_CLIENT_ID", "OC-AZvUVtxGhbOD")

def get_canva_client_secret():
    return os.getenv("CANVA_CLIENT_SECRET", "")

# サーバー上のcutout画像URL
CUTOUT_BASE_URL = "https://i-tategu-shop.com/wp-content/themes/i-tategu/assets/images/cutouts"
TREE_IMAGES_URL = "https://i-tategu-shop.com/wp-content/themes/i-tategu/assets/images"

# スライドサイズ
SLIDE_WIDTH_PX = 1000
SLIDE_HEIGHT_PX = 1000
EMU_PER_PX = 914400 / 96

# フォントマッピング
FONT_MAP = {
    'Alex Brush': 'Alex Brush',
    'Great Vibes': 'Great Vibes',
    'Pinyon Script': 'Pinyon Script',
    'Sacramento': 'Sacramento',
    'Dancing Script': 'Dancing Script',
    'Parisienne': 'Parisienne',
    'Allura': 'Allura',
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
    """Canvaトークンをリフレッシュ"""
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
        print(f"[Canva Token] Refresh successful!")
        return {
            'access_token': tokens.get('access_token'),
            'refresh_token': tokens.get('refresh_token', refresh_token)
        }

    print(f"[Canva Token] Refresh failed: {response.text[:500]}")
    return None


def import_to_canva(pptx_path, title, access_token, refresh_token, retry=False):
    """CanvaにPowerPointをインポート（エラー情報も返す）"""
    url = "https://api.canva.com/rest/v1/imports"
    title_base64 = base64.b64encode(title.encode("utf-8")).decode("utf-8")

    metadata = {
        "title_base64": title_base64,
        "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream",
        "Import-Metadata": json.dumps(metadata)
    }

    with open(pptx_path, "rb") as f:
        response = requests.post(url, headers=headers, data=f)

    print(f"[Canva Import] Status: {response.status_code}")

    # 401エラー時はトークンリフレッシュ
    if response.status_code == 401 and not retry:
        print("[INFO] Token expired, refreshing...")
        new_tokens = refresh_canva_token(refresh_token)
        if new_tokens:
            return import_to_canva(pptx_path, title, new_tokens['access_token'], new_tokens['refresh_token'], retry=True)
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
    font_info = sim_data.get('baseFont', 'Alex Brush')
    template_names = {'holy': '教会式①', 'happy': '教会式②', 'promise': '人前式', 'custom': 'カスタム'}
    template_info = template_names.get(sim_data.get('template', 'holy'), '教会式①')
    add_text_box(slide1, f"フォント: {font_info}　　本文: {template_info}",
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
    text_color = (42, 24, 16)

    # 背景画像をダウンロード（透明度を保持）
    cutout_path = download_image(cutout_urls.get(background, cutout_urls['product']), temp_dir, preserve_transparency=True)

    if cutout_path and os.path.exists(cutout_path):
        try:
            bg_img = Image.open(cutout_path)
            bg_width, bg_height = bg_img.size
            base_scale = 0.95
            img_ratio = bg_width / bg_height
            max_width = SLIDE_WIDTH_PX * base_scale
            max_height = SLIDE_HEIGHT_PX * base_scale

            if bg_width / max_width > bg_height / max_height:
                base_width = max_width
                base_height = max_width / img_ratio
            else:
                base_height = max_height
                base_width = max_height * img_ratio

            draw_width = base_width * board_size_pct
            draw_height = base_height * board_size_pct
            img_x = (SLIDE_WIDTH_PX - draw_width) / 2
            img_y = (SLIDE_HEIGHT_PX - draw_height) / 2

            slide2.shapes.add_picture(
                cutout_path,
                px_to_emu(img_x), px_to_emu(img_y),
                px_to_emu(draw_width), px_to_emu(draw_height)
            )
        except Exception as e:
            print(f"[WARN] Background image error: {e}")

    # タイトル
    title_font = get_element_font('title')
    title_key = sim_data.get('title', 'wedding')
    title_text = sim_data.get('customTitle', '') if title_key == 'custom' else TITLES.get(title_key, 'Wedding Certificate')
    title_x = SLIDE_WIDTH_PX * (sim_data.get('titleX', 50) / 100)
    title_y = SLIDE_HEIGHT_PX * (sim_data.get('titleY', 22) / 100)
    title_size = 24 * (sim_data.get('titleSize', 100) / 100) * FONT_SCALE
    add_text_box(slide2, title_text, title_x, title_y, title_font, title_size, center=True, color_rgb=text_color)

    # 本文
    body_font = get_element_font('body')
    template_key = sim_data.get('template', 'holy')
    body_text = sim_data.get('customText', '') if template_key == 'custom' else TEMPLATES.get(template_key, '')
    body_x = SLIDE_WIDTH_PX * (sim_data.get('bodyX', 50) / 100)
    body_y = SLIDE_HEIGHT_PX * (sim_data.get('bodyY', 32) / 100)
    body_size = 11 * (sim_data.get('bodySize', 115) / 100) * FONT_SCALE
    body_line_height = sim_data.get('bodyLineHeight', 1.4)
    if body_text:
        add_multiline_text_box(slide2, body_text, body_x, body_y, body_font, body_size,
                               line_height=body_line_height, color_rgb=text_color)

    # 日付
    date_font = get_element_font('date')
    date_format_key = sim_data.get('dateFormat', 'western')
    formatted_date = sim_data.get('customDate', '') if date_format_key == 'custom' else format_date(order_data['wedding_date'], date_format_key)
    date_x = SLIDE_WIDTH_PX * (sim_data.get('dateX', 50) / 100)
    date_y = SLIDE_HEIGHT_PX * (sim_data.get('dateY', 60) / 100)
    date_size = 18 * (sim_data.get('dateSize', 85) / 100) * FONT_SCALE
    if formatted_date:
        add_text_box(slide2, formatted_date, date_x, date_y, date_font, date_size, center=True, color_rgb=text_color)

    # 名前
    name_font = get_element_font('name')
    name_x_pct = sim_data.get('nameX', 50) / 100
    name_y_pct = sim_data.get('nameY', 74) / 100
    name_size = 32 * (sim_data.get('nameSize', 90) / 100) * FONT_SCALE
    name_center_x = SLIDE_WIDTH_PX * name_x_pct
    name_y = SLIDE_HEIGHT_PX * name_y_pct

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
        tree_x_pct = sim_data.get('treeX', 0.75)
        tree_y_pct = sim_data.get('treeY', 0.65)
        tree_size_pct = sim_data.get('treeSize', 80) / 100

        tree_url = f"{TREE_IMAGES_URL}/tree-{tree_type}.png"
        tree_path = download_image(tree_url, temp_dir)

        if tree_path and os.path.exists(tree_path):
            try:
                tree_img = Image.open(tree_path)
                tree_width, tree_height = tree_img.size
                base_tree_size = min(SLIDE_WIDTH_PX, SLIDE_HEIGHT_PX) * 0.3
                draw_tree_width = base_tree_size * tree_size_pct
                draw_tree_height = draw_tree_width * (tree_height / tree_width)
                tree_x = SLIDE_WIDTH_PX * tree_x_pct - draw_tree_width / 2
                tree_y = SLIDE_HEIGHT_PX * tree_y_pct - draw_tree_height / 2

                slide2.shapes.add_picture(
                    tree_path,
                    px_to_emu(tree_x), px_to_emu(tree_y),
                    px_to_emu(draw_tree_width), px_to_emu(draw_tree_height)
                )
            except Exception as e:
                print(f"[WARN] Tree image error: {e}")

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


def send_discord_notification(order_data, design, webhook_url):
    """Discord通知送信"""
    if not webhook_url:
        return False

    edit_url = design.get("urls", {}).get("edit_url", "")
    groom = order_data['sim_data'].get('groomName', '')
    bride = order_data['sim_data'].get('brideName', '')

    embed = {
        "title": f"Canvaデザイン準備完了 #{order_data['order_id']}",
        "color": 5814783,
        "fields": [
            {"name": "商品", "value": f"{order_data['board_name']} No.{order_data['board_number']}", "inline": True},
            {"name": "お客様", "value": f"{groom} & {bride}", "inline": True},
            {"name": "日付", "value": order_data['wedding_date'], "inline": True},
            {"name": "Canva編集", "value": f"[デザインを編集する]({edit_url})"},
        ],
        "footer": {"text": "i.tategu Canva自動化（Railway）"},
    }

    response = requests.post(
        webhook_url,
        json={"embeds": [embed]},
        headers={"Content-Type": "application/json"}
    )

    return response.status_code == 204


def mark_order_processed(order_id, design_url, wc_url, wc_key, wc_secret):
    """注文を処理済みにマーク"""
    url = f"{wc_url}/wp-json/wc/v3/orders/{order_id}"

    data = {
        "meta_data": [
            {"key": "_canva_automation_done", "value": "1"},
            {"key": "_canva_design_url", "value": design_url},
        ]
    }

    response = requests.put(url, json=data, auth=(wc_key, wc_secret))
    return response.status_code == 200


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

    # 既に処理済みかチェック
    meta = {m['key']: m['value'] for m in order.get('meta_data', [])}
    if meta.get('_canva_automation_done'):
        print(f"[SKIP] Already processed: {order_id}")
        return False

    # 注文データ解析
    order_data = parse_order_data(order)

    if not order_data['board_name']:
        print(f"[SKIP] No board info: {order_id}")
        return False

    print(f"[Canva] Product: {order_data['board_name']} No.{order_data['board_number']}")
    print(f"[Canva] Names: {order_data['sim_data'].get('groomName', '')} & {order_data['sim_data'].get('brideName', '')}")

    # 一時ディレクトリ作成
    with tempfile.TemporaryDirectory() as temp_dir:
        # PowerPoint作成
        pptx_path = create_pptx(order_data, temp_dir)

        # Canvaタイトル
        groom = order_data['sim_data'].get('groomName', '')
        bride = order_data['sim_data'].get('brideName', '')
        canva_title = f"注文{order_id} {order_data['board_name']} No.{order_data['board_number']} {groom}＆{bride} {order_data['wedding_date']}"

        # Canvaインポート
        print(f"[Canva] Importing to Canva...")
        design, new_tokens = import_to_canva(
            pptx_path, canva_title,
            config['canva_access_token'],
            config['canva_refresh_token']
        )

        if not design:
            print(f"[ERROR] Canva import failed")
            return False

        design_id = design.get('id')
        print(f"[Canva] Design ID: {design_id}")

        # Discord通知
        print(f"[Canva] Sending Discord notification...")
        send_discord_notification(order_data, design, config['discord_webhook'])
        print(f"[Canva] Discord notification sent")

        # 処理済みマーク
        design_url = design.get('urls', {}).get('edit_url', '')
        mark_order_processed(order_id, design_url, config['wc_url'], config['wc_key'], config['wc_secret'])
        print(f"[Canva] Order marked as processed")

    print(f"[Canva] Order #{order_id} completed!")
    return True
