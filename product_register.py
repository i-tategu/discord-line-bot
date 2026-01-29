"""
商品登録モジュール — Railway版
スマホ・PCのブラウザから商品をWooCommerceに登録する
"""
import os
import base64
import requests
from functools import wraps
from flask import request, jsonify, session, redirect, url_for, render_template_string

# ===== 環境変数 =====
def get_wp_url():
    return os.environ.get("WP_URL", "https://i-tategu-shop.com")

def get_wp_user():
    return os.environ.get("WP_USER", "")

def get_wp_password():
    return os.environ.get("WP_APP_PASSWORD", "")

def get_register_password():
    return os.environ.get("PRODUCT_REGISTER_PASSWORD", "")

# ===== 木材データ =====
WOOD_INFO = {
    "ケヤキ": {"meaning": "幸運・長寿・健康", "type": "広葉樹", "story": "「際立って優れた木」を意味する名を持つ日本の銘木。清水寺の舞台を1000年以上支え続けるその強さは、ふたりの人生を末永く見守る力そのものです。", "recommend": "堂々と、力強く歩んでいきたいカップルに。"},
    "サクラ": {"meaning": "精神美・優美な女性", "type": "広葉樹", "story": "日本人が最も愛する花の木。「精神の美しさ」を表すその花言葉は、心の美しさを誓い合うおふたりにぴったりの樹種です。", "recommend": "内面の美しさを大切にするカップルに。"},
    "エンジュ": {"meaning": "幸福・慕情・上品", "type": "広葉樹", "story": "「木」偏に「鬼」と書き、邪気を払う霊木。「幸福を招く木」として親しまれています。", "recommend": "幸福を招き入れたいカップルに。"},
    "ヤマモモ": {"meaning": "一途な愛・ただひとりを愛する", "type": "広葉樹", "story": "花言葉の「ただひとりを愛する」は、結婚の誓いそのもの。", "recommend": "「ただひとりを愛する」という誓いを形にしたいカップルに。"},
    "カイヅカイブキ": {"meaning": "永遠の美・変わらぬ心", "type": "針葉樹", "story": "常緑の針葉樹として「変わらぬ心」の象徴。", "recommend": "永遠に変わらない想いを誓いたいカップルに。"},
    "トチ": {"meaning": "博愛・贅沢", "type": "広葉樹", "story": "縄文時代から日本人の命を支えてきた「食べられる木」。", "recommend": "すべてを包み込む大きな愛を誓いたいカップルに。"},
    "ホウ": {"meaning": "誠実・友情", "type": "広葉樹", "story": "朴の木は日本最大級の葉を持つ落葉樹。人の暮らしに寄り添い続けてきた「誠実」な木です。", "recommend": "飾らない誠実さを大切にするカップルに。"},
    "タモ": {"meaning": "揺るぎない絆", "type": "広葉樹", "story": "野球のバットに使われるほどの強靭さと弾力性。困難にも折れない夫婦の絆そのものです。", "recommend": "どんな困難も二人で乗り越えたいカップルに。"},
    "ヤマザクラ": {"meaning": "純潔・高尚", "type": "広葉樹", "story": "山野に自生する日本古来の桜。", "recommend": "凛とした美しさを大切にするカップルに。"},
    "クス": {"meaning": "忍耐・活力", "type": "広葉樹", "story": "樹齢1000年を超える巨木も珍しくない、生命力あふれる常緑樹。", "recommend": "力強く前向きに歩みたいカップルに。"},
    "クリ": {"meaning": "真心・満足", "type": "広葉樹", "story": "縄文時代から人々の暮らしを支えてきた実りの木。", "recommend": "実り多い人生を願うカップルに。"},
    "イチョウ": {"meaning": "長寿・荘厳", "type": "針葉樹", "story": "2億年以上前から姿を変えずに生き続ける「生きた化石」。", "recommend": "悠久の時を超える愛を誓いたいカップルに。"},
    "ヒノキ": {"meaning": "不滅・不変の心", "type": "針葉樹", "story": "伊勢神宮の式年遷宮にも使われる神聖な木。", "recommend": "清らかな心で新たな門出を迎えたいカップルに。"},
    "スギ": {"meaning": "深みのある心・高貴", "type": "針葉樹", "story": "日本固有の木であり、「真っ直ぐな木」が語源。", "recommend": "真っ直ぐな心で歩みたいカップルに。"},
    "ウォールナット": {"meaning": "勝利・子孫繁栄", "type": "広葉樹", "story": "世界三大銘木のひとつ。深く落ち着いたチョコレート色は、成熟した大人の愛を表現します。", "recommend": "大人の落ち着いた雰囲気を好むカップルに。"},
    "ブラックチェリー": {"meaning": "真実の愛", "type": "広葉樹", "story": "「真実の愛」という花言葉を持つ、結婚証明書にこれ以上ないほどふさわしい樹種です。", "recommend": "年月とともに深まる愛を形にしたいカップルに。"},
    "オーク": {"meaning": "永遠・おもてなし", "type": "広葉樹", "story": "ヨーロッパでは「森の王」と呼ばれ、力強さと寛容さの象徴。", "recommend": "格調高い雰囲気を求めるカップルに。"},
    "クワ": {"meaning": "知恵・彼女の全てが好き", "type": "広葉樹", "story": "「彼女の全てが好き」という花言葉が結婚の想いと重なります。", "recommend": "パートナーの全てを愛するカップルに。"},
    "カエデ": {"meaning": "調和・大切な思い出", "type": "広葉樹", "story": "秋に美しく色づく楓は「大切な思い出」の象徴。", "recommend": "思い出を大切にするカップルに。"},
    "メープル": {"meaning": "調和・大切な思い出", "type": "広葉樹", "story": "カエデの英名。ハードメープルはバイオリンにも使われる音響木材です。", "recommend": "音楽やアートを愛するカップルに。"},
    "キハダ": {"meaning": "回想・癒やし", "type": "広葉樹", "story": "漢方薬「黄檗」の原料として古来より人を癒してきた木。", "recommend": "穏やかで温かい家庭を築きたいカップルに。"},
    "イチイ": {"meaning": "高潔・家庭円満", "type": "針葉樹", "story": "正一位の「一位」に通じることから、高貴な木として珍重されてきました。", "recommend": "高い志を持って新生活を始めたいカップルに。"},
    "アカシア": {"meaning": "真実の愛・友情", "type": "広葉樹", "story": "「真実の愛」を花言葉に持つアカシア。硬く丈夫な性質は、強い愛の絆を象徴しています。", "recommend": "揺るぎない真実の愛を誓いたいカップルに。"},
}

WOOD_ROMAJI = {
    "ケヤキ": "Zelkova", "トチ": "Horse Chestnut", "サクラ": "Cherry",
    "ヤマザクラ": "Wild Cherry", "クス": "Camphor", "クリ": "Chestnut",
    "イチョウ": "Ginkgo", "ヒノキ": "Cypress", "スギ": "Cedar",
    "ウォールナット": "Walnut", "ブラックチェリー": "Black Cherry",
    "オーク": "Oak", "タモ": "Ash", "クワ": "Mulberry",
    "カエデ": "Maple", "メープル": "Maple", "エンジュ": "Pagoda Tree",
    "ホウ": "Magnolia", "キハダ": "Amur Cork", "イチイ": "Yew",
    "アカシア": "Acacia", "ヤマモモ": "Bayberry", "カイヅカイブキ": "Kaizuka Juniper",
}

PRICE_MAP = {"A": 30000, "B": 34000, "C": 38000, "D": 42000}

# ===== ヘルパー関数 =====
def _wp_auth_headers():
    credentials = f"{get_wp_user()}:{get_wp_password()}"
    token = base64.b64encode(credentials.encode()).decode()
    return {
        'Authorization': f'Basic {token}',
        'User-Agent': 'i-tategu-product-register/1.0'
    }

def calculate_guest_category(width, height):
    area = width * height
    if area < 100000:
        return "S（~40名）"
    elif area < 180000:
        return "M（40~60名）"
    elif area < 280000:
        return "L（60~80名）"
    else:
        return "XL（80名~）"

def calculate_recommended_guests(width_mm, height_mm):
    usable_area = (width_mm * height_mm) * 0.40
    stamp_area = 20 ** 2
    max_stamps = int(usable_area / stamp_area)
    if max_stamps < 40:
        return "~40", max_stamps
    elif max_stamps < 60:
        return "40~60", max_stamps
    elif max_stamps < 80:
        return "60~80", max_stamps
    else:
        return "80~", max_stamps

def get_wc_term_id(term_name, taxonomy="categories"):
    endpoint = "products/categories" if taxonomy == "categories" else "products/tags"
    url = f"{get_wp_url()}/wp-json/wc/v3/{endpoint}"
    headers = _wp_auth_headers()
    try:
        res = requests.get(url, headers=headers, params={"search": term_name, "per_page": 100}, timeout=15)
        if res.status_code == 200:
            for item in res.json():
                if item["name"] == term_name:
                    return item["id"]
        # 存在しない場合は作成
        create_res = requests.post(url, headers=headers, json={"name": term_name}, timeout=15)
        if create_res.status_code == 201:
            return create_res.json()["id"]
    except Exception as e:
        print(f"[WC Term] Error: {e}")
    return None

def get_next_number(wood_type):
    """WordPress上の既存商品から次の番号を取得"""
    url = f"{get_wp_url()}/wp-json/wc/v3/products"
    headers = _wp_auth_headers()
    max_num = 0
    try:
        res = requests.get(url, headers=headers, params={
            "search": wood_type, "per_page": 100, "status": "any"
        }, timeout=15)
        if res.status_code == 200:
            import re
            for p in res.json():
                match = re.search(r'No\.(\d+)', p.get('name', ''))
                if match and wood_type in p.get('name', ''):
                    max_num = max(max_num, int(match.group(1)))
    except Exception:
        pass
    return max_num + 1

def generate_description(wood_type, width, height, number, thickness=20):
    """商品説明HTMLを自動生成"""
    wood_romaji = WOOD_ROMAJI.get(wood_type, "Natural Wood")
    info = WOOD_INFO.get(wood_type, {})
    meaning = info.get("meaning", "自然の恵み・温もり")
    story = info.get("story", "")
    recommend = info.get("recommend", "")
    guests_text, _ = calculate_recommended_guests(width, height)

    story_block = f'<p style="line-height:1.8;color:#444;">{story}</p>' if story else ""
    recommend_block = ""
    if recommend:
        recommend_block = f"""
<div style="background:#f0ede8;padding:20px;border-radius:8px;margin-top:20px;">
    <p style="font-weight:600;margin-bottom:8px;color:#5D4E37;">&#128161; こんなおふたりに</p>
    <p style="color:#555;line-height:1.7;margin:0;">{recommend}</p>
</div>"""

    return f"""
<p style="text-align:center;margin-bottom:30px;">
    <span style="font-size:1.4em;letter-spacing:0.1em;font-weight:500;">「世界にひとつ」を選ぶ贅沢</span><br>
    <span style="font-size:0.8em;color:#888;font-family:serif;">Authentic Wedding Board - {wood_romaji}</span>
</p>
{story_block}
<p style="line-height:1.8;color:#444;">
    既製品にはない、自然が長い時間をかけて描いた木目と曲線。<br>
    誓いの言葉を記すそのキャンバスは、年を重ねるごとに味わいを増し、<br>
    10年後、20年後の記念日にも、当時の温もりを思い出させてくれるはずです。
</p>
<hr style="border:0;border-top:1px solid #ddd;margin:30px 0;width:30px;margin-left:auto;margin-right:auto;">
<div style="background-color:#f9f9f9;padding:25px;border-radius:4px;">
    <p style="margin:0 0 10px 0;font-size:0.9em;color:#666;">Dataset</p>
    <table style="width:100%;border-collapse:collapse;font-size:0.95em;">
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;width:30%;">樹種</td><td style="padding:8px 0;border-bottom:1px solid #eee;"><strong>{wood_type} ({wood_romaji})</strong></td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;">木言葉</td><td style="padding:8px 0;border-bottom:1px solid #eee;">{meaning}</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;">サイズ</td><td style="padding:8px 0;border-bottom:1px solid #eee;">W{width} × H{height} × D{thickness} mm</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;">推奨人数</td><td style="padding:8px 0;border-bottom:1px solid #eee;">約 {guests_text} 名様前後</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #eee;">No.</td><td style="padding:8px 0;border-bottom:1px solid #eee;">{number}</td></tr>
    </table>
</div>
{recommend_block}
<p style="font-size:0.85em;color:#888;margin-top:20px;text-align:right;">※ 表面は平滑にサンディング加工済みです。</p>
"""

def upload_media(image_data, filename):
    """画像をWordPressメディアライブラリにアップロード"""
    url = f"{get_wp_url()}/wp-json/wp/v2/media"
    headers = _wp_auth_headers()
    content_type = 'image/jpeg' if filename.lower().endswith(('.jpg', '.jpeg')) else 'image/png'
    files = {'file': (filename, image_data, content_type)}
    res = requests.post(url, headers=headers, files=files, timeout=60)
    if res.status_code == 201:
        return res.json()['id']
    raise Exception(f"Media upload failed: {res.status_code} {res.text[:200]}")

def create_product(wood_type, width, height, price, image_ids, number, thickness=20):
    """WooCommerce商品を作成"""
    url = f"{get_wp_url()}/wp-json/wc/v3/products"
    headers = _wp_auth_headers()

    info = WOOD_INFO.get(wood_type, {})
    meaning = info.get("meaning", "自然の恵み・温もり")
    wood_romaji = WOOD_ROMAJI.get(wood_type, "Natural Wood")
    guests_text, _ = calculate_recommended_guests(width, height)

    product_name = f"【一点物】 {wood_type} 一枚板 ({width}x{height}mm) No.{number:02d}"
    description = generate_description(wood_type, width, height, f"{number:02d}", thickness)
    short_desc = f"【世界にひとつ】{wood_type}の一枚板ウェディングボード。木言葉は「{meaning}」。"

    # カテゴリ・タグ
    cat_ids = []
    tag_ids = []
    wc_id = get_wc_term_id(wood_type, "categories")
    if wc_id:
        cat_ids.append({"id": wc_id})
    wt_id = get_wc_term_id(wood_type, "tags")
    if wt_id:
        tag_ids.append({"id": wt_id})
    tree_type = info.get("type", "広葉樹")
    tt_id = get_wc_term_id(tree_type, "tags")
    if tt_id:
        tag_ids.append({"id": tt_id})
    size_cat = calculate_guest_category(width, height)
    sc_id = get_wc_term_id(size_cat, "tags")
    if sc_id:
        tag_ids.append({"id": sc_id})

    product_data = {
        "name": product_name,
        "type": "simple",
        "status": "publish",
        "regular_price": str(price),
        "description": description,
        "short_description": short_desc,
        "images": [{"id": img_id} for img_id in image_ids],
        "categories": cat_ids,
        "tags": tag_ids,
        "meta_data": [
            {"key": "wood_type", "value": wood_type},
            {"key": "board_width", "value": str(width)},
            {"key": "board_height", "value": str(height)},
            {"key": "board_thickness", "value": str(thickness)},
            {"key": "_recommended_guests", "value": guests_text},
        ]
    }

    res = requests.post(url, headers=headers, json=product_data, timeout=60)
    if res.status_code == 201:
        data = res.json()
        return {"id": data["id"], "name": data["name"], "permalink": data["permalink"]}
    raise Exception(f"Product creation failed: {res.status_code} {res.text[:300]}")


# ===== Flask ルート登録 =====
def register_routes(app):
    """Flask appにルートを登録する"""

    @app.route('/product-register', methods=['GET'])
    def product_register_page():
        # パスワード認証チェック
        if session.get('pr_auth') != True:
            return render_template_string(LOGIN_HTML)
        return render_template_string(REGISTER_HTML, wood_types=sorted(WOOD_INFO.keys()))

    @app.route('/product-register/login', methods=['POST'])
    def product_register_login():
        password = request.form.get('password', '')
        if password == get_register_password():
            session['pr_auth'] = True
            return redirect('/product-register')
        return render_template_string(LOGIN_HTML, error="パスワードが違います")

    @app.route('/product-register/api/register', methods=['POST'])
    def product_register_api():
        if session.get('pr_auth') != True:
            return jsonify({"success": False, "message": "認証が必要です"}), 401

        try:
            wood_type = request.form.get('wood_type')
            width = int(request.form.get('width', 0))
            height = int(request.form.get('height', 0))
            thickness = int(request.form.get('thickness', 20))
            price_grade = request.form.get('price_grade', 'A')
            price = PRICE_MAP.get(price_grade, 30000)

            if not wood_type or width <= 0 or height <= 0:
                return jsonify({"success": False, "message": "樹種・サイズを入力してください"})

            # 画像アップロード
            image_ids = []
            labels = ['塗装あり表', '塗装あり裏', '無塗装表', '無塗装裏']
            for i in range(1, 5):
                file = request.files.get(f'image_{i}')
                if file and file.filename:
                    fname = f"{wood_type}_{i}_{file.filename}"
                    img_data = file.read()
                    media_id = upload_media(img_data, fname)
                    image_ids.append(media_id)
                    print(f"[Product Register] Image {labels[i-1]} uploaded (ID: {media_id})")

            if not image_ids:
                return jsonify({"success": False, "message": "画像を1枚以上アップロードしてください"})

            # 次の番号を取得
            number = get_next_number(wood_type)

            # 商品作成
            result = create_product(wood_type, width, height, price, image_ids, number, thickness)
            print(f"[Product Register] Created: {result['name']} (ID: {result['id']})")

            return jsonify({
                "success": True,
                "message": f"{result['name']} を登録しました",
                "product_url": result['permalink'],
                "product_id": result['id']
            })

        except Exception as e:
            print(f"[Product Register] Error: {e}")
            return jsonify({"success": False, "message": str(e)})

    @app.route('/product-register/api/wood-info', methods=['GET'])
    def product_register_wood_info():
        if session.get('pr_auth') != True:
            return jsonify({}), 401
        wood = request.args.get('wood', '')
        info = WOOD_INFO.get(wood, {})
        return jsonify(info)


# ===== HTML テンプレート =====
LOGIN_HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>商品登録 - i.tategu</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f3ef; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
.login-box { background: #fff; padding: 40px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); max-width: 360px; width: 90%; text-align: center; }
.login-box h1 { font-size: 1.3rem; color: #2f4f4f; margin-bottom: 8px; }
.login-box p { font-size: 0.85rem; color: #888; margin-bottom: 24px; }
.login-box input { width: 100%; padding: 12px 16px; border: 1px solid #ddd; border-radius: 8px; font-size: 1rem; margin-bottom: 16px; }
.login-box button { width: 100%; padding: 12px; background: #2f4f4f; color: #fff; border: none; border-radius: 8px; font-size: 1rem; cursor: pointer; }
.login-box button:hover { background: #1a3a3a; }
.error { color: #c44; font-size: 0.85rem; margin-bottom: 12px; }
</style>
</head>
<body>
<div class="login-box">
    <h1>i.tategu 商品登録</h1>
    <p>パスワードを入力してください</p>
    {% if error %}<p class="error">{{ error }}</p>{% endif %}
    <form method="POST" action="/product-register/login">
        <input type="password" name="password" placeholder="パスワード" autofocus required>
        <button type="submit">ログイン</button>
    </form>
</div>
</body>
</html>
"""

REGISTER_HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>商品登録 - i.tategu</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f3ef; }
.header { background: #2f4f4f; color: #fff; padding: 16px 20px; text-align: center; }
.header h1 { font-size: 1.1rem; font-weight: 500; }
.container { max-width: 600px; margin: 0 auto; padding: 20px; }
.card { background: #fff; border-radius: 12px; padding: 24px; margin-bottom: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
.card h2 { font-size: 1rem; color: #2f4f4f; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #b8860b; }
label { display: block; font-size: 0.85rem; color: #666; margin-bottom: 6px; font-weight: 500; }
select, input[type="number"] { width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 1rem; margin-bottom: 16px; background: #fff; }
.size-row { display: flex; gap: 12px; }
.size-row > div { flex: 1; }
.price-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 16px; }
.price-btn { padding: 12px 8px; border: 2px solid #ddd; border-radius: 8px; background: #fff; cursor: pointer; text-align: center; transition: all 0.2s; }
.price-btn.active { border-color: #b8860b; background: #fef9ef; }
.price-btn .grade { font-size: 1.1rem; font-weight: 600; color: #2f4f4f; }
.price-btn .amount { font-size: 0.8rem; color: #888; }
.image-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.image-slot { border: 2px dashed #ddd; border-radius: 8px; padding: 16px; text-align: center; cursor: pointer; transition: all 0.2s; position: relative; min-height: 120px; display: flex; flex-direction: column; align-items: center; justify-content: center; }
.image-slot:hover { border-color: #b8860b; }
.image-slot.has-image { border-style: solid; border-color: #2f4f4f; }
.image-slot input { position: absolute; opacity: 0; width: 100%; height: 100%; top: 0; left: 0; cursor: pointer; }
.image-slot .label { font-size: 0.75rem; color: #888; margin-bottom: 4px; }
.image-slot .icon { font-size: 1.5rem; color: #ccc; }
.image-slot img { max-width: 100%; max-height: 80px; border-radius: 4px; }
.submit-btn { width: 100%; padding: 14px; background: #2f4f4f; color: #fff; border: none; border-radius: 8px; font-size: 1.1rem; cursor: pointer; font-weight: 500; }
.submit-btn:hover { background: #1a3a3a; }
.submit-btn:disabled { background: #999; cursor: not-allowed; }
.result { padding: 16px; border-radius: 8px; margin-top: 16px; display: none; }
.result.success { background: #f0f7f0; color: #2f4f4f; border: 1px solid #2f4f4f; }
.result.error { background: #fff0f0; color: #c44; border: 1px solid #c44; }
.result a { color: #b8860b; font-weight: 500; }
.wood-info { background: #fef9ef; padding: 12px; border-radius: 8px; margin-bottom: 16px; font-size: 0.85rem; color: #5D4E37; display: none; }
.wood-info .meaning { font-weight: 600; color: #b8860b; }
.loading { display: none; text-align: center; padding: 20px; }
.loading.show { display: block; }
.spinner { border: 3px solid #eee; border-top: 3px solid #2f4f4f; border-radius: 50%; width: 30px; height: 30px; animation: spin 0.8s linear infinite; margin: 0 auto 10px; }
@keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<div class="header"><h1>i.tategu 商品登録</h1></div>
<div class="container">
    <form id="registerForm" enctype="multipart/form-data">
    <!-- 樹種 -->
    <div class="card">
        <h2>樹種</h2>
        <label>樹種を選択</label>
        <select name="wood_type" id="woodType" required>
            <option value="">選択してください</option>
            {% for w in wood_types %}
            <option value="{{ w }}">{{ w }}</option>
            {% endfor %}
            <option value="__custom__">その他（手入力）</option>
        </select>
        <input type="text" id="customWood" name="custom_wood" placeholder="樹種名を入力" style="display:none;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:1rem;width:100%;margin-bottom:16px;">
        <div class="wood-info" id="woodInfo">
            <span class="meaning" id="woodMeaning"></span><br>
            <span id="woodStory"></span>
        </div>
    </div>

    <!-- サイズ -->
    <div class="card">
        <h2>サイズ (mm)</h2>
        <div class="size-row">
            <div>
                <label>幅 (W)</label>
                <input type="number" name="width" placeholder="例: 400" required min="100" max="2000">
            </div>
            <div>
                <label>高さ (H)</label>
                <input type="number" name="height" placeholder="例: 600" required min="100" max="2000">
            </div>
            <div>
                <label>厚み (D)</label>
                <input type="number" name="thickness" placeholder="例: 20" min="5" max="200">
            </div>
        </div>
    </div>

    <!-- 価格帯 -->
    <div class="card">
        <h2>価格帯</h2>
        <div class="price-grid">
            <div class="price-btn active" data-grade="A" onclick="selectPrice(this)">
                <div class="grade">A</div>
                <div class="amount">&yen;30,000</div>
            </div>
            <div class="price-btn" data-grade="B" onclick="selectPrice(this)">
                <div class="grade">B</div>
                <div class="amount">&yen;34,000</div>
            </div>
            <div class="price-btn" data-grade="C" onclick="selectPrice(this)">
                <div class="grade">C</div>
                <div class="amount">&yen;38,000</div>
            </div>
            <div class="price-btn" data-grade="D" onclick="selectPrice(this)">
                <div class="grade">D</div>
                <div class="amount">&yen;42,000</div>
            </div>
        </div>
        <input type="hidden" name="price_grade" id="priceGrade" value="A">
    </div>

    <!-- 画像 -->
    <div class="card">
        <h2>画像（4枚）</h2>
        <div class="image-grid">
            <div class="image-slot" id="slot1">
                <div class="label">塗装あり 表</div>
                <div class="icon">+</div>
                <input type="file" name="image_1" accept="image/*" onchange="previewImage(this, 'slot1')">
            </div>
            <div class="image-slot" id="slot2">
                <div class="label">塗装あり 裏</div>
                <div class="icon">+</div>
                <input type="file" name="image_2" accept="image/*" onchange="previewImage(this, 'slot2')">
            </div>
            <div class="image-slot" id="slot3">
                <div class="label">無塗装 表</div>
                <div class="icon">+</div>
                <input type="file" name="image_3" accept="image/*" onchange="previewImage(this, 'slot3')">
            </div>
            <div class="image-slot" id="slot4">
                <div class="label">無塗装 裏</div>
                <div class="icon">+</div>
                <input type="file" name="image_4" accept="image/*" onchange="previewImage(this, 'slot4')">
            </div>
        </div>
    </div>

    <button type="submit" class="submit-btn" id="submitBtn">商品を登録</button>
    </form>

    <div class="loading" id="loading">
        <div class="spinner"></div>
        <p>登録中...</p>
    </div>
    <div class="result" id="result"></div>
</div>

<script>
function selectPrice(el) {
    document.querySelectorAll('.price-btn').forEach(b => b.classList.remove('active'));
    el.classList.add('active');
    document.getElementById('priceGrade').value = el.dataset.grade;
}

function previewImage(input, slotId) {
    const slot = document.getElementById(slotId);
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const existing = slot.querySelector('img');
            if (existing) existing.remove();
            const img = document.createElement('img');
            img.src = e.target.result;
            slot.querySelector('.icon').style.display = 'none';
            slot.appendChild(img);
            slot.classList.add('has-image');
        };
        reader.readAsDataURL(input.files[0]);
    }
}

// 樹種選択時の情報表示
document.getElementById('woodType').addEventListener('change', function() {
    const custom = document.getElementById('customWood');
    const info = document.getElementById('woodInfo');
    if (this.value === '__custom__') {
        custom.style.display = 'block';
        info.style.display = 'none';
    } else {
        custom.style.display = 'none';
        if (this.value) {
            fetch('/product-register/api/wood-info?wood=' + encodeURIComponent(this.value))
                .then(r => r.json())
                .then(data => {
                    if (data.meaning) {
                        document.getElementById('woodMeaning').textContent = '木言葉: ' + data.meaning;
                        document.getElementById('woodStory').textContent = data.story || '';
                        info.style.display = 'block';
                    } else {
                        info.style.display = 'none';
                    }
                });
        } else {
            info.style.display = 'none';
        }
    }
});

// フォーム送信
document.getElementById('registerForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const form = this;
    const btn = document.getElementById('submitBtn');
    const loading = document.getElementById('loading');
    const result = document.getElementById('result');

    // カスタム樹種の処理
    let woodType = form.wood_type.value;
    if (woodType === '__custom__') {
        woodType = form.custom_wood.value;
        if (!woodType) { alert('樹種名を入力してください'); return; }
    }

    const formData = new FormData(form);
    if (woodType !== form.wood_type.value) {
        formData.set('wood_type', woodType);
    }

    btn.disabled = true;
    loading.classList.add('show');
    result.style.display = 'none';

    fetch('/product-register/api/register', {
        method: 'POST',
        body: formData
    })
    .then(r => r.json())
    .then(data => {
        loading.classList.remove('show');
        btn.disabled = false;
        result.style.display = 'block';
        if (data.success) {
            result.className = 'result success';
            let html = data.message;
            if (data.product_url) {
                html += '<br><a href="' + data.product_url + '" target="_blank">商品ページを確認 →</a>';
            }
            result.innerHTML = html;
            form.reset();
            document.querySelectorAll('.image-slot').forEach(s => {
                const img = s.querySelector('img');
                if (img) img.remove();
                s.querySelector('.icon').style.display = '';
                s.classList.remove('has-image');
            });
        } else {
            result.className = 'result error';
            result.textContent = 'エラー: ' + data.message;
        }
    })
    .catch(err => {
        loading.classList.remove('show');
        btn.disabled = false;
        result.style.display = 'block';
        result.className = 'result error';
        result.textContent = '通信エラー: ' + err.message;
    });
});
</script>
</body>
</html>
"""
