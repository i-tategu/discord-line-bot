"""
来店客アプリ用：板写真の背景透過処理（rembg）＋余白付与＋リサイズ。

- 透過：rembg(u2net) で背景除去
- 余白：板が見切れないよう周囲に透過マージンを付与（macOS Vision版と同等の挙動）
- モデル：DATA_DIR（永続ボリューム）配下にキャッシュし再DLを回避

rembg/onnxruntime は requirements.txt に追加が必要。未インストールでも本モジュールの
import 自体は失敗しないよう、重い依存は関数内で遅延 import する。
"""
import io
import os

from PIL import Image

# u2net モデルを永続ボリュームにキャッシュ（再起動ごとの再ダウンロード回避）
_DATA_DIR = os.environ.get("DATA_DIR", "/tmp")
os.environ.setdefault("U2NET_HOME", os.path.join(_DATA_DIR, ".u2net"))

_session = None


def _get_session():
    global _session
    if _session is None:
        from rembg import new_session  # 遅延 import
        _session = new_session("u2net")
    return _session


def process_board_image(img_bytes, max_side=1200, pad_frac=0.14):
    """1枚の板写真を透過＋余白付与した PNG バイト列にして返す。"""
    from rembg import remove  # 遅延 import

    cut = remove(img_bytes, session=_get_session())
    img = Image.open(io.BytesIO(cut)).convert("RGBA")

    # 板（不透明領域）にトリミング
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)

    # 長辺を max_side に縮小
    longest = max(img.width, img.height)
    if longest > max_side:
        s = max_side / float(longest)
        img = img.resize((max(1, int(img.width * s)), max(1, int(img.height * s))), Image.LANCZOS)

    # 透過の余白を周囲に付与（見切れ防止）
    pad = int(round(max(img.width, img.height) * pad_frac))
    canvas = Image.new("RGBA", (img.width + 2 * pad, img.height + 2 * pad), (0, 0, 0, 0))
    canvas.paste(img, (pad, pad), img)

    buf = io.BytesIO()
    canvas.save(buf, "PNG", optimize=True)
    return buf.getvalue()
