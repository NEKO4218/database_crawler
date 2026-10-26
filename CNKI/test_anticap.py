import base64
import sys
from io import BytesIO

sys.stdout.reconfigure(encoding='utf-8')

import AntiCAP
from PIL import Image, ImageDraw, ImageFont

IMG = r'captcha_debug/点选测试.png'
TARGETS = ['打', '出', '服']


def to_b64(path):
    with open(path, 'rb') as f:
        return base64.b64encode(f.read()).decode()


def crop_b64(path, box):
    img = Image.open(path).convert('RGB')
    x1, y1, x2, y2 = box
    crop = img.crop((x1, y1, x2, y2))
    buf = BytesIO()
    crop.save(buf, format='PNG')
    return base64.b64encode(buf.getvalue()).decode()


def render_order_image(chars, cell=72, gap=12):
    width = len(chars) * cell + (len(chars) + 1) * gap
    height = cell + gap * 2
    img = Image.new('RGB', (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype('msyh.ttc', 48)
    except Exception:
        font = ImageFont.load_default()
    for i, ch in enumerate(chars):
        x0 = gap + i * (cell + gap)
        y0 = gap
        draw.rectangle((x0, y0, x0 + cell, y0 + cell), outline=(200, 200, 200), width=1)
        bbox = draw.textbbox((0, 0), ch, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((x0 + (cell - tw) / 2, y0 + (cell - th) / 2), ch, fill=(20, 20, 20), font=font)
    out = 'captcha_debug/order_hint.png'
    img.save(out)
    return out


if __name__ == '__main__':
    Atc = AntiCAP.Handler(show_banner=False)
    img_b64 = to_b64(IMG)

    print('=== Detection_Text ===')
    dets = Atc.Detection_Text(img_base64=img_b64)
    for item in dets:
        ch = Atc.OCR(img_base64=crop_b64(IMG, item['box']))
        print(item['box'], repr(ch))

    print('\n=== ClickText_Order ===')
    order_path = render_order_image(TARGETS)
    order_b64 = to_b64(order_path)
    boxes = Atc.ClickText_Order(order_img_base64=order_b64, target_img_base64=img_b64)
    print('boxes', boxes)
    centers = [((b[0] + b[2]) / 2, (b[1] + b[3]) / 2) for b in boxes if any(b)]
    print('centers', centers)
