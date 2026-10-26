"""测试点选：RapidOCR 锁定 → ddddocr 补缺 → 相似度兜底。"""
import sys

sys.stdout.reconfigure(encoding='utf-8')

from cnki_captcha import (
    identify_text_click_positions,
    preload_click_captcha,
)

if __name__ == '__main__':
    preload_click_captcha()

    image_path = 'captcha_debug/点选测试.png'
    targets = ['打', '出', '服']
    positions = identify_text_click_positions(image_path, targets)
    print('目标:', targets)
    print('坐标:', positions)
