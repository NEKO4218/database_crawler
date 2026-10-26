"""CNKI blockPuzzle 滑块验证码独立测试脚本。"""
import argparse
import os
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from cnki_captcha import (
    has_block_puzzle,
    identify_gap_distance,
    is_captcha_success,
    refresh_captcha,
    save_captcha_images,
    solve_block_puzzle,
    wait_for_block_puzzle,
)

DEFAULT_VERIFY_URL = (
    'https://kns.cnki.net/verify/home?captchaType=blockPuzzle'
    '&ident=d3624d'
    '&captchaId=bf99ed0a-db61-4790-ab5c-d1669941b84d'
    '&returnUrl=eyDvWCt4s2-U0foBpgUc98_4QfdbPTouBuXUv1LwPpsKloLgcGerOg%3D%3D'
)
DEFAULT_SAVE_DIR = './captcha_debug'
DEFAULT_MAX_RETRIES = 5
DEFAULT_HOLD_SECONDS = 10


def parse_args():
    parser = argparse.ArgumentParser(description='CNKI blockPuzzle 滑块验证码测试')
    parser.add_argument('--url', default=DEFAULT_VERIFY_URL, help='验证码页面 URL')
    parser.add_argument('--save-dir', default=DEFAULT_SAVE_DIR, help='调试图像保存目录')
    parser.add_argument('--max-retries', type=int, default=DEFAULT_MAX_RETRIES, help='最大重试次数')
    parser.add_argument('--save-only', action='store_true', help='仅保存图片并识别，不拖动滑块')
    parser.add_argument('--hold-seconds', type=int, default=DEFAULT_HOLD_SECONDS, help='结束后浏览器停留秒数')
    return parser.parse_args()


def create_driver():
    chrome_options = Options()
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    return driver


def run_save_only(driver, save_dir):
    wait_for_block_puzzle(driver)
    attempt_dir = os.path.join(save_dir, 'attempt_1')
    hkpath, qkpath = save_captcha_images(driver, attempt_dir)
    gap_x, match_res, distance = identify_gap_distance(hkpath, qkpath)
    result_path = os.path.join(attempt_dir, 'result.txt')
    with open(result_path, 'w', encoding='utf-8') as f:
        f.write(f'gap_x={gap_x}\n')
        f.write(f'match_res={match_res}\n')
        f.write(f'distance={distance}\n')
        f.write(f'hkpath={hkpath}\n')
        f.write(f'qkpath={qkpath}\n')
    print(f'图片已保存到: {attempt_dir}')
    print(f'背景图: {qkpath}')
    print(f'滑块图: {hkpath}')
    return True


def run_full_test(driver, save_dir, max_retries):
    if not has_block_puzzle(driver):
        print('页面未检测到 blockPuzzle 验证码，请检查 URL 是否有效')
        return False

    initial_url = driver.current_url
    for attempt in range(1, max_retries + 1):
        print(f'\n===== 第 {attempt}/{max_retries} 次尝试 =====')
        attempt_dir = os.path.join(save_dir, f'attempt_{attempt}')
        ok = solve_block_puzzle(driver, save_dir=attempt_dir)
        if ok:
            print('验证码处理成功')
            return True
        print('验证码处理失败，准备刷新重试')
        refresh_captcha(driver)
        wait_for_block_puzzle(driver)

    print(f'已重试 {max_retries} 次仍未通过，当前 URL: {driver.current_url}')
    print(f'初始 URL: {initial_url}')
    return is_captcha_success(driver, initial_url)


def main():
    args = parse_args()
    os.makedirs(args.save_dir, exist_ok=True)

    print('参数配置:')
    print(f'  save_dir={args.save_dir}')
    print(f'  save_only={args.save_only}')
    print('  识别方式 = ddddocr 双图 slide_match')
    print('  滑动距离 = 匹配结果 target[0]')

    driver = create_driver()
    try:
        driver.get(args.url)
        time.sleep(3)
        if args.save_only:
            success = run_save_only(driver, args.save_dir)
        else:
            success = run_full_test(driver, args.save_dir, args.max_retries)
        print('测试结果:', '成功' if success else '失败')
        print(f'浏览器将保持打开 {args.hold_seconds} 秒，便于目视检查')
        time.sleep(args.hold_seconds)
    finally:
        driver.quit()


if __name__ == '__main__':
    main()
