"""CNKI 验证码公共模块：腾讯滑块 tencent-captcha-dy + 知网 blockPuzzle + 文字点选。

文字点选匹配顺序（按目标字逐步锁定坐标）：
1. RapidOCR 精确匹配 → 命中即确定坐标
2. 缺失目标再回退 ddddocr 精确匹配
3. 仍缺失则对剩余备选做相似度匹配

模型加载默认 GPU 优先（ONNX Runtime CUDA），不可用时回退 CPU。
可通过环境变量 OCR_DEVICE=gpu|cpu|auto 强制设备（默认 auto）。
"""
import base64
import atexit
import multiprocessing
import os
import random
import re
import time
from io import BytesIO
from urllib.parse import urljoin

import ddddocr
import requests
from PIL import Image
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

if 'OMP_NUM_THREADS' not in os.environ:
    os.environ['OMP_NUM_THREADS'] = '1'

_slide_ocr = None
_det_ocr = None
_cls_ocr = None
_rapid_ocr = None
_rapid_load_failed = False
_use_gpu = None  # None=未检测；True/False=已判定
_gpu_runtime_ready = False

# 点选图较小，放大可提高 RapidOCR 散字召回（样例上 2 倍效果最好）
RAPIDOCR_UPSCALE = int(os.environ.get('RAPIDOCR_UPSCALE', '2'))


def _prepare_gpu_runtime():
    """把 PyTorch 自带的 CUDA/cuDNN DLL 加入搜索路径，供 onnxruntime-gpu 使用。"""
    global _gpu_runtime_ready
    if _gpu_runtime_ready:
        return True
    try:
        import torch
        torch_lib = os.path.join(os.path.dirname(torch.__file__), 'lib')
        if os.path.isdir(torch_lib):
            if hasattr(os, 'add_dll_directory'):
                os.add_dll_directory(torch_lib)
            os.environ['PATH'] = torch_lib + os.pathsep + os.environ.get('PATH', '')
            print(f'已加入 GPU 运行库路径: {torch_lib}')
        _gpu_runtime_ready = True
        return True
    except Exception as e:
        print('准备 GPU 运行库失败（将尝试系统 CUDA）:', e)
        _gpu_runtime_ready = True  # 不再重复尝试
        return False


def _probe_cuda_usable():
    """实际创建 Session，确认 CUDAExecutionProvider 可用。"""
    _prepare_gpu_runtime()
    try:
        import onnxruntime as ort
        providers = ort.get_available_providers()
        if 'CUDAExecutionProvider' not in providers:
            print('onnxruntime 未提供 CUDAExecutionProvider，可用:', providers)
            return False
        # 用 ddddocr 自带小模型做真实探测
        model = os.path.join(os.path.dirname(ddddocr.__file__), 'common_old.onnx')
        if not os.path.isfile(model):
            model = os.path.join(os.path.dirname(ddddocr.__file__), 'common_det.onnx')
        sess = ort.InferenceSession(
            model,
            providers=[
                ('CUDAExecutionProvider', {'device_id': int(os.environ.get('OCR_GPU_ID', '0'))}),
                'CPUExecutionProvider',
            ],
        )
        active = sess.get_providers()
        del sess
        ok = 'CUDAExecutionProvider' in active
        print(f'ONNX Runtime 探测 providers={active} -> GPU={"可用" if ok else "不可用"}')
        return ok
    except Exception as e:
        print('CUDA 探测失败，回退 CPU:', e)
        return False


def prefer_gpu():
    """是否使用 GPU 推理。OCR_DEVICE=gpu|cpu|auto（默认 auto）。"""
    global _use_gpu
    if _use_gpu is not None:
        return _use_gpu
    choice = (os.environ.get('OCR_DEVICE') or 'auto').strip().lower()
    if choice in ('cpu', '0', 'false', 'no'):
        _use_gpu = False
        print('OCR 设备: CPU（OCR_DEVICE 强制）')
        return _use_gpu
    if choice in ('gpu', 'cuda', '1', 'true', 'yes'):
        _use_gpu = _probe_cuda_usable()
        if not _use_gpu:
            print('OCR_DEVICE=gpu 但 CUDA 不可用，已回退 CPU')
        else:
            print('OCR 设备: GPU（OCR_DEVICE 强制）')
        return _use_gpu
    # auto：有 CUDA 用 GPU
    _use_gpu = _probe_cuda_usable()
    print(f'OCR 设备: {"GPU" if _use_gpu else "CPU"}（auto）')
    return _use_gpu


def _fix_ddddocr_cuda_session(ocr):
    """将 ddddocr 的 ORT Session 切换为 CUDA（新版 ORT 兼容参数）。"""
    if not prefer_gpu():
        return ocr
    graph = getattr(ocr, '_DdddOcr__graph_path', None)
    if not graph or not os.path.isfile(graph):
        return ocr
    try:
        import onnxruntime as ort
        device_id = int(os.environ.get('OCR_GPU_ID', '0'))
        providers = [
            ('CUDAExecutionProvider', {
                'device_id': device_id,
                'arena_extend_strategy': 'kNextPowerOfTwo',
                'gpu_mem_limit': 2 * 1024 * 1024 * 1024,
                'cudnn_conv_algo_search': 'DEFAULT',
                'do_copy_in_default_stream': '1',
            }),
            'CPUExecutionProvider',
        ]
        sess = ort.InferenceSession(graph, providers=providers)
        active = sess.get_providers()
        if 'CUDAExecutionProvider' not in active:
            print(f'ddddocr Session 仍无 CUDA，保持 CPU: {active}')
            return ocr
        ocr._DdddOcr__providers = providers
        ocr._DdddOcr__ort_session = sess
        print(f'ddddocr Session 已切换 GPU: {active}')
    except Exception as e:
        print('ddddocr Session 切换 GPU 失败，沿用 CPU:', e)
    return ocr


def _make_ddddocr(**kwargs):
    """创建 ddddocr 实例；GPU 模式下重建为 CUDA Session。"""
    init_kw = dict(kwargs)
    init_kw.setdefault('show_ad', False)
    init_kw['use_gpu'] = False
    ocr = ddddocr.DdddOcr(**init_kw)
    if getattr(ocr, '_DdddOcr__ort_session', None) is not None:
        _fix_ddddocr_cuda_session(ocr)
    return ocr


def get_slide_ocr():
    global _slide_ocr
    if _slide_ocr is None:
        _slide_ocr = _make_ddddocr(det=False, ocr=False)
        print(f'ddddocr 滑块模型已加载 ({"GPU" if prefer_gpu() else "CPU"})')
    return _slide_ocr


def get_det_ocr():
    global _det_ocr
    if _det_ocr is None:
        _det_ocr = _make_ddddocr(det=True, ocr=False)
        print(f'ddddocr 检测模型已加载 ({"GPU" if prefer_gpu() else "CPU"})')
    return _det_ocr


def get_cls_ocr():
    global _cls_ocr
    if _cls_ocr is None:
        _cls_ocr = _make_ddddocr(det=False, ocr=True)
        print(f'ddddocr 识别模型已加载 ({"GPU" if prefer_gpu() else "CPU"})')
    return _cls_ocr


def get_rapid_ocr():
    """懒加载 RapidOCR（ONNX）；GPU 优先，不可用时返回 None。"""
    global _rapid_ocr, _rapid_load_failed, _use_gpu
    if _rapid_ocr is not None:
        return _rapid_ocr
    if _rapid_load_failed:
        return None
    try:
        from rapidocr_onnxruntime import RapidOCR
        use_cuda = prefer_gpu()
        # GPU 时限制 ORT CPU 线程，避免占满 CPU
        threads = int(os.environ.get('OCR_ORT_THREADS', '2' if use_cuda else '4'))
        kwargs = {
            'det_use_cuda': use_cuda,
            'cls_use_cuda': use_cuda,
            'rec_use_cuda': use_cuda,
            'intra_op_num_threads': threads,
            'inter_op_num_threads': threads,
        }
        try:
            _rapid_ocr = RapidOCR(**kwargs)
        except Exception as e:
            if use_cuda:
                print('RapidOCR GPU 加载失败，回退 CPU:', e)
                _use_gpu = False
                _rapid_ocr = RapidOCR(
                    det_use_cuda=False, cls_use_cuda=False, rec_use_cuda=False,
                    intra_op_num_threads=4, inter_op_num_threads=4,
                )
            else:
                raise
        device = 'GPU' if prefer_gpu() else 'CPU'
        print(f'RapidOCR 已加载 ({device})')
    except Exception as e:
        print('RapidOCR 加载失败，将使用 ddddocr:', e)
        _rapid_load_failed = True
        return None
    return _rapid_ocr


def base64_to_image(base64_string, savepath):
    try:
        imagedata = base64.b64decode(base64_string)
        image = Image.open(BytesIO(imagedata))
        image.save(savepath)
        return True
    except Exception as e:
        print(e)
        return False


def gettrack(distance):
    track = []
    current = 0
    mid = distance * 3 / 5
    t = 0.2
    v = 0

    while current + 3 < distance:
        if current < mid:
            a = 6
        else:
            a = -9
        v0 = v
        v = v0 + a * t
        move = v0 * t + 1 / 2 * a * t * t
        current += round(move)
        track.append(round(move))
    if distance > current:
        track.append(round(distance - current))
    return track


TENCENT_SLIDER_START_X = 22
TENCENT_SLIDE_STEP_SLEEP_MIN = 0.012
TENCENT_SLIDE_STEP_SLEEP_MAX = 0.028
TENCENT_DOWNLOAD_WAIT = 8
TENCENT_RATE_LIMIT_SLEEP = 300
TENCENT_RELOGIN_RETRIES = 3
# 额外固定偏差（px，显示坐标）。拼图相对识别框的内边距会自动按中心对齐计算。
TENCENT_GAP_BIAS = float(os.environ.get('TENCENT_GAP_BIAS', '0') or '0')
GAP_IDENTIFY_TIMEOUT = 90
DOWNLOAD_FILE_SUFFIXES = ('.pdf', '.xls', '.xlsx', '.caj', '.crdownload', '.tmp')
_TENCENT_RATE_KEYS = ('过于频繁', '操作频繁', '次数过多', '稍后再试', '请稍后再', '尝试过多')
_tencent_last_fail = None

_gap_proc = None
_gap_in_q = None
_gap_out_q = None


def _gap_infer_device():
    raw = (os.environ.get('GAP_INFER_DEVICE') or '0').strip().lower()
    if raw in ('cpu', 'none', 'false'):
        return 'cpu'
    try:
        return int(raw)
    except ValueError:
        return 0


def _gap_server(in_q, out_q, device):
    """常驻子进程：模型只加载一次，循环识别缺口。"""
    os.environ.setdefault('OMP_NUM_THREADS', '1')
    try:
        from captcha_recognizer.recognizer import Recognizer

        def predict_with_device(self, source=None, **kwargs):
            conf = kwargs.pop('conf', 0.45)
            imgsz = kwargs.pop('imgsz', [416, 416])
            results = self.model.predict(
                source=source, device=device, conf=conf, imgsz=imgsz, **kwargs)
            if len(results):
                return results[0]

        Recognizer.predict = predict_with_device
        rec = Recognizer()
        out_q.put(('ready', None))
    except Exception as e:
        try:
            out_q.put(('boot_err', str(e)))
        except Exception:
            pass
        return
    while True:
        try:
            item = in_q.get()
        except Exception:
            break
        if not item or item[0] == 'stop':
            break
        temp_path = item[1]
        conf = item[2] if len(item) > 2 else 0.45
        try:
            box, confidence = rec.identify_gap(
                source=temp_path, verbose=False, conf=conf)
            out_q.put(('ok', box, confidence))
        except Exception as e:
            try:
                out_q.put(('err', str(e)))
            except Exception:
                pass


def _stop_gap_worker():
    global _gap_proc, _gap_in_q, _gap_out_q
    proc, in_q = _gap_proc, _gap_in_q
    _gap_proc = _gap_in_q = _gap_out_q = None
    if in_q is not None:
        try:
            in_q.put(('stop', None))
        except Exception:
            pass
    if proc is not None and proc.is_alive():
        proc.join(3)
        if proc.is_alive():
            proc.terminate()
            proc.join(5)


def _ensure_gap_worker():
    global _gap_proc, _gap_in_q, _gap_out_q
    if _gap_proc is not None and _gap_proc.is_alive():
        return True
    _stop_gap_worker()
    ctx = multiprocessing.get_context('spawn')
    in_q = ctx.Queue()
    out_q = ctx.Queue()
    proc = ctx.Process(
        target=_gap_server, args=(in_q, out_q, _gap_infer_device()), daemon=True)
    proc.start()
    try:
        msg = out_q.get(timeout=60)
    except Exception as e:
        print(f'    ✗ 缺口识别进程启动超时：{e}', flush=True)
        if proc.is_alive():
            proc.terminate()
            proc.join(5)
        return False
    if not msg or msg[0] != 'ready':
        print(f'    ✗ 缺口识别进程启动失败：{msg[1] if msg else msg}', flush=True)
        if proc.is_alive():
            proc.terminate()
            proc.join(5)
        return False
    _gap_proc, _gap_in_q, _gap_out_q = proc, in_q, out_q
    print('    缺口识别进程已就绪（模型常驻）', flush=True)
    return True


atexit.register(_stop_gap_worker)


def identify_gap_isolated(temp_path, timeout=GAP_IDENTIFY_TIMEOUT):
    """子进程识别腾讯滑块缺口；模型常驻，避免每次重新加载 ONNX。"""
    if not _ensure_gap_worker():
        return None, None
    try:
        while True:
            try:
                _gap_out_q.get_nowait()
            except Exception:
                break
        _gap_in_q.put(('run', temp_path, 0.45))
        msg = _gap_out_q.get(timeout=timeout)
        if msg and msg[0] == 'ok' and not msg[1]:
            _gap_in_q.put(('run', temp_path, 0.28))
            msg = _gap_out_q.get(timeout=timeout)
    except Exception as e:
        print(f'    ✗ 未收到识别结果：{e}', flush=True)
        _stop_gap_worker()
        return None, None
    if not msg:
        return None, None
    if msg[0] == 'ok':
        box, conf = msg[1], msg[2]
        if not box:
            print(f'    ✗ 缺口识别空框 conf={conf}', flush=True)
        return box, conf
    print(f'    ✗ 识别异常：{msg[1]}', flush=True)
    return None, None


def _extract_left_px(style):
    match = re.search(r'left:\s*([\d.]+)px', style or '')
    if match:
        return float(match.group(1))
    return None


def _extract_url_from_css_background(style_or_bg_image):
    if not style_or_bg_image or str(style_or_bg_image).strip() == 'none':
        return None
    s = str(style_or_bg_image).strip()
    m = re.search(r'url\s*\(\s*["\']([^"\']+)["\']\s*\)', s, re.I)
    if m:
        return m.group(1).strip()
    m = re.search(r'url\s*\(\s*([^)]+)\s*\)', s, re.I)
    if m:
        return m.group(1).strip().strip('"\'')
    return None


def snapshot_download_names(download_dir):
    if not download_dir or not os.path.isdir(download_dir):
        return set()
    try:
        return {
            name for name in os.listdir(download_dir)
            if not name.endswith('.lock') and not name.endswith('.done')
        }
    except Exception:
        return set()


def has_new_download_file(download_dir, before_files):
    """出现新的 PDF/Excel/临时下载文件即视为下载已开始。"""
    before_files = before_files or set()
    for name in snapshot_download_names(download_dir) - before_files:
        lower = name.lower()
        if any(lower.endswith(s) for s in DOWNLOAD_FILE_SUFFIXES):
            return True
    return False


def _css_any_displayed(driver, css):
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, css):
            try:
                if el.is_displayed():
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _tencent_wrap_visible(driver):
    return (
        _css_any_displayed(driver, '#tCaptchaDyMainWrap')
        or _css_any_displayed(driver, '.tencent-captcha-dy__warp')
    )


def _tencent_slider_solvable(driver):
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR, '.tencent-captcha-dy__slider-block'):
            try:
                if el.is_displayed():
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _tencent_wrap_text(driver):
    """腾讯验证码弹窗内的可见文案，用于识别成功/错误提示。"""
    texts = []
    selectors = (
        '#tCaptchaDyMainWrap',
        '.tencent-captcha-dy__warp',
        '[class*="tencent-captcha-dy"]',
    )
    for css in selectors:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, css):
                try:
                    if not el.is_displayed():
                        continue
                    text = (el.text or el.get_attribute('innerText') or '').strip()
                except Exception:
                    continue
                if text and text not in texts:
                    texts.append(text)
        except Exception:
            continue
    return ' '.join(texts)


def _tencent_success_ui_visible(driver):
    if is_cnki_verify_done_page(driver):
        return True
    if _css_any_displayed(driver, '.tencent-captcha-dy__verify-status-img--success'):
        return True
    for css in (
            '[class*="tencent-captcha-dy"][class*="success"]',
            '[class*="verify-status"][class*="success"]',
            '[class*="verify-success"]'):
        if _css_any_displayed(driver, css):
            return True
    text = _tencent_wrap_text(driver)
    compact = text.replace(' ', '')
    if compact and any(k in compact for k in (
            '验证成功', '验证通过', '通过验证', '成功验证',
            '验证完成', '已通过验证', '已通过')):
        return True
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR, '.tencent-captcha-dy__verify-speed'):
            text = (el.text or '').strip()
            if text and ('超过' in text or '成功' in text) and el.is_displayed():
                return True
    except Exception:
        pass
    return False


def is_cnki_verify_done_page(driver) -> bool:
    """下载验证通过后的「验证完成，已进入下载」页（无滑块，不应再解）。"""
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR, '.verify-box, .verIcon, .verTop'):
            try:
                if not el.is_displayed():
                    continue
            except Exception:
                continue
            text = (el.text or '')
            if '验证完成' in text or '已进入下载' in text:
                return True
    except Exception:
        pass
    try:
        for img in driver.find_elements(
                By.CSS_SELECTOR, '.verify-box img, .verIcon img'):
            src = (img.get_attribute('src') or '').lower()
            if 'succeed.png' in src or 'success.png' in src:
                try:
                    if img.is_displayed():
                        return True
                except Exception:
                    return True
    except Exception:
        pass
    return False


def click_verify_done_close(driver) -> bool:
    """点「关闭窗口」，结束验证完成页。"""
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR, '.verify-box span, .verIcon span, .verTop span'):
            if '关闭' not in (el.text or ''):
                continue
            driver.execute_script('arguments[0].click();', el)
            time.sleep(0.5)
            return True
    except Exception:
        pass
    return False


def is_tencent_slider_captcha(driver):
    """仅在仍需拖动滑块时视为验证码。通过后弹窗不关，不再当成未过。"""
    if _tencent_success_ui_visible(driver):
        return False
    return _tencent_slider_solvable(driver)


def has_visible_captcha_widget(driver):
    """页面上是否仍有可交互的验证码控件（不依赖 URL）。"""
    selectors = (
        '.tencent-captcha-dy__slider-block',
        '.verify-sub-block',
        '.verify-img-panel img',
        '.verifybox',
    )
    for css in selectors:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, css):
                try:
                    if el.is_displayed():
                        return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def wait_for_tencent_slider(driver, timeout=12):
    WebDriverWait(driver, timeout, 0.4).until(
        EC.visibility_of_element_located(
            (By.CSS_SELECTOR, '.tencent-captcha-dy__slider-block'))
    )


def _find_tencent_bg_element(driver):
    selectors = [
        '.tencent-captcha-dy__verify-bg-img',
        '.tencent-captcha-dy__verify-bg',
        '[class*="tencent-captcha"][class*="verify-bg"]',
        '[class*="verify-bg-img"]',
    ]
    for css in selectors:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, css):
                if el.is_displayed():
                    return el
        except Exception:
            continue
    return None


def resolve_tencent_background_url(driver, bg_el):
    if not bg_el:
        return None
    url = _extract_url_from_css_background(bg_el.get_attribute('style') or '')
    if url:
        return url
    try:
        comp = driver.execute_script(
            'return window.getComputedStyle(arguments[0]).backgroundImage;', bg_el)
        url = _extract_url_from_css_background(comp or '')
        if url:
            return url
    except Exception:
        pass
    try:
        for img in bg_el.find_elements(By.CSS_SELECTOR, 'img'):
            src = (img.get_attribute('src') or '').strip()
            if src:
                return src
    except Exception:
        pass
    return None


def _fetch_tencent_bg_bytes(driver, bg_url):
    if not bg_url:
        raise ValueError('背景地址为空')
    bg_url = bg_url.strip()
    if bg_url.startswith('data:'):
        if 'base64,' not in bg_url:
            raise ValueError('data URL 中无 base64 段')
        raw = re.sub(r'\s+', '', bg_url.split('base64,', 1)[1])
        return base64.b64decode(raw)
    if bg_url.startswith('//'):
        bg_url = 'https:' + bg_url
    elif not re.match(r'^https?://', bg_url, re.I):
        bg_url = urljoin(driver.current_url, bg_url)
    try:
        ua = driver.execute_script('return navigator.userAgent;') or ''
    except Exception:
        ua = ''
    if not ua.strip():
        ua = (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
            '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )
    headers = {
        'User-Agent': ua,
        'Referer': driver.current_url,
        'Accept': 'image/avif,image/webp,image/apng,image/png,image/*,*/*;q=0.8',
    }
    cookies = {c['name']: c['value'] for c in driver.get_cookies()}
    resp = requests.get(
        bg_url, headers=headers, cookies=cookies, timeout=20, allow_redirects=True)
    if resp.status_code != 200:
        raise ValueError(f'HTTP {resp.status_code}，响应长度={len(resp.content)}')
    if not resp.content or len(resp.content) < 32:
        raise ValueError(f'图片内容过短（{len(resp.content)} 字节）')
    return resp.content


def _fetch_tencent_bg_bytes_via_browser(driver, bg_url):
    if bg_url.strip().startswith('data:'):
        return _fetch_tencent_bg_bytes(driver, bg_url)
    abs_url = bg_url.strip()
    if abs_url.startswith('//'):
        abs_url = 'https:' + abs_url
    elif not re.match(r'^https?://', abs_url, re.I):
        abs_url = urljoin(driver.current_url, abs_url)
    b64 = driver.execute_async_script(
        """
        var u = arguments[0];
        var done = arguments[arguments.length - 1];
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', u, true);
            xhr.responseType = 'blob';
            xhr.onload = function() {
                if (xhr.status !== 200) { done(null); return; }
                var fr = new FileReader();
                fr.onload = function() { done(fr.result); };
                fr.onerror = function() { done(null); };
                fr.readAsDataURL(xhr.response);
            };
            xhr.onerror = function() { done(null); };
            xhr.send();
        } catch (e) { done(null); }
        """,
        abs_url,
    )
    if not b64 or 'base64,' not in str(b64):
        raise ValueError('XHR 未返回有效图片（可能被跨域拦截）')
    raw = re.sub(r'\s+', '', str(b64).split('base64,', 1)[1])
    data = base64.b64decode(raw)
    if len(data) < 32:
        raise ValueError('XHR 解码后内容过短')
    return data


def get_tencent_slider_start_x(driver):
    try:
        slider = driver.find_element(
            By.CSS_SELECTOR, '.tencent-captcha-dy__slider-block')
        left = _extract_left_px(slider.get_attribute('style') or '')
        if left is not None:
            return left
    except Exception:
        pass
    return TENCENT_SLIDER_START_X


def get_tencent_piece_geom(driver):
    """拼图块显示坐标：left / width。滑动应对齐拼图而非滑块按钮。"""
    try:
        piece = driver.find_element(By.CSS_SELECTOR, '.tencent-captcha-dy__fg-item')
        left = _extract_left_px(piece.get_attribute('style') or '')
        width = 0
        try:
            width = float(piece.size.get('width') or 0)
        except Exception:
            width = 0
        if left is None:
            m = re.search(r'width:\s*([\d.]+)px', piece.get_attribute('style') or '')
            if m and not width:
                width = float(m.group(1))
            left = TENCENT_SLIDER_START_X
        return left, width
    except Exception:
        return TENCENT_SLIDER_START_X, 0


def click_tencent_confirm(driver):
    """滑块对齐后需点「确定」才会开始下载。"""
    selectors = (
        '.tencent-captcha-dy__verify-confirm-btn',
        '[class*="verify-confirm"]',
        '[class*="tencent-captcha-dy"] button',
        '.tencent-captcha-dy__warp button',
    )
    confirm_words = ('确定', '确认', '继续', '提交', '完成', 'ok')
    for css in selectors:
        try:
            btns = driver.find_elements(By.CSS_SELECTOR, css)
        except Exception:
            continue
        for btn in btns:
            try:
                cls = (btn.get_attribute('class') or '').lower()
                if 'disabled' in cls:
                    continue
                if not btn.is_displayed():
                    continue
                text = (btn.text or btn.get_attribute('innerText') or '').strip()
                is_confirm_cls = any(
                    k in cls for k in ('confirm', 'submit', 'sure'))
                if text:
                    if text.lower() not in confirm_words and not any(
                            w in text for w in confirm_words):
                        continue
                elif not is_confirm_cls:
                    continue
                driver.execute_script('arguments[0].click();', btn)
                print('    ✓ 已点击验证「确定」')
                return True
            except Exception:
                continue
    return False


def gettrack_tencent(distance):
    """拟人滑动：缓入缓出、少量过冲再回拉，总时长约 0.6~1.2s。"""
    distance = int(round(distance))
    if distance <= 0:
        return []
    overshoot = random.randint(2, 6) if distance > 20 else 0
    total = distance + overshoot
    n = random.randint(32, 48)
    track = []
    current = 0
    for i in range(1, n + 1):
        t = i / n
        if t < 0.5:
            eased = 4 * t * t * t
        else:
            eased = 1 - ((-2 * t + 2) ** 3) / 2
        target = int(round(total * eased))
        step = target - current
        if step:
            track.append(step)
            current = target
    if current != total:
        track.append(total - current)
        current = total
    back = current - distance
    while back > 0:
        step = min(back, random.randint(1, 3))
        track.append(-step)
        back -= step
    return [s for s in track if s]


def _drag_tencent_slider(driver, slider, distance):
    track = gettrack_tencent(distance)
    ac = ActionChains(driver)
    ac.click_and_hold(slider)
    ac.pause(random.uniform(0.16, 0.32))
    y_drift = 0
    for offset in track:
        dy = 0
        if random.random() < 0.16 and abs(y_drift) < 2:
            dy = random.choice((-1, 1))
            y_drift += dy
        ac.move_by_offset(int(offset), int(dy))
        ac.pause(random.uniform(
            TENCENT_SLIDE_STEP_SLEEP_MIN, TENCENT_SLIDE_STEP_SLEEP_MAX))
    ac.pause(random.uniform(0.12, 0.22))
    ac.release()
    ac.perform()


def tencent_last_fail():
    return _tencent_last_fail


def _set_tencent_last_fail(reason):
    global _tencent_last_fail
    _tencent_last_fail = reason


def wait_tencent_cooldown_if_needed() -> bool:
    """兼容旧调用：限流改为关浏览器后等待，此处不再空等。"""
    return False


def tencent_feedback_text(driver) -> str:
    texts = []
    selectors = (
        '.tencent-captcha-dy__verify-error-text',
        '.tencent-captcha-dy__verify-error-tip',
        '.tencent-captcha-dy__verify-msg',
        '[class*="verify-error"]',
    )
    for css in selectors:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, css):
                try:
                    if not el.is_displayed():
                        continue
                except Exception:
                    continue
                text = (el.text or el.get_attribute('innerText') or '').strip()
                if text:
                    texts.append(text)
        except Exception:
            continue
    seen = []
    for text in texts:
        if text not in seen:
            seen.append(text)
    return ' '.join(seen)


def is_tencent_rate_limited(driver) -> bool:
    text = tencent_feedback_text(driver)
    return any(k in text for k in _TENCENT_RATE_KEYS)


def is_tencent_slider_error(driver):
    if is_tencent_rate_limited(driver):
        return True
    if _css_any_displayed(driver, '.tencent-captcha-dy__slider-block--fail'):
        return True
    if _css_any_displayed(driver, '.tencent-captcha-dy__verify-error-tip'):
        return True
    text = tencent_feedback_text(driver)
    return bool(text) and any(k in text for k in ('错误', '重试', '失败', '不正确'))


def _wait_tencent_after_fail(driver, timeout=4.0):
    """失败后等自动换图，避免立刻再滑同一张或连点刷新。"""
    if is_tencent_rate_limited(driver):
        return
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _tencent_slider_solvable(driver) and not is_tencent_slider_error(driver):
            time.sleep(0.45)
            return
        time.sleep(0.25)


def wait_tencent_pass_by_download(
        driver, download_dir, before_files, timeout=TENCENT_DOWNLOAD_WAIT):
    """松开滑块后点「确定」，以下载文件或成功态判定通过。"""
    deadline = time.time() + timeout
    saw_success_ui = False
    clicked_confirm = False
    while time.time() < deadline:
        if has_new_download_file(download_dir, before_files):
            print('    ✓ 已检测到新下载文件，判定验证通过')
            return True
        if is_cnki_verify_done_page(driver):
            print('    ✓ 验证完成，已进入下载')
            click_verify_done_close(driver)
            return True
        if is_tencent_slider_error(driver):
            if has_new_download_file(download_dir, before_files):
                print('    ✓ 已检测到新下载文件，判定验证通过')
                return True
            return False
        if click_tencent_confirm(driver):
            clicked_confirm = True
        if _tencent_success_ui_visible(driver):
            if not saw_success_ui:
                print('    成功提示已出现，点击确定并等待下载...')
            saw_success_ui = True
            click_tencent_confirm(driver)
        time.sleep(0.25)
    if has_new_download_file(download_dir, before_files):
        print('    ✓ 已检测到新下载文件，判定验证通过')
        return True
    if saw_success_ui or clicked_confirm:
        print('    ✓ 滑块已通过（弹窗未关），交由下载等待确认文件')
        return True
    return False


def is_tencent_slider_success(driver, initial_url, download_dir=None, before_files=None):
    if has_new_download_file(download_dir, before_files):
        return True
    if is_cnki_verify_done_page(driver):
        return True
    if _tencent_success_ui_visible(driver):
        return True
    try:
        current = driver.current_url or ''
    except Exception:
        current = ''
    if is_cnki_verify_url(initial_url) and not is_cnki_verify_url(current):
        return True
    return False


def refresh_tencent_captcha(driver):
    selectors = (
        '.tencent-captcha-dy__footer-icon--refresh img',
        '.tencent-captcha-dy__footer-icon--refresh',
    )
    for css in selectors:
        btns = driver.find_elements(By.CSS_SELECTOR, css)
        for btn in btns:
            try:
                if btn.is_displayed():
                    driver.execute_script('arguments[0].click();', btn)
                    time.sleep(2)
                    return True
            except Exception:
                continue
    return False


def solve_tencent_slider(driver, save_dir=None, download_dir=None, before_files=None):
    """处理最新腾讯滑块 tencent-captcha-dy（bar.cnki.net/bar/verify）。

    通过后弹窗不会自动关闭，以下载目录出现新文件作为通过依据。
    """
    _set_tencent_last_fail(None)
    if has_new_download_file(download_dir, before_files):
        print('    ✓ 滑动前已检测到新下载文件，判定验证通过')
        return True
    if is_cnki_verify_done_page(driver):
        print('    ✓ 验证完成，已进入下载')
        click_verify_done_close(driver)
        return True
    if is_tencent_rate_limited(driver):
        err_text = tencent_feedback_text(driver)
        _set_tencent_last_fail('rate_limit')
        print(f'    已提示过于频繁，本次不滑动{("：" + err_text) if err_text else ""}')
        return False
    if _tencent_success_ui_visible(driver) and not _tencent_slider_solvable(driver):
        if download_dir:
            return wait_tencent_pass_by_download(driver, download_dir, before_files)
        print('    ✓ 腾讯滑块已显示成功，弹窗未关')
        return True

    try:
        wait_for_tencent_slider(driver)
        time.sleep(0.8)
    except Exception as e:
        if has_new_download_file(download_dir, before_files):
            print('    ✓ 未找到滑块，但已检测到新下载文件，判定验证通过')
            return True
        if is_cnki_verify_done_page(driver):
            print('    ✓ 未找到滑块，验证已完成并进入下载')
            click_verify_done_close(driver)
            return True
        if _tencent_success_ui_visible(driver):
            if download_dir:
                return wait_tencent_pass_by_download(driver, download_dir, before_files)
            print('    ✓ 未找到滑块，成功提示已出现')
            return True
        print('未找到腾讯滑块元素:', e)
        return False

    initial_url = driver.current_url
    slider_start_x = get_tencent_slider_start_x(driver)
    print(f'    腾讯滑块初始位置：{slider_start_x}px')

    bg_el = _find_tencent_bg_element(driver)
    if not bg_el:
        print('    ✗ 未找到腾讯验证码背景')
        return False
    bg_url = resolve_tencent_background_url(driver, bg_el)
    if not bg_url:
        print('    ✗ 无法提取背景图 URL')
        return False
    print(f'    背景图：{bg_url[:120]}{"..." if len(bg_url) > 120 else ""}')

    try:
        bg_bytes = _fetch_tencent_bg_bytes(driver, bg_url)
    except Exception as e_req:
        print(f'    requests 下载背景失败：{e_req}')
        try:
            bg_bytes = _fetch_tencent_bg_bytes_via_browser(driver, bg_url)
            print('    ✓ 已通过页面内请求获取背景图')
        except Exception as e_xhr:
            print(f'    页面内 XHR 获取失败：{e_xhr}')
            return False

    save_dir = save_dir or os.path.join('.', 'captcha_debug', 'tencent')
    os.makedirs(save_dir, exist_ok=True)
    temp_bg_path = os.path.join(save_dir, 'tencent_bg.png')
    with open(temp_bg_path, 'wb') as f:
        f.write(bg_bytes)

    print('    使用 captcha_recognizer 识别缺口（子进程）...', flush=True)
    box, confidence = identify_gap_isolated(temp_bg_path)
    if not box:
        _set_tencent_last_fail('no_gap')
        print('    ✗ 缺口识别失败')
        return False
    print(f'    识别结果：box={box}, confidence={confidence}', flush=True)

    target = float(box[0])
    gap_right = float(box[2]) if len(box) > 2 else target
    display_w = bg_el.size.get('width') or 0
    style = bg_el.get_attribute('style') or ''
    m_size = re.search(r'background-size:\s*([\d.]+)px', style)
    if m_size:
        display_w = float(m_size.group(1)) or display_w
    try:
        img_w = Image.open(BytesIO(bg_bytes)).size[0]
    except Exception:
        img_w = 0
    scale = (img_w / display_w) if display_w and img_w else 2.187
    gap_display = target / scale
    gap_w_display = (gap_right - target) / scale if scale else 0
    piece_left, piece_w = get_tencent_piece_geom(driver)
    start_x = piece_left if piece_left is not None else slider_start_x
    # 识别框通常小于拼图：按中心对齐，避免只贴左缘导致偏右约 8~12px
    pad = 0.0
    if piece_w > 0 and gap_w_display > 4:
        pad = (piece_w - gap_w_display) / 2.0
    slide_distance = int(round(gap_display - start_x - pad + TENCENT_GAP_BIAS))
    print(
        f'    缺口={gap_display:.1f}px scale={scale:.3f} 拼图左={piece_left:.1f} '
        f'滑块左={slider_start_x} 识别宽={gap_w_display:.1f} 拼图宽={piece_w:.1f} '
        f'内边距={pad:.1f} 固定偏差={TENCENT_GAP_BIAS:.1f} 滑动={slide_distance}px'
    )
    if slide_distance < 5:
        print('    ✗ 滑动距离过小，放弃本次')
        return False

    try:
        slider = driver.find_element(
            By.CSS_SELECTOR, '.tencent-captcha-dy__slider-block')
    except Exception as e:
        print('    ✗ 未找到滑块按钮:', e)
        return False

    try:
        _drag_tencent_slider(driver, slider, slide_distance)
    except Exception as e:
        print(f'    ✗ 拖动滑块失败: {e}')
        return False
    print('    ✓ 已释放腾讯滑块', flush=True)
    time.sleep(0.6)
    click_tencent_confirm(driver)
    last_confirm = time.time()

    # 腾讯成功态不一定有固定成功图标：滑块消失、弹窗关闭、URL 离开验证页都算通过
    deadline = time.time() + (TENCENT_DOWNLOAD_WAIT if download_dir else 12)
    gone_since = None
    saw_success = False
    while time.time() < deadline:
        if has_new_download_file(download_dir, before_files):
            print('    ✓ 已检测到新下载文件，判定验证通过')
            return True
        if is_cnki_verify_done_page(driver):
            print('    ✓ 验证完成，已进入下载')
            click_verify_done_close(driver)
            return True
        err_text = tencent_feedback_text(driver)
        if is_tencent_rate_limited(driver):
            _set_tencent_last_fail('rate_limit')
            print(f'    ✗ 腾讯滑块过于频繁{("：" + err_text) if err_text else ""}')
            return False
        if is_tencent_slider_error(driver):
            _set_tencent_last_fail('verify_error')
            print(f'    ✗ 腾讯滑块验证错误{("：" + err_text) if err_text else ""}')
            _wait_tencent_after_fail(driver)
            return False
        if _tencent_success_ui_visible(driver):
            saw_success = True
            click_tencent_confirm(driver)
            print('    ✓ 腾讯滑块已显示成功')
            return True
        try:
            current_url = driver.current_url or ''
        except Exception:
            current_url = initial_url
        if is_cnki_verify_url(initial_url) and not is_cnki_verify_url(current_url):
            print('    ✓ 已离开验证页，判定验证通过')
            return True
        if not _tencent_slider_solvable(driver):
            gone_since = gone_since or time.time()
            if time.time() - gone_since >= 2.0:
                print('    ✓ 腾讯滑块控件已消失且无错误提示，判定验证通过')
                return True
        else:
            gone_since = None
        if time.time() - last_confirm >= 1.5:
            click_tencent_confirm(driver)
            last_confirm = time.time()
        time.sleep(0.25)

    if has_new_download_file(download_dir, before_files):
        print('    ✓ 已检测到新下载文件，判定验证通过')
        return True
    if saw_success:
        print('    ✓ 腾讯滑块已显示成功（弹窗未关）')
        return True
    if download_dir and not _tencent_slider_solvable(driver):
        # 已无滑块可操作，下载交给上层 wait_pdf_or_unavailable 继续等待
        print('    ✓ 腾讯滑块已不可操作，交由下载等待确认文件')
        return True
    _set_tencent_last_fail('no_pass')
    print('    ✗ 腾讯滑块未出现成功提示')
    return False


def identify_gap_distance(hkpath, qkpath):
    """使用 ddddocr 双图匹配定位滑块缺口。"""
    det = get_slide_ocr()
    with open(hkpath, 'rb') as f:
        target_bytes = f.read()
    with open(qkpath, 'rb') as f:
        background_bytes = f.read()

    res = det.slide_match(target_bytes, background_bytes, simple_target=True)
    print(f'滑块位置: {res}')
    gap_x = 0
    try:
        gap_x = res['target'][0]
    except (TypeError, KeyError, IndexError):
        print('没有找到缺口目标')
    distance = round(gap_x)
    print(f'gap_x={gap_x}, distance={distance}')
    return gap_x, res, distance


def is_cnki_verify_url(url):
    """知网验证码页：检索站 /verify/home 与下载站 bar/verify（含 WebVPN 包装）。"""
    u = (url or '').lower()
    return (
        'kns.cnki.net/verify' in u
        or 'bar.cnki.net/bar/verify' in u
        or '/bar/verify' in u
        or '/verify/index' in u
        or '/verify/home' in u
    )


def captcha_type_from_url(url):
    """从 URL 推断验证码类型；控件未加载时用于等待，不把 blockPuzzle 误判为点选。"""
    u = (url or '').lower().replace(' ', '')
    if not u:
        return None
    if 'captchatype=blockpuzzle' in u or 'verify-sub-block' in u:
        return 'slider'
    if 'captchatype=click' in u or 'clickword' in u:
        return 'text_click'
    if '/bar/verify' in u:
        return 'tencent_slider'
    if is_cnki_verify_url(url):
        return 'slider'
    return None


def is_slider_captcha(driver):
    return bool(driver.find_elements(By.CLASS_NAME, 'verify-sub-block'))


def is_text_click_captcha(driver):
    if not driver.find_elements(By.CSS_SELECTOR, '.verify-img-panel img'):
        return False
    if is_slider_captcha(driver) or is_tencent_slider_captcha(driver):
        return False
    msgs = driver.find_elements(By.CLASS_NAME, 'verify-msg')
    if msgs and '点击' in msgs[0].text:
        return True
    return bool(driver.find_elements(By.CLASS_NAME, 'verifybox'))


def get_captcha_type(driver):
    if is_cnki_verify_done_page(driver):
        return None
    if _tencent_success_ui_visible(driver) and not _tencent_slider_solvable(driver):
        return None
    if is_tencent_slider_captcha(driver):
        return 'tencent_slider'
    if is_slider_captcha(driver):
        return 'slider'
    if is_text_click_captcha(driver):
        return 'text_click'
    url = ''
    try:
        url = driver.current_url or ''
    except Exception:
        url = ''
    return captcha_type_from_url(url)


def has_block_puzzle(driver):
    """兼容旧名：检测页面是否仍需处理验证码。通过后残留弹窗不算。"""
    if is_cnki_verify_done_page(driver):
        return False
    if get_captcha_type(driver):
        return True
    try:
        url = driver.current_url or ''
    except Exception:
        url = ''
    if is_cnki_verify_url(url):
        return True
    try:
        for el in driver.find_elements(By.CLASS_NAME, 'verifybox'):
            if el.is_displayed():
                return True
    except Exception:
        pass
    return False


def wait_for_any_captcha_widget(driver, timeout=12):
    """等待滑块 / 点选 / 腾讯滑块控件出现；验证码页跳转后控件常晚于 URL 出现。"""
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        if is_cnki_verify_done_page(driver):
            return None
        if is_tencent_slider_captcha(driver):
            return 'tencent_slider'
        if is_slider_captcha(driver):
            return 'slider'
        if is_text_click_captcha(driver):
            return 'text_click'
        try:
            last = captcha_type_from_url(driver.current_url or '')
        except Exception:
            last = None
        time.sleep(0.4)
    return last


def wait_for_block_puzzle(driver, timeout=20):
    WebDriverWait(driver, timeout, 0.5).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'verify-sub-block'))
    )


def wait_for_text_click_captcha(driver, timeout=20):
    WebDriverWait(driver, timeout, 0.5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '.verify-msg'))
    )


def is_captcha_success(driver, initial_url):
    time.sleep(2)
    try:
        current = driver.current_url or ''
    except Exception:
        current = ''
    if is_cnki_verify_url(initial_url) and not is_cnki_verify_url(current):
        return True
    if is_cnki_verify_done_page(driver):
        return True
    if _tencent_success_ui_visible(driver):
        return True
    if is_tencent_slider_captcha(driver):
        return is_tencent_slider_success(driver, initial_url)
    if not is_cnki_verify_url(current) and current != initial_url:
        if not driver.find_elements(By.CLASS_NAME, 'verifybox'):
            return True
    if not driver.find_elements(By.CLASS_NAME, 'verifybox'):
        return True
    if is_slider_captcha(driver):
        tips = driver.find_elements(By.CLASS_NAME, 'verify-tips')
        for tip in tips:
            text = tip.text.strip()
            if text and ('成功' in text or '通过' in text):
                return True
        return False
    return is_text_click_success(driver, initial_url)


def is_text_click_success(driver, initial_url, max_wait=8):
    """点选后轮询：verifybox 消失或提示成功即为通过。"""
    for _ in range(max_wait):
        time.sleep(1)
        try:
            current = driver.current_url or ''
        except Exception:
            current = ''
        if is_cnki_verify_url(initial_url) and not is_cnki_verify_url(current):
            if not driver.find_elements(By.CLASS_NAME, 'verifybox'):
                return True
        if not is_cnki_verify_url(current) and current != initial_url:
            if not driver.find_elements(By.CLASS_NAME, 'verifybox'):
                return True
        if not driver.find_elements(By.CLASS_NAME, 'verifybox'):
            return True
        msgs = driver.find_elements(By.CLASS_NAME, 'verify-msg')
        for msg in msgs:
            text = msg.text.strip()
            if '成功' in text or '通过' in text:
                return True
            if '失败' in text or '错误' in text:
                return False
    return not driver.find_elements(By.CLASS_NAME, 'verifybox')


def captcha_puzzle_fingerprint(driver):
    """当前验证码题面指纹；用于判断刷新是否真正换图。"""
    try:
        if is_tencent_slider_captcha(driver):
            imgs = driver.find_elements(
                By.CSS_SELECTOR,
                '.tencent-captcha-dy img, #tcaptcha_transform img, .tc-fg-item img')
            srcs = []
            for el in imgs[:4]:
                try:
                    src = el.get_attribute('src') or ''
                except Exception:
                    src = ''
                if src:
                    srcs.append(src[-96:])
            return ('tencent', tuple(srcs)) if srcs else None
        if driver.find_elements(By.CLASS_NAME, 'verifybox'):
            try:
                msg = driver.find_element(By.CLASS_NAME, 'verify-msg').text
            except Exception:
                msg = ''
            try:
                img = driver.find_element(
                    By.CSS_SELECTOR, '.verify-img-panel img')
                src = img.get_attribute('src') or ''
            except Exception:
                src = ''
            return ('text_click', msg, src[-96:] if src else '')
        if has_block_puzzle(driver):
            hk, qk = _extract_captcha_images(driver)
            return ('block', hk[-96:], qk[-96:])
    except Exception:
        return None
    return None


def refresh_captcha(driver, force=False):
    """刷新验证码题面。

    返回:
      True  — 已换题（或腾讯路径已触发刷新）
      False — 未换题 / 限流未刷 / 刷新失败
    """
    if is_tencent_rate_limited(driver):
        return False
    if is_tencent_slider_captcha(driver):
        # 失败后控件常自动换图；限流时不要点刷新。识别失败则强制换图。
        if not force:
            time.sleep(1.2)
            if _tencent_slider_solvable(driver) and not is_tencent_slider_error(driver):
                return True
        refresh_tencent_captcha(driver)
        return True
    before = captcha_puzzle_fingerprint(driver)
    refresh_btns = driver.find_elements(By.CLASS_NAME, 'verify-refresh')
    if refresh_btns:
        try:
            refresh_btns[0].click()
        except Exception:
            try:
                driver.execute_script('arguments[0].click();', refresh_btns[0])
            except Exception as e:
                print('点击验证码刷新失败:', e)
                return False
        time.sleep(2)
        after = captcha_puzzle_fingerprint(driver)
        if before is not None and after == before:
            print('验证码刷新按钮已点，但题面未变化')
            return False
        if before is None and after is None:
            print('验证码刷新后仍无法读取题面')
            return False
        print('验证码已刷新换题')
        return True
    print('未找到验证码刷新按钮，尝试整页刷新')
    driver.refresh()
    time.sleep(3)
    return True


def _extract_captcha_images(driver):
    hkimg = driver.find_element(By.CLASS_NAME, 'verify-sub-block') \
        .find_element(By.CSS_SELECTOR, 'img').get_attribute('src')
    hkimg = hkimg.replace('data:image/png;base64,', '')

    qkimg = driver.find_element(By.CLASS_NAME, 'verify-img-panel') \
        .find_element(By.CSS_SELECTOR, 'img').get_attribute('src')
    qkimg = qkimg.replace('data:image/png;base64,', '')

    return hkimg, qkimg


def save_captcha_images(driver, save_dir):
    os.makedirs(save_dir, exist_ok=True)
    hkimg, qkimg = _extract_captcha_images(driver)
    hkpath = os.path.join(save_dir, '滑块图.png')
    qkpath = os.path.join(save_dir, '背景图.png')
    base64_to_image(hkimg, hkpath)
    base64_to_image(qkimg, qkpath)
    return hkpath, qkpath


def parse_text_click_targets(msg_text):
    """从 verify-msg 解析点选顺序，如：请依次点击【打,出,服】。"""
    match = re.search(r'【([^】]+)】', msg_text)
    if not match:
        return []
    inner = match.group(1).strip()
    parts = re.split(r'[,，、\s]+', inner)
    return [p for p in parts if p]


def save_text_click_image(driver, save_dir):
    os.makedirs(save_dir, exist_ok=True)
    img_el = driver.find_element(By.CSS_SELECTOR, '.verify-img-panel img')
    src = img_el.get_attribute('src') or ''
    image_path = os.path.join(save_dir, '点选验证码.png')
    if 'base64,' in src:
        b64 = src.split('base64,', 1)[1]
        base64_to_image(b64, image_path)
    else:
        raise RuntimeError('验证码图片不是 base64 格式')
    return image_path, img_el


def _render_target_char(char, cell=64):
    """将目标字渲染为参考图，供相似度比对。"""
    from PIL import ImageDraw, ImageFont

    img = Image.new('RGB', (cell, cell), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype('msyh.ttc', 48)
    except OSError:
        try:
            font = ImageFont.truetype('simhei.ttf', 48)
        except OSError:
            font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), char, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(
        ((cell - tw) / 2, (cell - th) / 2), char, fill=(20, 20, 20), font=font)
    return img


def _image_similarity(img_a, img_b, size=(64, 64)):
    """两图灰度归一化互相关，返回约 [-1, 1] 的相似度。"""
    import numpy as np

    a = np.asarray(img_a.convert('L').resize(size), dtype=np.float32).ravel()
    b = np.asarray(img_b.convert('L').resize(size), dtype=np.float32).ravel()
    a = (a - a.mean()) / (a.std() + 1e-6)
    b = (b - b.mean()) / (b.std() + 1e-6)
    return float(np.dot(a, b) / len(a))


def _similarity_match_from_candidates(target_char, detections, used):
    """从未使用的备选中按相似度排序，取最高者。"""
    target_img = _render_target_char(target_char)
    ranked = []
    for idx, det in enumerate(detections):
        if idx in used:
            continue
        if 'crop' not in det:
            continue
        sim = _image_similarity(target_img, det['crop'])
        ranked.append((sim, idx))
        label = det['char'] or '?'
        print(f'相似度 {target_char} vs 备选项{idx} ({label!r}): {sim:.3f}')
    if not ranked:
        return None, -1.0
    ranked.sort(key=lambda x: x[0], reverse=True)
    best_sim, best_idx = ranked[0]
    if best_sim < 0.15:
        return None, best_sim
    return best_idx, best_sim


def _is_near_position(det, pos, thresh=20.0):
    """判断检测框中心是否靠近已锁定坐标。"""
    if pos is None:
        return False
    cx, cy = pos
    return abs(det['cx'] - cx) < thresh and abs(det['cy'] - cy) < thresh


def _exact_match_one(target, detections, used):
    """在 detections 中精确匹配一个目标字，返回 index 或 None。"""
    for idx, det in enumerate(detections):
        if idx in used:
            continue
        if det.get('char') == target:
            return idx
    return None


def _collect_rapid_detections(image_path):
    """RapidOCR 检测+识别；小图放大以提高散字召回。"""
    import cv2
    import numpy as np

    ocr = get_rapid_ocr()
    if ocr is None:
        return []

    img = Image.open(image_path).convert('RGB')
    arr = np.asarray(img)
    scale = max(1, RAPIDOCR_UPSCALE)
    if scale > 1:
        h, w = arr.shape[:2]
        arr = cv2.resize(arr, (w * scale, h * scale), interpolation=cv2.INTER_CUBIC)

    result, _ = ocr(arr)
    if not result:
        print('RapidOCR 未检测到文字')
        return []

    detections = []
    for box, text, conf in result:
        text = (text or '').strip()
        if not text:
            continue
        xs = [p[0] / scale for p in box]
        ys = [p[1] / scale for p in box]
        x1, x2 = min(xs), max(xs)
        y1, y2 = min(ys), max(ys)
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        crop = img.crop((
            max(0, int(x1)), max(0, int(y1)),
            min(img.width, int(x2) + 1), min(img.height, int(y2) + 1),
        ))
        chars = [ch for ch in text if '\u4e00' <= ch <= '\u9fff']
        if not chars:
            print(f'RapidOCR 跳过非汉字: {text!r} conf={float(conf):.3f}')
            continue
        if len(chars) == 1:
            detections.append({
                'char': chars[0], 'cx': cx, 'cy': cy, 'crop': crop,
                'conf': float(conf), 'source': 'rapid',
            })
            print(
                f'RapidOCR 检测到字符: {chars[0]!r} conf={float(conf):.3f} '
                f'@ ({cx:.1f}, {cy:.1f})')
        else:
            width = max(x2 - x1, 1)
            for i, ch in enumerate(chars):
                ch_cx = x1 + width * (i + 0.5) / len(chars)
                detections.append({
                    'char': ch, 'cx': ch_cx, 'cy': cy, 'crop': crop,
                    'conf': float(conf), 'source': 'rapid',
                })
                print(
                    f'RapidOCR 检测到字符: {ch!r}(from {text!r}) '
                    f'conf={float(conf):.3f} @ ({ch_cx:.1f}, {cy:.1f})')
    return detections


def _collect_ddddocr_detections(image_path):
    """ddddocr 检测框 + OCR。"""
    with open(image_path, 'rb') as f:
        img_bytes = f.read()

    img = Image.open(BytesIO(img_bytes))
    boxes = get_det_ocr().detection(img_bytes) or []
    detections = []
    for box in boxes:
        x1, y1, x2, y2 = box
        crop = img.crop((x1, y1, x2, y2))
        buf = BytesIO()
        crop.save(buf, format='PNG')
        char = get_cls_ocr().classification(buf.getvalue()).strip()
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        detections.append({
            'char': char,
            'cx': cx,
            'cy': cy,
            'crop': crop,
            'source': 'ddddocr',
        })
        print(f'ddddocr 检测到字符: {char!r} @ ({cx:.1f}, {cy:.1f})')
    return detections


def identify_text_click_positions(image_path, target_chars):
    """按目标字逐步锁定坐标：RapidOCR → ddddocr 回退 → 相似度补全。"""
    if not target_chars:
        dets = _collect_rapid_detections(image_path) or _collect_ddddocr_detections(image_path)
        return [(d['cx'], d['cy']) for d in dets if d.get('char')]

    n = len(target_chars)
    positions = [None] * n

    # 1) RapidOCR 精确匹配：命中即锁定
    rapid_dets = []
    used_rapid = set()
    try:
        rapid_dets = _collect_rapid_detections(image_path)
    except Exception as e:
        print('RapidOCR 识别异常:', e)

    for i, target in enumerate(target_chars):
        idx = _exact_match_one(target, rapid_dets, used_rapid)
        if idx is not None:
            used_rapid.add(idx)
            det = rapid_dets[idx]
            positions[i] = (det['cx'], det['cy'])
            print(f'锁定(RapidOCR): {target} @ ({det["cx"]:.1f}, {det["cy"]:.1f})')

    missing = [i for i, p in enumerate(positions) if p is None]
    if not missing:
        return positions

    print(f'RapidOCR 未命中目标: {[target_chars[i] for i in missing]}，回退 ddddocr')

    # 2) ddddocr 精确匹配缺失目标
    dddd_dets = []
    used_dddd = set()
    try:
        dddd_dets = _collect_ddddocr_detections(image_path)
    except Exception as e:
        print('ddddocr 识别异常:', e)

    # 已锁定坐标附近的 ddddocr 框视为已占用，避免重复点同一字
    for idx, det in enumerate(dddd_dets):
        if any(_is_near_position(det, p) for p in positions if p is not None):
            used_dddd.add(idx)

    for i in missing:
        target = target_chars[i]
        idx = _exact_match_one(target, dddd_dets, used_dddd)
        if idx is not None:
            used_dddd.add(idx)
            det = dddd_dets[idx]
            positions[i] = (det['cx'], det['cy'])
            print(f'锁定(ddddocr): {target} @ ({det["cx"]:.1f}, {det["cy"]:.1f})')

    missing = [i for i, p in enumerate(positions) if p is None]
    if not missing:
        return positions

    print(f'仍缺失目标: {[target_chars[i] for i in missing]}，相似度匹配剩余备选')

    # 3) 相似度：合并两路未使用备选
    candidates = []
    for idx, det in enumerate(rapid_dets):
        if idx in used_rapid:
            continue
        if any(_is_near_position(det, p) for p in positions if p is not None):
            continue
        candidates.append(det)
    for idx, det in enumerate(dddd_dets):
        if idx in used_dddd:
            continue
        if any(_is_near_position(det, p) for p in positions if p is not None):
            continue
        if any(_is_near_position(det, (c['cx'], c['cy'])) for c in candidates):
            continue
        candidates.append(det)

    used_cand = set()
    for i in missing:
        target = target_chars[i]
        cand_idx, sim = _similarity_match_from_candidates(
            target, candidates, used_cand)
        if cand_idx is None:
            print(f'相似度仍未匹配目标字: {target}')
            return []
        used_cand.add(cand_idx)
        det = candidates[cand_idx]
        positions[i] = (det['cx'], det['cy'])
        print(
            f'锁定(相似度): {target} @ ({det["cx"]:.1f}, {det["cy"]:.1f}) '
            f'相似度 {sim:.3f}')

    if any(p is None for p in positions):
        return []
    return positions


def click_text_positions(driver, img_el, positions, image_size):
    """在验证码图片上依次点击。"""
    img_w, img_h = image_size
    display_w = img_el.size['width']
    display_h = img_el.size['height']
    for cx, cy in positions:
        x = cx * display_w / img_w
        y = cy * display_h / img_h
        driver.execute_script(
            """
            var img = arguments[0], x = arguments[1], y = arguments[2];
            var rect = img.getBoundingClientRect();
            var evt = new MouseEvent('click', {
                view: window, bubbles: true, cancelable: true,
                clientX: rect.left + x, clientY: rect.top + y
            });
            img.dispatchEvent(evt);
            """,
            img_el, x, y,
        )
        time.sleep(random.uniform(0.4, 0.9))
    time.sleep(random.uniform(1.5, 2.5))


def drag_slider(driver, distance):
    huakuai = driver.find_element(By.CLASS_NAME, 'verify-move-block')
    action = ActionChains(driver)
    action.click_and_hold(huakuai).perform()
    time.sleep(random.randint(1, 2))
    print('滑块处理，目标距离:', distance)
    track = gettrack(distance)
    for x in track:
        action.move_by_offset(xoffset=x, yoffset=0).perform()
        print(x, huakuai.location['x'])
        time.sleep(0.01)
        action = ActionChains(driver)
    action.release().perform()
    time.sleep(random.randint(2, 4))


def solve_block_puzzle(driver, save_dir=None):
    try:
        wait_for_block_puzzle(driver)
    except Exception as e:
        print('未找到 blockPuzzle 验证码元素:', e)
        return False

    initial_url = driver.current_url
    if save_dir:
        hkpath, qkpath = save_captcha_images(driver, save_dir)
        result_path = os.path.join(save_dir, 'result.txt')
    else:
        hkpath, qkpath = save_captcha_images(
            driver, os.path.join('.', 'captcha_debug', 'tmp'))
        result_path = None

    gap_x, match_res, distance = identify_gap_distance(hkpath, qkpath)
    if distance <= 0:
        print('识别距离无效')
        return False

    if result_path:
        with open(result_path, 'w', encoding='utf-8') as f:
            f.write(f'gap_x={gap_x}\n')
            f.write(f'match_res={match_res}\n')
            f.write(f'distance={distance}\n')
            f.write(f'hkpath={hkpath}\n')
            f.write(f'qkpath={qkpath}\n')

    drag_slider(driver, distance)
    return is_captcha_success(driver, initial_url)


def solve_text_click_captcha(driver, save_dir=None):
    """处理文字点选验证码（verifybox + verify-msg）。"""
    try:
        wait_for_text_click_captcha(driver)
    except Exception as e:
        print('未找到文字点选验证码元素:', e)
        return False

    initial_url = driver.current_url
    msg_el = driver.find_element(By.CLASS_NAME, 'verify-msg')
    targets = parse_text_click_targets(msg_el.text)
    if not targets:
        print('未能解析点选目标:', msg_el.text)
        return False

    print(f'点选目标: {targets}')
    if not save_dir:
        save_dir = os.path.join('.', 'captcha_debug', 'text_click')

    image_path, img_el = save_text_click_image(driver, save_dir)
    positions = identify_text_click_positions(image_path, targets)
    if not positions:
        print('未能识别点选坐标')
        return False

    with open(os.path.join(save_dir, 'text_click_result.txt'), 'w', encoding='utf-8') as f:
        f.write(f'targets={targets}\n')
        f.write(f'positions={positions}\n')

    img = Image.open(image_path)
    click_text_positions(driver, img_el, positions, img.size)
    return is_text_click_success(driver, initial_url)


def preload_click_captcha():
    """预加载点选引擎（GPU 优先：RapidOCR + ddddocr）。"""
    prefer_gpu()  # 尽早探测并打印设备
    rapid = get_rapid_ocr()
    get_det_ocr()
    get_cls_ocr()
    get_slide_ocr()
    device = 'GPU' if prefer_gpu() else 'CPU'
    if rapid is not None:
        print(f'点选模型已预加载：RapidOCR + ddddocr（{device}）')
    else:
        print(f'点选模型已预加载：ddddocr（{device}，RapidOCR 不可用）')
    return True


def preload_text_select_captcha():
    """兼容旧名，等价于 preload_click_captcha。"""
    return preload_click_captcha()


def solve_captcha(driver, save_dir=None, download_dir=None, before_files=None):
    """按验证码类型自动选择处理方式。"""
    captcha_type = get_captcha_type(driver)
    if not captcha_type:
        if is_cnki_verify_done_page(driver) or _tencent_success_ui_visible(driver):
            print('验证码已完成')
            return True
        if not has_visible_captcha_widget(driver):
            print('未找到可交互验证码控件，按已通过处理')
            return True
        print('未知验证码类型')
        return False
    if captcha_type == 'tencent_slider':
        return solve_tencent_slider(
            driver, save_dir, download_dir=download_dir, before_files=before_files)
    if captcha_type == 'slider':
        return solve_block_puzzle(driver, save_dir)
    if captcha_type == 'text_click':
        return solve_text_click_captcha(driver, save_dir)
    print('未知验证码类型')
    return False
