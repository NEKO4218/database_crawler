'''
CNKI论文pdf下载工具

download_dir 为下载地址和滑块图片。

cnkipaperdownpdf：第一版的下载方式，将时间段所有文件下载到指定地址download_dir。

cnkipaperdownpdfv2：增加选择学术期刊等的下载方式。

cnkipaperdownpdfv3：增加功能点：根据搜索结果下载到所属学科目录下,增加北大核心筛选。

上述三个版本，实际上是同一个流程，根据需求可以通过注释来选择使用对应功能。

1、需要设置路径download_dir，改为本地需要下载的路径，图片解析下载的路径等。
2、可以设置时间参数，bgdate为初始日期，enddate为结束日期。
'''

from selenium import webdriver
import re
import time
import random
import gc
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from PIL import Image
from selenium.webdriver import ActionChains
from io import BytesIO
import base64
import ddddocr
import os
import requests
import math
from captcha_recognizer.recognizer import Recognizer

# OCR / 缺口识别单例，避免每次验证码新建模型占内存
_slide_ocr = None
_gap_recognizer = None


def _get_slide_ocr():
    global _slide_ocr
    if _slide_ocr is None:
        _slide_ocr = ddddocr.DdddOcr(det=False, ocr=False, show_ad=False)
    return _slide_ocr


def _get_gap_recognizer():
    global _gap_recognizer
    if _gap_recognizer is None:
        _gap_recognizer = Recognizer()
    return _gap_recognizer


def _build_chrome_options(download_dir):
    """创建带内存限制参数的 Chrome Options。"""
    chrome_options = Options()
    prefs = {
        'profile.default_content_settings.popups': 0,
        'download.default_directory': download_dir,
        'directory_upgrade': True,
    }
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--disk-cache-size=1')
    chrome_options.add_argument('--disable-gpu-shader-disk-cache')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--no-default-browser-check')
    return chrome_options


def _release_memory():
    gc.collect()

try:
    from cnkiforexcel import (
        ensure_cnki_captcha_after_nav,
        ensure_unit_login_if_needed,
        get_adv_search_url,
        open_cnki_url,
    )
except Exception:
    ensure_cnki_captcha_after_nav = None
    ensure_unit_login_if_needed = None
    get_adv_search_url = None
    open_cnki_url = None


def _open_cnki_adv_search(driver, sleep_sec=2, save_dir=None):
    """打开高级检索页，并按配置完成 VPN / 机构登录，随后处理验证码。"""
    url = (
        get_adv_search_url() if callable(get_adv_search_url)
        else 'https://kns.cnki.net/kns8s/AdvSearch')
    if callable(open_cnki_url):
        open_cnki_url(driver, url)
    else:
        driver.get(url)
    try:
        driver.maximize_window()
    except Exception:
        pass
    if sleep_sec:
        time.sleep(sleep_sec)
    if callable(ensure_unit_login_if_needed):
        try:
            ensure_unit_login_if_needed(
                driver, restore_handle=driver.current_window_handle)
        except Exception as e:
            print('机构/VPN 登录失败:', e)
    if callable(ensure_cnki_captcha_after_nav):
        try:
            ensure_cnki_captcha_after_nav(
                driver, save_dir or '.', driver.current_window_handle)
        except Exception as e:
            print('登录后验证码处理失败:', e)


global next_flag
next_flag=False
global bgdate
global initdata
bgdate = '2024-01-01'
enddate = '2024-07-25'

##########从这里开始
##########滑块验证码，更新模式

def base64_to_image(base64_string, savepath):
    image = None
    try:
        imagedata = base64.b64decode(base64_string)
        image = Image.open(BytesIO(imagedata))
        image.save(savepath)
    except Exception as e:
        print(e)
    finally:
        if image is not None:
            try:
                image.close()
            except Exception:
                pass


def url_to_image(url, savepath):
    image = None
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        image = Image.open(BytesIO(response.content))
        image.save(savepath)
    except Exception as e:
        print(e)
    finally:
        if image is not None:
            try:
                image.close()
            except Exception:
                pass

def gettrack(distance):
        # 移动轨迹
        track = []
        # 当前位移
        current = 0
        # 减速阈值
        mid = distance * 3 / 5
        # 计算间隔
        t = 0.2
        # 初速度
        v = 0

        while current + 3 < distance:
            if current < mid:
                # 加速度为正6
                a = 6
            else:
                # 加速度为负9
                a = -9
            # 初速度v0
            v0 = v
            # 当前速度v = v0 + at
            v = v0 + a * t
            # 移动距离x = v0t + 1/2 * a * t^2
            move = v0 * t + 1 / 2 * a * t * t
            # 当前位移
            current += round(move)
            # 加入轨迹
            track.append(round(move))
        sum = 0
        if distance > current:
            end = distance - current
            track.append(round(end))
        for i in range(0, len(track)):
            sum = sum + track[i]
        # print(track)
        # print(sum)
        return track

def distantcon(hkpath, qkpath):
    det = _get_slide_ocr()
    with open(hkpath, 'rb') as f:
        target_bytes = f.read()
    with open(qkpath, 'rb') as f:
        background_bytes = f.read()
    res = det.slide_match(target_bytes, background_bytes, simple_target=True)
    print(res)
    target = res.get('target')[0]
    print(target)
    rendersize = target * 400 / 310
    print('rendersize:', rendersize)
    del target_bytes, background_bytes
    return rendersize

def distant_trans(target):
    return (10 * math.sqrt(14200 * target + 5929) - 770) / 71

def distantcon_v2(qkpath):
    recognizer = _get_gap_recognizer()
    box, confidence = recognizer.identify_gap(source=qkpath, verbose=False)
    target = 0
    try:
        target = box[0]
    except Exception:
        print('没有找到目标')
    print(f'Target position: {target}')
    rendersize = distant_trans(target)
    print(f'Render size: {rendersize}')
    return rendersize


def _cleanup_captcha_images(*paths):
    for p in paths:
        try:
            if p and os.path.isfile(p):
                os.remove(p)
        except Exception:
            pass

hkpathroot='D\\2024年工作\知网数据库下载\\'
qkpathroot='D\\2024年工作\知网数据库下载\\'

def huakuaicon(driver,pdfpath):
    #先找一下有没有滑块
    try:
        blockeles = driver.find_elements(By.CLASS_NAME, 'verify-sub-block')
        if len(blockeles) < 1:
            print("没有找到滑块")
            return False
    except Exception as e:
        print('没有找到滑块',e)
        return False
    #1、先保存滑块图片
    #2、保存缺口图片
    hkimg = driver.find_element(By.CLASS_NAME, 'verify-sub-block') \
        .find_element(By.CSS_SELECTOR, 'img').get_attribute('src')
    hkimg = hkimg.replace('data:image/png;base64,', '')

    qkimg = driver.find_element(By.CLASS_NAME, 'verify-img-panel') \
        .find_element(By.CSS_SELECTOR, 'img').get_attribute('src')
    qkimg = qkimg.replace('data:image/png;base64,', '')

    #3保存图片
    #hkpath=r'D:\2024年工作\知网数据库下载\图片\滑块图片.png'
    #qkpath=r'D:\2024年工作\知网数据库下载\图片\缺口图片.png'
    #picpath=hkpathroot+'\\'+keywords+'\pdf'+"\图片"
    picpath = pdfpath + "\图片"

    if not os.path.exists(picpath):
        os.makedirs(picpath)
    hkpath = picpath + '\滑块图片.png'
    qkpath = picpath + '\缺口图片.png'

    base64_to_image(hkimg, hkpath)
    base64_to_image(qkimg, qkpath)
    distant = distantcon(hkpath, qkpath)
    _cleanup_captcha_images(hkpath, qkpath)
    huakuai = driver.find_element(By.CLASS_NAME, "verify-move-block")
    action = ActionChains(driver)
    action.click_and_hold(huakuai).perform()  # 按住滑块
    time.sleep(random.randint(1, 2))
    print('滑块处理')
    track = gettrack(distant)
    for x in track:
         action.move_by_offset(xoffset=x, yoffset=0).perform()
         print(x, huakuai.location['x'])
         time.sleep(0.01)
         action = ActionChains(driver)  # 重要，不然，offset会累加导致移动的距离不对
    action.release().perform()
    time.sleep(random.randint(2, 4))
    _release_memory()
    return True

def huakuaicon_v2(driver, pdfpath):
    # 先找一下有没有滑块
    try:
        blockeles = driver.find_elements(By.ID, 'aliyunCaptcha-sliding-slider')
        if len(blockeles) < 1:
            print("没有找到滑块")
            return False
    except Exception as e:
        print('没有找到滑块', e)
        return False

    # 保存滑块图片和缺口图片
    hkimg = driver.find_element(By.ID, 'aliyunCaptcha-puzzle').get_attribute('src')
    qkimg = driver.find_element(By.ID, 'aliyunCaptcha-img').get_attribute('src')

    # 保存图片
    picpath = pdfpath + "\图片"

    if not os.path.exists(picpath):
        os.makedirs(picpath)
    hkpath = picpath + '\滑块图片.png'
    qkpath = picpath + '\缺口图片.png'

    url_to_image(hkimg, hkpath)
    url_to_image(qkimg, qkpath)

    distance = distantcon_v2(qkpath)
    _cleanup_captcha_images(hkpath, qkpath)
    huakuai = driver.find_element(By.ID, "aliyunCaptcha-sliding-slider")
    action = ActionChains(driver)
    action.click_and_hold(huakuai).perform()  # 按住滑块
    time.sleep(random.randint(1, 2))
    print('滑块处理')
    distance = round(distance)
    print('滑块距离：', distance)
    track = gettrack(distance)
    for x in track:
        action.move_by_offset(xoffset=x, yoffset=0).perform()
        print(x, huakuai.location['x'])
        time.sleep(0.01)
        action = ActionChains(driver)  # 重要，不然，offset会累加导致移动的距离不对
    action.release().perform()
    time.sleep(random.randint(2, 4))
    _release_memory()
    return True

def downloadpdf(driver, pdfpath):  
    buttons = driver.find_elements(By.CLASS_NAME, 'name')  
    i = 0  
    while i < len(buttons):  
        print('\n该页第%s篇文章下载' % (i + 1))  
        ele = buttons[i]  
        try:  
            ele_a = ele.find_element(By.CSS_SELECTOR, 'a')  
            ele_a.click()  
        except Exception:  
            print('没有找到下载按钮')
        time.sleep(random.randint(3, 5))  
        windows = driver.window_handles  
        print('当前窗口个数：', len(windows))  
        if len(windows) > 1:  
            driver.switch_to.window(driver.window_handles[1]) 
            downpdfele = None
            try: 
                downpdfele = driver.find_element(By.CLASS_NAME, 'operate-btn').find_elements(By.ID, 'pdfDown')  
            except Exception:  
                print('没有下载pdf按钮')

            if downpdfele is not None and len(downpdfele) > 0:  # 优先下载pdf  
                retry_count = 0  
                max_retries = 3  # 最大重试次数  
                flag = False  
                while not flag and retry_count < max_retries:  # 下载最多尝试3次  
                    try:  
                        downpdfele[0].click()  
                        time.sleep(random.randint(3, 5))  
                        flag = True  
                    except Exception as e:  
                        retry_count += 1  
                        print(f'下载PDF失败，第{retry_count}次尝试: {e}')  

                if retry_count == max_retries:  
                    print('达到最大重试次数，跳过当前PDF下载')  

            else:  
                downcajele = driver.find_elements(By.CLASS_NAME, 'btn-dlcaj')  
                if len(downcajele) > 1:  # 优先下载caj  
                    downcajele[0].click()  
                    time.sleep(5)  

            time.sleep(random.randint(2, 4))  
            windows = driver.window_handles  # 是否需要滑块验证码  
            print('点击pdf下载链接后窗口个数：', len(windows))  
            if len(windows) > 2:  
                print('*********可能需要滑块处理**********')  
                driver.switch_to.window(windows[2])  
                print('currenturl:', driver.current_url)  
                time.sleep(random.randint(2, 4))  
                result = huakuaicon_v2(driver, pdfpath)

                time.sleep(random.randint(2, 4))  
                try:  
                    driver.close()  
                except Exception:  
                    pass
                driver.switch_to.window(driver.window_handles[1])  
                if result is True:  
                    i = i - 1  # If success, retry current item  

            try:
                driver.close()  
            except Exception:  
                pass
            # 只保留主检索窗，关掉多余标签
            try:
                main = driver.window_handles[0]
                for h in list(driver.window_handles):
                    if h != main:
                        driver.switch_to.window(h)
                        driver.close()
                driver.switch_to.window(main)
            except Exception:
                try:
                    driver.switch_to.window(driver.window_handles[0])
                except Exception:
                    pass
        i += 1
        if i % 5 == 0:
            _release_memory()


#download_dir=r'D:\2024年工作\知网数据库下载\\'
#1.0版本，没有将学科分类下载到指定的文件夹
def cnkipaperdownpdf(keywords, keywords2):
        global enddate
        global next_flag
        #download_dir = f'D://2024年工作/知网数据库下载/{keywords}/'  # 路径一定是斜杠，不能是反斜杠
        download_dir = r'D:\2024年工作\知网数据库下载\{}\pdf'.format(keywords)  
        print(download_dir)
        if not os.path.exists(download_dir):
            # 如果文件夹不存在，则创建文件夹
            os.makedirs(download_dir)

        chrome_options = _build_chrome_options(download_dir)
        driver = webdriver.Chrome(options=chrome_options)

        # driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
        # driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
        _open_cnki_adv_search(driver, sleep_sec=random.randint(2, 4))

        select_eleddlist = driver.find_element(By.ID, 'gradetxt').find_elements(By.CSS_SELECTOR, 'dd')
        # 获取第一个选择项（保持不变）
        sel_dd_1 = select_eleddlist[0]
        dd_cli_1 = sel_dd_1.find_element(By.CLASS_NAME, 'sort.reopt')
        dd_cli_1.click()
        time.sleep(random.randint(2, 4))

        # 选择第一个下拉选项: 主题（保持不变）
        sel_li_1 = dd_cli_1.find_element(By.CSS_SELECTOR, '.sort-list li[data-val="SU"]')
        print(sel_li_1.get_attribute('outerHTML'))
        try:
            sel_li_1.click()
        except:
            print('没有找到元素')
        time.sleep(random.randint(2, 4))

        # 输入关键词到第一个文本框（保持不变）
        text_element_1 = sel_dd_1.find_element(By.CSS_SELECTOR, 'input[type="text"]')
        print('text_element_1:', text_element_1.get_attribute('outerHTML'))
        time.sleep(random.randint(2, 4))
        text_element_1.send_keys(keywords)
        time.sleep(random.randint(2, 4))

        # 修改逻辑操作为 NOT
        logical_sort = driver.find_element(By.CSS_SELECTOR, '.gradeSearch .sort.logical')
        logical_sort_default = logical_sort.find_element(By.CLASS_NAME, 'sort-default')
        logical_sort_default.click()
        time.sleep(random.randint(2, 4))

        not_option = logical_sort.find_element(By.CSS_SELECTOR, '.sort-list li:last-child a')
        print(not_option.get_attribute('outerHTML'))
        try:
            not_option.click()
        except:
            print('没有找到元素')
        time.sleep(random.randint(2, 4))

        # 获取第二个选择项
        sel_dd_2 = select_eleddlist[1]
        dd_cli_2 = sel_dd_2.find_element(By.CLASS_NAME, 'sort.reopt')
        dd_cli_2.click()
        time.sleep(random.randint(2, 4))

        # 选择第二个下拉选项: 关键词
        sel_li_2 = dd_cli_2.find_element(By.CSS_SELECTOR, '.sort-list li[data-val="SU"]')
        print(sel_li_2.get_attribute('outerHTML'))
        try:
            sel_li_2.click()
        except:
            print('没有找到元素')
        time.sleep(random.randint(2, 4))

        # 输入关键词到第二个文本框
        text_element_2 = sel_dd_2.find_element(By.CSS_SELECTOR, 'input[type="text"]')
        print('text_element_2:', text_element_2.get_attribute('outerHTML'))
        time.sleep(random.randint(2, 4))
        text_element_2.send_keys(keywords2)
        time.sleep(random.randint(2, 4))

        # 去掉英文文献
        input = driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
        try:
            input.click()
        except:
            print('没有找到元素')
        time.sleep(10)
        # 填写开始时间
        times = driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
        # start_time=times[0]#.send_keys('2011-01-01')
        # start_time.send_keys('2011-01-01')
        date_txt_bg = driver.find_element(By.ID, 'datebox0')
        driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
        date_txt_bg.clear()  # 先清除原来的日期值
        date_txt_bg.send_keys(bgdate)
        # date_txt_bg.send_keys(enddate)
        # print('更新开始论文时间：',enddate)

        date_txt_end = driver.find_element(By.ID, 'datebox1')
        driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
        date_txt_end.clear()  # 先清除原来的日期值
        # date_txt.send_keys('2010-01-01')
        date_txt_end.send_keys(enddate)
        print('更新结束论文时间：', enddate)

        print(times[0].get_attribute('outerHTML'))
        time.sleep(10)
        # start_time = time[0].send_keys('2011-01-01')
        searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
        # searchbtn.click()
        driver.execute_script("arguments[0].click();", searchbtn)
        time.sleep(10)
        now_url = driver.current_url  # 当前页面
        print('主页:', now_url)

        # 只采集核心
        print('只选择学术期刊')
        journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
        journal.click()
        time.sleep(random.randint(2, 4))



        # # 只采集核心
        # print('每个选项选择学科属性')
        # cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        # ccllist=cclele.find_elements(By.CSS_SELECTOR, 'li')
        # #for j in range(0,len(ccllist)):
        # j=0
        # while j<len(ccllist):
        #     if j > 9:
        #         driver.execute_script("arguments[0].style.display = 'list-item';", ccllist[j])
        #     #print(ccllist[j].get_attribute('outerHTML'))
        #     checkbox=ccllist[j].find_element(By.CSS_SELECTOR, 'input')
        #     print(j,checkbox.get_attribute('title'))
        #              #print(checkbox.get_attribute('outerHTML'))
        #     checkbox.click() #点击选中
        #     time.sleep(10)
        #     #需要再取消一下,需要重新获取一下当前的节点
        #     cclelev2 = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        #     ccllistv2 = cclelev2.find_elements(By.CSS_SELECTOR, 'input[checked="checked"]')
        #     #checkbox = ccllistv2[0].find_element(By.CSS_SELECTOR, 'input')
        #     ccllistv2[0].click()
        #     time.sleep(2)
        #
        #     #重新获取一下自选框,取消完之后重新获取
        #     cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        #     ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
        #     j=j+1


        #driver.window_handles()
        main_handle = driver.current_window_handle
        nextbutton = True
        count = 1
        page_count = 0



        while nextbutton == True:
            nextbutton = False
            print("===========第%s页===========\n" % count)
            #下载pdf
            downloadpdf(driver,keywords)
            # 3、翻页
            try:
                next_ele = driver.find_element(By.ID, 'Page_next_top')
                # next_ele.click()
                driver.execute_script("arguments[0].click();", next_ele)
                nextbutton = True
                time.sleep(random.randint(2, 4))
            except Exception as e:
                print("翻页失败", e)
                next_flag = True
                nextbutton = False
            count = count + 1

        try:
            driver.quit()
        except Exception:
            pass
        _release_memory()




#设置学科文件夹
def rename_folder(old_path, new_name):
    # 构造新的文件夹路径  
    new_path = os.path.join(os.path.dirname(old_path), new_name)  
    
    # 修改文件夹名  
    if os.path.exists(old_path):  
        os.rename(old_path, new_path)  
    else:  
        print(f"文件夹 {old_path} 不存在，无法重命名。")  
def setsubcon(driver,pdfpath,subject):
    #driver.window_handles()
        main_handle = driver.current_window_handle
        nextbutton = True
        count = 1
        page_count = 0
        while nextbutton == True:
            nextbutton = False
            print("===========第%s页===========\n" % count)
            #下载pdf
            downloadpdf(driver,pdfpath)
            # 3、翻页
            try:
                #next_ele = driver.find_element(By.ID, 'Page_next_top')
                # next_ele.click()
                #driver.execute_script("arguments[0].click();", next_ele)
                next_ele = driver.find_element(By.ID, 'PageNext')
                next_ele.click()
                nextbutton = True
                time.sleep(5)
            except Exception as e:
                print("翻页失败", e)
                next_flag = True
                nextbutton = False
            count = count + 1

        #下载完成
        #pdfpath='D\\2024年工作\知网数据库下载\\'+keywords+'\\pdf'
        rename_folder(pdfpath, subject)


# download_dir='D\\2024年工作\知网数据库下载\\'
def cnkipaperdownpdfv2(keywords):#2024.6.18完成，增加学科分类
    global enddate
    global next_flag
    # download_dir = f'D://2024年工作/知网数据库下载/{keywords}/'  # 路径一定是斜杠，不能是反斜杠
    pdfpath = 'D\\2024年工作\知网数据库下载'+'\\'+ keywords+"\\pdf"
    print(pdfpath)
    if not os.path.exists(pdfpath):
        # 如果文件夹不存在，则创建文件夹
        os.makedirs(pdfpath)

    chrome_options = _build_chrome_options(pdfpath)
    driver = webdriver.Chrome(options=chrome_options)

    # driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
    # driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
    _open_cnki_adv_search(driver, sleep_sec=5)
    select_eleddlist = driver.find_element(By.ID, 'gradetxt').find_elements(By.CSS_SELECTOR, 'dd')
    # 获取第一个选择项
    sel_dd = select_eleddlist[0]
    dd_cli = sel_dd.find_element(By.CLASS_NAME, 'sort.reopt')
    dd_cli.click()
    time.sleep(5)
    # 获取下拉选项:文献来源
    # sel_li=dd_cli.find_element(By.CLASS_NAME, 'sort-list').find_element(By.CSS_SELECTOR, 'li[data-val="LY"]')
    sel_li = dd_cli.find_element(By.CLASS_NAME, 'sort-list').find_element(By.CSS_SELECTOR, 'li[data-val="SU"]')
    print(sel_li.get_attribute('outerHTML'))
    sel_li.click()
    time.sleep(5)
    text_element = sel_dd.find_element(By.CSS_SELECTOR, 'input[type="text"]')
    print('text_element:', text_element.get_attribute('outerHTML'))
    time.sleep(5)
    text_element.send_keys(keywords)
    time.sleep(5)

    # 去掉英文文献
    input = driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
    input.click()
    time.sleep(10)
    # 填写开始时间
    times = driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
    # start_time=times[0]#.send_keys('2011-01-01')
    # start_time.send_keys('2011-01-01')
    date_txt_bg = driver.find_element(By.ID, 'datebox0')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
    date_txt_bg.clear()  # 先清除原来的日期值
    date_txt_bg.send_keys(bgdate)
    # date_txt_bg.send_keys(enddate)
    # print('更新开始论文时间：',enddate)

    date_txt_end = driver.find_element(By.ID, 'datebox1')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
    date_txt_end.clear()  # 先清除原来的日期值
    # date_txt.send_keys('2010-01-01')
    date_txt_end.send_keys(enddate)
    print('更新结束论文时间：', enddate)

    print(times[0].get_attribute('outerHTML'))
    time.sleep(10)
    # start_time = time[0].send_keys('2011-01-01')
    searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
    # searchbtn.click()
    driver.execute_script("arguments[0].click();", searchbtn)
    time.sleep(10)
    now_url = driver.current_url  # 当前页面
    print('主页:', now_url)

    # 只采集学术期刊
    print('只选择学术期刊')
    journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
    journal.click()
    time.sleep(3)

    #只选择核心期刊
    print('只选择北大核心期刊')
    LYBSMele= driver.find_element(By.CSS_SELECTOR, 'dl[groupid="LYBSM"]')
    driver.execute_script("arguments[0].class= '  ';", LYBSMele)
    #LYBSMele.click()
    #time.sleep(5)


    #central=LYBSMele.find_element(By.CSS_SELECTOR, 'input[text="北大核心"]')
    #central.click()
    #time.sleep(3)


    #分学科进行采集
    print('每个选项选择学科属性')
    cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
    ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
    #aaa

    j = 0
    while j < len(ccllist):
        if j > 9:
            driver.execute_script("arguments[0].style.display = 'list-item';", ccllist[j])
        # print(ccllist[j].get_attribute('outerHTML'))
        checkbox = ccllist[j].find_element(By.CSS_SELECTOR, 'input')
        print("\n===========选择学科%s,名称：%s"%(j, checkbox.get_attribute('title')))
        subject=checkbox.get_attribute('title') #获取学科类别
        # print(checkbox.get_attribute('outerHTML'))
        checkbox.click()  # 点击选中
        time.sleep(3)

        #下载pdf文件
        setsubcon(driver, pdfpath, subject)

        #下载完成之后需要再取消一下,需要重新获取一下当前的节点
        cclelev2 = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        ccllistv2 = cclelev2.find_elements(By.CSS_SELECTOR, 'input[checked="checked"]')
        # checkbox = ccllistv2[0].find_element(By.CSS_SELECTOR, 'input')
        ccllistv2[0].click()
        time.sleep(2)

        # 重新获取一下自选框,取消完之后重新获取
        cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
        j = j + 1

    # driver.window_handles()
    # main_handle = driver.current_window_handle
    # nextbutton = True
    # count = 1
    # page_count = 0
    #
    # while nextbutton == True:
    #     nextbutton = False
    #     print("===========第%s页===========\n" % count)
    #     # 下载pdf
    #     downloadpdf(driver, keywords)
    #     # 3、翻页
    #     try:
    #         next_ele = driver.find_element(By.ID, 'Page_next_top')
    #         # next_ele.click()
    #         driver.execute_script("arguments[0].click();", next_ele)
    #         nextbutton = True
    #         time.sleep(5)
    #     except Exception as e:
    #         print("翻页失败", e)
    #         next_flag = True
    #         nextbutton = False
    #     count = count + 1
    #
    # driver.quit()


def cnkipaperdownpdfv3(keywords):#2024.6.18完成，增加学科分类
    global enddate
    global next_flag
    # download_dir = f'D://2024年工作/知网数据库下载/{keywords}/'  # 路径一定是斜杠，不能是反斜杠
    pdfpath = r'D:\2024年工作\知网数据库下载\{}\pdf'.format(keywords)  
    print(f"PDF 下载路径: {pdfpath}")  
    if not os.path.exists(pdfpath):
        # 如果文件夹不存在，则创建文件夹
        os.makedirs(pdfpath)

    chrome_options = _build_chrome_options(pdfpath)
    driver = webdriver.Chrome(options=chrome_options)

    # driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
    # driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
    _open_cnki_adv_search(driver, sleep_sec=5)
    select_eleddlist = driver.find_element(By.ID, 'gradetxt').find_elements(By.CSS_SELECTOR, 'dd')
    # 获取第一个选择项
    sel_dd = select_eleddlist[0]
    dd_cli = sel_dd.find_element(By.CLASS_NAME, 'sort.reopt')
    dd_cli.click()
    time.sleep(5)
    # 获取下拉选项:文献来源
    # sel_li=dd_cli.find_element(By.CLASS_NAME, 'sort-list').find_element(By.CSS_SELECTOR, 'li[data-val="LY"]')
    sel_li = dd_cli.find_element(By.CLASS_NAME, 'sort-list').find_element(By.CSS_SELECTOR, 'li[data-val="SU"]')
    print(sel_li.get_attribute('outerHTML'))
    # sel_li.click()
    time.sleep(5)
    text_element = sel_dd.find_element(By.CSS_SELECTOR, 'input[type="text"]')
    print('text_element:', text_element.get_attribute('outerHTML'))
    time.sleep(5)
    text_element.send_keys(keywords)
    time.sleep(5)

    # 去掉英文文献
    input = driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
    input.click()
    time.sleep(10)
    # 填写开始时间
    times = driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
    # start_time=times[0]#.send_keys('2011-01-01')
    # start_time.send_keys('2011-01-01')
    date_txt_bg = driver.find_element(By.ID, 'datebox0')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
    date_txt_bg.clear()  # 先清除原来的日期值
    date_txt_bg.send_keys(bgdate)
    # date_txt_bg.send_keys(enddate)
    # print('更新开始论文时间：',enddate)

    date_txt_end = driver.find_element(By.ID, 'datebox1')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
    date_txt_end.clear()  # 先清除原来的日期值
    # date_txt.send_keys('2010-01-01')
    date_txt_end.send_keys(enddate)
    print('更新结束论文时间：', enddate)

    print(times[0].get_attribute('outerHTML'))
    time.sleep(10)
    # start_time = time[0].send_keys('2011-01-01')
    searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
    # searchbtn.click()
    driver.execute_script("arguments[0].click();", searchbtn)
    time.sleep(10)
    now_url = driver.current_url  # 当前页面
    print('主页:', now_url)

    # 只采集学术期刊
    print('只选择学术期刊')
    journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
    journal.click()
    time.sleep(3)

    #只选择核心期刊
    print('只选择北大核心期刊')
    LYBSMele= driver.find_element(By.CSS_SELECTOR, 'dl[groupid="LYBSM"]')
    driver.execute_script("arguments[0].class= '  ';", LYBSMele)
    LYBSMele.click()
    time.sleep(5)
    #print(LYBSMele.get_attribute('outerHTML'))
    #print(LYBSMele.get_attribute('innerHTML'))
    central=LYBSMele.find_element(By.CSS_SELECTOR, 'input[text="北大核心"]')
    central.click()
    time.sleep(3)


    #分学科进行采集
    print('每个选项选择学科属性')
    cclele = driver.find_element(By.CSS_SELECTOR, 'dl[groupid="CCL"]')
    driver.execute_script("arguments[0].class= '  ';", cclele)
    cclele.click()
    time.sleep(5)
    #cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
    ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
    print(len(ccllist))


    j = 0
    while j < len(ccllist):
        if j > 9:
            driver.execute_script("arguments[0].style.display = 'list-item';", ccllist[j])
        # print(ccllist[j].get_attribute('outerHTML'))
        checkbox = ccllist[j].find_element(By.CSS_SELECTOR, 'input')
        print("\n===========选择学科%s,名称：%s"%(j, checkbox.get_attribute('title')))
        subject=checkbox.get_attribute('title') #获取学科类别
        # print(checkbox.get_attribute('outerHTML'))
        checkbox.click()  # 点击选中
        time.sleep(3)

        #下载pdf文件
        setsubcon(driver, pdfpath, subject)

        #下载完成之后需要再取消一下,需要重新获取一下当前的节点
        cclelev2 = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        ccllistv2 = cclelev2.find_elements(By.CSS_SELECTOR, 'input[checked="checked"]')
        # checkbox = ccllistv2[0].find_element(By.CSS_SELECTOR, 'input')
        ccllistv2[0].click()
        time.sleep(2)

        # 重新获取一下自选框,取消完之后重新获取
        cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
        j = j + 1

def cnkipaperdownpdfv4(keywords): # 手动设置检索条件
        global enddate
        global next_flag
        #download_dir = f'D://2024年工作/知网数据库下载/{keywords}/'  # 路径一定是斜杠，不能是反斜杠
        download_dir = r'D:\2024年工作\知网数据库下载\{}\pdf'.format(keywords)  
        print(download_dir)
        if not os.path.exists(download_dir):
            # 如果文件夹不存在，则创建文件夹
            os.makedirs(download_dir)

        chrome_options = _build_chrome_options(download_dir)
        driver = webdriver.Chrome(options=chrome_options)

        # driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
        # driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
        _open_cnki_adv_search(driver, sleep_sec=2)

        # 等待手动输入检索条件
        search_condition = input("请输入检索条件后按回车继续: ")

        # 去掉英文文献
        input_en = driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
        try:
            input_en.click()
        except:
            print('没有找到元素')
        time.sleep(10)
        # 填写开始时间
        times = driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
        # start_time=times[0]#.send_keys('2011-01-01')
        # start_time.send_keys('2011-01-01')
        date_txt_bg = driver.find_element(By.ID, 'datebox0')
        driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
        date_txt_bg.clear()  # 先清除原来的日期值
        date_txt_bg.send_keys(bgdate)
        # date_txt_bg.send_keys(enddate)
        # print('更新开始论文时间：',enddate)

        date_txt_end = driver.find_element(By.ID, 'datebox1')
        driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
        date_txt_end.clear()  # 先清除原来的日期值
        # date_txt.send_keys('2010-01-01')
        date_txt_end.send_keys(enddate)
        print('更新结束论文时间：', enddate)

        print(times[0].get_attribute('outerHTML'))
        time.sleep(10)
        # start_time = time[0].send_keys('2011-01-01')
        searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
        # searchbtn.click()
        driver.execute_script("arguments[0].click();", searchbtn)
        time.sleep(10)
        now_url = driver.current_url  # 当前页面
        print('主页:', now_url)

        # 只采集核心
        print('只选择学术期刊')
        journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
        journal.click()
        time.sleep(random.randint(2, 4))



        # # 只采集核心
        # print('每个选项选择学科属性')
        # cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        # ccllist=cclele.find_elements(By.CSS_SELECTOR, 'li')
        # #for j in range(0,len(ccllist)):
        # j=0
        # while j<len(ccllist):
        #     if j > 9:
        #         driver.execute_script("arguments[0].style.display = 'list-item';", ccllist[j])
        #     #print(ccllist[j].get_attribute('outerHTML'))
        #     checkbox=ccllist[j].find_element(By.CSS_SELECTOR, 'input')
        #     print(j,checkbox.get_attribute('title'))
        #              #print(checkbox.get_attribute('outerHTML'))
        #     checkbox.click() #点击选中
        #     time.sleep(10)
        #     #需要再取消一下,需要重新获取一下当前的节点
        #     cclelev2 = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        #     ccllistv2 = cclelev2.find_elements(By.CSS_SELECTOR, 'input[checked="checked"]')
        #     #checkbox = ccllistv2[0].find_element(By.CSS_SELECTOR, 'input')
        #     ccllistv2[0].click()
        #     time.sleep(2)
        #
        #     #重新获取一下自选框,取消完之后重新获取
        #     cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        #     ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
        #     j=j+1


        #driver.window_handles()
        main_handle = driver.current_window_handle
        nextbutton = True
        count = 1
        page_count = 0



        while nextbutton == True:
            nextbutton = False
            print("===========第%s页===========\n" % count)
            #下载pdf
            downloadpdf(driver,keywords)
            # 3、翻页
            try:
                next_ele = driver.find_element(By.ID, 'Page_next_top')
                # next_ele.click()
                driver.execute_script("arguments[0].click();", next_ele)
                nextbutton = True
                time.sleep(random.randint(2, 4))
            except Exception as e:
                print("翻页失败", e)
                next_flag = True
                nextbutton = False
            count = count + 1

        try:
            driver.quit()
        except Exception:
            pass
        _release_memory()

def cnkipaperdownpdfv5(keywords): # 手动设置检索条件
        global enddate
        global next_flag
        #download_dir = f'D://2024年工作/知网数据库下载/{keywords}/'  # 路径一定是斜杠，不能是反斜杠
        download_dir = r'D:\2024年工作\知网数据库下载\{}\pdf'.format(keywords[0])  
        print(download_dir)
        if not os.path.exists(download_dir):
            # 如果文件夹不存在，则创建文件夹
            os.makedirs(download_dir)

        chrome_options = _build_chrome_options(download_dir)
        driver = webdriver.Chrome(
            options=chrome_options,
            service=Service(
                r"D:\多模态\chromedriver-win32\chromedriver-win32\chromedriver.exe"))

        # driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
        # driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
        try:
            _open_cnki_adv_search(driver, sleep_sec=2)
        except Exception:
            print('没有找到元素')
        # 添加额外的检索框（如果需要）
        if len(keywords) > 3:
            for _ in range(len(keywords) - 3):
                add_button = driver.find_element(By.CSS_SELECTOR, '.icon-btn.add-group')
                add_button.click()
                time.sleep(random.randint(2, 4))  # 等待新元素加载

        # 获取所有的dd元素
        select_eleddlist = driver.find_element(By.ID, 'gradetxt').find_elements(By.CSS_SELECTOR, 'dd')

        # 填充第一个检索条件
        sel_dd_1 = select_eleddlist[0]
        dd_cli_1 = sel_dd_1.find_element(By.CLASS_NAME, 'sort.reopt')
        dd_cli_1.click()
        time.sleep(random.randint(2, 4))

        # 选择第一个下拉选项: 主题
        sel_li_1 = dd_cli_1.find_element(By.CSS_SELECTOR, '.sort-list li[data-val="SU"]')
        print(sel_li_1.get_attribute('outerHTML'))
        try:
            sel_li_1.click()
        except:
            print('没有找到元素')
        time.sleep(random.randint(2, 4))

        # 输入关键词到第一个文本框
        text_element_1 = sel_dd_1.find_element(By.CSS_SELECTOR, 'input[type="text"]')
        print('text_element_1:', text_element_1.get_attribute('outerHTML'))
        time.sleep(random.randint(2, 4))
        text_element_1.send_keys(keywords[0])
        time.sleep(random.randint(2, 4))

        # 填充其他检索条件
        for i in range(1, len(keywords)):
            sel_dd_i = select_eleddlist[i]

            # 修改逻辑操作为 NOT
            logical_sort = sel_dd_i.find_element(By.CSS_SELECTOR, '.sort.logical')
            logical_sort_default = logical_sort.find_element(By.CLASS_NAME, 'sort-default')
            logical_sort_default.click()
            # time.sleep(random.randint(2, 4))

            not_option = logical_sort.find_element(By.CSS_SELECTOR, '.sort-list li:last-child a')
            print(not_option.get_attribute('outerHTML'))
            try:
                not_option.click()
            except:
                print('没有找到元素')

            # 设置字段类型为“主题”
            dd_cli_i = sel_dd_i.find_element(By.CLASS_NAME, 'sort.reopt')
            dd_cli_i.click()
            # time.sleep(random.randint(2, 4))

            sel_li_i = dd_cli_i.find_element(By.CSS_SELECTOR, '.sort-list li[data-val="SU"]')
            print(sel_li_i.get_attribute('outerHTML'))
            try:
                sel_li_i.click()
            except:
                print('没有找到元素')

            # 输入关键词到新的文本框
            text_element_i = sel_dd_i.find_element(By.CSS_SELECTOR, 'input[type="text"]')
            print('text_element_i:', text_element_i.get_attribute('outerHTML'))
            # time.sleep(random.randint(2, 4))
            text_element_i.send_keys(keywords[i])
            time.sleep(random.randint(2, 4))

        # 去掉英文文献
        input_en = driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
        try:
            input_en.click()
        except:
            print('没有找到元素')
        time.sleep(10)
        # 填写开始时间
        times = driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
        # start_time=times[0]#.send_keys('2011-01-01')
        # start_time.send_keys('2011-01-01')
        date_txt_bg = driver.find_element(By.ID, 'datebox0')
        driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
        date_txt_bg.clear()  # 先清除原来的日期值
        date_txt_bg.send_keys(bgdate)
        # date_txt_bg.send_keys(enddate)
        # print('更新开始论文时间：',enddate)

        date_txt_end = driver.find_element(By.ID, 'datebox1')
        driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
        date_txt_end.clear()  # 先清除原来的日期值
        # date_txt.send_keys('2010-01-01')
        date_txt_end.send_keys(enddate)
        print('更新结束论文时间：', enddate)

        print(times[0].get_attribute('outerHTML'))
        time.sleep(10)
        # start_time = time[0].send_keys('2011-01-01')
        searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
        # searchbtn.click()
        driver.execute_script("arguments[0].click();", searchbtn)
        time.sleep(10)
        now_url = driver.current_url  # 当前页面
        print('主页:', now_url)

        # 只采集核心
        print('只选择学术期刊')
        journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
        journal.click()
        time.sleep(random.randint(2, 4))



        # # 只采集核心
        # print('每个选项选择学科属性')
        # cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        # ccllist=cclele.find_elements(By.CSS_SELECTOR, 'li')
        # #for j in range(0,len(ccllist)):
        # j=0
        # while j<len(ccllist):
        #     if j > 9:
        #         driver.execute_script("arguments[0].style.display = 'list-item';", ccllist[j])
        #     #print(ccllist[j].get_attribute('outerHTML'))
        #     checkbox=ccllist[j].find_element(By.CSS_SELECTOR, 'input')
        #     print(j,checkbox.get_attribute('title'))
        #              #print(checkbox.get_attribute('outerHTML'))
        #     checkbox.click() #点击选中
        #     time.sleep(10)
        #     #需要再取消一下,需要重新获取一下当前的节点
        #     cclelev2 = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        #     ccllistv2 = cclelev2.find_elements(By.CSS_SELECTOR, 'input[checked="checked"]')
        #     #checkbox = ccllistv2[0].find_element(By.CSS_SELECTOR, 'input')
        #     ccllistv2[0].click()
        #     time.sleep(2)
        #
        #     #重新获取一下自选框,取消完之后重新获取
        #     cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        #     ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
        #     j=j+1


        #driver.window_handles()
        main_handle = driver.current_window_handle
        nextbutton = True
        count = 1
        page_count = 0



        while nextbutton == True:
            nextbutton = False
            print("===========第%s页===========\n" % count)
            #下载pdf
            downloadpdf(driver,keywords[0])
            # 3、翻页
            try:
                next_ele = driver.find_element(By.ID, 'Page_next_top')
                # next_ele.click()
                driver.execute_script("arguments[0].click();", next_ele)
                nextbutton = True
                time.sleep(random.randint(2, 4))
            except Exception as e:
                print("翻页失败", e)
                next_flag = True
                nextbutton = False
            count = count + 1

        try:
            driver.quit()
        except Exception:
            pass
        _release_memory()

if __name__ == "__main__":
    # text = '机器人'
    # cnkipaperdownpdfv4(text)
    keywords = ['机器人','人工智能','机器视觉','自然语言处理 + NLP','计算机视觉']
    cnkipaperdownpdfv5(keywords)