'''
SCI / Web of Science 论文题录自动导出

参考 SCI自动化登录流程.md、SCI自动化导出题录流程.md，
结构对齐 cnkiforexcel.py：可配置起止日期与检索粒度，
默认 Excel 自定义导出并勾选全部可选项，每批最多 1000 条。
'''
from __future__ import annotations

import calendar
import json
import os
import random
import re
import shutil
import tempfile
import time
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

# ---------------------------------------------------------------------------
# 基本配置（可改）
# ---------------------------------------------------------------------------
bgdate = '2023-01-01'
initdate = '2026-06-30'

WOS_HOME_URL = 'https://www.webofscience.com'
WOS_HOST_CN = 'https://webofscience.clarivate.cn'
WOS_HOST_COM = 'https://www.webofscience.com'
WOS_SEARCH_URL = f'{WOS_HOST_CN}/wos/alldb/basic-search'
WOS_WOSCC_SEARCH_URL = f'{WOS_HOST_CN}/wos/woscc/basic-search'
# 文档入口：www.webofscience.com 会跳到 Clarivate 机构登录
WOS_INSTITUTION_LOGIN_URL = (
    'https://access.clarivate.com/login?app=wos&alternative=true'
    '&goto=https:%2F%2Fwww.webofknowledge.com'
    '&shibShireURL=https:%2F%2Fwww.webofknowledge.com%2F%3Fauth%3DShibboleth'
    '&shibReturnURL=https:%2F%2Fwww.webofknowledge.com%2F'
    '&roaming=true'
)
# 登录入口：先官网/机构登录页，避免 clarivate.cn 检索 SPA 渲染超时把认证流程跳过
WOS_LOGIN_ENTRY_URLS = (
    WOS_HOME_URL,
    WOS_INSTITUTION_LOGIN_URL,
    f'{WOS_HOST_COM}/wos/alldb/basic-search',
    WOS_SEARCH_URL,
    WOS_WOSCC_SEARCH_URL,
)
# SSO 完成后优先进中国镜像；.com / webofknowledge 回跳常触发 ERR_INVALID_REDIRECT
WOS_POST_SSO_URLS = (
    WOS_SEARCH_URL,
    WOS_WOSCC_SEARCH_URL,
    f'{WOS_HOST_COM}/wos/alldb/basic-search',
    f'{WOS_HOST_COM}/wos/woscc/basic-search',
)
DOWNLOAD_ROOT = r'D:\2026年工作\SCI数据库下载'

# 检索粒度：'day'=按天 | 'month'=按月 | 'year'=按年 | 'all'=整段一次
CRAWL_GRANULARITY = 'day'

# 数据库固定「所有数据库」；只勾选合集（6 项，与订阅库对应）
SEARCH_DATABASE = '所有数据库'
SEARCH_COLLECTIONS = '全选'

# 单次 Excel 导出上限；命中超过此值需分批；超过 WOS_MAX_HITS 则对半切分日期
EXPORT_BATCH_SIZE = 1000
# 失败只减半一次，不再低于此条数
EXPORT_MIN_SPLIT = 125
WOS_MAX_HITS = 100000
# 日期检索后是否按 Research Areas 逐类 Refine 再导出（也可在任务里单独设置）
EXPORT_BY_RESEARCH_AREAS = False
# 页面加载 / 检索失败重试次数
PAGE_LOAD_RETRIES = 3
# 点击 Search 后等待进入结果页的最长时间；超时关浏览器重开再试
SEARCH_RESULT_WAIT = 10 * 60  # 秒
SEARCH_TIMEOUT_RETRIES = 3
# 检索触发 Server.unexpectedError（操作过频）时：关浏览器等待后重试
SERVER_UNEXPECTED_ERROR_WAIT = 5 * 60  # 秒
SERVER_UNEXPECTED_ERROR_RETRIES = 3
# 单批下载等待秒数；超时后再扫一次已落地的默认文件名
EXPORT_DOWNLOAD_TIMEOUT = 60

# ---------- 超时设计（秒）----------
# 原则：网页「还在加载」不计入交互超时；关键元素出现后再开始计时。
# PAGE_LOAD_TIMEOUT      driver.get / 等 document.ready + 关键元素。网络超时 30s
# LOGIN_TOTAL_TIMEOUT    整段登录交互安全上限（各页加载等待会顺延此截止时间）
# STEP_FEDERATION        联邦下拉已出现后，选机构
# STEP_GOTO_INST         点「转到机构」
# STEP_CARSI             CARSI 输入框出现后，选学校并登录
# STEP_CAS               账密框出现后，填写并点登录
# NAV_AFTER_CLICK_WAIT   表单已就绪且已点击后，等跳转；未跳则重点（不是等加载）
# STEP_IDP               同意页控件出现后，勾选提交
# SEARCH_FORM_TIMEOUT    检索表单出现
# SEARCH_RESULT_WAIT     点击 Search 后等结果页（最长 10 分钟）
# HIT_COUNT_TIMEOUT      已在结果页后解析命中数
# EXPORT_DIALOG_TIMEOUT  导出弹窗出现
# EXPORT_DOWNLOAD_TIMEOUT 点 Export 后等 Excel 落地（上面单独定义）
PAGE_LOAD_TIMEOUT = 30
LOGIN_TOTAL_TIMEOUT = 180
STEP_FEDERATION_TIMEOUT = 20
STEP_GOTO_INST_TIMEOUT = 15
STEP_CARSI_TIMEOUT = 30
STEP_CAS_TIMEOUT = 30
NAV_AFTER_CLICK_WAIT = 12
STEP_IDP_TIMEOUT = 20
SEARCH_FORM_TIMEOUT = 15
HIT_COUNT_TIMEOUT = 90
EXPORT_DIALOG_TIMEOUT = 20

# 与 CNKI 并行：独立 Chrome 用户目录 + 调试端口，避免抢同一 profile
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CHROME_PROFILE_ROOT = os.path.join(_SCRIPT_DIR, '.chrome_profiles', 'sci')
# 9223 起，按进程号偏移，降低与 CNKI(9222 段)冲突概率
CHROME_DEBUG_PORT_BASE = 9223

CREDENTIALS_FILE = os.path.join(_SCRIPT_DIR, 'sci_credentials.json')

CORE_COLLECTION_NAME = 'Web of Science 核心合集'

# 「所有数据库」下的 6 个合集（页面勾选项，不含「全选」）
WOS_COLLECTIONS = [
    CORE_COLLECTION_NAME,
    'KCI-Korean Journal Database',
    'Preprint Citation Index',
    'ProQuest Dissertations & Theses Citation Index',
    'Research Commons',
    'SciELO Citation Index',
]

# 仅核心合集直达 /wos/woscc/basic-search；
# 其余 5 个合集在 /wos/alldb/basic-search 勾选 Collections（独立库网址无日期筛选）
WOS_COLLECTION_URL_SLUGS = {
    CORE_COLLECTION_NAME: 'woscc',
    'KCI-Korean Journal Database': 'kjd',
    'Preprint Citation Index': 'pprn',
    'ProQuest Dissertations & Theses Citation Index': 'pqdt',
    'Research Commons': 'rc',
    'SciELO Citation Index': 'scielo',
    '全选': 'alldb',
}

# 核心合集下的引文索引（选择数据库切到核心合集后出现）
WOS_CITATION_INDEXES = [
    'SCI-EXPANDED',
    'SSCI',
    'A&HCI',
    'ESCI',
    'CPCI-S',
    'CPCI-SSH',
    'BKCI-S',
    'BKCI-SSH',
]

# 兼容旧名
WOS_DATABASE_TREE = {
    SEARCH_DATABASE: {name: {} for name in ['全选'] + WOS_COLLECTIONS},
}

# 中英文对照，用于点选合集勾选框
NAME_ALIASES = {
    '所有数据库': ['所有数据库', 'All Databases'],
    '全选': ['全选', 'Select All', 'All'],
    'Web of Science 核心合集': [
        'Web of Science 核心合集',
        'Web of Science Core Collection',
        '核心合集',
    ],
    'KCI-Korean Journal Database': [
        'KCI-Korean Journal Database',
        'KCI',
        'Korean Journal Database',
    ],
    'Preprint Citation Index': [
        'Preprint Citation Index',
        'Preprint',
    ],
    'ProQuest Dissertations & Theses Citation Index': [
        'ProQuest Dissertations & Theses Citation Index',
        'ProQuest™ Dissertations & Theses Citation Index',
        'ProQuest',
        'Dissertations & Theses',
    ],
    'Research Commons': ['Research Commons'],
    'SciELO Citation Index': [
        'SciELO Citation Index',
        'SciELO',
    ],
    'All': ['All', '全部', '引文索引: All', '引文索引: 全部', '全选'],
    'SCI-EXPANDED': [
        'SCI-EXPANDED',
        'Science Citation Index Expanded',
        '科学引文索引扩展版',
    ],
    'SSCI': [
        'SSCI',
        'Social Sciences Citation Index',
        '社会科学引文索引',
    ],
    'A&HCI': [
        'A&HCI',
        'AHCI',
        'Arts & Humanities Citation Index',
        '艺术与人文引文索引',
    ],
    'ESCI': [
        'ESCI',
        'Emerging Sources Citation Index',
        '新兴来源引文索引',
    ],
    'CPCI-S': [
        'CPCI-S',
        'Conference Proceedings Citation Index - Science',
        '会议录引文索引-科学',
        '会议录引文索引 - 科学',
    ],
    'CPCI-SSH': [
        'CPCI-SSH',
        'Conference Proceedings Citation Index - Social Science & Humanities',
        '会议录引文索引-社会科学',
    ],
    'BKCI-S': [
        'BKCI-S',
        'Book Citation Index – Science',
        'Book Citation Index - Science',
        '图书引文索引-科学',
    ],
    'BKCI-SSH': [
        'BKCI-SSH',
        'Book Citation Index – Social Sciences & Humanities',
        '图书引文索引-社会科学',
    ],
}

# 预定义任务。为空则启动时交互配置。
# 字段：
#   db_path     : ['所有数据库', '全选'] 或 ['所有数据库', 'Web of Science 核心合集', ...]
#   collections : '全选'，或合集名 / 合集名 list（仅勾选合集，数据库固定为所有数据库）
#   bgdate / initdate / granularity
#   by_research_areas : True 时日期检索后按 Research Areas 逐类 Refine 导出
EXPORT_TASKS = [
    # {
    #     'name': '仅核心合集',
    #     'db_path': ['所有数据库', 'Web of Science 核心合集'],
    #     'bgdate': '2024-01-01',
    #     'initdate': '2024-01-31',
    #     'granularity': 'day',
    #     'by_research_areas': False,
    # },
]

EXPORT_FAILURE_LOG = 'export_failures.txt'

# 本进程 Chrome profile，退出时清理
_chrome_profile_dir = None
_chrome_restart_count = 0


# ---------------------------------------------------------------------------
# 工具：凭据 / 日期 / 驱动
# ---------------------------------------------------------------------------
def load_credentials():
    """从 sci_credentials.json 读取登录配置。"""
    defaults = {
        'federation': 'CHINA CERNET Federation',
        'org_name': '对外经济贸易大学',
        'username': '',
        'password': '',
    }
    if not os.path.isfile(CREDENTIALS_FILE):
        print(f'未找到凭据文件: {CREDENTIALS_FILE}，请复制 sci_credentials.example.json')
        return defaults
    try:
        with open(CREDENTIALS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return defaults
        out = dict(defaults)
        out.update({
            k: data[k] for k in (
                'federation', 'org_name', 'username', 'password')
            if k in data
        })
        return out
    except Exception as e:
        print('读取 SCI 凭据失败:', e)
        return defaults


def iter_crawl_date_ranges(start, end, granularity):
    """按粒度切分 [start, end]。生成 (range_start, range_end) 的 date 对。"""
    if start > end:
        return

    g = (granularity or 'all').strip().lower()
    if g in ('all', 'total', '整段', '总'):
        yield start, end
        return

    current = start
    while current <= end:
        if g in ('day', 'd', '天'):
            yield current, current
            current += timedelta(days=1)
            continue

        if g in ('month', 'm', '月'):
            last_day = calendar.monthrange(current.year, current.month)[1]
            period_end = min(current.replace(day=last_day), end)
            yield current, period_end
            current = period_end + timedelta(days=1)
            continue

        if g in ('year', 'y', '年'):
            period_end = min(current.replace(month=12, day=31), end)
            yield current, period_end
            current = period_end + timedelta(days=1)
            continue

        raise ValueError(f'未知检索粒度: {granularity}')


def granularity_label(granularity):
    return {
        'day': '按天',
        'month': '按月',
        'year': '按年',
        'all': '总（整段一次）',
    }.get((granularity or '').lower(), str(granularity))


def split_date_range_halves(start, end):
    """将日期区间对半切分，返回 [(s1,e1), (s2,e2)]；单日则返回原区间。"""
    if start >= end:
        return [(start, end)]
    days = (end - start).days
    mid = start + timedelta(days=days // 2)
    return [(start, mid), (mid + timedelta(days=1), end)]


def _safe_dirname(name):
    """文件/目录名安全化。"""
    text = re.sub(r'[<>:"/\\|?*]+', '_', str(name or '').strip())
    text = re.sub(r'\s+', ' ', text).strip(' ._')
    return text[:80] or 'unnamed'


def db_path_to_dirname(db_path):
    """库/合集路径转为目录名。"""
    if isinstance(db_path, str):
        return _safe_dirname(db_path)
    if not db_path:
        return 'All_Databases'
    return '_'.join(_safe_dirname(p) for p in db_path)


def resolve_search_mode(database, collections):
    """判定检索方式。

    返回 ('core', 引文索引) 或 ('alldb', 合集list或None)。
    只勾核心合集时切到「选择数据库=核心合集 + 引文索引」，Excel 弹窗才是标准中文界面。
    """
    if isinstance(collections, (list, tuple)):
        raw = [c for c in collections if c]
    elif collections:
        raw = [collections]
    else:
        raw = []

    db_cit = _canonical_citation(database)
    db_col = _canonical_collection(database)
    if db_cit:
        return 'core', db_cit
    if db_col == CORE_COLLECTION_NAME or database in (
            CORE_COLLECTION_NAME, 'Web of Science Core Collection'):
        if not raw:
            return 'core', 'All'
        cit = _canonical_citation(raw[0]) if len(raw) == 1 else None
        if cit:
            return 'core', cit
        if len(raw) == 1 and _canonical_collection(raw[0]) in ('全选', None):
            if raw[0] in ('All', '全部', '全选', '引文索引: All'):
                return 'core', 'All'
        if len(raw) == 1 and _canonical_collection(raw[0]) == CORE_COLLECTION_NAME:
            return 'core', 'All'
        cit = _canonical_citation(raw[0]) if raw else None
        return 'core', cit or 'All'

    if len(raw) == 1:
        cit = _canonical_citation(raw[0])
        if cit:
            return 'core', cit
        col = _canonical_collection(raw[0])
        if col == CORE_COLLECTION_NAME:
            return 'core', 'All'
        if col == '全选' or raw[0] in ('全选', 'All', '全部'):
            return 'alldb', None

    if not raw:
        return 'alldb', None

    canons = []
    for item in raw:
        col = _canonical_collection(item)
        if col == '全选':
            return 'alldb', None
        if col and col not in canons:
            canons.append(col)
    if canons == [CORE_COLLECTION_NAME]:
        return 'core', 'All'
    if set(canons) >= set(WOS_COLLECTIONS):
        return 'alldb', None
    return 'alldb', canons or None


def resolve_db_path(database, collections):
    """database + collections → 路径 list。"""
    mode, extra = resolve_search_mode(database, collections)
    if mode == 'core':
        return [CORE_COLLECTION_NAME, extra or 'All']
    path = [SEARCH_DATABASE]
    if extra is None:
        path.append('全选')
    else:
        path.extend(extra)
    return path


def _canonical_citation(name):
    """归一到核心合集引文索引短名，如 SCI-EXPANDED。"""
    text = re.sub(r'\s+', ' ', str(name or '')).strip()
    if not text or text in ('All', '全部', '全选', '引文索引: All', '引文索引: 全部'):
        return None
    for canon in WOS_CITATION_INDEXES:
        if _text_matches_aliases(text, option_aliases(canon)):
            return canon
    return None


def _canonical_collection(name):
    """把页面/旧配置里的合集名归一到 WOS_COLLECTIONS 或「全选」。"""
    text = re.sub(r'\s+', ' ', str(name or '')).strip()
    if not text:
        return None
    if text in ('所有数据库', 'All Databases'):
        return None
    if text in ('全选', 'Select All', 'All', '全部'):
        return '全选'
    for canon in WOS_COLLECTIONS:
        aliases = option_aliases(canon)
        if _text_matches_aliases(text, aliases):
            return canon
    if _text_matches_aliases(text, option_aliases('全选')):
        return '全选'
    return None


def _normalize_wanted_collections(collections):
    """返回要勾选的合集 list；None 表示全选 6 项。"""
    if collections is None:
        return None
    if isinstance(collections, (list, tuple)):
        raw = [c for c in collections if c]
    else:
        raw = [collections]
    canons = []
    saw_all = False
    for item in raw:
        canon = _canonical_collection(item)
        if canon == '全选':
            saw_all = True
            continue
        if canon and canon not in canons:
            canons.append(canon)
    if not canons:
        return None if (saw_all or not raw) else None
    if set(canons) >= set(WOS_COLLECTIONS):
        return None
    return canons


def collections_label(collections):
    if isinstance(collections, str):
        cit = _canonical_citation(collections)
        if cit:
            return cit
        if collections in ('All', '全部', '引文索引: All'):
            return 'All'
    wanted = _normalize_wanted_collections(collections)
    if wanted is None:
        return '全选'
    return ', '.join(wanted)


def collection_url_slug(collection):
    """合集名 → URL 路径段。仅核心合集用 woscc，其余用 alldb。"""
    canon = collection
    if collection not in WOS_COLLECTION_URL_SLUGS:
        canon = _canonical_collection(collection)
    if canon == CORE_COLLECTION_NAME:
        return 'woscc'
    return 'alldb'


def collection_name_for_slug(slug):
    slug = (slug or '').lower()
    for name, value in WOS_COLLECTION_URL_SLUGS.items():
        if value == slug:
            return name
    return '全选'


def collection_search_urls(collection=None):
    """返回某合集 basic-search 的候选地址（中国站优先，.com 兜底）。

    核心合集：/wos/woscc/basic-search
    其余合集与全选：/wos/alldb/basic-search（独立库页无日期筛选）
    """
    slug = collection_url_slug(collection or '全选')
    return [
        f'{WOS_HOST_CN}/wos/{slug}/basic-search',
        f'{WOS_HOST_COM}/wos/{slug}/basic-search',
    ]


def search_url_for(database, collections):
    """当前任务应打开的检索页。核心合集走 woscc，其余走 alldb。"""
    mode, extra = resolve_search_mode(database, collections)
    if mode == 'core':
        return collection_search_urls(CORE_COLLECTION_NAME)[0]
    return collection_search_urls('全选')[0]


def iter_collection_search_jobs(database, collections):
    """多合集时拆成多次检索。产出 (显示名, collections 参数)。

    核心合集各自直达 woscc；其余合集在 alldb 勾选对应 Collections。
    """
    mode, extra = resolve_search_mode(database, collections)
    if mode == 'core':
        yield CORE_COLLECTION_NAME, extra or 'All'
        return
    wanted = _normalize_wanted_collections(collections)
    if wanted is None:
        yield '全选', '全选'
        return
    for name in wanted:
        yield name, name


def search_url_slug_of(url):
    m = re.search(r'/wos/([^/]+)/', url or '', re.I)
    return (m.group(1) or '').lower() if m else ''


def is_on_target_search_page(driver, target_url):
    """是否已在目标库的 basic-search / advanced-search（不是结果页）。"""
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        return False
    if 'basic-search' not in url and 'advanced-search' not in url:
        return False
    if '/summary/' in url:
        return False
    want = search_url_slug_of(target_url)
    got = search_url_slug_of(url)
    return bool(want) and want == got


def prompt_database_collection():
    """只勾选合集（数据库固定「所有数据库」）。返回路径 list。"""
    print('\n请勾选合集（只选「核心合集」时打开核心合集检索页；其余在所有数据库页勾选以便日期筛选）：')
    print('  0. 全选（所有数据库下 6 个合集）')
    for index, name in enumerate(WOS_COLLECTIONS, 1):
        print(f'  {index}. {name}')
    while True:
        choice = input(
            '输入编号（多个用逗号/空格；回车或 0 = 全选）: '
        ).strip()
        if choice in ('', '0'):
            print('已选合集: 全选')
            return [SEARCH_DATABASE, '全选']
        tokens = re.split(r'[,，\s]+', choice)
        picked = []
        bad = False
        for token in tokens:
            if not token:
                continue
            if token == '0':
                print('已选合集: 全选')
                return [SEARCH_DATABASE, '全选']
            if not token.isdigit():
                bad = True
                break
            idx = int(token)
            if idx < 1 or idx > len(WOS_COLLECTIONS):
                bad = True
                break
            name = WOS_COLLECTIONS[idx - 1]
            if name not in picked:
                picked.append(name)
        if bad or not picked:
            print('输入无效，请重新选择')
            continue
        if set(picked) >= set(WOS_COLLECTIONS):
            print('已选合集: 全选')
            return [SEARCH_DATABASE, '全选']
        print(f'已选合集: {", ".join(picked)}')
        return [SEARCH_DATABASE] + picked


def create_chrome_driver(download_dir):
    """创建独立 Chrome 实例，可与 CNKI 脚本并行。"""
    global _chrome_profile_dir, _chrome_restart_count
    os.makedirs(CHROME_PROFILE_ROOT, exist_ok=True)
    _chrome_profile_dir = tempfile.mkdtemp(
        prefix=f'sci_{os.getpid()}_', dir=CHROME_PROFILE_ROOT)
    debug_port = (
        CHROME_DEBUG_PORT_BASE
        + (os.getpid() % 400)
        + (_chrome_restart_count % 50)
    )
    _chrome_restart_count += 1

    chrome_options = Options()
    prefs = {
        'profile.default_content_settings.popups': 0,
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'directory_upgrade': True,
        'safebrowsing.enabled': True,
    }
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument(f'--user-data-dir={_chrome_profile_dir}')
    chrome_options.add_argument(f'--remote-debugging-port={debug_port}')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--no-default-browser-check')
    # Clarivate SSO 回跳偶发 ERR_INVALID_REDIRECT，关闭 QUIC 可减轻
    chrome_options.add_argument('--disable-quic')
    # 不等全部图片/脚本，避免 clarivate.cn 检索页 renderer timeout 卡死登录
    chrome_options.page_load_strategy = 'eager'
    print(f'Chrome profile: {_chrome_profile_dir} | debug-port={debug_port}')
    driver = webdriver.Chrome(options=chrome_options)
    try:
        driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    except Exception:
        pass
    set_chrome_download_dir(driver, download_dir)
    return driver


def cleanup_chrome_profile():
    """退出时清理本进程临时 profile。"""
    global _chrome_profile_dir
    if not _chrome_profile_dir:
        return
    path = _chrome_profile_dir
    _chrome_profile_dir = None
    try:
        shutil.rmtree(path, ignore_errors=True)
        print(f'已清理 Chrome profile: {path}')
    except Exception as e:
        print('清理 Chrome profile 失败:', e)


def quit_chrome_driver(driver):
    """关闭当前 Chrome 并清理临时 profile。"""
    if driver is not None:
        try:
            driver.quit()
        except Exception:
            pass
    cleanup_chrome_profile()
    time.sleep(1.2)
    return None


def start_browser_session(download_dir, database, collections):
    """新开浏览器、登录并进入对应检索页。失败仍返回 driver（可能未登录）。"""
    driver = None
    try:
        driver = create_chrome_driver(download_dir)
        try:
            driver.maximize_window()
        except Exception:
            pass
        time.sleep(2.0)
        if not complete_wos_login(driver):
            print('浏览器启动后登录失败')
            return driver
        try:
            open_search_page(
                driver, target_url=search_url_for(database, collections))
        except Exception as e:
            print('登录后打开检索页失败:', e)
        return driver
    except Exception as e:
        print('启动浏览器会话异常:', e)
        return driver


def restart_browser_session(driver, download_dir, database, collections, reason=''):
    """关掉当前浏览器，重新启动并登录。网络抖动时会再试几次，不向外抛异常。"""
    msg = reason or '开始新一段，重启浏览器'
    print(msg)
    try:
        quit_chrome_driver(driver)
    except Exception:
        pass
    time.sleep(2.5)
    last = None
    for attempt in range(1, 4):
        new_driver = None
        try:
            print(f'重启登录尝试 {attempt}/3 ...')
            new_driver = start_browser_session(
                download_dir, database, collections)
            if is_driver_alive(new_driver) and is_wos_logged_in(new_driver):
                bind_browser_ctx(
                    new_driver, download_dir, database, collections)
                return new_driver
            print(f'重启后未登录 ({attempt}/3)，关闭后再试')
        except Exception as e:
            last = e
            print(f'重启浏览器异常 ({attempt}/3):', e)
        try:
            quit_chrome_driver(new_driver)
        except Exception:
            pass
        time.sleep(4 * attempt)
    if last:
        print('重启浏览器多次失败:', last)
    try:
        new_driver = start_browser_session(download_dir, database, collections)
        bind_browser_ctx(new_driver, download_dir, database, collections)
        return new_driver
    except Exception as e:
        print('最后一次启动浏览器仍失败:', e)
        return None


# 当前浏览器会话（检索遇 Server.unexpectedError 重启后供深层调用同步 driver）
_browser_ctx = {
    'driver': None,
    'download_dir': None,
    'database': None,
    'collections': None,
}


def bind_browser_ctx(driver, download_dir=None, database=None, collections=None):
    """绑定/更新当前任务使用的浏览器与检索上下文。"""
    if driver is not None:
        _browser_ctx['driver'] = driver
    if download_dir is not None:
        _browser_ctx['download_dir'] = download_dir
    if database is not None:
        _browser_ctx['database'] = database
    if collections is not None:
        _browser_ctx['collections'] = collections
    return _browser_ctx.get('driver')


def ctx_driver(fallback=None):
    """取当前可用 driver（重启后自动切到新实例）。"""
    d = _browser_ctx.get('driver')
    if is_driver_alive(d):
        return d
    if is_driver_alive(fallback):
        _browser_ctx['driver'] = fallback
        return fallback
    return d if d is not None else fallback


def restart_browser_after_server_error(driver=None, reason=''):
    """关闭浏览器，等待 SERVER_UNEXPECTED_ERROR_WAIT 秒后重新登录。"""
    dl = _browser_ctx.get('download_dir')
    db = _browser_ctx.get('database')
    cols = _browser_ctx.get('collections')
    old = driver if driver is not None else _browser_ctx.get('driver')
    print(reason or '检测到 Server.unexpectedError，关闭浏览器')
    try:
        quit_chrome_driver(old)
    except Exception:
        pass
    _browser_ctx['driver'] = None
    wait_sec = int(SERVER_UNEXPECTED_ERROR_WAIT)
    print(f'等待 {wait_sec} 秒（{wait_sec // 60} 分钟）后重新检索...')
    time.sleep(wait_sec)
    if not dl or db is None or cols is None:
        print('缺少 download_dir/database/collections，无法重启会话')
        return None
    new_driver = start_browser_session(dl, db, cols)
    bind_browser_ctx(new_driver, dl, db, cols)
    if not is_driver_alive(new_driver) or not is_wos_logged_in(new_driver):
        print('等待后重新登录失败')
    return new_driver


def set_chrome_download_dir(driver, download_dir):
    """运行中切换 Chrome 下载目录。"""
    abs_dir = os.path.abspath(download_dir)
    os.makedirs(abs_dir, exist_ok=True)
    payload = {'behavior': 'allow', 'downloadPath': abs_dir}
    errors = []
    for cmd in ('Page.setDownloadBehavior', 'Browser.setDownloadBehavior'):
        try:
            driver.execute_cdp_cmd(cmd, payload)
            print(f'已设置下载目录: {abs_dir}')
            return True
        except Exception as e:
            errors.append(f'{cmd}: {e}')
    print('设置下载目录失败: ' + ' | '.join(errors))
    return False


def is_driver_alive(driver):
    if driver is None:
        return False
    try:
        _ = driver.window_handles
        return True
    except Exception:
        return False


def _is_nav_timeout(exc):
    text = str(exc or '').lower()
    return ('timed out' in text or 'timeout' in text
            or 'renderer' in text or 'pageloadtimeout' in text)


def _stop_page_load(driver):
    try:
        driver.execute_script('window.stop();')
    except Exception:
        pass


def _nav_landed_ok(driver):
    """driver.get 超时后，页面是否已经跳到可用地址。"""
    if not is_driver_alive(driver):
        return False
    if is_chrome_error_page(driver):
        return False
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        return False
    if (not url or url.startswith('about:') or url.startswith('data:')
            or url.startswith('chrome-error') or url.startswith('chrome://')):
        return False
    return True


def safe_get(driver, url, retries=3, wait=3.0):
    """打开 URL；连接重置 / SSL 失败时重试，不向外抛网络异常。

    渲染超时若已跳到登录/检索页，视为成功，避免卡在 clarivate.cn SPA。
    """
    if not is_driver_alive(driver):
        return False
    last = None
    for attempt in range(1, retries + 1):
        try:
            driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
        except Exception:
            pass
        try:
            driver.get(url)
            time.sleep(wait)
            return True
        except Exception as e:
            last = e
            print(f'打开页面失败 ({attempt}/{retries}): {url}')
            print(f'  {str(e).splitlines()[0][:200]}')
            if not is_driver_alive(driver):
                return False
            if _is_nav_timeout(e):
                _stop_page_load(driver)
                if _nav_landed_ok(driver):
                    print(f'页面已到达（忽略渲染超时）: {driver.current_url}')
                    time.sleep(wait)
                    return True
            time.sleep(1.5 * attempt)
    if last:
        print('多次打开页面仍失败')
    return False


def record_export_failure(save_dir, label):
    """将失败标签追加写入导出目录下的 txt（去重）。"""
    path = os.path.join(save_dir, EXPORT_FAILURE_LOG)
    existing = set()
    if os.path.isfile(path):
        with open(path, 'r', encoding='utf-8') as f:
            existing = {line.strip() for line in f if line.strip()}
    if label in existing:
        print(f'失败记录已存在: {label}')
        return
    with open(path, 'a', encoding='utf-8') as f:
        f.write(label + '\n')
    print(f'已记录失败: {label} -> {path}')


def snapshot_download_files(save_dir):
    """快照下载目录中的 xls/xlsx 文件名（兼容旧调用）。"""
    return set(snapshot_download_meta(save_dir))


def snapshot_download_meta(save_dir):
    """快照下载文件 {文件名: (mtime, size)}，不含临时文件。"""
    result = {}
    if not os.path.isdir(save_dir):
        return result
    for name in os.listdir(save_dir):
        if name.endswith('.crdownload') or name.endswith('.tmp'):
            continue
        lower = name.lower()
        if not lower.endswith(('.xls', '.xlsx', '.csv')):
            continue
        path = os.path.join(save_dir, name)
        try:
            if os.path.isfile(path):
                result[name] = (os.path.getmtime(path), os.path.getsize(path))
        except Exception:
            continue
    return result


def _download_has_temp(save_dir, name):
    return (
        os.path.exists(os.path.join(save_dir, name + '.crdownload'))
        or os.path.exists(os.path.join(save_dir, name + '.tmp'))
    )


def _is_stable_download(path):
    try:
        size1 = os.path.getsize(path)
        time.sleep(0.8)
        size2 = os.path.getsize(path)
        return size1 == size2 and size1 > 0
    except Exception:
        return False


def pick_new_download(save_dir, before_meta):
    """找出新下载或被覆盖的默认文件（如 savedrecs.xls）。"""
    after = snapshot_download_meta(save_dir)
    candidates = []
    for name, (mtime, size) in after.items():
        if _download_has_temp(save_dir, name):
            continue
        prev = before_meta.get(name)
        if prev is None:
            candidates.append(name)
            continue
        prev_mtime, prev_size = prev
        if mtime > prev_mtime + 0.05 or size != prev_size:
            candidates.append(name)
    if not candidates:
        return None

    def sort_key(n):
        nl = n.lower()
        path = os.path.join(save_dir, n)
        try:
            mtime = os.path.getmtime(path)
        except Exception:
            mtime = 0
        preferred = 0 if nl.startswith('savedrecs') else 1
        return (preferred, -mtime)

    name = min(candidates, key=sort_key)
    path = os.path.join(save_dir, name)
    if _is_stable_download(path):
        return path
    return None


def get_new_download_file(save_dir, before_meta, timeout=120):
    """等待新下载或同名覆盖文件。before_meta 为 snapshot_download_meta 结果。"""
    if isinstance(before_meta, set):
        before_meta = {n: (0, 0) for n in before_meta}
    deadline = time.time() + timeout
    while time.time() < deadline:
        hit = pick_new_download(save_dir, before_meta)
        if hit:
            return hit
        time.sleep(1)
    # 超时后再认领一次：文件可能刚好落地，或覆盖了 savedrecs.xls
    return pick_new_download(save_dir, before_meta)


def rename_export_file(new_file, save_dir, label_prefix, start_n, end_n):
    """把默认文件名改成批次名。成功返回目标路径。"""
    ext = os.path.splitext(new_file)[1] or '.xls'
    dest = os.path.join(save_dir, f'{label_prefix}_{start_n}-{end_n}{ext}')
    if os.path.exists(dest) and os.path.abspath(dest) != os.path.abspath(new_file):
        stamp = datetime.now().strftime('%H%M%S')
        dest = os.path.join(
            save_dir, f'{label_prefix}_{start_n}-{end_n}_{stamp}{ext}')
    try:
        if os.path.abspath(new_file) != os.path.abspath(dest):
            os.replace(new_file, dest)
        print(f'导出成功: {dest}')
        return dest
    except Exception as e:
        print(f'重命名失败，保留原文件 {new_file}:', e)
        return new_file


def collection_check_root(save_dir):
    """已导出检查根目录：任务日期文件夹的上一层（合集目录）。

    例：.../所有数据库_KCI-Korean Journal Database/2023-01-02_2026-06-30_day
        → .../所有数据库_KCI-Korean Journal Database
    """
    abs_dir = os.path.abspath(save_dir)
    parent = os.path.dirname(abs_dir)
    if parent and os.path.isdir(parent):
        return parent
    return abs_dir


def find_existing_batch_file(save_dir, label_prefix, start_n, end_n):
    """合集目录下各日期批次子文件夹中，是否已有该批次改名后的文件。"""
    root = collection_check_root(save_dir)
    if not os.path.isdir(root):
        return None
    prefix = f'{label_prefix}_{start_n}-{end_n}'
    hits = []
    for dirpath, _dirnames, filenames in os.walk(root):
        for name in filenames:
            if name.endswith('.crdownload') or name.endswith('.tmp'):
                continue
            lower = name.lower()
            if not lower.endswith(('.xls', '.xlsx', '.csv')):
                continue
            stem = os.path.splitext(name)[0]
            if stem == prefix or stem.startswith(prefix + '_'):
                path = os.path.join(dirpath, name)
                if os.path.isfile(path):
                    hits.append(path)
    if not hits:
        return None
    return max(hits, key=os.path.getmtime)


def safe_click(driver, el):
    """优先 JS 点击，失败再普通 click。元素过期时静默失败以便重找。"""
    try:
        driver.execute_script(
            'arguments[0].scrollIntoView({block:"center"});', el)
        time.sleep(0.15)
        driver.execute_script('arguments[0].click();', el)
        return True
    except StaleElementReferenceException:
        return False
    except Exception:
        try:
            el.click()
            return True
        except StaleElementReferenceException:
            return False
        except Exception as e:
            print('点击失败:', str(e).splitlines()[0][:180])
            return False


def click_by_text(driver, texts, tag='*', timeout=8, partial=True):
    """按可见文本点击元素。texts 可为 str 或 list。"""
    if isinstance(texts, str):
        texts = [texts]
    deadline = time.time() + timeout
    while time.time() < deadline:
        for text in texts:
            if partial:
                xpath = (
                    f'//{tag}[contains(normalize-space(.), "{text}")]'
                )
            else:
                xpath = (
                    f'//{tag}[normalize-space(.)="{text}"]'
                )
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                try:
                    if el.is_displayed():
                        if safe_click(driver, el):
                            return True
                except Exception:
                    continue
        time.sleep(0.4)
    return False


def js_click_by_text(driver, texts, exact=False):
    """点击 overlay/菜单项。不用 offsetParent（CDK 菜单常为 null）。"""
    if isinstance(texts, str):
        texts = [texts]
    return driver.execute_script(
        """
        const wants = arguments[0].map(t => t.toLowerCase());
        const exact = arguments[1];
        function ok(n) {
          const r = n.getBoundingClientRect();
          const st = window.getComputedStyle(n);
          if (st.visibility === 'hidden' || st.display === 'none') return false;
          if (r.width < 2 || r.height < 2) return false;
          return true;
        }
        const nodes = Array.from(document.querySelectorAll(
          'button, a, span, li, div, [role="menuitem"], [role="option"]'));
        for (const n of nodes) {
          if (!ok(n)) continue;
          const t = (n.innerText || n.textContent || '').trim();
          if (!t || t.length > 80) continue;
          const tl = t.toLowerCase().replace(/\\s+/g, ' ');
          for (const w of wants) {
            if (exact ? tl === w : (tl === w || tl.includes(w))) {
              n.click();
              return t;
            }
          }
        }
        return null;
        """,
        texts, exact)


def overflow_collection_splits(database, collections):
    """单日超 10 万时再拆：核心合集按引文索引，所有数据库按 6 个合集。"""
    mode, extra = resolve_search_mode(database, collections)
    if mode == 'core':
        if extra and extra != 'All':
            return []
        return [(CORE_COLLECTION_NAME, name) for name in WOS_CITATION_INDEXES]
    names = WOS_COLLECTIONS if extra is None else list(extra)
    if names == [CORE_COLLECTION_NAME]:
        return [(CORE_COLLECTION_NAME, name) for name in WOS_CITATION_INDEXES]
    if len(names) <= 1:
        return []
    return [(SEARCH_DATABASE, name) for name in names]


def fill_input(driver, el, value):
    """清空并填入输入框。"""
    try:
        driver.execute_script(
            "arguments[0].removeAttribute('readonly');"
            "arguments[0].value='';",
            el)
        el.clear()
    except Exception:
        pass
    el.send_keys(Keys.CONTROL, 'a')
    el.send_keys(Keys.BACKSPACE)
    el.send_keys(value)
    # 触发 Angular 变更检测
    try:
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('blur', {bubbles:true}));",
            el)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 登录
# ---------------------------------------------------------------------------
def is_wos_search_page(url):
    u = (url or '').lower()
    return 'webofscience' in u and ('basic-search' in u or 'advanced-search' in u
                                    or '/wos/' in u)


def is_chrome_error_page(driver):
    """Chrome 错误页（含 ERR_INVALID_REDIRECT）。"""
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        url = ''
    if url.startswith('chrome-error://') or url.startswith('chrome://'):
        return True
    try:
        title = (driver.title or '')
    except Exception:
        title = ''
    if '无法正常运作' in title or 'this page isn' in title.lower():
        return True
    try:
        body = driver.find_element(By.TAG_NAME, 'body').text or ''
    except Exception:
        body = ''
    return any(k in body for k in (
        'ERR_INVALID_REDIRECT',
        'ERR_TOO_MANY_REDIRECTS',
        'ERR_CONNECTION',
        'ERR_EMPTY_RESPONSE',
        'ERR_TIMED_OUT',
        '发送的响应无效',
        '该网页无法正常运作',
        'This page isn’t working',
        "This page isn't working",
    ))


def is_page_unusable(driver):
    """页面无法继续操作（错误页、死掉的 driver）。"""
    if not is_driver_alive(driver):
        return True
    return is_chrome_error_page(driver)


def has_wos_no_results(driver):
    """检索后出现 Your search found no results（仍停在检索页，不会进结果页）。"""
    if not is_driver_alive(driver):
        return False
    try:
        found = driver.execute_script(
            r'''
            const nodes = Array.from(document.querySelectorAll(
              'div.search-error, .search-error, .error-code,'
              + ' [role="alert"], .light-red-bg'));
            const keys = [
              'your search found no results',
              'check the spelling and/or broaden',
              '没有找到',
              '未找到结果',
              '未检索到结果',
              '没有检索到',
              'no results were found'
            ];
            for (const n of nodes) {
              const t = ((n.innerText || n.textContent || '') + '').toLowerCase();
              if (!t) continue;
              if (/server\.unexpectederror/i.test(t)) continue;
              if (keys.some(k => t.includes(k))) return true;
            }
            return false;
            ''')
        return bool(found)
    except Exception:
        pass
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR,
                'div.search-error, .search-error, .error-code, [role="alert"]'):
            t = (el.text or '').strip().lower()
            if not t or 'unexpectederror' in t:
                continue
            if ('your search found no results' in t
                    or '没有找到' in t
                    or '未找到结果' in t
                    or '未检索到结果' in t):
                return True
    except Exception:
        pass
    return False


def has_wos_server_unexpected_error(driver):
    """是否出现 WoS Server.unexpectedError（操作过频等服务端错误）。"""
    if not is_driver_alive(driver):
        return False
    try:
        found = driver.execute_script(
            r'''
            const nodes = Array.from(document.querySelectorAll(
              'div.search-error, .search-error, .error-code,'
              + ' [role="alert"], .light-red-bg'));
            for (const n of nodes) {
              const t = (n.innerText || n.textContent || '').trim();
              if (!t) continue;
              if (/Server\.unexpectedError/i.test(t)) return true;
              if (/unexpectedError/i.test(t) && /server/i.test(t)) return true;
            }
            const body = (document.body && document.body.innerText) || '';
            return /Server\.unexpectedError/i.test(body);
            ''')
        return bool(found)
    except Exception:
        pass
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR,
                'div.search-error, .search-error, .error-code, [role="alert"]'):
            t = (el.text or '').strip()
            if 'Server.unexpectedError' in t or (
                    'unexpectedError' in t and 'Server' in t):
                return True
    except Exception:
        pass
    return False


def wait_search_result_or_server_error(driver, timeout=None):
    """点击检索后等待进入结果页。

    返回: 'ok' | 'no_results' | 'server_error' | 'unusable' | 'timeout'
    默认最长等待 SEARCH_RESULT_WAIT（10 分钟），期间每分钟打印进度。
    """
    timeout = SEARCH_RESULT_WAIT if timeout is None else timeout
    deadline = time.time() + timeout
    started = time.time()
    next_log = started + 60
    print(f'等待进入结果页（最长 {int(timeout)} 秒 / {int(timeout) // 60} 分钟）...')
    while time.time() < deadline:
        if has_wos_no_results(driver):
            print('检索无结果: Your search found no results')
            return 'no_results'
        if has_wos_server_unexpected_error(driver):
            print('检测到服务错误提示: Server.unexpectedError')
            return 'server_error'
        if is_wos_results_page(driver):
            elapsed = int(time.time() - started)
            print(f'已进入结果页（耗时 {elapsed} 秒）')
            return 'ok'
        if is_page_unusable(driver):
            return 'unusable'
        now = time.time()
        if now >= next_log:
            left = int(deadline - now)
            print(f'  仍在等待结果页... 已等 {int(now - started)} 秒，剩余约 {left} 秒')
            next_log = now + 60
        time.sleep(0.5)
    if has_wos_no_results(driver):
        print('检索无结果: Your search found no results')
        return 'no_results'
    if has_wos_server_unexpected_error(driver):
        print('检测到服务错误提示: Server.unexpectedError')
        return 'server_error'
    if is_wos_results_page(driver):
        return 'ok'
    if is_page_unusable(driver):
        return 'unusable'
    print(f'等待结果页超时（{int(timeout)} 秒）')
    return 'timeout'


def restart_browser_after_search_timeout(driver=None, reason=''):
    """检索超时未出结果页：关闭浏览器并重新登录（不再盲等，避免死循环）。"""
    dl = _browser_ctx.get('download_dir')
    db = _browser_ctx.get('database')
    cols = _browser_ctx.get('collections')
    old = driver if driver is not None else _browser_ctx.get('driver')
    print(reason or '检索超时，关闭浏览器后重开')
    try:
        quit_chrome_driver(old)
    except Exception:
        pass
    _browser_ctx['driver'] = None
    time.sleep(2.5)
    if not dl or db is None or cols is None:
        print('缺少 download_dir/database/collections，无法重启会话')
        return None
    new_driver = restart_browser_session(
        None, dl, db, cols, reason='检索超时后重新启动浏览器')
    bind_browser_ctx(new_driver, dl, db, cols)
    return new_driver


def is_search_form_ready(driver):
    """基础/高级检索表单是否已出现（必须在检索页，不能是结果页/错误页）。"""
    if is_page_unusable(driver):
        return False
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        return False
    if '/summary/' in url:
        return False
    if 'basic-search' not in url and 'advanced-search' not in url:
        return False
    for sel in (
        'button[aria-label*="Select search field"]',
        'button[data-ta="run-search"]',
        'input[data-ta="search-criteria-input"]',
        'app-search-select-database',
        'wos-select#snSelectDb',
        'app-search-row',
        '.split-field-input',
    ):
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                if el.is_displayed():
                    return True
        except Exception:
            continue
    return False


def is_wos_results_page(driver):
    """是否在检索结果页（可点 Export）。"""
    if is_page_unusable(driver):
        return False
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        return False
    if '/summary/' in url or '/search-results' in url:
        return True
    try:
        if driver.find_elements(By.CSS_SELECTOR, 'app-export-menu button'):
            if 'basic-search' not in url and 'advanced-search' not in url:
                return True
    except Exception:
        pass
    # 新版结果页：标题 brand-blue / Documents 标签计数
    try:
        for sel in (
            'h1.search-info-title .brand-blue',
            '.search-info-title .brand-blue',
            'span.tab-results-count',
            'a#parent-summary-link .tab-results-count',
            '[data-ta="select-page-checkbox"]',
        ):
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                try:
                    if el.is_displayed():
                        return True
                except Exception:
                    continue
    except Exception:
        pass
    return False


def wait_search_form_ready(driver, timeout=None):
    timeout = SEARCH_FORM_TIMEOUT if timeout is None else timeout
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_page_unusable(driver):
            return False
        if is_search_form_ready(driver):
            return True
        time.sleep(0.4)
    return is_search_form_ready(driver)


def recover_wos_page(driver, prefer_core=False, retries=None, target_url=None):
    """错误页 / 无效重定向后重新打开检索页，必要时重登。"""
    retries = PAGE_LOAD_RETRIES if retries is None else retries
    print('页面未成功加载，开始恢复...')
    if not is_driver_alive(driver):
        print('浏览器已断开，无法恢复页面')
        return False
    if not target_url:
        target_url = WOS_WOSCC_SEARCH_URL if prefer_core else WOS_SEARCH_URL
    slug = search_url_slug_of(target_url)
    targets = []
    for dest in collection_search_urls(collection_name_for_slug(slug)):
        if dest not in targets:
            targets.append(dest)
    if target_url not in targets:
        targets.insert(0, target_url)
    for dest in WOS_POST_SSO_URLS:
        if dest not in targets:
            targets.append(dest)

    for attempt in range(1, retries + 1):
        print(f'  恢复尝试 {attempt}/{retries}')
        if is_page_unusable(driver) or not is_wos_logged_in(driver):
            recovered = False
            try:
                recovered = goto_wos_after_sso(driver)
            except Exception as e:
                print('  改开检索页失败:', e)
            if not recovered or is_page_unusable(driver) or not is_wos_logged_in(driver):
                if not complete_wos_login(driver):
                    time.sleep(2)
                    continue
        dest = targets[(attempt - 1) % len(targets)]
        print(f'  重新打开: {dest}')
        if not safe_get(driver, dest, retries=2, wait=3.5):
            time.sleep(2)
            continue
        if is_page_unusable(driver):
            print('  仍是错误页')
            time.sleep(1.5)
            continue
        accept_cookies(driver, timeout=6)
        if wait_search_form_ready(driver, timeout=12):
            print('页面已恢复:', driver.current_url)
            return True
        if is_wos_logged_in(driver) and not is_page_unusable(driver):
            if wait_search_form_ready(driver, timeout=8):
                print('页面已恢复:', driver.current_url)
                return True
    print('页面恢复失败')
    return False


def is_wos_logged_in(driver):
    """粗略判断是否已进入 WoS 主站（非登录页）。

    仅打开 www.webofscience.com 首页不算已登录，必须走机构认证。
    """
    if is_chrome_error_page(driver):
        return False
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        return False
    if 'access.clarivate.com' in url:
        return False
    if 'cas.uibe.edu.cn' in url or 'ds.carsi.edu.cn' in url:
        return False
    if 'idp-lib.uibe.edu.cn' in url:
        return False
    if 'webofscience' not in url and 'webofknowledge' not in url:
        return False
    path = url.split('?', 1)[0].rstrip('/')
    if path.endswith('webofscience.com') or path.endswith('webofknowledge.com'):
        return False
    if '/wos/' not in url and '/summary/' not in url:
        return False
    return True


def goto_wos_after_sso(driver, reason=None):
    """避开 access.clarivate.com 无效回跳，改开检索页。"""
    msg = reason or 'SSO 完成，改开 WoS 检索页（避开 ERR_INVALID_REDIRECT）...'
    print(msg)
    for dest in WOS_POST_SSO_URLS:
        print(f'  尝试: {dest}')
        if not safe_get(driver, dest, retries=2, wait=4.0):
            continue
        if is_chrome_error_page(driver):
            print('  仍是错误页，换下一地址')
            continue
        url = (driver.current_url or '').lower()
        if is_wos_logged_in(driver) or is_wos_search_page(url):
            accept_cookies(driver, timeout=8)
            wait_search_form_ready(driver, timeout=8)
            print('已进入 WoS:', driver.current_url)
            return True
        if 'webofscience' in url and 'access.clarivate.com' not in url:
            accept_cookies(driver, timeout=5)
            wait_search_form_ready(driver, timeout=8)
            print('已打开 WoS 站点:', driver.current_url)
            return True
    print('未能进入 WoS 检索页')
    return False


def accept_cookies(driver, timeout=10):
    """点击 OneTrust Accept all。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        btns = driver.find_elements(By.ID, 'onetrust-accept-btn-handler')
        for btn in btns:
            try:
                if btn.is_displayed():
                    safe_click(driver, btn)
                    print('已点击 Accept all cookies')
                    time.sleep(1)
                    return True
            except Exception:
                pass
        time.sleep(0.5)
    return False


def wait_document_ready(driver, timeout=None):
    """等 document.readyState 为 complete。加载中不计交互超时。"""
    timeout = PAGE_LOAD_TIMEOUT if timeout is None else timeout
    end = time.time() + max(timeout, 0.5)
    while time.time() < end:
        if not is_driver_alive(driver):
            return False
        try:
            if driver.execute_script('return document.readyState') == 'complete':
                return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def _el_visible(el):
    try:
        return el.is_displayed()
    except Exception:
        return False


def _login_url(driver):
    try:
        return (driver.current_url or '').lower()
    except Exception:
        return ''


def _on_institution_login_page(driver):
    """是否在 Clarivate 机构登录页（选联邦 / 转到机构）。"""
    url = _login_url(driver)
    if 'access.clarivate.com' in url:
        return True
    try:
        for sel in (
            'mat-select[formcontrolname="federationName"]',
            'mat-select[aria-label="机构"]',
        ):
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                if _el_visible(el):
                    return True
    except Exception:
        pass
    return False


def _open_wos_login_entry(driver):
    """按文档打开登录入口：官网 → 机构登录页，再回退检索页。"""
    for home in WOS_LOGIN_ENTRY_URLS:
        print(f'打开登录入口: {home}')
        if not safe_get(driver, home, retries=2, wait=2.5):
            print(f'登录入口打不开，换地址: {home}')
            continue
        # 官网常自动跳到 access.clarivate.com，稍等重定向
        for _ in range(10):
            url = _login_url(driver)
            if _on_institution_login_page(driver) or _on_carsi_page(url) or _on_cas_page(url):
                print('已进入机构认证页:', driver.current_url)
                return True
            if is_wos_logged_in(driver) and is_search_form_ready(driver):
                print('已在已登录检索页:', driver.current_url)
                return True
            time.sleep(0.4)
        url = _login_url(driver)
        if 'access.clarivate.com' in url or _on_carsi_page(url) or _on_cas_page(url):
            print('已进入机构认证页:', driver.current_url)
            return True
        # 官网首页未自动跳转时，直达机构登录
        if home != WOS_INSTITUTION_LOGIN_URL:
            print('未自动跳到机构登录，改开 Clarivate 登录页')
            if safe_get(driver, WOS_INSTITUTION_LOGIN_URL, retries=2, wait=2.5):
                print('已打开机构登录页:', driver.current_url)
                return True
        print(f'登录入口未进入认证页，换地址: {home}')
    print('无法打开登录入口')
    return False


def _on_carsi_page(url):
    u = url or ''
    return 'ds.carsi.edu.cn' in u or 'carsi' in u


def _on_cas_page(url):
    u = url or ''
    return 'cas.uibe.edu.cn' in u or '/cas/login' in u


def _on_idp_page(url):
    u = url or ''
    return 'idp-lib.uibe.edu.cn' in u or ('/idp/' in u and 'uibe' in u)


def wait_selectors_ready(driver, selectors, timeout=None):
    """等文档加载完成且任一选择器可见。"""
    timeout = PAGE_LOAD_TIMEOUT if timeout is None else timeout
    if isinstance(selectors, str):
        selectors = [selectors]
    end = time.time() + max(timeout, 0.5)
    while time.time() < end:
        # 选择器已出现即可，不强制 complete（eager 加载下图片可能一直 pending）
        for sel in selectors:
            try:
                for el in driver.find_elements(By.CSS_SELECTOR, sel):
                    if _el_visible(el):
                        return True
            except StaleElementReferenceException:
                break
            except Exception:
                continue
        try:
            state = driver.execute_script('return document.readyState')
        except Exception:
            time.sleep(0.25)
            continue
        if state not in ('complete', 'interactive'):
            time.sleep(0.25)
            continue
        time.sleep(0.25)
    return False


def wait_leave_page(driver, still_on_fn, timeout=None, retry_after=None, on_tick=None):
    """等 URL 离开当前阶段。

    下一页仍在 loading：等到 PAGE_LOAD_TIMEOUT。
    当前页已 complete 却还在原地址：retry_after 秒后返回 False，以便重试点击。
    on_tick：每轮回调，返回 True 表示刚处理了弹窗等，应继续等待跳转。
    """
    timeout = PAGE_LOAD_TIMEOUT if timeout is None else timeout
    retry_after = NAV_AFTER_CLICK_WAIT if retry_after is None else retry_after
    end = time.time() + max(timeout, 0.5)
    retry_at = time.time() + max(retry_after, 0.5)
    while time.time() < end:
        if on_tick:
            try:
                if on_tick():
                    retry_at = time.time() + max(retry_after, 0.5)
                    end = max(end, time.time() + PAGE_LOAD_TIMEOUT)
            except Exception:
                pass
        if not still_on_fn(_login_url(driver)):
            wait_document_ready(
                driver, timeout=min(5.0, max(0.5, end - time.time())))
            return True
        if time.time() >= retry_at:
            try:
                state = driver.execute_script('return document.readyState') or ''
            except Exception:
                state = 'complete'
            if state == 'complete':
                return False
        time.sleep(0.35)
    return False


def _extend_deadline(deadline, started_at):
    """加载等待不计入交互超时：把截止时间顺延。"""
    return deadline + (time.time() - started_at)


def _carsi_click_continue_login(driver):
    """CARSI 身份认证异常提示弹窗：点击「继续登录」（idpLogin）。"""
    last = getattr(_carsi_click_continue_login, '_last', 0)
    if time.time() - last < 4:
        return False
    try:
        clicked = driver.execute_script(
            """
            const isShown = (el) => {
              if (!el) return false;
              const s = window.getComputedStyle(el);
              if (!s || s.display === 'none' || s.visibility === 'hidden') return false;
              if (Number(s.opacity) === 0) return false;
              const r = el.getBoundingClientRect();
              return r.width > 2 && r.height > 2;
            };
            const links = Array.from(document.querySelectorAll(
              '.pop-dialog a.logout_button, a[href*="idpLogin"], a.logout_button'));
            for (const a of links) {
              const t = (a.innerText || a.textContent || '').replace(/\\s+/g, '');
              const href = a.getAttribute('href') || '';
              if (!isShown(a)) continue;
              if (t.indexOf('继续登录') >= 0 || href.indexOf('idpLogin') >= 0) {
                a.click();
                return 'link';
              }
            }
            const pop = document.querySelector('.pop-dialog, #announceText');
            if (pop && isShown(pop) && typeof idpLogin === 'function') {
              idpLogin();
              return 'js';
            }
            return false;
            """)
        if clicked:
            _carsi_click_continue_login._last = time.time()
            print('CARSI 弹窗已点击「继续登录」')
            time.sleep(0.8)
            return True
    except Exception:
        pass
    try:
        for a in driver.find_elements(
                By.CSS_SELECTOR,
                '.pop-dialog a.logout_button, a[href*="idpLogin"]'):
            try:
                if not a.is_displayed():
                    continue
                text = (a.text or '').strip()
                href = a.get_attribute('href') or ''
                if '继续登录' in text or 'idpLogin' in href:
                    if safe_click(driver, a):
                        _carsi_click_continue_login._last = time.time()
                        print('CARSI 弹窗已点击「继续登录」')
                        time.sleep(0.8)
                        return True
            except StaleElementReferenceException:
                continue
    except Exception:
        pass
    return False


def wait_cas_form_ready(driver, timeout=None):
    """等 CAS 用户名和密码框都出现。加载完成前不计交互超时。"""
    timeout = PAGE_LOAD_TIMEOUT if timeout is None else timeout
    print('等待 CAS 账密页加载完成...')
    end = time.time() + max(timeout, 0.5)
    while time.time() < end:
        url = _login_url(driver)
        if not _on_cas_page(url):
            if (_on_idp_page(url) or 'webofscience' in url
                    or is_wos_logged_in(driver)):
                print('等待加载时已离开 CAS 页')
                return True
            time.sleep(0.3)
            continue
        try:
            if driver.execute_script('return document.readyState') != 'complete':
                time.sleep(0.3)
                continue
        except Exception:
            time.sleep(0.3)
            continue
        uns = driver.find_elements(
            By.CSS_SELECTOR, 'input#un, input[name="username"]')
        pds = driver.find_elements(
            By.CSS_SELECTOR,
            'input#pd, input[name="password"], input[type="password"]')
        try:
            if uns and pds and _el_visible(uns[0]) and _el_visible(pds[0]):
                print('CAS 账密页已就绪')
                return True
        except StaleElementReferenceException:
            pass
        time.sleep(0.3)
    print('CAS 账密页未加载完成')
    return False


def _select_federation(driver, federation_name, timeout=None):
    """在 Clarivate 登录页选择机构联邦。"""
    print('等待机构登录页加载完成...')
    load_ok = wait_selectors_ready(
        driver,
        ['mat-select[formcontrolname="federationName"]',
         'mat-select[aria-label="机构"]', 'mat-select'],
        timeout=PAGE_LOAD_TIMEOUT)
    if load_ok:
        print('机构登录页已就绪')
    else:
        print('机构登录页加载超时，仍尝试选择联邦')
    timeout = STEP_FEDERATION_TIMEOUT if timeout is None else timeout
    deadline = time.time() + timeout
    while time.time() < deadline:
        selects = driver.find_elements(
            By.CSS_SELECTOR, 'mat-select[formcontrolname="federationName"]')
        if not selects:
            selects = driver.find_elements(
                By.CSS_SELECTOR, 'mat-select[aria-label="机构"], mat-select')
        if selects:
            try:
                safe_click(driver, selects[0])
                time.sleep(0.8)
            except Exception:
                pass

        # overlay 中的选项
        options = driver.find_elements(By.CSS_SELECTOR, 'mat-option')
        for opt in options:
            try:
                text = (opt.text or '').strip()
                if federation_name.lower() in text.lower():
                    safe_click(driver, opt)
                    print(f'已选择机构联邦: {text}')
                    time.sleep(0.8)
                    return True
            except Exception:
                continue

        # JS 兜底：按文本找并点击
        clicked = driver.execute_script(
            """
            const name = arguments[0].toLowerCase();
            const opts = Array.from(document.querySelectorAll('mat-option'));
            for (const o of opts) {
              const t = (o.innerText || o.textContent || '').trim().toLowerCase();
              if (t.includes(name)) { o.click(); return true; }
            }
            return false;
            """,
            federation_name,
        )
        if clicked:
            print(f'已通过 JS 选择机构联邦: {federation_name}')
            time.sleep(0.8)
            return True
        time.sleep(0.5)
    print(f'未找到机构联邦选项: {federation_name}')
    return False


def _click_go_to_institution(driver, timeout=None):
    """点击「转到机构」/ Go to institution。"""
    timeout = STEP_GOTO_INST_TIMEOUT if timeout is None else timeout
    texts = ['转到机构', 'Go to institution', 'GO TO INSTITUTION', 'Continue']
    if click_by_text(driver, texts, tag='button', timeout=timeout):
        print('已点击转到机构')
        return True
    # 备用：任意 submit / primary 按钮
    for sel in (
        'button[type="submit"]',
        'button.mat-mdc-unelevated-button',
        'button.mat-flat-button',
    ):
        btns = driver.find_elements(By.CSS_SELECTOR, sel)
        for btn in btns:
            try:
                t = (btn.text or '').strip()
                if any(x.lower() in t.lower() for x in texts) or not t:
                    if btn.is_displayed():
                        safe_click(driver, btn)
                        print('已点击转到机构（备用）')
                        return True
            except Exception:
                continue
    return False


def _carsi_select_org(driver, org_name, timeout=None):
    """CARSI：输入高校名并选择、登录。必须离开 CARSI 页才返回 True。"""
    print('等待 CARSI 页加载完成...')
    load_ok = wait_selectors_ready(
        driver, ['input#show', 'button#idpSkipButton'], timeout=PAGE_LOAD_TIMEOUT)
    if load_ok:
        print('CARSI 页已就绪')
    else:
        print('CARSI 页加载超时，仍尝试选择机构')
    timeout = STEP_CARSI_TIMEOUT if timeout is None else timeout
    deadline = time.time() + timeout
    clicks = 0
    while time.time() < deadline:
        url = _login_url(driver)
        if not _on_carsi_page(url):
            if _on_cas_page(url) or _on_idp_page(url):
                print('CARSI 已离开选择页')
                return True
            time.sleep(0.4)
            continue

        _carsi_click_continue_login(driver)

        show_inputs = driver.find_elements(By.CSS_SELECTOR, 'input#show')
        if show_inputs:
            inp = show_inputs[0]
            try:
                cur = (inp.get_attribute('value') or '').strip()
                if org_name not in cur:
                    fill_input(driver, inp, org_name)
                    time.sleep(1.0)
            except StaleElementReferenceException:
                time.sleep(0.4)
                continue
            except Exception as e:
                print('CARSI 输入机构失败:', e)

            items = driver.find_elements(By.CSS_SELECTOR, 'a.dropdown-item')
            for item in items:
                try:
                    text = (item.text or '').strip()
                    if org_name in text:
                        if safe_click(driver, item):
                            print(f'CARSI 已选择: {text[:60]}')
                            time.sleep(0.6)
                        break
                except StaleElementReferenceException:
                    break
                except Exception:
                    continue

            clicked = False
            try:
                clicked = bool(driver.execute_script(
                    'const b=document.querySelector("#idpSkipButton");'
                    'if(!b) return false; b.click(); return true;'))
            except Exception:
                clicked = False
            if not clicked:
                login_btns = driver.find_elements(
                    By.CSS_SELECTOR, 'button#idpSkipButton')
                if login_btns:
                    clicked = safe_click(driver, login_btns[0])
            if clicked:
                clicks += 1
                print('CARSI 已点击登录，等待 CAS 页加载...')
                time.sleep(0.6)
                _carsi_click_continue_login(driver)
                t_leave = time.time()
                if wait_leave_page(
                        driver, _on_carsi_page,
                        timeout=PAGE_LOAD_TIMEOUT,
                        retry_after=PAGE_LOAD_TIMEOUT,
                        on_tick=_carsi_click_continue_login):
                    print('CARSI 已离开选择页')
                    return True
                deadline = _extend_deadline(deadline, t_leave)
                print('仍停在 CARSI 页，将重试选择机构')
                if clicks >= 3:
                    print('CARSI 多次点击仍未跳转')
                    return False
        time.sleep(0.5)
    print('CARSI 选择超时')
    return False


def _cas_login(driver, username, password, timeout=None):
    """UIBE CAS：填写账号密码并登录。必须离开账密页才返回 True。"""
    timeout = STEP_CAS_TIMEOUT if timeout is None else timeout
    if not username or not password:
        print('凭据缺少 username/password')
        return False
    if not wait_cas_form_ready(driver, timeout=PAGE_LOAD_TIMEOUT):
        return False
    if not _on_cas_page(_login_url(driver)):
        print('CAS 已离开账密页:', _login_url(driver)[:80])
        return True
    print('开始填写 CAS 账密...')
    deadline = time.time() + timeout
    while time.time() < deadline:
        url = _login_url(driver)
        if not _on_cas_page(url):
            if (_on_idp_page(url) or 'webofscience' in url
                    or 'clarivate' in url or is_wos_logged_in(driver)):
                print('CAS 已离开账密页:', url[:80])
                return True
            time.sleep(0.4)
            continue

        uns = driver.find_elements(
            By.CSS_SELECTOR, 'input#un, input[name="username"]')
        pds = driver.find_elements(
            By.CSS_SELECTOR,
            'input#pd, input[name="password"], input[type="password"]')
        if not uns or not pds:
            time.sleep(0.4)
            continue
        try:
            un, pd = uns[0], pds[0]
            fill_input(driver, un, username)
            time.sleep(0.2)
            fill_input(driver, pd, password)
            time.sleep(0.2)
            rem = driver.find_elements(By.CSS_SELECTOR, 'input#rememberName')
            if rem:
                try:
                    if not rem[0].is_selected():
                        safe_click(driver, rem[0])
                        print('已勾选记住用户名')
                except StaleElementReferenceException:
                    pass
                except Exception:
                    pass
            clicked = False
            try:
                clicked = bool(driver.execute_script(
                    """
                    const btn = document.querySelector(
                      'input.login_box_landing_btn, input[value="立刻登录"]');
                    if (btn) { btn.click(); return true; }
                    if (typeof login === 'function') { login(); return true; }
                    const f = document.querySelector('form');
                    if (f) { f.submit(); return true; }
                    return false;
                    """))
            except Exception:
                clicked = False
            if not clicked:
                btns = driver.find_elements(
                    By.CSS_SELECTOR,
                    'input.login_box_landing_btn, input[type="button"].login_box_landing_btn,'
                    ' input[value="立刻登录"], button.login_box_landing_btn')
                if not btns:
                    btns = driver.find_elements(
                        By.XPATH,
                        '//input[@type="button" or @type="submit"][contains(@value,"登录")]'
                        '|//button[contains(normalize-space(.), "登录")]')
                if btns:
                    clicked = safe_click(driver, btns[0])
            if not clicked:
                try:
                    pd.send_keys(Keys.ENTER)
                    clicked = True
                except Exception:
                    pass
            if clicked:
                print('已点击立刻登录，等待下一页加载...')
            t_leave = time.time()
            if wait_leave_page(
                    driver, _on_cas_page,
                    timeout=PAGE_LOAD_TIMEOUT,
                    retry_after=NAV_AFTER_CLICK_WAIT):
                print('CAS 登录成功，已离开账密页')
                return True
            waited = time.time() - t_leave
            if waited > NAV_AFTER_CLICK_WAIT + 0.5:
                deadline = _extend_deadline(deadline, t_leave)
            print('仍停在 CAS 账密页，重填后再次登录')
        except StaleElementReferenceException:
            time.sleep(0.4)
            continue
        except Exception as e:
            print('CAS 登录失败:', e)
        time.sleep(0.6)
    print('CAS 登录超时，仍在账密页')
    return False


def _idp_consent(driver, timeout=None):
    """IDP：勾选同意并提交/接受。点完「接受」后立即返回，避免打断 SSO 回跳。"""
    print('等待学校 IdP 页加载完成...')
    load_ok = wait_selectors_ready(
        driver,
        ['input#accept[type="checkbox"]',
         'input[type="submit"][name="_eventId_proceed"]'],
        timeout=PAGE_LOAD_TIMEOUT)
    if load_ok:
        print('学校 IdP 页已就绪')
    else:
        print('学校 IdP 页加载超时，仍尝试勾选同意')
    timeout = STEP_IDP_TIMEOUT if timeout is None else timeout
    deadline = time.time() + timeout
    clicked_submit = False
    clicked_accept = False
    saw_page = False
    while time.time() < deadline:
        url = _login_url(driver)
        if not _on_idp_page(url):
            if is_wos_logged_in(driver) or is_chrome_error_page(driver):
                return True
            if saw_page or clicked_accept or clicked_submit:
                return True
            time.sleep(0.5)
            continue

        saw_page = True
        try:
            boxes = driver.find_elements(
                By.CSS_SELECTOR, 'input#accept[type="checkbox"]')
            for box in boxes:
                try:
                    if not box.is_selected():
                        safe_click(driver, box)
                        print('已勾选使用条款')
                except Exception:
                    pass

            submits = driver.find_elements(
                By.CSS_SELECTOR,
                'input[type="submit"][name="_eventId_proceed"]')
            if not clicked_submit:
                for btn in submits:
                    val = (btn.get_attribute('value') or '').strip()
                    if val == '提交':
                        safe_click(driver, btn)
                        print('已点击「提交」')
                        clicked_submit = True
                        time.sleep(2)
                        break
            submits = driver.find_elements(
                By.CSS_SELECTOR,
                'input[type="submit"][name="_eventId_proceed"]')
            if not clicked_accept:
                for btn in submits:
                    val = (btn.get_attribute('value') or '').strip()
                    if val == '接受':
                        safe_click(driver, btn)
                        print('已点击「接受」，等待 SSO 回跳...')
                        clicked_accept = True
                        time.sleep(3)
                        return True
        except StaleElementReferenceException:
            time.sleep(0.4)
            continue
        except Exception as e:
            print('IDP 同意步骤异常:', e)
        time.sleep(0.6)
    return clicked_accept or clicked_submit or is_wos_logged_in(driver)


def complete_wos_login(driver, timeout=None):
    """完整机构登录流程。成功返回 True。网络错误不向外抛。"""
    timeout = LOGIN_TOTAL_TIMEOUT if timeout is None else timeout
    try:
        return _complete_wos_login_impl(driver, timeout)
    except Exception as e:
        print('登录过程异常:', e)
        try:
            return is_wos_logged_in(driver)
        except Exception:
            return False


def _complete_wos_login_impl(driver, timeout=None):
    timeout = LOGIN_TOTAL_TIMEOUT if timeout is None else timeout
    creds = load_credentials()
    federation = creds.get('federation') or 'CHINA CERNET Federation'
    org_name = creds.get('org_name') or '对外经济贸易大学'
    username = creds.get('username') or ''
    password = creds.get('password') or ''

    if (is_wos_logged_in(driver) and is_search_form_ready(driver)
            and not _on_institution_login_page(driver)):
        print('已在 WoS 检索页，跳过登录')
        accept_cookies(driver, timeout=3)
        return True

    print('开始 WoS 机构登录（CERNET → 转到机构 → CARSI 选校 → CAS 账密 → IDP 同意）...')
    if not _open_wos_login_entry(driver):
        print('无法打开登录入口')
        return False

    deadline = time.time() + timeout
    federation_done = False
    idp_done = False
    forced_institution_login = False

    while time.time() < deadline:
        url = _login_url(driver)

        # 跳转空白页：等下一页加载，这段时间不计入交互超时
        if (not url or url.startswith('about:') or url.startswith('chrome-')
                or url.startswith('chrome:') or url.startswith('data:')):
            print('登录跳转中，等待下一页加载...')
            t0 = time.time()
            wait_document_ready(driver, timeout=PAGE_LOAD_TIMEOUT)
            waited = time.time() - t0
            if waited > 1.0:
                deadline = _extend_deadline(deadline, t0)
            time.sleep(0.5)
            continue

        # SSO 完成后的无效回跳 / Chrome 错误页：改开中国站检索页
        if idp_done and (
                is_chrome_error_page(driver)
                or 'access.clarivate.com' in url
                or url.startswith('chrome-error://')):
            print('检测到 SSO 回跳异常:', driver.current_url)
            if goto_wos_after_sso(driver):
                return True
            time.sleep(1)
            continue

        if is_wos_logged_in(driver) and not _on_institution_login_page(driver):
            print('登录成功:', driver.current_url)
            accept_cookies(driver, timeout=8)
            if not is_wos_search_page(driver.current_url):
                goto_wos_after_sso(driver)
            return True

        # 官网首页未跳转：按文档打开机构登录，不能当成已登录而跳过
        if (not idp_done) and (not forced_institution_login) and (
                'webofscience' in url or 'webofknowledge' in url):
            if (not is_search_form_ready(driver)
                    and not _on_institution_login_page(driver)
                    and not _on_carsi_page(url) and not _on_cas_page(url)):
                print('尚未进入机构认证，打开 Clarivate 机构登录页...')
                forced_institution_login = True
                safe_get(driver, WOS_INSTITUTION_LOGIN_URL, retries=2, wait=2.5)
                continue

        # Clarivate 登录页：SSO 完成后不要再点「转到机构」
        if (not idp_done) and (
                _on_institution_login_page(driver)
                or 'access.clarivate.com' in url
                or ('login' in url and 'clarivate' in url)):
            print('机构登录页:', driver.current_url)
            if not federation_done:
                if _select_federation(driver, federation):
                    federation_done = True
                    time.sleep(0.4)
                    _click_go_to_institution(driver)
                    time.sleep(1)
                    continue
            else:
                _click_go_to_institution(driver)
            time.sleep(1)
            continue

        # CARSI：函数内部先等页面加载完成，再开始选校计时
        if _on_carsi_page(url):
            print('CARSI 选校页:', driver.current_url)
            _carsi_select_org(driver, org_name)
            time.sleep(0.4)
            continue

        # CAS：函数内部先等账密框出现，再开始填写计时
        if _on_cas_page(url):
            print('CAS 账密页:', driver.current_url)
            _cas_login(driver, username, password)
            time.sleep(0.4)
            continue

        # IDP
        if _on_idp_page(url):
            print('IDP 同意页:', driver.current_url)
            if not idp_done:
                idp_done = _idp_consent(driver)
                time.sleep(1.5)
                if is_wos_logged_in(driver) and not is_chrome_error_page(driver):
                    accept_cookies(driver, timeout=8)
                    if not is_wos_search_page(driver.current_url):
                        goto_wos_after_sso(driver)
                    return True
                if goto_wos_after_sso(driver):
                    return True
            time.sleep(0.5)
            continue

        time.sleep(0.5)

    print('登录超时，最后尝试打开检索页')
    if goto_wos_after_sso(driver, reason='登录未完成，仍尝试打开 WoS 检索页...'):
        return is_wos_logged_in(driver)
    return is_wos_logged_in(driver)


def ensure_logged_in(driver):
    """确保已登录；掉线则重登。"""
    if is_wos_logged_in(driver):
        accept_cookies(driver, timeout=3)
        return True
    return complete_wos_login(driver)


def open_search_page(driver, prefer_core=False, target_url=None):
    """打开指定库的 basic-search。加载失败会换镜像并恢复重试。"""
    if not target_url:
        target_url = WOS_WOSCC_SEARCH_URL if prefer_core else WOS_SEARCH_URL
    return goto_search_url(driver, target_url)


def goto_search_url(driver, target_url, retries=None):
    """打开目标检索页（核心合集 woscc，其余 alldb）。"""
    retries = PAGE_LOAD_RETRIES if retries is None else retries
    candidates = [target_url]
    slug = search_url_slug_of(target_url)
    for dest in collection_search_urls(collection_name_for_slug(slug)):
        if dest not in candidates:
            candidates.append(dest)

    for attempt in range(1, retries + 1):
        if not is_driver_alive(driver):
            return False
        dest = candidates[(attempt - 1) % len(candidates)]
        if is_page_unusable(driver) or not is_wos_logged_in(driver):
            print(f'检索页不可用，恢复中 ({attempt}/{retries})...')
            if not recover_wos_page(driver, target_url=dest):
                if not complete_wos_login(driver):
                    time.sleep(2)
                    continue
        if (is_on_target_search_page(driver, dest)
                and is_search_form_ready(driver)
                and not is_page_unusable(driver)):
            accept_cookies(driver, timeout=3)
            print(f'已在目标检索页: {driver.current_url}')
            return True
        print(f'打开检索页: {dest} ({attempt}/{retries})')
        if not safe_get(driver, dest, retries=2, wait=2.5):
            print('打开检索页失败:', dest)
            time.sleep(1.5)
            continue
        if is_page_unusable(driver):
            print(f'检索页无效重定向 ({attempt}/{retries})')
            recover_wos_page(driver, target_url=dest)
            time.sleep(1.5)
            continue
        accept_cookies(driver, timeout=8)
        if wait_search_form_ready(driver, timeout=15):
            if is_on_target_search_page(driver, dest) or is_search_form_ready(driver):
                print(f'检索页已就绪: {driver.current_url}')
                return True
            print('表单出现但不在目标库，将重开目标链接')
            continue
        print('检索表单未出现，重试打开')
        if not complete_wos_login(driver) and not is_wos_logged_in(driver):
            continue
    print(f'多次打开检索页失败: {target_url}')
    return False


# ---------------------------------------------------------------------------
# 检索设置
# ---------------------------------------------------------------------------
def option_aliases(name):
    """返回某名称在页面上可能出现的文案列表。"""
    name = (name or '').strip()
    aliases = list(NAME_ALIASES.get(name) or [])
    if name and name not in aliases:
        aliases.insert(0, name)
    # 短码：括号内 SCI-EXPANDED
    m = re.search(r'\(([A-Z0-9&-]+)\)', name)
    if m and m.group(1) not in aliases:
        aliases.append(m.group(1))
    return aliases


def _text_matches_aliases(text, aliases):
    t = re.sub(r'\s+', ' ', (text or '')).strip().lower()
    if not t:
        return False
    for a in aliases:
        al = re.sub(r'\s+', ' ', a).strip().lower()
        if not al:
            continue
        if al == t or al in t or t in al:
            return True
    return False


def _click_option_by_aliases(driver, aliases):
    """在已打开的 overlay 中按别名点击选项（不用 offsetParent）。"""
    return driver.execute_script(
        """
        const aliases = arguments[0].map(a => a.toLowerCase());
        function vis(n) {
          const r = n.getBoundingClientRect();
          const st = window.getComputedStyle(n);
          if (st.display === 'none' || st.visibility === 'hidden') return false;
          return r.width > 2 && r.height > 2;
        }
        function match(t) {
          const s = (t || '').replace(/\\s+/g, ' ').trim().toLowerCase();
          if (!s) return false;
          return aliases.some(a => {
            if (s === a) return true;
            if (a.length >= 4 && s.includes(a)) return true;
            return false;
          });
        }
        const panes = Array.from(document.querySelectorAll(
          '.cdk-overlay-pane, [role="listbox"], [role="dialog"], [role="menu"]'));
        const roots = panes.length ? panes : [document.body];
        const sels = '[role="option"], [role="menuitem"], button, li, mat-option, label, span, div, a, mat-checkbox';
        for (const root of roots) {
          for (const n of root.querySelectorAll(sels)) {
            if (!vis(n)) continue;
            const t = (n.innerText || n.textContent || '').trim();
            if (!t || t.length > 120) continue;
            if (match(t)) {
              n.click();
              return t;
            }
          }
        }
        return null;
        """,
        aliases,
    )


def _wos_select_option(driver, select_id, option_text, timeout=18):
    """打开 wos-select#id 或「选择数据库」栏下拉，并选择选项。"""
    aliases = option_aliases(option_text)
    started = time.time()
    deadline = time.time() + timeout
    while time.time() < deadline:
        selects = driver.find_elements(By.CSS_SELECTOR, f'wos-select#{select_id}')
        if not selects:
            selects = driver.find_elements(By.ID, select_id)
        if not selects:
            if time.time() - started > 3:
                print(f'未找到 {select_id}，改用页面下拉')
                return False
            time.sleep(0.4)
            continue

        btn = None
        try:
            btn = selects[0].find_element(By.CSS_SELECTOR, 'button.dropdown')
        except Exception:
            btns = selects[0].find_elements(By.CSS_SELECTOR, 'button')
            btn = btns[0] if btns else None

        try:
            current = selects[0].find_element(
                By.CSS_SELECTOR, 'span.dropdown-text').text.strip()
            if _text_matches_aliases(current, aliases):
                print(f'{select_id} 已是「{current}」，跳过')
                return True
        except Exception:
            pass

        if btn:
            safe_click(driver, btn)
            time.sleep(0.7)

        clicked = _click_option_by_aliases(driver, aliases)
        if clicked:
            print(f'{select_id} 已选择: {clicked}')
            time.sleep(0.8)
            return True
        time.sleep(0.4)
    print(f'未能选择 {select_id} -> {option_text}')
    return False


_JS_COLLECTIONS_HELPERS = r'''
function paneVis(n) {
  if (!n) return false;
  const r = n.getBoundingClientRect();
  const st = window.getComputedStyle(n);
  if (!st || st.display === 'none' || st.visibility === 'hidden') return false;
  if (Number(st.opacity) === 0) return false;
  return r.width > 12 && r.height > 12;
}
function collectionsPane() {
  const panes = Array.from(document.querySelectorAll(
    '.cdk-overlay-container .cdk-overlay-pane, .cdk-overlay-pane,'
    + ' [role="listbox"], [role="menu"], [role="dialog"]'));
  const hits = [];
  for (const p of panes) {
    if (!paneVis(p)) continue;
    const t = p.innerText || '';
    const hasAll = t.includes('全选') || t.includes('Select All');
    const hasItem = t.includes('核心合集') || t.includes('Core Collection')
      || t.includes('SciELO') || t.includes('KCI');
    if (hasAll && hasItem) hits.push(p);
  }
  if (!hits.length) return null;
  hits.sort((a, b) => {
    const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
    return (rb.width * rb.height) - (ra.width * ra.height);
  });
  return hits[0];
}
function firstLine(t) {
  const raw = String(t || '');
  const line = raw.split(/\n/)[0] || '';
  return line.replace(/\s+/g, ' ').trim();
}
function rowLooksLikeWholeList(raw) {
  const s = raw || '';
  const n = s.split('\n').map(x => x.trim()).filter(Boolean).length;
  if (n > 2) return true;
  return (s.includes('Select All') || s.includes('全选'))
    && (s.includes('KCI') || s.includes('SciELO'))
    && (s.includes('Core Collection') || s.includes('核心合集'));
}
function rowOf(el, root) {
  let n = el;
  for (let i = 0; i < 10 && n && n !== root; i++) {
    const raw = ((n.innerText || n.textContent) || '').trim();
    const fl = firstLine(raw);
    if (fl && fl.length < 90 && !rowLooksLikeWholeList(raw)) {
      if (n.matches && n.matches(
        'label, li, [role="option"], [role="menuitem"], mat-checkbox,'
        + ' .mat-mdc-list-item, .mat-list-item, .mdc-list-item, [class*="list-item"]')) {
        return n;
      }
      if (n.querySelector && n.querySelector('input[type="checkbox"]')
          && fl.length < 80) return n;
    }
    n = n.parentElement;
  }
  return el.closest(
    'label, li, [role="option"], [role="menuitem"], mat-checkbox,'
    + ' .mat-mdc-list-item, .mdc-list-item') || el.parentElement;
}
function nativeInput(el) {
  if (!el) return null;
  if (el.tagName === 'INPUT') return el;
  return el.querySelector ? el.querySelector('input[type="checkbox"]') : null;
}
function boxChecked(el) {
  const inp = nativeInput(el);
  if (inp) return !!inp.checked;
  const aria = ((el.getAttribute && el.getAttribute('aria-checked')) || '').toLowerCase();
  if (aria === 'true') return true;
  if (aria === 'false') return false;
  const cls = String(el.className || '');
  return /mdc-checkbox--selected|mat-mdc-checkbox-checked|mat-checkbox-checked/.test(cls);
}
function clickCheckbox(el) {
  const row = el;
  const inp = nativeInput(row) || (row.tagName === 'INPUT' ? row : null);
  const inner = (row.querySelector && (
    row.querySelector('.mdc-checkbox')
    || row.querySelector('.mat-mdc-checkbox-inner-container')
    || row.querySelector('.mdc-checkbox__native-control')
    || row.querySelector('.mat-checkbox-inner-container'))) || null;
  const target = inner || inp || row;
  if (target.scrollIntoView) {
    try { target.scrollIntoView({block: 'nearest', inline: 'nearest'}); } catch (e) {}
  }
  if (inp) {
    inp.click();
    return true;
  }
  target.click();
  return true;
}
'''


def _is_collections_dropdown_open(driver):
    """合集下拉是否打开（可见「全选」+ 至少一个合集名）。"""
    return bool(driver.execute_script(
        _JS_COLLECTIONS_HELPERS + 'return !!collectionsPane();'))


def _open_collections_dropdown(driver, timeout=12):
    """打开「合集」下拉。已打开则直接返回。"""
    if _is_collections_dropdown_open(driver):
        return True
    opened = driver.execute_script(
        """
        function vis(b) {
          const r = b.getBoundingClientRect();
          return r.width > 2 && r.height > 2;
        }
        const keys = ['合集', 'collections', 'editions'];
        const btns = Array.from(document.querySelectorAll(
          'app-search-select-database button.dropdown, wos-select button.dropdown,'
          + ' button[role="combobox"], button.dropdown'));
        const hits = [];
        for (const b of btns) {
          if (!vis(b)) continue;
          const aria = (b.getAttribute('aria-label') || '').toLowerCase();
          const t = (b.innerText || '').toLowerCase();
          if (keys.some(k => aria.includes(k) || t.includes(k))) hits.push(b);
        }
        if (hits.length) { hits[hits.length - 1].click(); return true; }
        const bar = document.querySelector(
          'app-search-select-database, .select-database-on-simple-search, .select-db-group');
        if (bar) {
          const list = Array.from(
            bar.querySelectorAll('button.dropdown, button[role="combobox"]')
          ).filter(vis);
          if (list.length >= 2) { list[1].click(); return true; }
        }
        const ed = document.querySelector('wos-select#snSelectEd button.dropdown, #snSelectEd button');
        if (ed && vis(ed)) { ed.click(); return true; }
        return false;
        """)
    if not opened:
        opened = _click_named_dropdown(
            driver, ['合集', 'Collections', 'Editions', 'selectEdition'],
            None, which=1, open_only=True)
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _is_collections_dropdown_open(driver):
            time.sleep(0.25)
            return True
        time.sleep(0.25)
    return False


def _wait_collections_dropdown_closed(driver, timeout=3):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if not _is_collections_dropdown_open(driver):
            return True
        time.sleep(0.12)
    return not _is_collections_dropdown_open(driver)


def _read_collection_states(driver):
    """读取可见合集下拉的勾选状态，返回 {canonical_name: bool}，另含 '_select_all'。"""
    rows = driver.execute_script(
        _JS_COLLECTIONS_HELPERS + r'''
        const root = collectionsPane();
        if (!root) return [];
        const boxes = Array.from(root.querySelectorAll(
          'input[type="checkbox"], [role="checkbox"], mat-checkbox, .mat-mdc-checkbox'));
        const out = [];
        const used = new Set();
        for (const el of boxes) {
          const row = rowOf(el, root);
          if (!paneVis(el) && !paneVis(row)) continue;
          const t = firstLine((row && (row.innerText || row.textContent))
            || el.getAttribute('aria-label') || '');
          if (!t || t.length > 120 || rowLooksLikeWholeList(
                (row && (row.innerText || row.textContent)) || t)) continue;
          const key = t.toLowerCase();
          if (used.has(key)) continue;
          used.add(key);
          out.push({text: t, checked: boxChecked(row || el)});
        }
        return out;
        ''') or []
    states = {}
    for row in rows:
        text = (row or {}).get('text') or ''
        checked = bool((row or {}).get('checked'))
        first = text.split('\n', 1)[0].strip()
        if first in ('全选', 'Select All') or text in ('全选', 'Select All'):
            states['_select_all'] = checked
            continue
        canon = _canonical_collection(first) or _canonical_collection(text)
        if canon and canon != '全选':
            states[canon] = checked
    return states


def _click_collection_checkbox(driver, aliases, select_all=False, want_checked=None):
    """在可见合集下拉里点击勾选框。全选用精确匹配。

    want_checked: True/False 时若已是该状态则不点（避免把已勾选项点掉）。
    """
    return driver.execute_script(
        _JS_COLLECTIONS_HELPERS + r'''
        const aliases = arguments[0].map(a => (a || '').toLowerCase().replace(/\s+/g, ' ').trim());
        const selectAll = arguments[1];
        const wantChecked = arguments[2];
        function match(t) {
          const s = firstLine(t).toLowerCase();
          if (!s) return false;
          if (selectAll) return s === '全选' || s === 'select all';
          return aliases.some(a => {
            if (!a) return false;
            if (s === a) return true;
            if (a.length >= 4 && (s.includes(a) || a.includes(s))) return true;
            return false;
          });
        }
        const root = collectionsPane();
        if (!root) return null;
        const boxes = Array.from(root.querySelectorAll(
          'input[type="checkbox"], [role="checkbox"], mat-checkbox, .mat-mdc-checkbox'));
        let best = null;
        let bestLen = 1e9;
        let bestText = '';
        let bestRow = null;
        for (const el of boxes) {
          const row = rowOf(el, root);
          const t = ((row && (row.innerText || row.textContent))
            || el.getAttribute('aria-label') || '').trim();
          if (!match(t)) continue;
          if (rowLooksLikeWholeList(t) && !selectAll) continue;
          const fl = firstLine(t);
          if (fl.length < bestLen) {
            best = el;
            bestRow = row;
            bestLen = fl.length;
            bestText = fl;
          }
        }
        if (!best) return null;
        const host = bestRow || best;
        if (wantChecked === true || wantChecked === false) {
          if (boxChecked(host) === wantChecked) return 'already:' + bestText;
        }
        clickCheckbox(host);
        return bestText;
        ''',
        aliases, select_all, want_checked)



def _checked_collection_names(states):
    return [name for name in WOS_COLLECTIONS if states.get(name)]


def _ensure_collections_dropdown(driver):
    if _is_collections_dropdown_open(driver):
        return True
    return _open_collections_dropdown(driver)


def _toggle_select_all_collections(driver):
    """点击合集下拉中的「全选」。点「全选」不会关闭下拉。"""
    if not _ensure_collections_dropdown(driver):
        return None
    clicked = _click_collection_checkbox(
        driver, option_aliases('全选'), select_all=True)
    time.sleep(0.45)
    return clicked


def _try_clear_via_select_all(driver):
    """若「全选」能改变勾选，则全选再取消。返回是否已清空。"""
    if not _ensure_collections_dropdown(driver):
        return False
    before = _checked_collection_names(_read_collection_states(driver))
    print('合集：尝试全选 → 取消全选')
    if not _toggle_select_all_collections(driver):
        print('警告: 未点到「全选」')
        return False
    mid = _checked_collection_names(_read_collection_states(driver))
    if set(mid) == set(before):
        print('点「全选」未改变勾选，改为逐项取消多余合集')
        return False
    print(f'全选后勾选 {len(mid)} 项: {", ".join(mid) if mid else "无"}')
    _toggle_select_all_collections(driver)
    after = _checked_collection_names(_read_collection_states(driver))
    if after:
        print(f'取消全选后仍勾选 {", ".join(after)}')
        return False
    print('合集已全部取消勾选')
    return True


def _select_wos_collections(driver, collections):
    """只操作「合集」勾选。

    所有数据库页默认勾 4 个库。「全选」若无效，则逐项取消多余项。
    点「全选」不下拉关闭；点某一个合集会自动关闭下拉。
    已勾选的目标库不要再点，否则会被取消。
    """
    wanted = _normalize_wanted_collections(collections)
    _dismiss_open_overlays(driver)
    if not _open_collections_dropdown(driver):
        print('警告: 未能打开合集下拉')
        return False

    if wanted is None:
        before = _checked_collection_names(_read_collection_states(driver))
        clicked = _toggle_select_all_collections(driver)
        current = _checked_collection_names(_read_collection_states(driver))
        if len(current) < len(WOS_COLLECTIONS):
            print('全选后仍未齐，再点一次「全选」')
            clicked = _toggle_select_all_collections(driver) or clicked
            current = _checked_collection_names(_read_collection_states(driver))
        if clicked and len(current) == len(WOS_COLLECTIONS):
            print('已勾选合集: 全选')
        elif clicked:
            print(
                f'已点「全选」，当前勾选 {len(current)}/{len(WOS_COLLECTIONS)}'
                f'（原 {len(before)}）')
        else:
            print('警告: 未点到「全选」')
        _dismiss_open_overlays(driver)
        return bool(clicked)

    wanted_set = set(wanted)
    _try_clear_via_select_all(driver)

    for _ in range(len(WOS_COLLECTIONS) + 2):
        if not _ensure_collections_dropdown(driver):
            print('警告: 取消多余合集时下拉无法打开')
            return False
        current = _checked_collection_names(_read_collection_states(driver))
        extras = [name for name in current if name not in wanted_set]
        if not extras:
            break
        name = extras[0]
        print(f'取消勾选合集: {name}')
        clicked = _click_collection_checkbox(
            driver, option_aliases(name), want_checked=False)
        if clicked and str(clicked).startswith('already:'):
            clicked = _click_collection_checkbox(
                driver, option_aliases(name), want_checked=None)
        if not clicked:
            print(f'警告: 未能取消「{name}」')
            return False
        _wait_collections_dropdown_closed(driver)
        time.sleep(0.3)
        if _ensure_collections_dropdown(driver):
            now = _checked_collection_names(_read_collection_states(driver))
            if name in now and set(now) == set(current):
                print(f'警告: 点击后「{name}」仍勾选，勾选未变')
                return False

    for name in wanted:
        if not _ensure_collections_dropdown(driver):
            print(f'警告: 勾选「{name}」前未能打开合集下拉')
            return False
        current = _checked_collection_names(_read_collection_states(driver))
        if name in current:
            print(f'合集已勾选: {name}')
            continue
        print(f'勾选合集: {name}')
        clicked = _click_collection_checkbox(
            driver, option_aliases(name), want_checked=True)
        if not clicked:
            print(f'警告: 未勾选到合集「{name}」')
            return False
        _wait_collections_dropdown_closed(driver)
        time.sleep(0.3)

    if not _ensure_collections_dropdown(driver):
        return False
    current = _checked_collection_names(_read_collection_states(driver))
    _dismiss_open_overlays(driver)
    if set(current) != wanted_set:
        print(f'警告: 合集勾选结果为 {current or "无"}，期望 {wanted}')
        return False
    print(f'合集已是: {", ".join(wanted)}')
    return True


def _current_dropdown_text(driver, select_id):
    try:
        selects = driver.find_elements(By.CSS_SELECTOR, f'wos-select#{select_id}')
        if not selects:
            return ''
        span = selects[0].find_elements(By.CSS_SELECTOR, 'span.dropdown-text')
        return (span[0].text or '').strip() if span else ''
    except Exception:
        return ''


def _select_citation_index(driver, citation_index):
    """仅当需要非 All 的引文索引时，才点 Editions 下拉。"""
    index_name = citation_index or 'All'
    if not index_name or index_name in ('All', '全部', '全选'):
        return True
    _dismiss_open_overlays(driver)
    current_ed = _current_dropdown_text(driver, 'snSelectEd')
    if _text_matches_aliases(current_ed, option_aliases(index_name)):
        print(f'引文索引已是: {current_ed or index_name}')
        return True
    ed_ok = _wos_select_option(driver, 'snSelectEd', index_name)
    if not ed_ok:
        ed_ok = _click_named_dropdown(
            driver,
            ['引文索引', 'Collections', 'selectEdition', 'Editions'],
            option_aliases(index_name), which=1)
    if not ed_ok:
        print(f'警告: 未能切换引文索引到「{index_name}」')
        return False
    time.sleep(0.5)
    return True


def ensure_database_and_collections(
        driver, database=SEARCH_DATABASE, collections=SEARCH_COLLECTIONS):
    """核心合集打开 woscc；其余合集打开 alldb 再勾选对应 Collections。"""
    target = search_url_for(database, collections)
    mode, extra = resolve_search_mode(database, collections)
    print(f'跳转检索库: {target}')
    if not goto_search_url(driver, target):
        print(f'未能打开检索库: {target}')
        return False
    if mode == 'core':
        if extra and extra not in ('All', '全部', '全选'):
            return _select_citation_index(driver, extra)
        return True
    label = '全选' if extra is None else ', '.join(extra)
    print(f'在所有数据库页勾选合集: {label}')
    if not _select_wos_collections(driver, extra):
        print(f'未能勾选合集: {label}')
        return False
    return True


def _click_named_dropdown(driver, trigger_keys, aliases, which=0, open_only=False):
    """在选择数据库栏打开第 which 个下拉，并按 aliases 点选项。"""
    opened = driver.execute_script(
        """
        const keys = arguments[0].map(k => k.toLowerCase());
        const which = arguments[1];
        const btns = Array.from(document.querySelectorAll(
          'app-search-select-database button.dropdown, wos-select button.dropdown, button[role="combobox"]'));
        const hits = [];
        for (const b of btns) {
          const aria = (b.getAttribute('aria-label') || '').toLowerCase();
          const t = (b.innerText || '').toLowerCase();
          const wrap = (b.closest('app-search-select-database, .select-db-group, .select-database-on-simple-search') || b.parentElement);
          const wt = wrap ? (wrap.innerText || '').toLowerCase() : '';
          if (keys.some(k => aria.includes(k) || t.includes(k) || wt.includes(k))) {
            hits.push(b);
          }
        }
        // 兜底：选择栏里从左到右的下拉
        let list = hits;
        if (!list.length) {
          const bar = document.querySelector(
            'app-search-select-database, .select-database-on-simple-search');
          if (bar) list = Array.from(bar.querySelectorAll('button.dropdown, button[role="combobox"]'));
        }
        if (!list.length) return false;
        const idx = Math.min(which, list.length - 1);
        list[idx].click();
        return true;
        """,
        trigger_keys, which)
    if not opened:
        return False
    time.sleep(0.7)
    if open_only or not aliases:
        return True
    clicked = _click_option_by_aliases(driver, aliases)
    if clicked:
        print(f'已选择: {clicked}')
        time.sleep(0.8)
        return True
    return False


def _dismiss_open_overlays(driver):
    """关闭残留下拉/对话框，避免挡住检索字段选择。"""
    try:
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        time.sleep(0.25)
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
    except Exception:
        pass
    time.sleep(0.2)


def _date_range_inputs_present(driver):
    """Publication Date 选中后才会出现的起止日期框。"""
    if driver.find_elements(By.CSS_SELECTOR, 'input[data-ta="search-criteria-input-2"]'):
        return True
    if driver.find_elements(By.CSS_SELECTOR, 'input[aria-label="Search to box"]'):
        return True
    splits = driver.find_elements(By.CSS_SELECTOR, '.split-field-input input')
    return len(splits) >= 2


def _is_publication_date_field_active(driver):
    """仅当起止日期框已出现才视为已切到 Publication Date。

    不能凭 overlay 里选项的 data-ta=Publication Date 判断，
    否则会误以为已选中、跳过点击，日期框永远出不来。
    """
    return _date_range_inputs_present(driver)


def _open_search_field_dropdown(driver):
    """打开检索字段下拉（默认多为 Topic），不要点到库/合集下拉。"""
    _dismiss_open_overlays(driver)
    # 只点「Select search field」触发器
    for sel in (
        'button[aria-label*="Select search field"]',
        'button[aria-haspopup="dialog"][data-ta="Topic"]',
        'button[aria-haspopup="dialog"][data-ta="Publication Date"]',
    ):
        els = driver.find_elements(By.CSS_SELECTOR, sel)
        for el in els:
            try:
                if not el.is_displayed():
                    continue
                aria = (el.get_attribute('aria-label') or '').lower()
                if 'database' in aria or 'edition' in aria:
                    continue
                if 'select search field' not in aria and 'data-ta' not in sel:
                    # 无 aria 时仅允许 Topic / Publication Date 的 combobox
                    if el.get_attribute('aria-haspopup') != 'dialog':
                        continue
                if safe_click(driver, el):
                    time.sleep(0.8)
                    return True
            except Exception:
                continue
    return bool(driver.execute_script(
        """
        const btns = Array.from(document.querySelectorAll(
          'button[role="combobox"], button.dropdown'));
        for (const b of btns) {
          const aria = (b.getAttribute('aria-label') || '').toLowerCase();
          const ta = (b.getAttribute('data-ta') || '');
          const t = (b.innerText || '').trim();
          if (aria.includes('select search field') || aria.includes('检索字段')
              || t === '主题' || t === 'Topic' || t === '出版日期'
              || t === 'Publication Date') {
            const r = b.getBoundingClientRect();
            if (r.width < 2 || r.height < 2) continue;
            b.click();
            return true;
          }
          if (ta === 'Topic' || ta === 'Publication Date' || ta === 'Author'
              || ta === 'Title') {
            const r = b.getBoundingClientRect();
            if (r.width < 2 || r.height < 2) continue;
            b.click();
            return true;
          }
        }
        return false;
        """))


def _click_publication_date_option(driver):
    """在已打开的字段下拉里点 Publication Date 筛选项。"""
    clicked = driver.execute_script(
        """
        function visible(n) {
          const r = n.getBoundingClientRect();
          return r.width > 0 && r.height > 0;
        }
        function isTrigger(n) {
          const aria = (n.getAttribute('aria-label') || '');
          return aria.includes('Select search field') || aria.includes('检索字段');
        }
        const panes = Array.from(document.querySelectorAll(
          '.cdk-overlay-pane, [role="dialog"], [role="listbox"]'));
        const roots = panes.length ? panes : [document.body];
        const selectors = [
          'button[data-ta="Publication Date"]',
          '[data-ta="Publication Date"]',
          '[role="option"]',
          '[role="menuitem"]',
          'button', 'li', 'span', 'div', 'a'
        ];
        for (const root of roots) {
          for (const sel of selectors) {
            const nodes = Array.from(root.querySelectorAll(sel));
            for (const n of nodes) {
              if (!visible(n) || isTrigger(n)) continue;
              const ta = (n.getAttribute('data-ta') || '').trim();
              const t = (n.innerText || '').trim();
              if (ta === 'Publication Date' || t === 'Publication Date'
                  || t === '出版日期') {
                n.click();
                return t || ta;
              }
            }
          }
        }
        return null;
        """)
    if clicked:
        print(f'已点击筛选项: {clicked}')
        time.sleep(0.6)
        return True
    return False


def select_publication_date_field(driver, timeout=18):
    """先打开检索字段下拉，点 Publication Date，再等到起止日期框出现。"""
    if _date_range_inputs_present(driver):
        print('检索字段已是 Publication Date（日期框已在）')
        return True
    if is_page_unusable(driver):
        print('选择检索字段时页面已失效')
        return False
    if not is_search_form_ready(driver):
        if not wait_search_form_ready(driver, timeout=8):
            print('检索表单未就绪，无法选择 Publication Date')
            return False

    miss = 0
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_page_unusable(driver):
            print('选择检索字段时页面已失效')
            return False
        try:
            url = (driver.current_url or '').lower()
        except Exception:
            url = ''
        if 'basic-search' not in url and 'advanced-search' not in url:
            print('当前不是检索页，停止点检索字段')
            return False
        if not is_search_form_ready(driver) and not _date_range_inputs_present(driver):
            miss += 1
            if miss >= 4:
                print('检索表单未出现，停止空点检索字段')
                return False
            print('检索表单未就绪，等待...')
            wait_search_form_ready(driver, timeout=3)
            continue
        if not _open_search_field_dropdown(driver):
            miss += 1
            print(f'未找到检索字段下拉，重试 ({miss}/6)...')
            if miss >= 6:
                print('多次未找到检索字段下拉，停止')
                break
            time.sleep(0.6)
            continue
        miss = 0

        if not _click_publication_date_option(driver):
            print('下拉已开但未点到 Publication Date，重试...')
            _dismiss_open_overlays(driver)
            time.sleep(0.4)
            continue

        wait_deadline = time.time() + 6
        while time.time() < wait_deadline:
            if _date_range_inputs_present(driver):
                print('已选择检索字段: Publication Date（日期框已出现）')
                time.sleep(0.3)
                return True
            time.sleep(0.25)

        print('已点击 Publication Date，但日期框尚未出现，重试...')
        _dismiss_open_overlays(driver)
        time.sleep(0.4)

    print('改点「添加日期范围」...')
    if js_click_by_text(
            driver, ['添加日期范围', 'add date range', '+ 添加日期范围'],
            exact=False):
        wait_deadline = time.time() + 8
        while time.time() < wait_deadline:
            if _date_range_inputs_present(driver):
                print('已添加日期范围（日期框已出现）')
                return True
            time.sleep(0.3)
    print('未能切换到 Publication Date / 添加日期范围')
    return False


def set_publication_date_range(driver, start_date, end_date, timeout=20):
    """填写出版日期起止 YYYY-MM-DD；若无日期框则先切到 Publication Date。"""
    if not _date_range_inputs_present(driver):
        print('日期框未出现，先选择 Publication Date 筛选项...')
        if not select_publication_date_field(driver, timeout=max(timeout, 20)):
            raise RuntimeError(
                '找不到出版日期起止输入框（未能选择 Publication Date 筛选项）')

    deadline = time.time() + 8
    from_inputs, to_inputs = [], []
    while time.time() < deadline:
        from_inputs = driver.find_elements(
            By.CSS_SELECTOR, 'input[data-ta="search-criteria-input"]')
        to_inputs = driver.find_elements(
            By.CSS_SELECTOR, 'input[data-ta="search-criteria-input-2"]')
        if not from_inputs or not to_inputs:
            from_inputs = driver.find_elements(
                By.CSS_SELECTOR, 'input[aria-label="Search from box"]')
            to_inputs = driver.find_elements(
                By.CSS_SELECTOR, 'input[aria-label="Search to box"]')
        if from_inputs and to_inputs:
            break
        split = driver.find_elements(By.CSS_SELECTOR, '.split-field-input input')
        if len(split) >= 2:
            from_inputs, to_inputs = [split[0]], [split[1]]
            break
        time.sleep(0.3)

    if not from_inputs or not to_inputs:
        raise RuntimeError(
            '找不到出版日期起止输入框（请确认已选择 Publication Date）')

    fill_input(driver, from_inputs[0], start_date)
    time.sleep(0.3)
    fill_input(driver, to_inputs[0], end_date)
    print(f'设置出版日期: {start_date} ~ {end_date}')
    time.sleep(0.5)


def click_run_search(driver):
    """点击 Search。"""
    btns = driver.find_elements(By.CSS_SELECTOR, 'button[data-ta="run-search"]')
    if not btns:
        btns = driver.find_elements(
            By.CSS_SELECTOR, 'button[data-pendo="run-search"]')
    if not btns:
        clicked = js_click_by_text(driver, ['检索', 'search'], exact=True)
        if clicked:
            print(f'已点击 {clicked}')
            time.sleep(4)
            return True
        if click_by_text(driver, ['Search', '检索'], tag='button', timeout=5):
            print('已点击 Search/检索')
            time.sleep(4)
            return True
        raise RuntimeError('找不到 Search/检索 按钮')
    safe_click(driver, btns[0])
    print('已点击 Search')
    time.sleep(4)
    return True


def parse_hit_count(driver, timeout=None):
    """解析检索命中数。返回 int；无结果返回 0；超时返回 None。"""
    timeout = HIT_COUNT_TIMEOUT if timeout is None else timeout
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_page_unusable(driver):
            print('解析命中数时页面已失效')
            return None
        # 常见计数选择器
        selectors = [
            '#hitCount.hitCount',
            '#hitCount',
            'span.hitCount',
            '.brand-blue',
            '[data-ta="hit-count"]',
            'app-search-results-count',
        ]
        for sel in selectors:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                try:
                    text = (el.text or '').strip().replace(',', '').replace(' ', '')
                    m = re.search(r'(\d+)', text)
                    if m:
                        n = int(m.group(1))
                        print(f'命中条数: {n}')
                        return n
                except Exception:
                    continue

        # 页面正文兜底
        try:
            body = driver.find_element(By.TAG_NAME, 'body').text
        except Exception:
            body = ''

        for pat in (
            r'([\d,]+)\s*results?',
            r'([\d,]+)\s*条结果',
            r'Results?:\s*([\d,]+)',
            r'找到\s*([\d,]+)',
        ):
            m = re.search(pat, body, re.I)
            if m:
                n = int(m.group(1).replace(',', ''))
                print(f'命中条数(正文): {n}')
                return n

        if any(kw in body for kw in (
                'No results', '0 results', '没有找到', '未找到结果',
                'Your search found no results')):
            print('命中条数: 0')
            return 0

        # 结果列表出现也继续等计数
        time.sleep(1)

    print('解析命中条数超时')
    return None


def set_sort_date_ascending(driver, timeout=25):
    """切换排序为日期升序（最旧优先）。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        # 若已是 Date / Oldest
        labels = driver.find_elements(
            By.CSS_SELECTOR,
            'span.sort-selected, .sort-selected, [class*="sort-selected"]')
        for lab in labels:
            t = (lab.text or '').strip().lower()
            if any(k in t for k in ('oldest', 'date: oldest', '日期', '最旧')):
                if 'relevan' not in t:
                    print(f'当前排序已是日期相关: {lab.text.strip()}')
                    # 再确认是否升序；若是 Newest 则继续切换
                    if 'newest' in t or '最新' in t:
                        break
                    if 'oldest' in t or '最旧' in t or 'ascending' in t:
                        return True

        # 点击当前排序触发器
        clicked_trigger = False
        for sel in (
            'span.sort-selected',
            '.sort-selected',
            'button[aria-label*="Sort"]',
            '[data-ta*="sort"]',
        ):
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                try:
                    if el.is_displayed():
                        safe_click(driver, el)
                        clicked_trigger = True
                        time.sleep(0.7)
                        break
                except Exception:
                    continue
            if clicked_trigger:
                break

        if not clicked_trigger:
            # 点「Relevance」文本
            click_by_text(driver, ['Relevance', '相关度'], tag='span', timeout=2)

        # 选择日期升序选项
        option_texts = [
            'Date: oldest first',
            'Oldest first',
            'Date (oldest first)',
            'Publication Date — oldest first',
            '日期: 最旧优先',
            '最旧优先',
            'Date ascending',
        ]
        for opt in option_texts:
            if click_by_text(driver, [opt], tag='*', timeout=2):
                print(f'已切换排序: {opt}')
                time.sleep(3)
                return True

        clicked = driver.execute_script(
            """
            const keys = ['oldest', '最旧', 'ascending', '升序'];
            const nodes = Array.from(document.querySelectorAll(
              'button, a, span, li, div, [role="option"], [role="menuitem"]'));
            for (const n of nodes) {
              const t = (n.innerText || '').trim().toLowerCase();
              if (!t || t.length > 80) continue;
              if (t.includes('newest') || t.includes('最新')) continue;
              if (keys.some(k => t.includes(k)) && t.includes('date') ||
                  keys.some(k => t.includes(k))) {
                if (n.offsetParent === null) continue;
                // 优先含 date/日期
                if (t.includes('date') || t.includes('日期') || t.includes('oldest')
                    || t.includes('最旧')) {
                  n.click();
                  return n.innerText.trim();
                }
              }
            }
            return null;
            """)
        if clicked:
            print(f'已切换排序(JS): {clicked}')
            time.sleep(3)
            return True
        time.sleep(0.5)

    print('未能切换到日期升序，继续使用当前排序')
    return False


# ---------------------------------------------------------------------------
# 导出
# ---------------------------------------------------------------------------
def close_export_overlay(driver):
    """仅关闭导出相关侧栏/对话框。

    切勿点结果页精炼标签上的 Close，也勿在无导出层时盲目 Esc，
    否则会取消已应用的 Research Areas 等 Refine 条件。
    """
    closed = False

    # 1) 明确的导出弹窗 / 侧栏内的关闭按钮
    try:
        closed = bool(driver.execute_script(
            r'''
            function visible(el) {
              if (!el) return false;
              const s = window.getComputedStyle(el);
              if (s.display === 'none' || s.visibility === 'hidden') return false;
              if (parseFloat(s.opacity || '1') < 0.1) return false;
              const r = el.getBoundingClientRect();
              return r.width > 0 && r.height > 0;
            }
            function isFilterChipClose(btn) {
              const lab = ((btn.getAttribute('aria-label') || '') + ' '
                + (btn.textContent || '')).toLowerCase();
              // 精炼标签移除按钮
              if (lab.includes('research area') || lab.includes('研究方向')) return true;
              if (lab.includes('remove filter') || lab.includes('clear filter')) return true;
              if (lab.includes('移除') && lab.includes('筛选')) return true;
              const chip = btn.closest(
                'mat-chip, .mat-mdc-chip, [class*="selected-filter"],'
                + '[class*="applied-filter"], [class*="refine-pill"],'
                + '[class*="filter-value"], app-search-history, app-qrf');
              if (chip) return true;
              return false;
            }
            function isExportUi(el) {
              if (!el) return false;
              if (el.closest && el.closest(
                '#onetrust-pc-sdk, #onetrust-consent-sdk, #onetrust-banner-sdk')) {
                return false;
              }
              const root = el.closest(
                '[role="dialog"], mat-dialog-container, .mat-mdc-dialog-container,'
                + ' .cdk-overlay-pane, app-export-menu, .export-menu, #exportMenu');
              if (!root) return false;
              const t = ((root.innerText || '') + ' ' + (root.id || '')
                + ' ' + (root.className || '')).toLowerCase();
              return t.includes('export') || t.includes('导出')
                || t.includes('excel') || t.includes('record content')
                || t.includes('记录内容') || !!root.querySelector(
                  '#exportToExcelButton, [id*="exportToExcel"], [data-ta*="export"]');
            }
            const candidates = Array.from(document.querySelectorAll(
              'button[aria-label="Close"], button[aria-label="关闭"],'
              + ' button[aria-label*="Close dialog"], button[aria-label*="关闭对话框"],'
              + ' [data-ta="export-close"], button.close,'
              + ' mat-icon[data-mat-icon-name="close"],'
              + ' button .mat-icon'));
            for (const el of candidates) {
              const btn = el.closest('button') || el;
              if (!visible(btn)) continue;
              if (isFilterChipClose(btn)) continue;
              if (!isExportUi(btn)) continue;
              btn.click();
              return true;
            }
            return false;
            '''))
    except Exception:
        closed = False

    if closed:
        time.sleep(0.45)
        return True

    # 2) 仅当导出 Excel 弹窗仍打开时，才用 Esc 关闭（避免冲掉 Refine）
    if _export_excel_dialog_open(driver):
        try:
            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(0.4)
            if not _export_excel_dialog_open(driver):
                return True
        except Exception:
            pass

    return False


def open_excel_export(driver, timeout=None):
    """打开 Export → Excel（#exportToExcelButton）。"""
    timeout = EXPORT_DIALOG_TIMEOUT if timeout is None else timeout
    # 仅关闭已打开的导出层，避免 Esc/误点清掉 Refine
    if _export_excel_dialog_open(driver) or _export_dialog_el(driver):
        close_export_overlay(driver)
        time.sleep(0.3)

    # 1) 点结果栏 Export（精确匹配，避免点到别的 Export）
    export_opened = False
    clicked = js_click_by_text(driver, ['export', '导出'], exact=True)
    if clicked:
        export_opened = True
        print(f'已打开 Export 菜单: {clicked}')
        time.sleep(1.0)

    if not export_opened:
        for sel in (
            'button[data-ta="export-button"]',
            'app-export-menu button',
            'button[aria-label="Export"]',
            'button[aria-label="导出"]',
            '#exportMenu button',
            'button.cdx-but-md',
        ):
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                try:
                    t = (el.text or '').strip().lower()
                    if t and t not in ('export', '导出') and 'export' not in t:
                        continue
                    if safe_click(driver, el):
                        export_opened = True
                        print(f'已打开 Export 菜单（{sel}）')
                        time.sleep(1.0)
                        break
                except Exception:
                    continue
            if export_opened:
                break

    if not export_opened:
        # 有的版本 Excel 按钮直接在工具栏
        print('未点到 Export 按钮，尝试直接点 Excel')

    # 2) 选 Excel：优先官方 id
    deadline = time.time() + timeout
    while time.time() < deadline:
        for sel in (
            '#exportToExcelButton',
            'button#exportToExcelButton',
            '[id="exportToExcelButton"]',
            'button[data-ta="export-excel"]',
            'button[id*="Excel"]',
            'button[id*="excel"]',
        ):
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                try:
                    driver.execute_script('arguments[0].click();', el)
                    print(f'已选择 Excel 导出（{sel}）')
                    wait_deadline = time.time() + 6
                    while time.time() < wait_deadline:
                        if _export_excel_dialog_open(driver) and _export_dialog_el(driver):
                            return True
                        time.sleep(0.3)
                    if _export_excel_dialog_open(driver):
                        return True
                    continue
                except Exception:
                    continue

        if js_click_by_text(
                driver,
                ['excel', 'excel file', 'excel 文件', '导出为 excel', '导出到 excel',
                 'export to excel', 'microsoft excel', '导出至 excel', '导出excel'],
                exact=False):
            print('已选择 Excel 导出（文本）')
            time.sleep(0.8)
            if _export_excel_dialog_open(driver):
                return True
        time.sleep(0.35)

    # 调试：打印当前菜单文案
    try:
        menu_text = driver.execute_script(
            """
            const roots = document.querySelectorAll(
              '.cdk-overlay-pane, [role="menu"], app-export-menu, .mat-mdc-menu-panel');
            return Array.from(roots).map(r => (r.innerText || '').slice(0, 400)).join('\\n---\\n');
            """)
        print('未找到 Excel。当前菜单文案:', (menu_text or '')[:500])
    except Exception:
        print('未找到 Excel 导出选项')
    return False


def _export_excel_dialog_open(driver):
    """Export Records to Excel 弹窗是否已出现（英文界面）。"""
    try:
        return bool(driver.execute_script(
            """
            const t = (document.body.innerText || '').toLowerCase();
            return t.includes('export records to excel')
              || t.includes('将记录导出至 excel')
              || (t.includes('records from:') && t.includes('record content'))
              || (t.includes('all records on page') && t.includes('no more than 1000'))
              || t.includes('记录内容')
              || t.includes('一次不能超过');
            """))
    except Exception:
        return False


def _close_pendo_guides(driver):
    """关掉 Pendo 引导层，避免挡住导出弹窗。"""
    for sel in (
        'button._pendo-close-guide',
        'button._pendo-button-primaryButton',
        'button._pendo-button-secondaryButton',
        '[id^="pendo-close-guide"]',
    ):
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if el.is_displayed():
                    driver.execute_script('arguments[0].click();', el)
            except Exception:
                continue


def _close_cookie_preference_center(driver):
    """关掉 OneTrust Cookie Preference Center，避免当成导出弹窗。"""
    for sel in (
        '#close-pc-btn-handler',
        '#onetrust-close-btn-handler',
        '#onetrust-pc-sdk button[aria-label="Close"]',
        '#onetrust-accept-btn-handler',
        '.ot-pc-refuse-all-handler',
    ):
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if el.is_displayed():
                    driver.execute_script('arguments[0].click();', el)
                    time.sleep(0.3)
            except Exception:
                continue
    try:
        driver.execute_script(
            """
            ['#onetrust-pc-sdk', '#onetrust-banner-sdk'].forEach(sel => {
              const el = document.querySelector(sel);
              if (el) el.style.setProperty('display', 'none', 'important');
            });
            """)
    except Exception:
        pass


def _export_dialog_el(driver):
    """定位 Export Records to Excel 弹窗。排除 Cookie / 排序 overlay。"""
    return driver.execute_script(
        """
        function norm(s) { return (s || '').toLowerCase(); }
        function reject(el) {
          const id = ((el.id || '') + ' ' + (el.className || '')).toLowerCase();
          const t = norm(el.innerText || '');
          if (id.includes('onetrust') || id.includes('ot-pc') || id.includes('ot-sdk')) return true;
          if (t.includes('cookie preference') || t.includes('strictly necessary cookies')) return true;
          if (el.closest && el.closest('#onetrust-pc-sdk, #onetrust-consent-sdk, #onetrust-banner-sdk')) {
            return true;
          }
          return false;
        }
        function match(t) {
          t = norm(t);
          if (t.includes('cookie preference')) return false;
          if (t.includes('export records to excel') && t.includes('record content')) return true;
          if (t.includes('将记录导出至 excel')) return true;
          if (t.includes('records from:') && t.includes('record content')
              && t.includes('no more than 1000')) return true;
          if (t.includes('记录内容') && (t.includes('excel') || t.includes('导出'))) return true;
          return false;
        }
        const sels = [
          'app-export-overlay',
          'app-export-out-details',
          'app-export-records',
          'mat-dialog-container',
          '.mat-mdc-dialog-container',
          '.mdc-dialog__container',
          '[role="dialog"]',
          'app-export',
          'app-export-modal',
          'app-export-excel',
          '[class*="export"]',
          '.cdk-overlay-pane',
          '.mdc-dialog'
        ];
        let best = null;
        let bestLen = 1e12;
        const consider = (el) => {
          if (!el || reject(el) || !match(el.innerText || '')) return;
          const len = (el.innerText || '').length;
          if (len < bestLen) { best = el; bestLen = len; }
        };
        for (const sel of sels) {
          document.querySelectorAll(sel).forEach(consider);
        }
        if (best) return best;
        const all = document.querySelectorAll(
          'div, section, aside, form, app-export-overlay, app-export-out-details');
        for (const el of all) {
          if (el.childElementCount > 60) continue;
          consider(el);
        }
        return best;
        """)


def _set_angular_input(driver, el, value):
    """用 native setter + send_keys 写入 Angular/Material 输入框。"""
    value = str(value)
    try:
        driver.execute_script(
            """
            const el = arguments[0];
            const v = arguments[1];
            el.removeAttribute('readonly');
            el.removeAttribute('disabled');
            el.disabled = false;
            el.focus();
            const proto = window.HTMLInputElement.prototype;
            const desc = Object.getOwnPropertyDescriptor(proto, 'value');
            if (desc && desc.set) {
              desc.set.call(el, '');
              el.dispatchEvent(new Event('input', {bubbles:true}));
              desc.set.call(el, v);
            } else {
              el.value = v;
            }
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            """,
            el, value)
    except Exception:
        pass
    try:
        el.click()
        el.send_keys(Keys.CONTROL, 'a')
        el.send_keys(Keys.BACKSPACE)
        el.send_keys(value)
    except Exception:
        pass
    try:
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('blur', {bubbles:true}));",
            el)
    except Exception:
        pass
    try:
        return (el.get_attribute('value') or '').strip() == value
    except Exception:
        return False


def _select_export_range_radio(driver, dialog):
    """选中 Records from:（#radio3），避开 All records on page。"""
    for sel in (
        '#radio3 input[type="radio"]',
        '#radio3 label',
        'mat-radio-button#radio3',
        '#radio3',
        'label[for="radio3"]',
    ):
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if el.is_displayed() and safe_click(driver, el):
                    print(f'已选择记录范围单选: {sel}')
                    return True
            except Exception:
                continue
    picked = driver.execute_script(
        """
        const dialog = arguments[0] || document.body;
        function clickRadio(host) {
          const inp = host.querySelector && host.querySelector('input[type="radio"]');
          if (inp) { inp.click(); return; }
          const circle = host.querySelector && host.querySelector(
            '.mdc-radio, .mat-radio-container, .mdc-radio__native-control');
          (circle || host).click();
        }
        const radios = Array.from(dialog.querySelectorAll(
          'mat-radio-button, [role="radio"]'));
        // 优先点带「至」的「记录:」行（中文弹窗）
        for (const n of radios) {
          const t = (n.innerText || '').replace(/\\s+/g, ' ').trim();
          if (/页面上的所有|all records on page/i.test(t)) continue;
          if ((/记录[:：]/.test(t) && t.includes('至'))
              || /Records from/i.test(t)) {
            clickRadio(n);
            return t.slice(0, 40);
          }
        }
        for (const n of radios) {
          const t = (n.innerText || '').replace(/\\s+/g, ' ').trim();
          if (/页面上的所有|all records on page/i.test(t)) continue;
          if (/^记录[:：]/.test(t) || t === '记录:' || t === '记录：') {
            clickRadio(n);
            return t.slice(0, 40);
          }
        }
        const r3 = dialog.querySelector(
          '#radio3, mat-radio-button#radio3, [id="radio3"], label[for="radio3"]');
        if (r3) { clickRadio(r3); return 'radio3'; }
        return null;
        """,
        dialog)
    if picked:
        print(f'已选择记录范围单选: {picked}')
        return True
    for xp in (
        '//*[contains(text(), "Records from:")]',
        '//mat-radio-button[contains(., "Records from")]',
        '//label[contains(., "Records from")]',
        '//mat-radio-button[contains(., "记录") and contains(., "至")]',
        '//*[normalize-space(text())="记录:"]',
    ):
        for el in driver.find_elements(By.XPATH, xp):
            try:
                if not el.is_displayed():
                    continue
                if el.find_elements(By.XPATH, './ancestor::*[@id="onetrust-pc-sdk" or @id="onetrust-consent-sdk"]'):
                    continue
                if safe_click(driver, el):
                    print(f'已选择记录范围单选: {xp}')
                    return True
            except Exception:
                continue
    print('未点到「记录:」单选，继续尝试填数字框')
    return False


def _collect_export_range_inputs(driver, dialog):
    """在整页（排除 Cookie 层）找到 Records from: 旁的两个数字框。"""
    return driver.execute_script(
        """
        function noise(el) {
          return !!(el.closest && el.closest(
            '#onetrust-pc-sdk, #onetrust-banner-sdk, #onetrust-consent-sdk, .ot-sdk-container'
          ));
        }
        function usable(i) {
          if (!i || noise(i)) return false;
          const t = (i.type || 'text').toLowerCase();
          if (/^(radio|checkbox|hidden|button|submit|file)$/.test(t)) return false;
          const aria = (i.getAttribute('aria-label') || '').toLowerCase();
          if (aria.includes('cookie') || aria.includes('vendor') || aria.includes('search')) {
            return false;
          }
          return true;
        }
        function collect(root) {
          if (!root) return [];
          return Array.from(root.querySelectorAll(
            'input.mat-mdc-input-element, input.mat-input-element,'
            + ' input[id^="mat-input-"], input[type="text"], input[type="number"], input:not([type])'
          )).filter(usable);
        }
        const labels = Array.from(document.querySelectorAll(
          'mat-radio-button, label, span, div, [role="radio"]'));
        for (const n of labels) {
          if (noise(n)) continue;
          const t = (n.innerText || '').replace(/\\s+/g, ' ').trim();
          if (!/Records from/i.test(t) && !/记录[:：]/.test(t)) continue;
          if (t.length > 80) continue;
          let scope = n.closest(
            'mat-radio-button, mat-radio-group, form, [class*="export"], [role="dialog"]'
          ) || n.parentElement;
          for (let i = 0; i < 8 && scope; i++) {
            if (noise(scope)) break;
            const ins = collect(scope);
            if (ins.length >= 2) return ins.slice(0, 2);
            scope = scope.parentElement;
          }
        }
        const dialog = arguments[0];
        if (dialog && !noise(dialog)) {
          const ins = collect(dialog);
          if (ins.length >= 2) return ins.slice(0, 2);
        }
        const mats = collect(document.body).filter(i => {
          const id = (i.id || '');
          const cls = String(i.className || '');
          return id.startsWith('mat-input') || cls.includes('mat-mdc-input')
            || cls.includes('mat-input-element');
        });
        return mats.length >= 2 ? mats.slice(0, 2) : mats;
        """,
        dialog) or []


def _dump_export_range_debug(driver, dialog):
    """失败时打印弹窗内 input 概况，便于对照。"""
    try:
        info = driver.execute_script(
            """
            const dialog = arguments[0] || document.body;
            const inputs = Array.from(dialog.querySelectorAll('input')).slice(0, 12);
            const panes = Array.from(document.querySelectorAll(
              'app-export-overlay, app-export-out-details, mat-dialog-container,'
              + ' [role="dialog"], .cdk-overlay-pane, app-export'));
            return {
              text: (dialog.innerText || '').slice(0, 280),
              panes: panes.map(p => (p.innerText || '').slice(0, 80)),
              inputs: inputs.map(i => ({
                id: i.id, name: i.name, type: i.type,
                aria: i.getAttribute('aria-label') || '',
                disabled: !!i.disabled, value: i.value || ''
              }))
            };
            """,
            dialog)
        print('导出弹窗调试:', info)
    except Exception as e:
        print('导出弹窗调试失败:', e)


def set_export_record_range(driver, start_n, end_n):
    """选择 Records from: 并填写起止（每批 ≤EXPORT_BATCH_SIZE）。界面默认 English。"""
    _close_cookie_preference_center(driver)
    deadline = time.time() + 10
    dialog = None
    while time.time() < deadline:
        if _export_excel_dialog_open(driver):
            dialog = _export_dialog_el(driver)
            if dialog:
                break
        time.sleep(0.25)
    if dialog is None:
        dialog = _export_dialog_el(driver)
        if dialog is None and _export_excel_dialog_open(driver):
            print('导出弹窗容器未定位，将在页面中按 Records from: 查找数字框')

    _select_export_range_radio(driver, dialog)
    time.sleep(0.6)

    inputs = []
    for _ in range(8):
        inputs = _collect_export_range_inputs(driver, dialog)
        if len(inputs) >= 2:
            break
        time.sleep(0.35)
        dialog = _export_dialog_el(driver) or dialog

    if len(inputs) < 2:
        print('未能设置导出记录范围：弹窗内找不到起止数字框')
        _dump_export_range_debug(driver, dialog)
        return False

    ok_from = _set_angular_input(driver, inputs[0], start_n)
    time.sleep(0.15)
    ok_to = _set_angular_input(driver, inputs[1], end_n)
    time.sleep(0.2)

    got_from = ''
    got_to = ''
    try:
        got_from = (inputs[0].get_attribute('value') or '').strip()
        got_to = (inputs[1].get_attribute('value') or '').strip()
    except Exception:
        pass
    if got_from == str(start_n) and got_to == str(end_n):
        print(f'导出范围: {start_n} - {end_n}')
        return True
    print(
        f'未能设置导出记录范围（期望 {start_n}-{end_n}，'
        f'实际 {got_from or "?"}-{got_to or "?"}，'
        f'native={ok_from}/{ok_to}）')
    _dump_export_range_debug(driver, dialog)
    return False


def _dump_record_content_debug(driver):
    """下拉失败时打印 overlay / wos-select 概况。"""
    try:
        info = driver.execute_script(
            """
            function box(el) {
              if (!el) return null;
              const r = el.getBoundingClientRect();
              return {
                tag: el.tagName, id: el.id || '',
                cls: String(el.className || '').slice(0, 80),
                text: (el.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 80),
                w: Math.round(r.width), h: Math.round(r.height)
              };
            }
            const overlay = document.querySelector('app-export-overlay');
            const details = document.querySelector('app-export-out-details');
            const selects = Array.from(document.querySelectorAll(
              'app-export-overlay wos-select, app-export-out-details wos-select'
            )).map(box);
            const btns = Array.from(document.querySelectorAll(
              'app-export-overlay wos-select button, app-export-out-details wos-select button,'
              + ' app-export-overlay button.dropdown'
            )).map(box);
            return {
              overlay: !!overlay,
              details: !!details,
              overlayText: overlay ? (overlay.innerText || '').slice(0, 180) : '',
              selects, btns
            };
            """)
        print('Record Content 调试:', info)
    except Exception as e:
        print('Record Content 调试失败:', e)


def _record_content_menu_open(driver):
    """Record Content 下拉是否已展开（可见 Custom selection / Full Record）。"""
    try:
        return bool(driver.execute_script(
            """
            function vis(n) {
              const r = n.getBoundingClientRect();
              const st = window.getComputedStyle(n);
              if (st.display === 'none' || st.visibility === 'hidden') return false;
              return r.width > 8 && r.height > 10;
            }
            const roots = Array.from(document.querySelectorAll(
              'app-export-overlay wos-select, app-export-out-details wos-select,'
              + ' app-export-overlay, .cdk-overlay-pane, [role="listbox"], [role="menu"]'));
            for (const root of roots) {
              const nodes = root.querySelectorAll('span, button, div, li, a');
              for (const n of nodes) {
                if (!vis(n)) continue;
                const t = (n.innerText || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                if (t.length > 60) continue;
                if (t.includes('custom selection') || t.includes('自定义选择项')) {
                  return true;
                }
              }
            }
            return false;
            """))
    except Exception:
        return False


def _open_record_content_dropdown(driver):
    """点击导出层 app-export-overlay 里的 Record Content（wos-select）下拉。"""
    _close_cookie_preference_center(driver)
    if _record_content_menu_open(driver):
        return True
    opened = driver.execute_script(
        """
        function vis(n) {
          const r = n.getBoundingClientRect();
          const st = window.getComputedStyle(n);
          if (st.display === 'none' || st.visibility === 'hidden') return false;
          return r.width > 20 && r.height > 12;
        }
        function noise(el) {
          return !!(el.closest && el.closest(
            '#onetrust-pc-sdk, #onetrust-banner-sdk, #onetrust-consent-sdk, app-export-menu'));
        }
        function clickBtn(btn, how) {
          if (!btn || !vis(btn) || noise(btn)) return null;
          btn.scrollIntoView({block: 'center'});
          btn.click();
          return how;
        }
        const overlay = document.querySelector('app-export-out-details')
          || document.querySelector('app-export-overlay');
        if (overlay && !noise(overlay)) {
          const btns = Array.from(overlay.querySelectorAll(
            'wos-select button.dropdown, wos-select button[role="combobox"], wos-select button'
          ));
          for (const b of btns) {
            const hit = clickBtn(b, 'overlay wos-select button');
            if (hit) return hit;
          }
          const span = overlay.querySelector('wos-select span.dropdown-text');
          if (span) {
            const host = span.closest('button, wos-select') || span;
            const hit = clickBtn(host, 'overlay dropdown-text');
            if (hit) return hit;
          }
        }
        const labels = Array.from(document.querySelectorAll('label, span, div, mat-label'));
        for (const n of labels) {
          if (noise(n)) continue;
          const t = (n.innerText || '').replace(/\\s+/g, ' ').trim();
          if (t !== 'Record Content:' && t !== 'Record Content' && t !== '记录内容:') continue;
          let scope = n.parentElement;
          for (let i = 0; i < 8 && scope; i++) {
            const btn = scope.querySelector(
              'wos-select button, button.dropdown, button[role="combobox"]');
            if (btn) {
              const hit = clickBtn(btn, 'after Record Content:');
              if (hit) return hit;
            }
            scope = scope.parentElement;
          }
        }
        return null;
        """)
    if opened:
        print(f'已打开记录内容下拉: {opened}')
        time.sleep(0.9)
        if _record_content_menu_open(driver):
            return True
    for xp in (
        '//app-export-overlay//wos-select//button',
        '//app-export-out-details//wos-select//button',
        '//app-export-overlay//span[contains(@class,"dropdown-text")]',
        '//*[normalize-space(text())="Record Content:"]/following::button[1]',
        '//*[normalize-space(text())="Record Content:"]/following::wos-select[1]//button',
        '//span[contains(@class,"dropdown-text") and contains(., "Author, Title, Source")]',
        '//button[contains(@class,"dropdown") and contains(., "Author, Title, Source")'
        ' and not(contains(., "Abstract"))]',
        '//*[@role="combobox" and contains(., "Author, Title, Source")]',
    ):
        for el in driver.find_elements(By.XPATH, xp):
            try:
                if el.is_displayed() and safe_click(driver, el):
                    print(f'已打开记录内容下拉: {xp}')
                    time.sleep(0.9)
                    if _record_content_menu_open(driver):
                        return True
            except Exception:
                continue
    print('未打开 Record Content 下拉')
    _dump_record_content_debug(driver)
    return False


def _click_custom_selection_edit(driver):
    """在已展开的 Record Content 下拉里点 Custom selection 右侧的 Edit。"""
    clicked = driver.execute_script(
        """
        function vis(n) {
          const r = n.getBoundingClientRect();
          const st = window.getComputedStyle(n);
          if (st.display === 'none' || st.visibility === 'hidden') return false;
          return r.width > 4 && r.height > 8;
        }
        function noise(el) {
          return !!(el.closest && el.closest(
            '#onetrust-pc-sdk, #onetrust-banner-sdk, #onetrust-consent-sdk'));
        }
        const roots = Array.from(document.querySelectorAll(
          'app-export-overlay wos-select, app-export-out-details wos-select,'
          + ' app-export-overlay, .cdk-overlay-pane, [role="listbox"], [role="menu"]'));
        for (const root of roots) {
          if (noise(root)) continue;
          const tRoot = (root.innerText || '').toLowerCase();
          if (!tRoot.includes('custom selection') && !tRoot.includes('自定义选择项')) continue;
          const nodes = Array.from(root.querySelectorAll('button, a, span, div'));
          for (const b of nodes) {
            if (!vis(b) || noise(b)) continue;
            const t = (b.innerText || '').replace(/\\s+/g, ' ').trim();
            if (t !== 'Edit' && t !== '编辑') continue;
            const row = b.closest(
              'li, [role="option"], [role="menuitem"], .mat-mdc-option, div') || b.parentElement;
            const rt = ((row && row.innerText) || '').toLowerCase();
            if (rt.includes('custom') || rt.includes('自定义') || t === 'Edit' || t === '编辑') {
              b.scrollIntoView({block: 'center'});
              b.click();
              return t;
            }
          }
        }
        return null;
        """)
    if clicked:
        print(f'已点击自定义编辑: {clicked}')
        return True
    for xp in (
        '//app-export-overlay//*[contains(text(),"Custom selection")]'
        '/following::*[normalize-space()="Edit"][1]',
        '//app-export-out-details//*[contains(text(),"Custom selection")]'
        '/following::*[normalize-space()="Edit"][1]',
        '//*[contains(text(), "Custom selection")]/following::*[normalize-space()="Edit"][1]',
        '//*[contains(normalize-space(.), "Custom selection")]//button[normalize-space()="Edit"]',
        '//*[contains(normalize-space(.), "Custom selection")]//span[normalize-space()="Edit"]',
        '//*[contains(normalize-space(.), "自定义选择项")]//button[normalize-space()="编辑"]',
        '//button[normalize-space()="Edit"]',
        '//span[normalize-space()="Edit"]',
        '//button[normalize-space()="编辑"]',
        '//span[normalize-space()="编辑"]',
    ):
        for el in driver.find_elements(By.XPATH, xp):
            try:
                if not el.is_displayed():
                    continue
                row_text = (el.text or '')
                try:
                    parent = el.find_element(By.XPATH, './ancestor::*[self::li or self::div][1]')
                    row_text = parent.text or row_text
                except Exception:
                    pass
                low = row_text.lower()
                if 'custom' not in low and '自定义' not in low and row_text.strip() not in (
                        'Edit', '编辑'):
                    continue
                if safe_click(driver, el):
                    print(f'已点击自定义编辑: {xp}')
                    return True
            except Exception:
                continue
    print('未点到 Custom selection 的 Edit')
    return False


def _custom_fields_dialog_open(driver):
    try:
        return bool(driver.execute_script(
            """
            const t = (document.body.innerText || '').toLowerCase();
            return t.includes('save selections')
              || t.includes('保存选择项')
              || t.includes('my custom export')
              || t.includes('我的自定义导出选择项');
            """))
    except Exception:
        return False


def select_custom_all_fields(driver):
    """Record Content 下拉 → Custom selection 的 Edit → 勾选顶部四组 → Save selections。"""
    opened_edit = False
    for attempt in range(1, 4):
        if not _open_record_content_dropdown(driver):
            print(f'警告: Record Content 下拉未打开（第 {attempt} 次）')
        time.sleep(0.35)
        if _click_custom_selection_edit(driver):
            opened_edit = True
            break
        print(f'未点到 Edit，重试打开下拉（{attempt}/3）')
        time.sleep(0.4)
    if not opened_edit:
        return False
    wait_deadline = time.time() + 8
    while time.time() < wait_deadline and not _custom_fields_dialog_open(driver):
        time.sleep(0.25)
    time.sleep(0.4)

    header_n = driver.execute_script(
        """
        function noise(el) {
          return !!(el.closest && el.closest(
            '#onetrust-pc-sdk, #onetrust-banner-sdk, #onetrust-consent-sdk'));
        }
        const headers = [
          'author, title, source',
          'abstract, keyword, addresses',
          'cited references and use',
          'funding and other',
          '作者、标题、来源出版物',
          '摘要、关键词、地址',
          '引用的参考文献和使用',
          '基金资助和其他'
        ];
        function ownLabel(el) {
          const lab = el.querySelector(
            '.mdc-label, .mat-checkbox-label, .mdc-form-field > label, label');
          if (lab) {
            const clone = lab.cloneNode(true);
            clone.querySelectorAll('mat-checkbox, input').forEach(n => n.remove());
            return (clone.innerText || '').replace(/\\s+/g, ' ').trim().toLowerCase();
          }
          let t = '';
          for (const n of el.childNodes) {
            if (n.nodeType === 3) t += n.textContent;
          }
          return t.replace(/\\s+/g, ' ').trim().toLowerCase();
        }
        const panes = Array.from(document.querySelectorAll(
          'mat-dialog-container, [role="dialog"], .cdk-overlay-pane, .mdc-dialog'));
        let root = null;
        for (const p of panes) {
          if (noise(p)) continue;
          const t = (p.innerText || '').toLowerCase();
          if (t.includes('my custom export') || t.includes('save selections')
              || t.includes('我的自定义导出')) { root = p; break; }
        }
        if (!root) return 0;
        let n = 0;
        root.querySelectorAll('mat-checkbox').forEach(el => {
          const label = ownLabel(el);
          const all = (el.innerText || '').replace(/\\s+/g, ' ').trim().toLowerCase();
          const first = (label || all).split('\\n')[0];
          if (!headers.some(h => first === h || first.startsWith(h) || label === h)) return;
          const input = el.querySelector('input[type="checkbox"]');
          const state = el.getAttribute('aria-checked')
            || (input && input.checked ? 'true' : 'false');
          const checked = state === 'true'
            || el.classList.contains('mat-mdc-checkbox-checked');
          if (!checked || state === 'mixed') {
            (input || el).click();
            n += 1;
          }
        });
        return n;
        """)
    print(f'自定义顶部分组：新勾选 {header_n} 项')
    time.sleep(0.3)

    n_checked = driver.execute_script(
        """
        function noise(el) {
          return !!(el.closest && el.closest(
            '#onetrust-pc-sdk, #onetrust-banner-sdk, #onetrust-consent-sdk'));
        }
        const panes = Array.from(document.querySelectorAll(
          'mat-dialog-container, .cdk-overlay-pane, [role="dialog"], .mdc-dialog'));
        let root = null;
        for (const p of panes) {
          if (noise(p)) continue;
          const t = (p.innerText || '').toLowerCase();
          if (t.includes('cookie preference')) continue;
          if (t.includes('save selections') || t.includes('保存选择项')
              || t.includes('my custom export') || t.includes('我的自定义导出选择项')) {
            root = p; break;
          }
        }
        const hosts = root ? [root] : panes.filter(p => !noise(p));
        let n = 0;
        for (const pane of hosts) {
          pane.querySelectorAll('mat-checkbox, input[type="checkbox"]').forEach(el => {
            const input = el.tagName === 'INPUT' ? el
              : el.querySelector('input[type="checkbox"]');
            const checked = input ? input.checked
              : (el.getAttribute('aria-checked') === 'true'
                 || el.classList.contains('mat-mdc-checkbox-checked'));
            if (!checked) {
              el.click();
              n += 1;
            }
          });
        }
        return n;
        """)
    print(f'自定义字段：新勾选 {n_checked} 项（已勾选的保持）')
    time.sleep(0.4)

    saved = driver.execute_script(
        """
        function vis(n) {
          const r = n.getBoundingClientRect();
          return r.width > 4 && r.height > 8;
        }
        function noise(el) {
          return !!(el.closest && el.closest(
            '#onetrust-pc-sdk, #onetrust-banner-sdk, #onetrust-consent-sdk'));
        }
        const panes = Array.from(document.querySelectorAll(
          'mat-dialog-container, [role="dialog"], .cdk-overlay-pane, .mdc-dialog'));
        const roots = panes.filter(p => {
          if (noise(p)) return false;
          const t = (p.innerText || '').toLowerCase();
          return t.includes('save selections') || t.includes('my custom export')
            || t.includes('保存选择项') || t.includes('我的自定义导出');
        });
        const search = roots.length ? roots : [document.body];
        for (const root of search) {
          const btns = Array.from(root.querySelectorAll('button'));
          for (const b of btns) {
            if (!vis(b) || noise(b)) continue;
            const t = (b.innerText || '').replace(/\\s+/g, ' ').trim();
            if (t === 'Save Selections' || t === 'Save selections'
                || t === '保存选择项') {
              b.scrollIntoView({block: 'center'});
              b.click();
              return t;
            }
          }
        }
        return null;
        """)
    if saved:
        print(f'已保存自定义选择项: {saved}')
        # 等自定义弹窗关掉，回到 Export Records to Excel
        wait_deadline = time.time() + 8
        while time.time() < wait_deadline:
            if not _custom_fields_dialog_open(driver):
                break
            time.sleep(0.25)
        time.sleep(0.5)
        return True
    print('警告: 未点到 Save selections')
    return False


def _wait_export_overlay_ready(driver, timeout=8):
    """Save selections 后，等 Excel 导出层重新可点（含底部 Export）。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _custom_fields_dialog_open(driver):
            time.sleep(0.25)
            continue
        if not _export_excel_dialog_open(driver):
            time.sleep(0.25)
            continue
        ready = driver.execute_script(
            """
            function vis(n) {
              if (!n) return false;
              const r = n.getBoundingClientRect();
              const st = window.getComputedStyle(n);
              if (st.display === 'none' || st.visibility === 'hidden') return false;
              return r.width > 2 && r.height > 2;
            }
            function labelOf(b) {
              return ((b.innerText || b.textContent || '')
                + ' ' + (b.getAttribute('aria-label') || '')
                + ' ' + (b.getAttribute('data-ta') || ''))
                .replace(/\\s+/g, ' ').trim().toLowerCase();
            }
            const roots = Array.from(document.querySelectorAll(
              'app-export-overlay, app-export-out-details'));
            const search = roots.length ? roots : [document.body];
            for (const root of search) {
              for (const b of root.querySelectorAll('button, a[role="button"]')) {
                if (!vis(b)) continue;
                if (b.closest('app-export-menu')) continue;
                const t = labelOf(b);
                if (t === 'export' || t === '导出' || t.startsWith('export ')
                    || t.includes('export records')) {
                  return true;
                }
              }
            }
            return false;
            """)
        if ready:
            return True
        time.sleep(0.3)
    return _export_excel_dialog_open(driver)


def click_export_confirm(driver):
    """点 Export Records to Excel 弹窗底部、与 Cancel 并排的 Export。"""
    _wait_export_overlay_ready(driver, timeout=8)
    deadline = time.time() + 10
    last_debug = ''
    while time.time() < deadline:
        clicked = driver.execute_script(
            """
            function vis(n) {
              if (!n) return false;
              const r = n.getBoundingClientRect();
              const st = window.getComputedStyle(n);
              if (st.display === 'none' || st.visibility === 'hidden'
                  || Number(st.opacity) === 0) return false;
              return r.width > 2 && r.height > 2;
            }
            function noise(el) {
              return !!(el.closest && el.closest(
                '#onetrust-pc-sdk, #onetrust-consent-sdk, app-export-menu,'
                + ' #onetrust-banner-sdk'));
            }
            function labelOf(b) {
              const raw = ((b.innerText || b.textContent || '')
                + ' ' + (b.getAttribute('aria-label') || '')
                + ' ' + (b.getAttribute('data-ta') || ''));
              return raw.replace(/\\s+/g, ' ').trim();
            }
            function isExportBtn(b) {
              if (!vis(b) || noise(b) || b.disabled) return false;
              const t = labelOf(b).toLowerCase();
              if (!t) return false;
              if (t === 'export' || t === '导出') return true;
              if (t === 'export records' || t === '导出记录') return true;
              // 勿点工具栏/菜单里的 Export
              if (t.includes('excel') || t.includes('cancel') || t.includes('取消')) {
                return false;
              }
              return false;
            }
            function clickBtn(b) {
              try { b.scrollIntoView({block: 'center'}); } catch (e) {}
              try { b.click(); } catch (e) {
                try {
                  b.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true}));
                } catch (e2) { return null; }
              }
              return labelOf(b) || 'Export';
            }
            const overlayRoots = Array.from(document.querySelectorAll(
              'app-export-overlay, app-export-out-details'));
            const roots = overlayRoots.length ? overlayRoots : [];
            if (!roots.length) {
              const panes = Array.from(document.querySelectorAll(
                'mat-dialog-container, [role="dialog"], .cdk-overlay-pane'));
              for (const p of panes) {
                if (noise(p) || !vis(p)) continue;
                const t = (p.innerText || '').toLowerCase();
                if (t.includes('export records to excel')
                    || t.includes('将记录导出至 excel')
                    || (t.includes('record content') && t.includes('cancel'))) {
                  roots.push(p);
                }
              }
            }
            if (!roots.length) roots.push(document.body);

            for (const root of roots) {
              const btns = Array.from(root.querySelectorAll(
                'button, a[role="button"], [mat-flat-button], [mat-button]'));
              // 1) 从 Cancel 向上找同级 Export（footer 可能隔一层）
              const cancel = btns.find(b => {
                if (!vis(b) || noise(b)) return false;
                const t = labelOf(b).toLowerCase();
                return t === 'cancel' || t === '取消';
              });
              if (cancel) {
                let node = cancel.parentElement;
                for (let i = 0; i < 5 && node; i++) {
                  const candidates = Array.from(
                    node.querySelectorAll('button, a[role="button"]'));
                  const exp = candidates.find(isExportBtn);
                  if (exp) return clickBtn(exp);
                  node = node.parentElement;
                }
              }
              // 2) 弹窗内任意可见 Export（优先 primary / 底部）
              const exports = btns.filter(isExportBtn);
              if (exports.length) {
                exports.sort((a, b) =>
                  b.getBoundingClientRect().top - a.getBoundingClientRect().top);
                const primary = exports.find(b => {
                  const cls = String(b.className || '');
                  return cls.includes('primary') || cls.includes('unelevated')
                    || cls.includes('flat');
                });
                return clickBtn(primary || exports[0]);
              }
            }
            // 调试：列出弹窗底部按钮文案
            const debug = [];
            for (const root of roots) {
              for (const b of root.querySelectorAll('button')) {
                if (!vis(b) || noise(b)) continue;
                debug.push(labelOf(b).slice(0, 40));
              }
            }
            return debug.length ? ('__miss__:' + debug.slice(-8).join(' | ')) : null;
            """)
        if clicked and not str(clicked).startswith('__miss__'):
            print(f'已确认导出: {clicked}')
            time.sleep(1.5)
            return True
        if clicked and str(clicked).startswith('__miss__'):
            last_debug = str(clicked)[8:]
        time.sleep(0.4)
    if last_debug:
        print(f'未找到导出确认按钮，弹窗可见按钮: {last_debug}')
    else:
        print('未找到导出确认按钮')
    return False


def export_batch(driver, save_dir, start_n, end_n, label_prefix, expected_hits=None):
    """导出一批记录。成功返回文件路径，失败返回 None。

    expected_hits: 若给定，打开导出前再次核对页面命中数，不一致则中止本批。
    """
    driver = ctx_driver(driver)
    if is_page_unusable(driver):
        print('导出前页面已失效')
        return None

    if expected_hits is not None:
        actual = parse_main_hit_count(driver, timeout=20, quiet=True)
        if actual is None or int(actual) != int(expected_hits):
            print(
                f'打开导出前数量不符: 页面={actual} 预期={expected_hits}，中止本批')
            return None
        print(f'打开导出前数量确认: {actual} == {expected_hits}')

    before = snapshot_download_meta(save_dir)
    # 仅在确实有导出层时关闭，避免误清 Refine
    if _export_excel_dialog_open(driver) or _export_dialog_el(driver):
        close_export_overlay(driver)
    _close_pendo_guides(driver)
    _close_cookie_preference_center(driver)
    time.sleep(0.5)

    if not open_excel_export(driver):
        return None

    # 打开导出菜单/弹窗后再次确认精炼未丢（Esc/误点可能导致回退全体）
    if expected_hits is not None:
        # 弹窗打开时主结果数通常仍可读
        actual2 = parse_main_hit_count(driver, timeout=6, quiet=True)
        if actual2 is not None and int(actual2) != int(expected_hits):
            print(
                f'导出弹窗打开后数量不符: 页面={actual2} 预期={expected_hits}，关闭并中止')
            close_export_overlay(driver)
            return None

    # 等「将记录导出至 Excel」弹窗渲染完再填范围
    wait_deadline = time.time() + 8
    while time.time() < wait_deadline and not _export_excel_dialog_open(driver):
        time.sleep(0.25)
    time.sleep(0.4)
    if not set_export_record_range(driver, start_n, end_n):
        close_export_overlay(driver)
        return None
    custom_ok = False
    for attempt in range(3):
        if select_custom_all_fields(driver):
            custom_ok = True
            break
        print(f'自定义字段未完成，重试 ({attempt + 1}/3)')
        time.sleep(0.7)
    if not custom_ok:
        print('未能完成 Edit / 全选顶部四组 / Save selections，本批不导出默认字段')
        close_export_overlay(driver)
        return None
    if not click_export_confirm(driver):
        # 再等一次：自定义层有时关掉较慢，Export 稍后才可点
        time.sleep(1.0)
        if not click_export_confirm(driver):
            close_export_overlay(driver)
            return None

    new_file = get_new_download_file(
        save_dir, before, timeout=EXPORT_DOWNLOAD_TIMEOUT)
    if not new_file:
        print(f'批次 {start_n}-{end_n} 下载超时，再检查是否已有新文件')
        new_file = pick_new_download(save_dir, before)
    if not new_file:
        print(f'批次 {start_n}-{end_n} 未检测到下载文件')
        close_export_overlay(driver)
        return None

    dest = rename_export_file(new_file, save_dir, label_prefix, start_n, end_n)
    close_export_overlay(driver)
    time.sleep(random.uniform(1.5, 3.0))
    return dest


def recover_for_export(driver, restore_fn=None):
    """导出失败后尽量回到结果页：关 overlay；若已是错误页则重新检索。"""
    driver = ctx_driver(driver)
    try:
        # 无导出层时不要点 Close / 发 Esc（会清掉 Research Areas 精炼）
        if _export_excel_dialog_open(driver) or _export_dialog_el(driver):
            close_export_overlay(driver)
        _close_pendo_guides(driver)
        _close_cookie_preference_center(driver)
    except Exception:
        pass
    driver = ctx_driver(driver)
    if is_wos_results_page(driver):
        return True
    print('导出过程中结果页丢失，尝试恢复...')
    if restore_fn:
        try:
            ok = bool(restore_fn())
            return ok and is_driver_alive(ctx_driver(driver))
        except Exception as e:
            print('重新检索失败:', e)
            return False
    if is_page_unusable(driver):
        return recover_wos_page(driver)
    return False


def _verify_hits_before_export(
        driver, expected_hits, ensure_hits_fn=None, retries=2):
    """导出前校验页面命中数；不一致则调用 ensure_hits_fn 重新精炼。"""
    if expected_hits is None:
        return True
    expected = int(expected_hits)
    for attempt in range(1, retries + 1):
        driver = ctx_driver(driver)
        actual = parse_main_hit_count(driver, timeout=20, quiet=True)
        if actual is not None and int(actual) == expected:
            print(f'导出前数量校验通过: {actual} == {expected}')
            return True
        print(
            f'导出前数量不一致 ({attempt}/{retries}): '
            f'页面={actual} 预期={expected}，重新精炼')
        if not ensure_hits_fn:
            return False
        try:
            if not ensure_hits_fn():
                print('重新精炼失败')
                continue
        except Exception as e:
            print(f'重新精炼异常: {e}')
            continue
        # 精炼函数内部已用 wait_hit_count_equals 校验过；此处用新选择器再读一次
        time.sleep(0.6)
        driver = ctx_driver(driver)
        actual = parse_main_hit_count(driver, timeout=25, quiet=True)
        if actual is not None and int(actual) == expected:
            print(f'导出前数量校验通过: {actual} == {expected}')
            return True
        time.sleep(0.5)
    driver = ctx_driver(driver)
    actual = parse_main_hit_count(driver, timeout=20, quiet=True)
    ok = actual is not None and int(actual) == expected
    if ok:
        print(f'导出前数量校验通过: {actual} == {expected}')
    else:
        print(f'导出前数量仍不一致: 页面={actual} 预期={expected}')
    return ok


def export_all_hits(
        driver, save_dir, total, label_prefix, restore_fn=None,
        expected_hits=None, ensure_hits_fn=None):
    """按 EXPORT_BATCH_SIZE 分批导出；失败只减半一次到 EXPORT_MIN_SPLIT，再失败则记录。

    expected_hits / ensure_hits_fn: 每批导出前校验命中数，不一致则重新精炼。
    """
    if total <= 0:
        return 0
    # (start, end, already_split)
    pending = []
    start_n = 1
    while start_n <= total:
        end_n = min(start_n + EXPORT_BATCH_SIZE - 1, total)
        pending.append((start_n, end_n, False))
        start_n = end_n + 1

    success = 0
    seen = set()
    while pending:
        start_n, end_n, already_split = pending.pop(0)
        key = (start_n, end_n)
        if key in seen:
            print(f'批次 {start_n}-{end_n} 已处理过，跳过以免循环')
            continue
        seen.add(key)
        count = end_n - start_n + 1
        print(f'\n----- 导出批次 {start_n}-{end_n} / {total} -----')

        existed = find_existing_batch_file(save_dir, label_prefix, start_n, end_n)
        if existed:
            print(f'批次文件已存在（合集目录内），跳过: {existed}')
            success += 1
            continue

        driver = ctx_driver(driver)
        if not recover_for_export(driver, restore_fn):
            print(f'结果页无法恢复，记录失败 {start_n}-{end_n}')
            record_export_failure(save_dir, f'{label_prefix}_{start_n}-{end_n}')
            continue

        driver = ctx_driver(driver)
        if not _verify_hits_before_export(
                driver, expected_hits,
                ensure_hits_fn=ensure_hits_fn or restore_fn):
            print(f'导出前校验失败，记录失败 {start_n}-{end_n}')
            record_export_failure(save_dir, f'{label_prefix}_{start_n}-{end_n}')
            continue

        driver = ctx_driver(driver)
        before = snapshot_download_meta(save_dir)
        path = None
        try:
            path = export_batch(
                driver, save_dir, start_n, end_n, label_prefix,
                expected_hits=expected_hits)
        except Exception as e:
            print(f'批次异常 {start_n}-{end_n}:', e)
            try:
                close_export_overlay(ctx_driver(driver))
            except Exception:
                pass

        if not path:
            # 导出过程超时，但默认名文件可能已落地
            orphan = pick_new_download(save_dir, before)
            if orphan:
                path = rename_export_file(
                    orphan, save_dir, label_prefix, start_n, end_n)

        if path:
            success += 1
            continue

        print(f'批次失败: {start_n}-{end_n}，恢复后重试一次')
        driver = ctx_driver(driver)
        recover_for_export(driver, restore_fn)
        time.sleep(2)
        driver = ctx_driver(driver)
        if not _verify_hits_before_export(
                driver, expected_hits,
                ensure_hits_fn=ensure_hits_fn or restore_fn):
            print(f'重试前校验失败，记录失败 {start_n}-{end_n}')
            record_export_failure(save_dir, f'{label_prefix}_{start_n}-{end_n}')
            continue
        driver = ctx_driver(driver)
        before_retry = snapshot_download_meta(save_dir)
        try:
            path = export_batch(
                driver, save_dir, start_n, end_n, label_prefix,
                expected_hits=expected_hits)
        except Exception as e:
            print(f'批次重试异常 {start_n}-{end_n}:', e)
            path = None
            try:
                close_export_overlay(ctx_driver(driver))
            except Exception:
                pass

        if not path:
            orphan = pick_new_download(save_dir, before_retry) or pick_new_download(
                save_dir, before)
            if orphan:
                path = rename_export_file(
                    orphan, save_dir, label_prefix, start_n, end_n)

        if path:
            success += 1
            continue

        existed = find_existing_batch_file(save_dir, label_prefix, start_n, end_n)
        if existed:
            print(f'重试后发现已有批次文件，视为成功: {existed}')
            success += 1
            continue

        can_split = (
            (not already_split)
            and count > EXPORT_MIN_SPLIT
        )
        if can_split:
            mid = start_n + count // 2 - 1
            left = (start_n, mid, True)
            right = (mid + 1, end_n, True)
            print(
                f'批次 {start_n}-{end_n} 仍失败/超时，减半一次为 '
                f'{left[0]}-{left[1]} 与 {right[0]}-{right[1]}'
            )
            pending.insert(0, right)
            pending.insert(0, left)
            continue

        print(f'批次 {start_n}-{end_n} 失败，已达减半下限，记录并跳过')
        record_export_failure(
            save_dir, f'{label_prefix}_{start_n}-{end_n}')
    return success


# ---------------------------------------------------------------------------
# 单段检索 + 导出（含超限切分）
# ---------------------------------------------------------------------------
def prepare_search_form(driver, database, collections):
    """打开对应检索页（核心合集 woscc / 其余 alldb 勾选合集），再选 Publication Date。"""
    target = search_url_for(database, collections)
    last_err = None
    for attempt in range(1, PAGE_LOAD_RETRIES + 1):
        if is_page_unusable(driver) or not is_on_target_search_page(driver, target) \
                or not is_search_form_ready(driver):
            print(f'检索表单未就绪，重新打开 {target} ({attempt}/{PAGE_LOAD_RETRIES})')
            if not goto_search_url(driver, target):
                last_err = RuntimeError('检索页恢复失败')
                recover_wos_page(driver, target_url=target)
                time.sleep(1.5)
                continue
        if not ensure_database_and_collections(driver, database, collections):
            last_err = RuntimeError('未能打开目标检索库')
            recover_wos_page(driver, target_url=target)
            continue
        if is_page_unusable(driver):
            last_err = RuntimeError('打开检索库后面页失效')
            continue
        _dismiss_open_overlays(driver)
        time.sleep(0.5)
        if select_publication_date_field(driver):
            return True
        last_err = RuntimeError('无法选择 Publication Date 筛选项，日期输入框未出现')
        print(f'未能选择 Publication Date，重开检索页 ({attempt}/{PAGE_LOAD_RETRIES})')
        recover_wos_page(driver, target_url=target)
        time.sleep(1)
    raise last_err or RuntimeError('无法选择 Publication Date 筛选项，日期输入框未出现')


def run_search_once(
        driver, start_date, end_date, database, collections, skip_sort=False,
        download_dir=None):
    """执行一次检索，返回命中数（int）；失败返回 None。页面失效会恢复重试。

    skip_sort: Research Areas 模式下为 True，避免排序冲掉后续精炼条件。
    点击 Search 后最长等待 SEARCH_RESULT_WAIT（默认 10 分钟）；
    超时则关闭浏览器重开再试，最多 SEARCH_TIMEOUT_RETRIES 次。
    若出现 Server.unexpectedError：关闭浏览器 → 等待 5 分钟 → 重试，
    最多 SERVER_UNEXPECTED_ERROR_RETRIES 次。
    """
    if download_dir:
        bind_browser_ctx(driver, download_dir, database, collections)
    else:
        bind_browser_ctx(driver, database=database, collections=collections)

    target = search_url_for(database, collections)
    last_err = None
    server_round = 0
    timeout_round = 0

    while True:
        driver = ctx_driver(driver)
        if not is_driver_alive(driver):
            print('检索前浏览器不可用')
            last_err = 'driver_dead'
            break

        hit_server_error = False
        hit_search_timeout = False
        form_ok = False

        for attempt in range(1, PAGE_LOAD_RETRIES + 1):
            driver = ctx_driver(driver)
            if not goto_search_url(driver, target):
                print(f'无法打开检索页 / 登录失败 ({attempt}/{PAGE_LOAD_RETRIES})')
                recover_wos_page(driver, target_url=target)
                last_err = 'open_search_page'
                continue
            try:
                prepare_search_form(driver, database, collections)
                set_publication_date_range(driver, start_date, end_date)
                click_run_search(driver)
                accept_cookies(driver, timeout=3)
                form_ok = True

                status = wait_search_result_or_server_error(
                    driver, timeout=SEARCH_RESULT_WAIT)
                if status == 'no_results':
                    print('检索无结果，跳过本段')
                    bind_browser_ctx(driver)
                    return 0
                if status == 'server_error':
                    hit_server_error = True
                    last_err = 'server_unexpected_error'
                    break
                if status == 'timeout':
                    if has_wos_server_unexpected_error(driver):
                        hit_server_error = True
                        last_err = 'server_unexpected_error'
                        break
                    hit_search_timeout = True
                    last_err = 'no_results_page_timeout'
                    print(
                        f'检索后 {SEARCH_RESULT_WAIT // 60} 分钟仍未进入结果页，'
                        f'将关闭浏览器重试')
                    break
                if status == 'unusable':
                    print('检索后页面失效，恢复后重试')
                    recover_wos_page(driver, target_url=target)
                    last_err = 'page_after_search'
                    continue

                if is_page_unusable(driver):
                    print('检索后页面失效，恢复后重试')
                    recover_wos_page(driver, target_url=target)
                    last_err = 'page_after_search'
                    continue

                hits = parse_hit_count(driver, timeout=HIT_COUNT_TIMEOUT)
                if hits is None:
                    if has_wos_server_unexpected_error(driver):
                        hit_server_error = True
                        last_err = 'server_unexpected_error'
                        break
                    if is_page_unusable(driver):
                        print('解析命中数时页面失效，恢复后重试')
                        recover_wos_page(driver, target_url=target)
                        last_err = 'hit_count_page'
                        continue
                    print(f'解析命中条数失败 ({attempt}/{PAGE_LOAD_RETRIES})')
                    last_err = 'hit_count'
                    time.sleep(2)
                    continue
                if hits > 0 and not skip_sort:
                    set_sort_date_ascending(driver)
                elif hits > 0 and skip_sort:
                    print('RA 模式：跳过排序（避免冲掉 Research Areas 精炼）')
                bind_browser_ctx(driver)
                return hits
            except Exception as e:
                last_err = e
                print(f'检索异常 ({attempt}/{PAGE_LOAD_RETRIES}):', e)
                if has_wos_server_unexpected_error(driver):
                    hit_server_error = True
                    last_err = 'server_unexpected_error'
                    break
                if is_page_unusable(driver):
                    recover_wos_page(driver, target_url=target)
                time.sleep(2)

            if hit_server_error or hit_search_timeout:
                break

        if hit_server_error:
            server_round += 1
            if server_round > SERVER_UNEXPECTED_ERROR_RETRIES:
                print(
                    f'Server.unexpectedError 已重试 '
                    f'{SERVER_UNEXPECTED_ERROR_RETRIES} 次仍失败，标记本段失败')
                break
            print(
                f'Server.unexpectedError：第 {server_round}/'
                f'{SERVER_UNEXPECTED_ERROR_RETRIES} 次冷却重试')
            driver = restart_browser_after_server_error(
                driver,
                reason=(
                    f'检索触发 Server.unexpectedError'
                    f'（{server_round}/{SERVER_UNEXPECTED_ERROR_RETRIES}）'
                ))
            if not is_driver_alive(driver) or not is_wos_logged_in(driver):
                print('冷却后浏览器未能就绪，终止本检索重试')
                last_err = 'restart_after_server_error'
                break
            continue

        if hit_search_timeout:
            timeout_round += 1
            if timeout_round > SEARCH_TIMEOUT_RETRIES:
                print(
                    f'检索超时已重开浏览器 {SEARCH_TIMEOUT_RETRIES} 次仍失败，'
                    f'标记本段失败（避免死循环）')
                break
            print(
                f'检索超时：第 {timeout_round}/{SEARCH_TIMEOUT_RETRIES} 次'
                f'关闭浏览器并重试')
            driver = restart_browser_after_search_timeout(
                driver,
                reason=(
                    f'点击 Search 后超时未出结果页'
                    f'（{timeout_round}/{SEARCH_TIMEOUT_RETRIES}）'
                ))
            if not is_driver_alive(driver) or not is_wos_logged_in(driver):
                print('超时重开后浏览器未能就绪，终止本检索重试')
                last_err = 'restart_after_search_timeout'
                break
            continue

        # 表单/解析类失败：不再无限循环
        if not form_ok:
            print('检索表单准备多次失败，终止本检索')
        break

    print('检索最终失败:', last_err)
    return None


# ---------------------------------------------------------------------------
# Research Areas 精炼导出
# ---------------------------------------------------------------------------
def _research_areas_see_all_panel(driver):
    """定位可见的 Research Areas See all 弹窗（含 input[name=SJ]）。"""
    panels = driver.find_elements(By.CSS_SELECTOR, '.shared-refine-see-all')
    for panel in panels:
        try:
            ok = driver.execute_script(
                r'''
                const el = arguments[0];
                if (!el) return false;
                const r = el.getBoundingClientRect();
                if (r.width < 80 || r.height < 80) return false;
                const s = window.getComputedStyle(el);
                if (s.display === 'none' || s.visibility === 'hidden') return false;
                if (el.querySelector('input[name="SJ"]')) return true;
                const t = (el.textContent || '');
                return /Refine by Research Areas|Research Areas|研究方向/i.test(t);
                ''', panel)
            if ok:
                return panel
        except Exception:
            continue
    return None


def _see_all_sj_count(driver):
    try:
        return int(driver.execute_script(
            r'''
            const root = document.querySelector('.shared-refine-see-all');
            if (!root) return 0;
            const r = root.getBoundingClientRect();
            if (r.width < 80 || r.height < 80) return 0;
            return root.querySelectorAll('input[name="SJ"]').length;
            ''') or 0)
    except Exception:
        return 0


def expand_research_areas_section(driver, timeout=20):
    """展开侧栏 Research Areas，使 See all 可见。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        for btn in driver.find_elements(By.CSS_SELECTOR, 'button[data-ta="see-all"]'):
            try:
                aria = (
                    (btn.get_attribute('aria-label') or '')
                    + ' ' + (btn.text or ''))
                if ('Research Areas' in aria or '研究方向' in aria) and btn.is_displayed():
                    return True
            except Exception:
                continue

        candidates = []
        candidates.extend(driver.find_elements(
            By.CSS_SELECTOR,
            '[data-ta="filter-section-SJ"] button.filter-heading, '
            'button[aria-controls="filter-section-SJ"], '
            '.filter-section-border-5 button.filter-heading'))
        for btn in driver.find_elements(
                By.CSS_SELECTOR, 'button.filter-heading, button.legend-button'):
            try:
                t = (btn.text or '') + ' ' + (btn.get_attribute('aria-label') or '')
                if 'Research Areas' in t or '研究方向' in t:
                    candidates.append(btn)
            except Exception:
                continue
        for btn in candidates:
            try:
                if not btn.is_displayed():
                    continue
                expanded = (btn.get_attribute('aria-expanded') or '').lower()
                if expanded == 'true':
                    time.sleep(0.4)
                    break
                safe_click(driver, btn)
                time.sleep(0.7)
                break
            except Exception:
                continue
        time.sleep(0.4)
    return False


def open_research_areas_see_all(driver, timeout=30):
    """展开 Research Areas 并打开 See all 弹窗，等到出现 SJ 选项。"""
    if not expand_research_areas_section(driver, timeout=min(12, timeout)):
        print('未能展开 Research Areas 筛选区')
        return False
    deadline = time.time() + timeout
    clicked_once = False
    while time.time() < deadline:
        n = _see_all_sj_count(driver)
        if n > 0:
            print(f'已打开 Research Areas See all 弹窗（{n} 个选项）')
            return True
        if _research_areas_see_all_panel(driver) and not clicked_once:
            time.sleep(0.5)
            continue

        clicked = False
        for btn in driver.find_elements(By.CSS_SELECTOR, 'button[data-ta="see-all"]'):
            try:
                aria = (
                    (btn.get_attribute('aria-label') or '')
                    + ' ' + (btn.text or ''))
                if 'Research Areas' not in aria and '研究方向' not in aria:
                    continue
                if not btn.is_displayed():
                    continue
                print('点击 See all Research Areas')
                safe_click(driver, btn)
                clicked = True
                clicked_once = True
                time.sleep(1.0)
                break
            except Exception:
                continue
        if not clicked:
            for btn in driver.find_elements(By.CSS_SELECTOR, 'button[data-ta="see-all"]'):
                try:
                    if not btn.is_displayed():
                        continue
                    near = driver.execute_script(
                        r'''
                        const b = arguments[0];
                        const sec = document.querySelector(
                          '[data-ta="filter-section-SJ"], #filter-section-SJ');
                        if (!sec || !b) return false;
                        const br = b.getBoundingClientRect();
                        const sr = sec.getBoundingClientRect();
                        return Math.abs(br.top - sr.bottom) < 120
                          || (br.top >= sr.top - 40 && br.top <= sr.bottom + 80);
                        ''', btn)
                    if near:
                        print('点击 See all（靠近 Research Areas 区块）')
                        safe_click(driver, btn)
                        clicked = True
                        clicked_once = True
                        time.sleep(1.0)
                        break
                except Exception:
                    continue
        time.sleep(0.4)
    print('未能打开 Research Areas See all 弹窗')
    return False


def close_research_areas_see_all(driver, timeout=8):
    """关闭 See all（Cancel / Back）。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _see_all_sj_count(driver) == 0 and not _research_areas_see_all_panel(driver):
            return True
        panel = _research_areas_see_all_panel(driver)
        if not panel:
            return True
        for sel in (
                'button[data-ta="cancel-submit"]',
                'button#refineBackToAll',
                'button[data-ta="Back to all filters"]'):
            for btn in panel.find_elements(By.CSS_SELECTOR, sel):
                try:
                    if btn.is_displayed():
                        safe_click(driver, btn)
                        time.sleep(0.5)
                        break
                except Exception:
                    continue
            else:
                continue
            break
        try:
            driver.execute_script(
                r'''
                const root = document.querySelector('.shared-refine-see-all');
                if (!root) return;
                const b = root.querySelector(
                  'button[data-ta="cancel-submit"], button#refineBackToAll');
                if (b) b.click();
                ''')
        except Exception:
            pass
        time.sleep(0.3)
    return _see_all_sj_count(driver) == 0


def list_research_areas_from_see_all(driver):
    """从 See all 弹窗读取全部 Research Areas。"""
    script = r'''
    const root = document.querySelector('.shared-refine-see-all');
    if (!root) return [];
    const out = [];
    const seen = new Set();
    root.querySelectorAll('input[name="SJ"]').forEach(inp => {
      const val = (inp.value || '').trim();
      if (!val || seen.has(val)) return;
      seen.add(val);
      let name = '';
      const lab = root.querySelector('label[for="' + inp.id + '"]');
      const nameEl = lab && lab.querySelector('.filter-option-name');
      const countEl = lab && lab.querySelector('.filter-option-count');
      if (nameEl) name = (nameEl.textContent || '').trim();
      if (!name) {
        const aria = inp.getAttribute('aria-label') || '';
        name = aria.replace(/,\s*[\d,]+\s*$/, '').trim();
      }
      name = name.replace(/\s+/g, ' ').trim();
      let count = 0;
      if (countEl) {
        count = parseInt((countEl.textContent || '').replace(/,/g, ''), 10) || 0;
      }
      if (!count) {
        const aria = inp.getAttribute('aria-label') || '';
        const m = aria.match(/,\s*([\d,]+)\s*$/);
        if (m) count = parseInt(m[1].replace(/,/g, ''), 10) || 0;
      }
      out.push({value: val, name: name || val, count: count});
    });
    return out;
    '''
    try:
        rows = driver.execute_script(script) or []
    except Exception:
        rows = []
    cleaned = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        name = (row.get('name') or '').strip()
        value = (row.get('value') or '').strip()
        if not name and not value:
            continue
        cleaned.append({
            'name': name or value,
            'value': value,
            'count': int(row.get('count') or 0),
        })
    return cleaned


def get_research_area_count_from_see_all(driver, area):
    """从当前打开的 See all 弹窗刷新读取指定研究方向的最新数量。

    返回 int；找不到则返回 None。
    """
    target_value = (area.get('value') or '').strip()
    target_name = (area.get('name') or '').strip()
    try:
        n = driver.execute_script(
            r'''
            const targetValue = (arguments[0] || '').trim();
            const targetName = (arguments[1] || '').trim();
            const root = document.querySelector('.shared-refine-see-all');
            if (!root) return null;
            const nameL = targetName.toLowerCase();
            const inputs = Array.from(root.querySelectorAll('input[name="SJ"]'));
            const pickCount = (inp) => {
              const lab = root.querySelector('label[for="' + inp.id + '"]');
              const countEl = lab && lab.querySelector('.filter-option-count');
              if (countEl) {
                const n = parseInt(
                  (countEl.textContent || '').replace(/,/g, ''), 10);
                if (!isNaN(n)) return n;
              }
              const aria = inp.getAttribute('aria-label') || '';
              const m = aria.match(/,\s*([\d,]+)\s*$/);
              if (m) {
                const n = parseInt(m[1].replace(/,/g, ''), 10);
                if (!isNaN(n)) return n;
              }
              return null;
            };
            const optionName = (inp) => {
              const lab = root.querySelector('label[for="' + inp.id + '"]');
              const nameEl = lab && lab.querySelector('.filter-option-name');
              if (nameEl) {
                return (nameEl.textContent || '').replace(/\s+/g, ' ').trim();
              }
              const aria = inp.getAttribute('aria-label') || '';
              return aria.replace(/,\s*[\d,]+\s*$/, '').replace(/\s+/g, ' ').trim();
            };
            for (const inp of inputs) {
              const val = (inp.value || '').trim();
              if (targetValue && val === targetValue) return pickCount(inp);
            }
            if (nameL) {
              for (const inp of inputs) {
                const nm = optionName(inp).toLowerCase();
                const aria = (inp.getAttribute('aria-label') || '').toLowerCase();
                if (nm === nameL
                    || aria.startsWith(nameL + ',')
                    || aria.startsWith(nameL + ' ')) {
                  return pickCount(inp);
                }
              }
            }
            return null;
            ''', target_value, target_name)
        if n is None:
            return None
        return int(n)
    except Exception:
        return None


def refresh_area_count_from_see_all(driver, area, old_expected=None):
    """打开弹窗后刷新 area['count']，返回最新数量（失败则回退 old/原 count）。"""
    fresh = get_research_area_count_from_see_all(driver, area)
    if fresh is None:
        fallback = int(
            old_expected if old_expected is not None
            else (area.get('count') or 0))
        print(f'未能从弹窗刷新数量，沿用: {fallback}')
        return fallback
    prev = int(
        old_expected if old_expected is not None
        else (area.get('count') or 0))
    area['count'] = fresh
    if prev and prev != fresh:
        print(f'弹窗数量已刷新: {prev} → {fresh}')
    else:
        print(f'Refine 前弹窗数量: {fresh}')
    return fresh


def collect_all_research_areas(driver, timeout=30):
    """打开 See all，采集全部研究方向后关闭弹窗。"""
    if not open_research_areas_see_all(driver, timeout=timeout):
        return []
    time.sleep(0.6)
    areas = list_research_areas_from_see_all(driver)
    try:
        panel = _research_areas_see_all_panel(driver)
        if panel:
            opts = panel.find_elements(By.CSS_SELECTOR, '.options, ul.refine-options')
            for box in opts:
                driver.execute_script(
                    'arguments[0].scrollTop = arguments[0].scrollHeight;', box)
                time.sleep(0.4)
            more = list_research_areas_from_see_all(driver)
            if len(more) > len(areas):
                areas = more
    except Exception:
        pass
    close_research_areas_see_all(driver)
    print(f'Research Areas 共 {len(areas)} 类')
    return areas


def parse_main_hit_count(driver, timeout=60, quiet=False):
    """只读主结果区命中数，避免误读侧栏 Research Areas 的 filter-option-count。

    新版 WoS 常见位置：
    - h1.search-info-title > span.brand-blue
    - span.tab-results-count（Documents 标签）
    - 全选复选框 aria-label / 文案中的 0/N
    """
    deadline = time.time() + timeout
    selectors = [
        '#hitCount.hitCount',
        '#hitCount',
        'span#hitCount',
        '[data-ta="hit-count"]',
        'app-search-results-count #hitCount',
        'app-search-results-count .hitCount',
        'app-search-results-count',
        'h1.search-info-title .brand-blue',
        'h1.search-info-title-new .brand-blue',
        '.search-info-title .brand-blue',
        '.search-info-title-new .brand-blue',
        '.search-info-container-new .brand-blue',
        '.search-description-holder .brand-blue',
        'span.tab-results-count',
        'a#parent-summary-link .tab-results-count',
        '[data-ta="parent-summary-link"] .tab-results-count',
    ]

    def _parse_int_text(text):
        if not text:
            return None
        cleaned = (
            str(text).replace('\xa0', ' ').replace(',', '')
            .replace(' ', '').strip())
        m = re.search(r'(\d+)', cleaned)
        return int(m.group(1)) if m else None

    while time.time() < deadline:
        driver = ctx_driver(driver)
        if not is_driver_alive(driver):
            return None
        if is_page_unusable(driver):
            return None

        # 1) CSS 选择器
        for sel in selectors:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
            except Exception:
                continue
            for el in els:
                try:
                    if not el.is_displayed():
                        continue
                    skip = driver.execute_script(
                        '''
                        const n = arguments[0];
                        return !!(n.closest && (
                          n.closest('.filter-option') ||
                          n.closest('app-refine-option') ||
                          n.closest('.shared-refine-see-all') ||
                          n.closest('[data-ta="filter-section-SJ"]') ||
                          n.closest('.refine-options')
                        ));
                        ''', el)
                    if skip:
                        continue
                    n = _parse_int_text(el.text or '')
                    if n is None:
                        continue
                    if not quiet:
                        print(f'命中条数: {n}')
                    return n
                except Exception:
                    continue

        # 2) JS 汇总（标题 / Documents 标签 / 全选框）
        try:
            n = driver.execute_script(
                r'''
                function parseNum(s) {
                  if (!s) return null;
                  const m = String(s).replace(/\u00a0/g, ' ')
                    .replace(/,/g, '').match(/(\d+)/);
                  return m ? parseInt(m[1], 10) : null;
                }
                const inFilter = (n) => !!(n && n.closest && (
                  n.closest('.filter-option') ||
                  n.closest('app-refine-option') ||
                  n.closest('.shared-refine-see-all') ||
                  n.closest('[data-ta="filter-section-SJ"]') ||
                  n.closest('.refine-options')
                ));
                const q = (sel) => Array.from(document.querySelectorAll(sel));
                for (const sel of [
                  'h1.search-info-title .brand-blue',
                  '.search-info-title .brand-blue',
                  '.search-info-container-new .brand-blue',
                  '.search-description-holder .brand-blue',
                  'span.tab-results-count',
                  'a#parent-summary-link .tab-results-count',
                  '#hitCount',
                  '[data-ta="hit-count"]'
                ]) {
                  for (const el of q(sel)) {
                    if (inFilter(el)) continue;
                    const n = parseNum(el.textContent || el.innerText || '');
                    if (n != null) return n;
                  }
                }
                // 全选：aria-label="Select all records, 0 of 7,991" 或 label 0/7,991
                for (const inp of q(
                  '[data-ta="select-page-checkbox"] input,'
                  + ' mat-checkbox[data-ta="select-page-checkbox"] input,'
                  + ' input[aria-label*="Select all records"]')) {
                  const aria = inp.getAttribute('aria-label') || '';
                  let m = aria.match(/of\s*([\d,]+)/i);
                  if (m) return parseNum(m[1]);
                  const lab = document.querySelector(
                    'label[for="' + (inp.id || '') + '"]');
                  const lt = (lab && (lab.textContent || '')) || '';
                  m = lt.match(/\/\s*([\d,]+)/);
                  if (m) return parseNum(m[1]);
                }
                const h1 = document.querySelector(
                  'h1.search-info-title, h1.search-info-title-new');
                if (h1) {
                  const m = (h1.innerText || '').match(
                    /([\d,]+)\s*results?/i);
                  if (m) return parseNum(m[1]);
                }
                return null;
                ''')
            if n is not None:
                n = int(n)
                if not quiet:
                    print(f'命中条数: {n}')
                return n
        except Exception:
            pass

        # 3) 正文兜底
        try:
            body = driver.find_element(By.TAG_NAME, 'body').text or ''
        except Exception:
            body = ''
        for pat in (
            r'([\d,]+)\s*results?\s+from',
            r'([\d,]+)\s*results?',
            r'([\d,]+)\s*条结果',
            r'Results?:\s*([\d,]+)',
            r'找到\s*([\d,]+)',
        ):
            m = re.search(pat, body, re.I)
            if m:
                n = int(m.group(1).replace(',', ''))
                if not quiet:
                    print(f'命中条数(正文): {n}')
                return n
        if any(kw in body for kw in (
                'No results', '0 results', '没有找到', '未找到结果',
                'Your search found no results')):
            if not quiet:
                print('命中条数: 0')
            return 0
        time.sleep(0.4)
    return None


def wait_hit_count_equals(driver, expected, baseline=None, timeout=60, tolerance=1):
    """等待主结果命中数等于（或接近）精炼弹窗中的预期数量。

    tolerance: WoS 弹窗数与结果页偶发差 1，默认允许 ±1。
    """
    expected = int(expected or 0)
    tolerance = max(0, int(tolerance))
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        last = parse_main_hit_count(driver, timeout=3, quiet=True)
        if last is None:
            time.sleep(0.5)
            continue
        if expected > 0 and abs(int(last) - expected) <= tolerance:
            if int(last) == expected:
                print(f'精炼校验通过: 页面命中 {last} == 弹窗 {expected}')
            else:
                print(
                    f'精炼校验通过(容差±{tolerance}): '
                    f'页面命中 {last} ≈ 弹窗 {expected}')
            return last
        if expected <= 0 and last == 0:
            print('精炼校验通过: 页面命中 0')
            return 0
        time.sleep(0.6)
    print(
        f'精炼校验失败: 页面命中 {last}，弹窗预期 {expected}'
        + (f'，精炼前 {baseline}' if baseline is not None else ''))
    return None


def _mat_checkbox_is_on(driver, inp):
    """判断 Angular Material checkbox 是否已勾选（不单靠 input.checked）。"""
    try:
        return bool(driver.execute_script(
            r'''
            const inp = arguments[0];
            if (!inp) return false;
            if (inp.checked) return true;
            const mat = inp.closest('mat-checkbox');
            if (!mat) return false;
            if (mat.classList.contains('mat-mdc-checkbox-checked')) return true;
            if (mat.classList.contains('mat-checkbox-checked')) return true;
            const aria = (mat.getAttribute('aria-checked')
              || inp.getAttribute('aria-checked') || '').toLowerCase();
            return aria === 'true';
            ''', inp))
    except Exception:
        try:
            return bool(inp.is_selected())
        except Exception:
            return False


def _real_click_element(driver, el):
    """真实鼠标点击（ActionChains / CDP），必要时再 JS click。"""
    try:
        driver.execute_script(
            'arguments[0].scrollIntoView({block:"center", inline:"nearest"});', el)
        time.sleep(0.15)
    except Exception:
        pass

    # CDP 坐标点击（绕过遮挡层对 element.click 的拦截）
    try:
        rect = driver.execute_script(
            '''
            const r = arguments[0].getBoundingClientRect();
            return {x: r.left + r.width/2, y: r.top + r.height/2,
                    w: r.width, h: r.height};
            ''', el)
        if rect and rect.get('w', 0) > 0 and rect.get('h', 0) > 0:
            x, y = float(rect['x']), float(rect['y'])
            driver.execute_cdp_cmd('Input.dispatchMouseEvent', {
                'type': 'mousePressed', 'x': x, 'y': y,
                'button': 'left', 'clickCount': 1,
            })
            driver.execute_cdp_cmd('Input.dispatchMouseEvent', {
                'type': 'mouseReleased', 'x': x, 'y': y,
                'button': 'left', 'clickCount': 1,
            })
            return True
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(el).pause(0.05).click().perform()
        return True
    except Exception:
        pass
    try:
        el.click()
        return True
    except Exception:
        pass
    return safe_click(driver, el)


def _force_check_sj_input(driver, inp, want_checked=True):
    """强制切换 Angular MDC checkbox 选中态（最后手段）。"""
    try:
        return bool(driver.execute_script(
            r'''
            const inp = arguments[0];
            const want = !!arguments[1];
            if (!inp) return false;
            const mat = inp.closest('mat-checkbox');
            const lab = inp.id
              ? document.querySelector('label[for="' + inp.id + '"]') : null;
            const touch = mat && mat.querySelector('.mat-mdc-checkbox-touch-target');
            const box = mat && mat.querySelector('.mdc-checkbox');
            const targets = [touch, box, lab, mat, inp].filter(Boolean);

            const fire = (el) => {
              const opts = {bubbles: true, cancelable: true, view: window};
              ['pointerover','mouseover','pointerdown','mousedown',
               'pointerup','mouseup','click'].forEach(t => {
                try { el.dispatchEvent(new PointerEvent(t, opts)); } catch (e) {}
                try { el.dispatchEvent(new MouseEvent(t, opts)); } catch (e) {}
              });
            };

            const isOn = () => !!(
              inp.checked
              || (mat && mat.classList.contains('mat-mdc-checkbox-checked'))
              || (mat && (mat.getAttribute('aria-checked') || '') === 'true')
            );

            if (isOn() === want) return true;

            for (const t of targets) {
              fire(t);
              if (isOn() === want) return true;
            }

            // 直接写 native checked + Angular 友好事件
            try {
              const desc = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'checked');
              if (desc && desc.set) desc.set.call(inp, want);
              else inp.checked = want;
              inp.dispatchEvent(new Event('input', {bubbles: true}));
              inp.dispatchEvent(new Event('change', {bubbles: true}));
              if (mat) {
                mat.classList.toggle('mat-mdc-checkbox-checked', want);
                mat.classList.toggle('mat-checkbox-checked', want);
                mat.setAttribute('aria-checked', want ? 'true' : 'false');
              }
            } catch (e) {}
            return isOn() === want;
            ''', inp, want_checked))
    except Exception:
        return False


def _click_sj_checkbox(driver, inp):
    """按 WoS See all 结构点击单个 Research Area 复选框。"""
    # 优先点 touch-target / label / filter-option-name（MDC 原生 input 常被挡住）
    candidates = []
    try:
        mat = inp.find_element(By.XPATH, './ancestor::mat-checkbox[1]')
        candidates.extend(mat.find_elements(
            By.CSS_SELECTOR, '.mat-mdc-checkbox-touch-target, .mdc-checkbox'))
        candidates.append(mat)
    except Exception:
        mat = None
    try:
        inp_id = inp.get_attribute('id') or ''
        if inp_id:
            labs = driver.find_elements(By.CSS_SELECTOR, f'label[for="{inp_id}"]')
            candidates.extend(labs)
            for lab in labs:
                candidates.extend(lab.find_elements(
                    By.CSS_SELECTOR, '.filter-option-name'))
    except Exception:
        pass
    candidates.append(inp)

    for el in candidates:
        try:
            if not el.is_displayed() and el is not inp:
                continue
        except Exception:
            continue
        if _real_click_element(driver, el):
            time.sleep(0.35)
            if _mat_checkbox_is_on(driver, inp):
                return True

    if _force_check_sj_input(driver, inp, True):
        time.sleep(0.2)
        return _mat_checkbox_is_on(driver, inp)
    return _mat_checkbox_is_on(driver, inp)


def _find_sj_input_in_see_all(driver, area):
    """在 See all 弹窗中定位目标 input[name=SJ]。"""
    panel = _research_areas_see_all_panel(driver)
    if not panel:
        return None
    target_value = (area.get('value') or '').strip()
    target_name = (area.get('name') or '').strip()
    target_name_l = target_name.lower()

    if target_value:
        # value 含空格，用 XPath 更稳
        found = panel.find_elements(
            By.XPATH,
            f'.//input[@name="SJ" and @value="{target_value}"]')
        if found:
            return found[0]

    for inp in panel.find_elements(By.CSS_SELECTOR, 'input[name="SJ"]'):
        try:
            aria = (inp.get_attribute('aria-label') or '').strip()
            aria_l = aria.lower()
            name_el = None
            try:
                lab = panel.find_element(
                    By.CSS_SELECTOR,
                    f'label[for="{inp.get_attribute("id")}"]')
                name_els = lab.find_elements(By.CSS_SELECTOR, '.filter-option-name')
                if name_els:
                    name_el = (name_els[0].text or '').strip()
            except Exception:
                pass
            nm = (name_el or re.sub(r',\s*[\d,]+\s*$', '', aria)).strip()
            nm_l = nm.lower()
            if target_name_l and (
                    nm_l == target_name_l
                    or aria_l.startswith(target_name_l + ',')
                    or aria_l.startswith(target_name_l + ' ')):
                return inp
            if target_value and (inp.get_attribute('value') or '').strip() == target_value:
                return inp
        except StaleElementReferenceException:
            continue
        except Exception:
            continue
    return None


def select_one_research_area_in_see_all(driver, area, timeout=35):
    """在 See all 弹窗内仅勾选指定研究方向（真实点击 MDC checkbox）。"""
    target_value = (area.get('value') or '').strip()
    target_name = (area.get('name') or '').strip()
    deadline = time.time() + timeout
    last_reason = ''
    searched = False

    while time.time() < deadline:
        if _see_all_sj_count(driver) <= 0:
            last_reason = 'panel_missing'
            time.sleep(0.4)
            continue

        panel = _research_areas_see_all_panel(driver)
        if not panel:
            last_reason = 'panel_missing'
            time.sleep(0.4)
            continue

        # 清空搜索框，避免残留过滤
        try:
            search_boxes = panel.find_elements(
                By.CSS_SELECTOR,
                'input[aria-label="Search for Research Areas"], '
                'input[placeholder="Search for Research Areas"], '
                '.refine-search-bar input, .search-bar input')
            for box in search_boxes:
                val = box.get_attribute('value') or ''
                if val:
                    box.clear()
                    time.sleep(0.3)
        except Exception:
            pass

        inp = _find_sj_input_in_see_all(driver, area)
        if inp is None and target_name and not searched:
            # 用搜索框缩小列表后再找
            try:
                search_boxes = panel.find_elements(
                    By.CSS_SELECTOR,
                    'input[aria-label="Search for Research Areas"], '
                    'input[placeholder="Search for Research Areas"], '
                    '.refine-search-bar input')
                if search_boxes:
                    box = search_boxes[0]
                    box.clear()
                    box.send_keys(target_name)
                    searched = True
                    time.sleep(0.9)
                    inp = _find_sj_input_in_see_all(driver, area)
            except Exception:
                pass

        if inp is None:
            last_reason = 'not_found'
            sample = []
            try:
                for el in panel.find_elements(
                        By.CSS_SELECTOR, 'input[name="SJ"]')[:3]:
                    sample.append({
                        'value': el.get_attribute('value'),
                        'aria': el.get_attribute('aria-label'),
                    })
            except Exception:
                pass
            if sample:
                print(f'未找到目标，弹窗样例: {sample}')
            time.sleep(0.5)
            continue

        # 取消其它已勾选项
        try:
            for other in panel.find_elements(By.CSS_SELECTOR, 'input[name="SJ"]'):
                try:
                    if other == inp:
                        continue
                    if _mat_checkbox_is_on(driver, other):
                        _click_sj_checkbox(driver, other)
                        time.sleep(0.1)
                except StaleElementReferenceException:
                    break
                except Exception:
                    continue
        except Exception:
            pass

        # 重新定位目标（取消其它项可能导致 stale）
        inp = _find_sj_input_in_see_all(driver, area)
        if inp is None:
            last_reason = 'stale_after_uncheck'
            time.sleep(0.3)
            continue

        if _mat_checkbox_is_on(driver, inp):
            print(f'已在 See all 弹窗勾选: {target_name} [{target_value}]')
            return True

        if _click_sj_checkbox(driver, inp):
            # 再确认一次
            inp2 = _find_sj_input_in_see_all(driver, area) or inp
            if _mat_checkbox_is_on(driver, inp2):
                print(f'已在 See all 弹窗勾选: {target_name} [{target_value}]')
                return True

        last_reason = 'click_no_effect'
        # 最后手段：对 label 发完整鼠标事件
        try:
            inp3 = _find_sj_input_in_see_all(driver, area)
            if inp3 is not None:
                ok = driver.execute_script(
                    r'''
                    const inp = arguments[0];
                    const mat = inp.closest('mat-checkbox');
                    const touch = mat && mat.querySelector(
                      '.mat-mdc-checkbox-touch-target, .mdc-checkbox__native-control');
                    const lab = document.querySelector('label[for="' + inp.id + '"]');
                    const target = touch || lab || mat || inp;
                    const opts = {bubbles:true, cancelable:true, view:window};
                    ['pointerdown','mousedown','pointerup','mouseup','click'].forEach(t => {
                      target.dispatchEvent(new MouseEvent(t, opts));
                    });
                    return !!(inp.checked
                      || (mat && mat.classList.contains('mat-mdc-checkbox-checked')));
                    ''', inp3)
                if ok:
                    print(f'已在 See all 弹窗勾选(事件): {target_name}')
                    return True
        except Exception:
            pass
        time.sleep(0.4)

    print(
        f'勾选失败原因: {last_reason} | 目标={target_name!r} value={target_value!r}')
    return False


def click_research_areas_refine(driver, timeout=25):
    """点击 See all 弹窗底部 Refine。"""
    deadline = time.time() + timeout
    clicked = False
    while time.time() < deadline:
        if clicked and _see_all_sj_count(driver) == 0 and not _research_areas_see_all_panel(driver):
            accept_cookies(driver, timeout=2)
            time.sleep(1.0)
            return True

        panel = _research_areas_see_all_panel(driver)
        if not panel:
            if clicked:
                accept_cookies(driver, timeout=2)
                time.sleep(1.0)
                return True
            time.sleep(0.3)
            continue

        # 优先 Selenium 点 seeAll-footer 里的 Refine
        btns = panel.find_elements(
            By.CSS_SELECTOR,
            '.seeAll-footer button[data-ta="refine-submit"]')
        if not btns:
            btns = panel.find_elements(
                By.CSS_SELECTOR, 'button[data-ta="refine-submit"]')
        for btn in reversed(btns):
            try:
                style = (btn.get_attribute('style') or '').lower()
                if 'opacity: 0' in style or 'opacity:0' in style:
                    continue
                if (btn.get_attribute('aria-hidden') or '').lower() == 'true':
                    continue
                if not btn.is_displayed():
                    continue
                if _real_click_element(driver, btn):
                    clicked = True
                    print('已点击 See all 弹窗 Refine 按钮')
                    time.sleep(1.2)
                    wait_deadline = time.time() + 30
                    while time.time() < wait_deadline:
                        if (_see_all_sj_count(driver) == 0
                                and not _research_areas_see_all_panel(driver)):
                            accept_cookies(driver, timeout=2)
                            time.sleep(1.0)
                            return True
                        time.sleep(0.3)
                    return True
            except Exception:
                continue

        # JS 兜底
        try:
            result = driver.execute_script(
                r'''
                const root = document.querySelector('.shared-refine-see-all');
                if (!root) return {ok:false};
                const btns = Array.from(root.querySelectorAll(
                  '.seeAll-footer button[data-ta="refine-submit"], button[data-ta="refine-submit"]'));
                for (let i = btns.length - 1; i >= 0; i--) {
                  const b = btns[i];
                  const s = window.getComputedStyle(b);
                  if (s.display === 'none' || s.visibility === 'hidden') continue;
                  if (parseFloat(s.opacity || '1') < 0.2) continue;
                  if ((b.getAttribute('aria-hidden') || '') === 'true') continue;
                  if (b.disabled) continue;
                  b.click();
                  return {ok:true};
                }
                return {ok:false};
                ''') or {}
            if result.get('ok'):
                clicked = True
                print('已点击 See all 弹窗 Refine 按钮(JS)')
                time.sleep(1.2)
                return True
        except Exception:
            pass
        time.sleep(0.35)
    print('未能点击 See all 弹窗 Refine 按钮')
    return False


def research_area_filter_applied(driver, area_name):
    """结果页是否出现已应用的 Research Areas 精炼标签。"""
    name = (area_name or '').strip()
    if not name:
        return False
    try:
        body = driver.find_element(By.TAG_NAME, 'body').text
    except Exception:
        return False
    needles = [
        f'Research Areas: {name}',
        f'Research Areas:{name}',
        f'研究方向: {name}',
        f'研究方向:{name}',
    ]
    return any(n in body for n in needles)


def refine_by_one_research_area(
        driver, area, expected_count=None, baseline_hits=None,
        timeout=50, verify_retries=3):
    """打开 See all → 勾选一类 → Refine，并校验命中数与弹窗一致。

    每次打开弹窗时重新读取该类当前数量（不沿用初次采集的陈旧 count）。
    """
    seed_expected = int(
        expected_count if expected_count is not None
        else (area.get('count') or 0))
    area_name = area.get('name') or area.get('value') or ''

    for attempt in range(1, verify_retries + 1):
        baseline = baseline_hits
        if baseline is None:
            baseline = parse_main_hit_count(driver, timeout=8, quiet=True)

        print(f'打开 See all 并精炼 ({attempt}/{verify_retries}): {area_name}')
        if not open_research_areas_see_all(driver, timeout=min(30, timeout)):
            print(f'Refine 打开 See all 失败 ({attempt}/{verify_retries})')
            continue

        # 弹窗打开后立刻刷新数量（WoS 计数会变）
        expected = refresh_area_count_from_see_all(
            driver, area, old_expected=seed_expected)

        if not select_one_research_area_in_see_all(driver, area, timeout=35):
            print(f'勾选研究方向失败: {area_name}')
            close_research_areas_see_all(driver)
            continue

        # 勾选后再读一次，避免列表异步刷新后数量变化
        time.sleep(0.35)
        expected = refresh_area_count_from_see_all(
            driver, area, old_expected=expected)
        seed_expected = expected

        if not click_research_areas_refine(driver, timeout=25):
            print(f'点击 Refine 失败: {area_name}')
            close_research_areas_see_all(driver)
            continue

        hits = wait_hit_count_equals(
            driver, expected, baseline=baseline, timeout=55, tolerance=1)
        if hits is None:
            actual = parse_main_hit_count(driver, timeout=8, quiet=True)
            print(
                f'精炼后数量不一致 ({attempt}/{verify_retries}): '
                f'页面={actual} 弹窗={expected} ({area_name})')
            if attempt < verify_retries:
                clear_research_area_filters(driver)
                time.sleep(0.8)
            continue

        if not research_area_filter_applied(driver, area_name):
            print(f'提示: 未看到精炼标签文案，但命中数已与弹窗一致 ({hits})')
        # 以页面实际命中为准，供后续导出分批
        area['count'] = int(hits)
        return hits

    return None


def clear_research_area_filters(driver, timeout=12):
    """清除已应用的 Research Areas 精炼条件。"""
    deadline = time.time() + timeout
    cleared = False
    while time.time() < deadline:
        candidates = []
        for el in driver.find_elements(By.CSS_SELECTOR, 'button, a, [role="button"]'):
            try:
                if not el.is_displayed():
                    continue
                label = (
                    (el.get_attribute('aria-label') or '')
                    + ' ' + (el.text or '')).strip()
                low = label.lower()
                if not label:
                    continue
                if any(k in low for k in (
                        'remove research', 'clear research',
                        'remove filter', 'clear filter',
                        '清除', '移除')) and (
                        'research' in low or 'sj.' in low
                        or '研究方向' in label):
                    candidates.append(el)
                elif (('research areas' in low or '研究方向' in label)
                      and any(k in low for k in ('remove', 'clear', '×', 'x', '关闭'))):
                    candidates.append(el)
            except Exception:
                continue
        for el in driver.find_elements(
                By.CSS_SELECTOR,
                '[class*="selected-filter"] button, '
                '[class*="applied-filter"] button, '
                '[class*="refine-pill"] button, '
                'mat-chip button, .mat-mdc-chip button'):
            try:
                if el.is_displayed():
                    candidates.append(el)
            except Exception:
                continue

        if not candidates:
            break
        for el in candidates[:6]:
            try:
                safe_click(driver, el)
                cleared = True
                time.sleep(0.5)
            except Exception:
                continue
        time.sleep(0.4)
        page = ''
        try:
            page = driver.find_element(By.TAG_NAME, 'body').text[:4000]
        except Exception:
            pass
        if 'Research Areas:' not in page and '研究方向:' not in page:
            break
    return cleared


def export_range_by_research_areas(
        driver, save_dir, label, start_date, end_date,
        database, collections, indent=''):
    """日期段检索结果上，按 Research Areas 逐类 Refine 并导出。"""
    bind_browser_ctx(driver, save_dir, database, collections)
    driver = ctx_driver(driver)
    baseline_total = parse_main_hit_count(driver, timeout=15, quiet=True)
    areas = collect_all_research_areas(driver)
    if not areas:
        print(f'{indent}未获取到 Research Areas 列表，回退为整段导出')
        return None  # 调用方回退

    col_bit = collections_label(collections)
    base_prefix = (
        f'SCI_{label}_{_safe_dirname(SEARCH_DATABASE)}_{_safe_dirname(col_bit)}')
    ra_root = os.path.join(save_dir, 'ResearchAreas')
    os.makedirs(ra_root, exist_ok=True)

    ok_any = False
    for idx, area in enumerate(areas, 1):
        area_name = area.get('name') or area.get('value') or f'area_{idx}'
        area_dir_name = _safe_dirname(area_name)
        area_dir = os.path.join(ra_root, area_dir_name)
        os.makedirs(area_dir, exist_ok=True)
        prefix = f'{base_prefix}_{area_dir_name}'
        expected = int(area.get('count') or 0)
        print(
            f'\n{indent}----- Research Area {idx}/{len(areas)}: '
            f'{area_name} (弹窗 {expected}) -----')

        if idx > 1:
            cleared = clear_research_area_filters(driver)
            if not cleared:
                print(f'{indent}清除精炼失败，重新检索后再选下一类')
                hits_reset = run_search_once(
                    ctx_driver(driver), start_date, end_date, database,
                    collections, skip_sort=True, download_dir=save_dir)
                driver = ctx_driver(driver)
                if hits_reset is None or hits_reset <= 0:
                    print(f'{indent}重新检索失败，跳过 {area_name}')
                    record_export_failure(
                        save_dir, f'{label}_RA_{area_dir_name}_research')
                    continue
                baseline_total = hits_reset
            else:
                time.sleep(0.8)
                baseline_total = parse_main_hit_count(
                    driver, timeout=15, quiet=True) or baseline_total

        hits = refine_by_one_research_area(
            driver, area, expected_count=expected,
            baseline_hits=baseline_total, verify_retries=3)
        if hits is None:
            print(f'{indent}Refine/校验失败: {area_name}（弹窗预期 {expected}）')
            record_export_failure(
                save_dir, f'{label}_RA_{area_dir_name}_refine')
            # 尽量回到未精炼结果
            hits_reset = run_search_once(
                ctx_driver(driver), start_date, end_date, database,
                collections, skip_sort=True, download_dir=save_dir)
            driver = ctx_driver(driver)
            if hits_reset:
                baseline_total = hits_reset
            continue

        if hits == 0:
            print(f'{indent}{area_name} 无结果，跳过')
            continue

        export_hits = hits
        if hits > WOS_MAX_HITS:
            print(
                f'{indent}{area_name} 命中 {hits} > {WOS_MAX_HITS}，'
                f'仅导出前 {WOS_MAX_HITS} 条')
            record_export_failure(
                save_dir,
                f'{label}_RA_{area_dir_name}_partial_{WOS_MAX_HITS}_of_{hits}')
            export_hits = WOS_MAX_HITS

        # RA 模式不排序：切换 Date oldest first 会冲掉 Research Areas 精炼
        set_chrome_download_dir(driver, area_dir)

        def restore_results(
                area_ref=area, area_label=area_name, exp=expected):
            print(f'{indent}结果页丢失，重新检索并 Refine: {area_label}')
            hits2 = run_search_once(
                ctx_driver(driver), start_date, end_date, database,
                collections, skip_sort=True, download_dir=save_dir)
            if hits2 is None or hits2 <= 0:
                return False
            verified = refine_by_one_research_area(
                ctx_driver(driver), area_ref, expected_count=exp,
                baseline_hits=hits2, verify_retries=3)
            return verified is not None

        def ensure_refined(
                area_ref=area, area_label=area_name, exp=expected):
            """导出前数量不一致时重新 Refine（不排序）。"""
            print(f'{indent}重新 Refine: {area_label}')
            d = ctx_driver(driver)
            cleared = clear_research_area_filters(d)
            if not cleared:
                hits2 = run_search_once(
                    ctx_driver(driver), start_date, end_date, database,
                    collections, skip_sort=True, download_dir=save_dir)
                if hits2 is None or hits2 <= 0:
                    return False
                baseline = hits2
            else:
                time.sleep(0.6)
                baseline = parse_main_hit_count(
                    ctx_driver(driver), timeout=12, quiet=True)
            verified = refine_by_one_research_area(
                ctx_driver(driver), area_ref, expected_count=exp,
                baseline_hits=baseline, verify_retries=3)
            return verified is not None

        n = export_all_hits(
            ctx_driver(driver), area_dir, export_hits, prefix,
            restore_fn=restore_results,
            expected_hits=hits,
            ensure_hits_fn=ensure_refined)
        driver = ctx_driver(driver)
        print(f'{indent}{area_name} 完成，成功批次 {n}')
        ok_any = ok_any or n > 0

    set_chrome_download_dir(driver, save_dir)
    print(f'{indent}{label} Research Areas 分类导出结束')
    return ok_any


def _match_research_area_entry(areas, name):
    """在 See all 列表中按名称匹配研究方向。"""
    name = (name or '').strip()
    if not name or not areas:
        return None
    name_l = name.lower()
    name_safe = _safe_dirname(name).lower()
    for a in areas:
        an = (a.get('name') or '').strip()
        if an.lower() == name_l:
            return a
    for a in areas:
        an = (a.get('name') or '').strip()
        if _safe_dirname(an).lower() == name_safe:
            return a
    for a in areas:
        an = (a.get('name') or '').strip()
        al = an.lower()
        if name_l in al or al in name_l:
            return a
    return None


def export_ra_missing_batches(
        driver, save_dir, label, start_date, end_date,
        database, collections, jobs, indent=''):
    """按核对结果仅导出缺失批次。

    jobs: [{'name': str, 'count': int, 'batches': [(start, end), ...]}, ...]
    返回成功导出的批次数。
    """
    if not jobs:
        print(f'{indent}无缺失批次，跳过')
        return 0

    bind_browser_ctx(driver, save_dir, database, collections)
    driver = ctx_driver(driver)

    print(f'\n{indent}===== 补导缺失批次: {label}（{len(jobs)} 个研究方向） =====')
    hits = run_search_once(
        driver, start_date, end_date, database, collections,
        skip_sort=True, download_dir=save_dir)
    driver = ctx_driver(driver)
    if hits is None or hits <= 0:
        print(f'{indent}检索失败，无法补导')
        return 0

    areas = collect_all_research_areas(driver)
    if not areas:
        print(f'{indent}未获取到 Research Areas 列表')
        return 0

    col_bit = collections_label(collections)
    base_prefix = (
        f'SCI_{label}_{_safe_dirname(SEARCH_DATABASE)}_{_safe_dirname(col_bit)}')
    ra_root = os.path.join(save_dir, 'ResearchAreas')
    os.makedirs(ra_root, exist_ok=True)

    baseline_total = hits
    success = 0

    for idx, job in enumerate(jobs, 1):
        area_name = (job.get('name') or '').strip()
        expected = int(job.get('count') or 0)
        batches = list(job.get('batches') or [])
        if not area_name or not batches:
            continue

        area = _match_research_area_entry(areas, area_name)
        if area is None:
            print(f'{indent}未匹配到研究方向，跳过: {area_name}')
            record_export_failure(
                save_dir, f'{label}_RA_{_safe_dirname(area_name)}_nomatch')
            continue

        # 弹窗 count 优先
        if int(area.get('count') or 0) > 0:
            expected = int(area.get('count'))

        area_dir_name = _safe_dirname(area.get('name') or area_name)
        area_dir = os.path.join(ra_root, area_dir_name)
        os.makedirs(area_dir, exist_ok=True)
        prefix = f'{base_prefix}_{area_dir_name}'

        print(
            f'\n{indent}----- 补导 {idx}/{len(jobs)}: {area_name} '
            f'(预期 {expected}，缺失 {len(batches)} 批) -----')
        for s, e in batches:
            print(f'{indent}  待补: {s}-{e}')

        if idx > 1:
            cleared = clear_research_area_filters(driver)
            if not cleared:
                hits_reset = run_search_once(
                    ctx_driver(driver), start_date, end_date, database,
                    collections, skip_sort=True, download_dir=save_dir)
                driver = ctx_driver(driver)
                if hits_reset is None or hits_reset <= 0:
                    print(f'{indent}重新检索失败，跳过 {area_name}')
                    continue
                baseline_total = hits_reset
            else:
                time.sleep(0.8)
                baseline_total = parse_main_hit_count(
                    ctx_driver(driver), timeout=15, quiet=True) or baseline_total

        hits_ra = refine_by_one_research_area(
            ctx_driver(driver), area, expected_count=expected,
            baseline_hits=baseline_total, verify_retries=3)
        driver = ctx_driver(driver)
        if hits_ra is None:
            print(f'{indent}Refine 失败，跳过 {area_name}')
            record_export_failure(
                save_dir, f'{label}_RA_{area_dir_name}_refine_fill')
            hits_reset = run_search_once(
                ctx_driver(driver), start_date, end_date, database,
                collections, skip_sort=True, download_dir=save_dir)
            driver = ctx_driver(driver)
            if hits_reset:
                baseline_total = hits_reset
            continue

        set_chrome_download_dir(driver, area_dir)

        def restore_results(
                area_ref=area, area_label=area_name, exp=expected):
            print(f'{indent}结果页丢失，重新检索并 Refine: {area_label}')
            hits2 = run_search_once(
                ctx_driver(driver), start_date, end_date, database,
                collections, skip_sort=True, download_dir=save_dir)
            if hits2 is None or hits2 <= 0:
                return False
            verified = refine_by_one_research_area(
                ctx_driver(driver), area_ref, expected_count=exp,
                baseline_hits=hits2, verify_retries=3)
            return verified is not None

        def ensure_refined(
                area_ref=area, area_label=area_name, exp=expected):
            print(f'{indent}重新 Refine: {area_label}')
            d = ctx_driver(driver)
            cleared = clear_research_area_filters(d)
            if not cleared:
                hits2 = run_search_once(
                    ctx_driver(driver), start_date, end_date, database,
                    collections, skip_sort=True, download_dir=save_dir)
                if hits2 is None or hits2 <= 0:
                    return False
                baseline = hits2
            else:
                time.sleep(0.6)
                baseline = parse_main_hit_count(
                    ctx_driver(driver), timeout=12, quiet=True)
            verified = refine_by_one_research_area(
                ctx_driver(driver), area_ref, expected_count=exp,
                baseline_hits=baseline, verify_retries=3)
            return verified is not None

        for start_n, end_n in batches:
            driver = ctx_driver(driver)
            print(f'\n{indent}----- 补导批次 {start_n}-{end_n} / {hits_ra} -----')
            existed = find_existing_batch_file(
                area_dir, prefix, start_n, end_n)
            if existed:
                print(f'{indent}批次已存在，跳过: {existed}')
                success += 1
                continue

            if not recover_for_export(driver, restore_results):
                print(f'{indent}结果页无法恢复，记录失败 {start_n}-{end_n}')
                record_export_failure(
                    save_dir, f'{prefix}_{start_n}-{end_n}')
                continue

            driver = ctx_driver(driver)
            if not _verify_hits_before_export(
                    driver, hits_ra, ensure_hits_fn=ensure_refined):
                print(f'{indent}导出前校验失败，记录失败 {start_n}-{end_n}')
                record_export_failure(
                    save_dir, f'{prefix}_{start_n}-{end_n}')
                continue

            driver = ctx_driver(driver)
            before = snapshot_download_meta(area_dir)
            path = None
            try:
                path = export_batch(
                    driver, area_dir, start_n, end_n, prefix,
                    expected_hits=hits_ra)
            except Exception as e:
                print(f'{indent}补导异常 {start_n}-{end_n}:', e)
                try:
                    close_export_overlay(ctx_driver(driver))
                except Exception:
                    pass
            if not path:
                orphan = pick_new_download(area_dir, before)
                if orphan:
                    path = rename_export_file(
                        orphan, area_dir, prefix, start_n, end_n)
            if path:
                print(f'{indent}补导成功: {path}')
                success += 1
            else:
                print(f'{indent}补导失败: {start_n}-{end_n}')
                record_export_failure(
                    save_dir, f'{prefix}_{start_n}-{end_n}')

        driver = ctx_driver(driver)

    set_chrome_download_dir(ctx_driver(driver), save_dir)
    print(f'{indent}补导结束，成功批次 {success}')
    return success


def run_one_range_export(
        driver, save_dir, start_date, end_date, database, collections,
        depth=0, by_research_areas=False):
    """对一个日期段检索并导出；超 10 万则对半切分。返回是否有成功导出。"""
    label = (
        start_date if start_date == end_date
        else f'{start_date}_{end_date}')
    indent = '  ' * depth
    jobs = list(iter_collection_search_jobs(database, collections))
    if len(jobs) > 1 and depth < 8:
        print(f'\n{indent}===== 检索段: {label}（{len(jobs)} 个库） =====')
        ok_any = False
        for name, col in jobs:
            print(f'{indent}  -> {name}  {collection_search_urls(name)[0]}')
            ok = run_one_range_export(
                ctx_driver(driver), save_dir, start_date, end_date,
                SEARCH_DATABASE, col, depth=depth + 1,
                by_research_areas=by_research_areas)
            driver = ctx_driver(driver)
            ok_any = ok_any or ok
        return ok_any

    print(f'\n{indent}===== 检索段: {label} =====')
    print(f'{indent}检索页: {search_url_for(database, collections)}')
    if by_research_areas:
        print(f'{indent}模式: 按 Research Areas 分类精炼导出')

    bind_browser_ctx(driver, save_dir, database, collections)
    hits = run_search_once(
        ctx_driver(driver), start_date, end_date, database, collections,
        skip_sort=by_research_areas, download_dir=save_dir)
    driver = ctx_driver(driver)
    if hits is None:
        record_export_failure(save_dir, label)
        return False
    if hits == 0:
        print(f'{indent}{label} 无结果，跳过')
        return False

    if hits > WOS_MAX_HITS and by_research_areas:
        print(
            f'{indent}{label} 命中 {hits} > {WOS_MAX_HITS}，'
            f'RA 模式不切分日期，直接按 Research Areas 分类导出')
    elif hits > WOS_MAX_HITS:
        start_d = datetime.strptime(start_date, '%Y-%m-%d').date()
        end_d = datetime.strptime(end_date, '%Y-%m-%d').date()
        if start_d < end_d:
            print(f'{indent}{label} 命中 {hits} > {WOS_MAX_HITS}，对半切分日期')
            ok_any = False
            for s, e in split_date_range_halves(start_d, end_d):
                ok = run_one_range_export(
                    ctx_driver(driver), save_dir,
                    s.strftime('%Y-%m-%d'), e.strftime('%Y-%m-%d'),
                    database, collections, depth=depth + 1,
                    by_research_areas=by_research_areas)
                driver = ctx_driver(driver)
                ok_any = ok_any or ok
            return ok_any

        splits = overflow_collection_splits(database, collections)
        if splits:
            print(
                f'{indent}{label} 单日命中 {hits} > {WOS_MAX_HITS}，'
                f'按 {len(splits)} 个引文索引/合集拆分')
            ok_any = False
            for db_name, col_name in splits:
                print(f'{indent}  -> {db_name} / {col_name}')
                ok = run_one_range_export(
                    ctx_driver(driver), save_dir, start_date, end_date,
                    db_name, col_name, depth=depth + 1,
                    by_research_areas=by_research_areas)
                driver = ctx_driver(driver)
                ok_any = ok_any or ok
            return ok_any

        print(
            f'{indent}{label} 单日命中 {hits} > {WOS_MAX_HITS}，'
            f'无法再切分，先导出前 {WOS_MAX_HITS} 条')
        record_export_failure(
            save_dir, f'{label}_partial_{WOS_MAX_HITS}_of_{hits}')
        hits = WOS_MAX_HITS

    if by_research_areas:
        ra_result = export_range_by_research_areas(
            ctx_driver(driver), save_dir, label, start_date, end_date,
            database, collections, indent=indent)
        driver = ctx_driver(driver)
        if ra_result is not None:
            return ra_result
        # 列表失败时回退整段导出（并截断超限）
        if hits > WOS_MAX_HITS:
            hits = WOS_MAX_HITS

    col_bit = collections_label(collections)
    prefix = f'SCI_{label}_{_safe_dirname(SEARCH_DATABASE)}_{_safe_dirname(col_bit)}'

    def restore_results():
        print(f'{indent}结果页丢失，重新检索 {label}')
        hits2 = run_search_once(
            ctx_driver(driver), start_date, end_date, database, collections,
            download_dir=save_dir)
        if hits2 is None:
            return False
        if hits2 <= 0:
            print(f'{indent}重新检索无结果')
            return False
        if hits2 != hits:
            print(f'{indent}重新检索命中 {hits2}（原 {hits}），继续按原区间导出')
        return True

    n = export_all_hits(
        ctx_driver(driver), save_dir, hits, prefix, restore_fn=restore_results)
    print(f'{indent}{label} 完成，成功批次 {n}')
    return n > 0


# ---------------------------------------------------------------------------
# 任务配置与主循环
# ---------------------------------------------------------------------------
def prompt_by_research_areas(default=None):
    if default is None:
        default = EXPORT_BY_RESEARCH_AREAS
    default_key = 'Y' if default else 'N'
    while True:
        raw = input(
            f'是否按 Research Areas 分类精炼导出? [Y/n] '
            f'(直接回车默认 {default_key}): '
        ).strip().lower()
        if not raw:
            return bool(default)
        if raw in ('y', 'yes', '1', 'true', '是'):
            return True
        if raw in ('n', 'no', '0', 'false', '否'):
            return False
        print('请输入 Y 或 N')


def prompt_crawl_granularity(default=None):
    default = default or CRAWL_GRANULARITY
    options = [
        ('1', 'day', '按天'),
        ('2', 'month', '按月'),
        ('3', 'year', '按年'),
        ('4', 'all', '总（整段一次）'),
    ]
    default_key = '1'
    for key, value, _ in options:
        if value == default:
            default_key = key
            break
    print('\n检索粒度：')
    for key, _value, label in options:
        mark = ' (默认)' if key == default_key else ''
        print(f'  {key}. {label}{mark}')
    while True:
        choice = input(
            f'请选择检索粒度编号（直接回车用默认 {default_key}）: ').strip()
        if not choice:
            choice = default_key
        for key, value, label in options:
            if choice == key or choice == value or choice == label:
                print(f'已选择检索粒度: {label}')
                return value
        print('输入无效，请重新选择')


def prompt_task_date_range(default_bg=None, default_end=None):
    default_bg = default_bg or bgdate
    default_end = default_end or initdate
    print(f'\n日期范围默认: {default_bg} ~ {default_end}')
    raw_bg = input('起始日期 YYYY-MM-DD（回车用默认）: ').strip()
    raw_end = input('结束日期 YYYY-MM-DD（回车用默认）: ').strip()
    start_s = raw_bg or default_bg
    end_s = raw_end or default_end
    try:
        datetime.strptime(start_s, '%Y-%m-%d')
        datetime.strptime(end_s, '%Y-%m-%d')
    except ValueError:
        print('日期格式无效，改用默认全局日期')
        return default_bg, default_end
    if start_s > end_s:
        print('起始晚于结束，已交换')
        start_s, end_s = end_s, start_s
    return start_s, end_s


def normalize_export_task(task, index=None):
    if not isinstance(task, dict):
        raise ValueError(f'任务格式错误: {task!r}')
    start_s = task.get('bgdate') or bgdate
    end_s = task.get('initdate') or initdate
    datetime.strptime(start_s, '%Y-%m-%d')
    datetime.strptime(end_s, '%Y-%m-%d')

    db_path = task.get('db_path') or task.get('category') or task.get('category_path')
    if isinstance(db_path, str):
        db_path = [
            p.strip() for p in db_path.replace('>', '/').split('/') if p.strip()]

    database = task.get('database')
    collections = task.get('collections')
    if db_path and isinstance(db_path, list):
        database = database or db_path[0]
        if collections is None:
            collections = db_path[1] if len(db_path) > 1 else SEARCH_COLLECTIONS
            if len(db_path) > 2:
                collections = list(db_path[1:])

    # 旧任务若把「核心合集」写在 database，改记为合集勾选
    db_as_col = _canonical_collection(database)
    if database not in (None, '', SEARCH_DATABASE, 'All Databases', '所有数据库'):
        wanted = _normalize_wanted_collections(collections)
        if wanted is None and db_as_col and db_as_col != '全选':
            collections = [db_as_col]
        elif wanted:
            collections = wanted
        else:
            collections = SEARCH_COLLECTIONS
    collections = _normalize_wanted_collections(collections)
    if collections is None:
        collections = SEARCH_COLLECTIONS
    database = SEARCH_DATABASE
    db_path = resolve_db_path(database, collections)

    return {
        'name': task.get('name') or (
            f'任务{index}' if index is not None else '任务'),
        'granularity': task.get('granularity') or CRAWL_GRANULARITY,
        'bgdate': start_s,
        'initdate': end_s,
        'database': database,
        'collections': collections,
        'db_path': list(db_path),
        'by_research_areas': bool(
            task['by_research_areas']
            if 'by_research_areas' in task
            else EXPORT_BY_RESEARCH_AREAS),
    }


def print_export_tasks(tasks):
    print(f'\n======== SCI 导出任务列表（共 {len(tasks)} 个）========')
    for i, task in enumerate(tasks, 1):
        db_label = ' / '.join(task.get('db_path') or [
            task['database'], str(task['collections'])])
        ra = 'RA分类' if task.get('by_research_areas') else '整段'
        print(
            f'  {i}. [{task.get("name", f"任务{i}")}] '
            f'{granularity_label(task["granularity"])} | '
            f'{task["bgdate"]} ~ {task["initdate"]} | '
            f'{db_label} | {ra}'
        )
    print('================================================')


def prompt_export_tasks():
    tasks = []
    print('\n开始配置 SCI 导出任务（可添加多个 合集勾选 + 日期/粒度 组合）')
    while True:
        print(f'\n----- 添加第 {len(tasks) + 1} 个任务 -----')
        db_path = prompt_database_collection()
        granularity = prompt_crawl_granularity(CRAWL_GRANULARITY)
        start_s, end_s = prompt_task_date_range(bgdate, initdate)
        by_ra = prompt_by_research_areas(EXPORT_BY_RESEARCH_AREAS)
        task = normalize_export_task({
            'name': f'任务{len(tasks) + 1}',
            'db_path': db_path,
            'granularity': granularity,
            'bgdate': start_s,
            'initdate': end_s,
            'by_research_areas': by_ra,
        }, index=len(tasks) + 1)
        tasks.append(task)
        print(
            f'已加入: {" / ".join(task["db_path"])} | '
            f'{granularity_label(task["granularity"])} | '
            f'{task["bgdate"]}~{task["initdate"]} | '
            f'{"按Research Areas分类" if task["by_research_areas"] else "整段导出"}'
        )
        while True:
            cont = input(
                '\n操作: [Enter]继续添加  d=删除上一个  l=查看列表  s=开始执行: '
            ).strip().lower()
            if cont in ('', 'y', 'yes', 'a', 'add'):
                break
            if cont in ('s', 'start', 'go', 'q'):
                if not tasks:
                    print('至少需要 1 个任务')
                    continue
                print_export_tasks(tasks)
                confirm = input('确认按以上任务列表开始? [Y/n]: ').strip().lower()
                if confirm in ('', 'y', 'yes'):
                    return tasks
                continue
            if cont in ('d', 'del', 'delete'):
                if tasks:
                    removed = tasks.pop()
                    print(f'已删除: {" / ".join(removed["db_path"])}')
                else:
                    print('列表为空')
                if not tasks:
                    break
                continue
            if cont in ('l', 'list'):
                if tasks:
                    print_export_tasks(tasks)
                else:
                    print('列表为空')
                continue
            print('无效输入')
    return tasks


def load_export_tasks():
    if EXPORT_TASKS:
        tasks = [
            normalize_export_task(t, index=i)
            for i, t in enumerate(EXPORT_TASKS, 1)
        ]
        print('使用配置中的 EXPORT_TASKS')
        print_export_tasks(tasks)
        return tasks
    return prompt_export_tasks()


def run_export_task(driver, task, task_index, task_total):
    """执行单个导出任务。返回 driver。"""
    granularity = task['granularity']
    start = datetime.strptime(task['bgdate'], '%Y-%m-%d').date()
    end = datetime.strptime(task['initdate'], '%Y-%m-%d').date()
    ranges = list(iter_crawl_date_ranges(start, end, granularity))
    database = task['database']
    collections = task['collections']
    db_path = task.get('db_path') or resolve_db_path(database, collections)
    by_ra = bool(task.get('by_research_areas'))

    download_dir = os.path.join(
        DOWNLOAD_ROOT,
        db_path_to_dirname(db_path),
        f'{task["bgdate"]}_{task["initdate"]}_{granularity}')
    os.makedirs(download_dir, exist_ok=True)

    print(
        f'\n######## 任务 {task_index}/{task_total}: {task.get("name", "")} ########\n'
        f'合集: {" / ".join(db_path)}\n'
        f'粒度: {granularity_label(granularity)} | '
        f'日期: {task["bgdate"]} ~ {task["initdate"]} | 共 {len(ranges)} 段\n'
        f'导出: {"按 Research Areas 分类" if by_ra else "整段"}\n'
        f'保存: {download_dir}'
    )

    if is_driver_alive(driver):
        driver = restart_browser_session(
            driver, download_dir, database, collections,
            reason='新任务，重启浏览器')
    else:
        driver = start_browser_session(download_dir, database, collections)
    bind_browser_ctx(driver, download_dir, database, collections)
    if not is_driver_alive(driver) or not is_wos_logged_in(driver):
        print('首次登录失败')
        return driver

    for idx, (rs, re_) in enumerate(ranges, 1):
        start_str = rs.strftime('%Y-%m-%d')
        end_str = re_.strftime('%Y-%m-%d')
        label = start_str if start_str == end_str else f'{start_str}_{end_str}'
        print(f'\n----- 任务{task_index} 第 {idx}/{len(ranges)} 段: {label} -----')
        try:
            if idx > 1:
                driver = restart_browser_session(
                    driver, download_dir, database, collections,
                    reason=(
                        f'上一段已完成，重启浏览器后开始第 {idx}/{len(ranges)} 段'
                    ))
            if not is_driver_alive(driver):
                print('浏览器未就绪，重新启动')
                driver = start_browser_session(
                    download_dir, database, collections)
            if is_page_unusable(driver) or not is_wos_logged_in(driver):
                print('页面失效或掉线，重启浏览器')
                driver = restart_browser_session(
                    driver, download_dir, database, collections,
                    reason='页面失效，重启浏览器')
            bind_browser_ctx(driver, download_dir, database, collections)
            if not is_driver_alive(driver) or not ensure_logged_in(driver):
                print('登录失效且重登失败，记录并跳过本段')
                record_export_failure(download_dir, label)
                continue
            run_one_range_export(
                ctx_driver(driver), download_dir, start_str, end_str,
                database, collections, by_research_areas=by_ra)
            # 检索遇 Server.unexpectedError 可能已换新浏览器
            driver = ctx_driver(driver)
        except Exception as e:
            print(f'任务段 {label} 异常:', e)
            record_export_failure(download_dir, label)
            driver = ctx_driver(driver)
            try:
                close_export_overlay(driver)
            except Exception:
                pass

    return ctx_driver(driver)


def scipaperfordata():
    """SCI 多任务检索导出入口。"""
    tasks = load_export_tasks()
    if not tasks:
        print('无导出任务，结束')
        return

    driver = None
    try:
        for i, task in enumerate(tasks, 1):
            try:
                driver = run_export_task(driver, task, i, len(tasks))
            except Exception as e:
                print(f'任务 {i} 异常（将继续下一任务）:', e)
                try:
                    quit_chrome_driver(driver)
                except Exception:
                    pass
                driver = None
            print(f'\n######## 任务 {i}/{len(tasks)} 完成 ########')
    finally:
        quit_chrome_driver(driver)
    print('\n全部 SCI 导出任务已结束')


if __name__ == '__main__':
    scipaperfordata()
