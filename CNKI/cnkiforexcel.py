'''
CNKI论文题录下载

基本设置与cnki.py相同。验证码（滑块 / 文字点选）逻辑见 cnki_captcha.py。
'''
from selenium import webdriver
import re
import time
import random
import os
import json
import shutil
import tempfile
from datetime import datetime, timedelta
from collections import Counter
from pathlib import Path
import sys
from urllib.parse import urlparse
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from cnki_captcha import (
    captcha_puzzle_fingerprint,
    click_verify_done_close,
    get_captcha_type,
    has_block_puzzle,
    has_new_download_file,
    has_visible_captcha_widget,
    is_cnki_verify_url,
    is_tencent_rate_limited,
    preload_click_captcha,
    refresh_captcha,
    snapshot_download_names,
    solve_captcha,
    tencent_last_fail,
    wait_for_any_captcha_widget,
    wait_for_block_puzzle,
    wait_for_tencent_slider,
    wait_for_text_click_captcha,
)

global initdata
bgdate = '2023-01-01'
initdate = '2026-06-30'

ADV_SEARCH_URL = (
    'https://kns.cnki.net/kns8/AdvSearch?dbprefix=SCDB'
    '&&crossDbcodes=CJFQ%2CCDMD%2CCIPD%2CCCND%2CCISD%2CSNAD%2CBDZK%2CCCJD%2CCCVD%2CCJFN%2CCCJD'
)
# 专利高级检索（文献分类 + 日期；结果页筛 PATENT）
PATENT_CLASSID = 'VUDIXAIY'
ADV_SEARCH_URL_PATENT = (
    f'https://kns.cnki.net/kns8s/AdvSearch?classid={PATENT_CLASSID}&language=CHS'
)
DOWNLOAD_ROOT = r'D:\2026年工作\知网数据库下载'
# 检索粒度：'day'=按天 | 'month'=按月 | 'year'=按年 | 'all'=整段一次
# 不足一个月/一年的尾段自动截到 initdate
CRAWL_GRANULARITY = 'all'
# 下载类型：'zh'=中文 | 'en'=英文 | 'none'=不选（跳过中/外文筛选）
DOWNLOAD_LANG = 'zh'
# 导出文献类型默认：'journal'=学术期刊 | 'patent'=专利 | 'none'=不选（跳过资源类型筛选）
DOWNLOAD_DOC_TYPE = 'journal'
# 可选默认主题词（空=不限制；交互配置时作为默认提示，任务可覆盖）
DEFAULT_THEME = ''
# 专利时间字段默认：'AD'=申请日(datebox0/1) | 'OD'=公开日(datebox2/3)
DEFAULT_PATENT_DATE_FIELD = 'AD'

# 与 SCI 并行：独立 Chrome 用户目录 + 调试端口（SCI 使用 9223 段）
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CHROME_PROFILE_ROOT = os.path.join(_SCRIPT_DIR, '.chrome_profiles', 'cnki')
CHROME_DEBUG_PORT_BASE = 9222
_chrome_profile_dir = None

# 机构登录方式默认值（可被 cnki_unit_credentials.json 的 method 覆盖）
# 'vpn' = 学校 WebVPN | 'ip' = IP登录 | 'school_visit' = 校外访问
UNIT_LOGIN_METHOD = 'vpn'
UNIT_CREDENTIALS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'cnki_unit_credentials.json')

WEBVPN_ORIGIN = 'https://webvpn.uibe.edu.cn'
WEBVPN_LOGIN_URL = WEBVPN_ORIGIN + '/login'
WEBVPN_MAGIC = '77726476706e69737468656265737421'
WEBVPN_AES_KEY = b'wrdvpnisthebest!'
# 已知主机编码（与 AES-CFB 一致，无 pycryptodome 时也可改写常用知网域名）
WEBVPN_HOST_ENC = {
    'www.cnki.net': 'e7e056d2243e635930068cb8',
    'kns.cnki.net': 'fbf952d2243e635930068cb8',
    'link.cnki.net': 'fcfe4f976933665b774687a98c',
    'bar.cnki.net': 'f2f653d2243e635930068cb8',
    'nav.cnki.net': 'fef657d2243e635930068cb8',
    'kcms.cnki.net': 'fbf44c8f6933665b774687a98c',
    'kns8.cnki.net': 'fbf952c46933665b774687a98c',
    'fsso.cnki.net': 'f6e452936933665b774687a98c',
}
WEBVPN_ADV_SEARCH_PATH = '/kns8s/AdvSearch'

# 预定义导出任务列表（可选）。为空则启动时交互配置多个任务。
# 每项字段：
#   category   : 期刊=学科导航如 ['基础科学','数学']；专利=IPC 如 ['G','G06']；
#                doc_type=none 时不选分类，可省略或 []
#   doc_type   : journal / patent / none（不选=跳过分类勾选+结果页资源类型筛选；可省略）
#   theme      : 可选主题词（专利写入「主题」行，默认第2行已是 SU）
#   patent_date: 专利时间字段 AD=申请日 / OD=公开日（默认 AD）
#   lang       : zh / en / none（不选=跳过中/外文筛选；专利任务忽略）
#   granularity / bgdate / initdate : 同前
EXPORT_TASKS = [
    # {
    #     'category': ['基础科学', '数学', '数学分析', '泛函分析'],
    #     'doc_type': 'journal',
    #     'theme': '',
    #     'lang': 'zh',
    #     'granularity': 'day',
    #     'bgdate': '2023-01-01',
    #     'initdate': '2026-06-30',
    # },
    # {
    #     'category': ['G', 'G06'],
    #     'doc_type': 'patent',
    #     'theme': '人工智能',
    #     'patent_date': 'AD',
    #     'granularity': 'year',
    #     'bgdate': '2020-01-01',
    #     'initdate': '2024-12-31',
    # },
]

# 文献分类树：空 dict 表示可选叶子；非空可继续细分（对应页面点 i.icon 展开）
LITERATURE_CATEGORY_TREE = {
    '基础科学': {
        '自然科学理论与方法': {
            '自然科学研究的方针政策': {},
            '自然科学理论与方法论': {
                '科学的哲学原理': {},
                '科学的方法论': {},
            },
            '自然科学现状及发展': {},
            '自然科学研究及自然科学史': {},
        },
        '数学': {
            '数学概论': {},
            '古典数学': {},
            '初、高等数学': {},
            '数理逻辑、数学基础': {
                '总论': {},
                '数理逻辑(符号逻辑)': {},
                '应用数理逻辑': {},
                '数学基础': {},
                '集合论': {},
            },
            '代数、数论、组合理论': {
                '总论': {},
                '代数方程论、线性代数': {},
                '群论': {},
                '抽象代数(近世代数)': {},
                '范畴论、同调代数': {},
                '微分代数、差分代数': {},
                '数论': {},
                '组合数学': {},
                '离散数学': {},
                '模糊数学': {},
            },
            '数学分析': {
                '数学分析与应用': {},
                '分析基础': {},
                '微积分': {},
                '无穷级数论(级数论)': {},
                '函数论': {},
                '微分方程、积分方程': {},
                '变分法': {},
                '泛函分析': {},
                '不等式及其他': {},
            },
            '几何、拓扑': {
                '概述': {},
                '几何基础(几何学原理)': {},
                '解析几何': {},
                '向量(矢量)和张量分析': {},
                '非欧几何、多维空间几何': {},
                '射影(投影)几何、画法几何': {},
                '微分几何、积分几何': {},
                '代数几何': {},
                '拓扑(形势几何学)': {},
            },
            '动力系统理论': {},
            '概率论 、数理统计': {},
            '运筹学': {},
            '控制论、信息论(数学理论)': {},
            '计算数学': {},
            '应用数学': {},
        },
        '非线性科学与系统科学': {},
        '力学': {},
        '物理学': {
            '物理学总论': {},
            '物理学现状、概况': {},
            '物理学实验方法与设备': {},
            '物理测量': {},
            '理论物理学': {},
            '声学': {},
            '光学': {},
            '电磁学、电动力学': {},
            '无线电物理学': {},
            '真空电子学(电子物理学)': {},
            '半导体物理学': {},
            '固体物理学': {},
            '低温物理学': {},
            '高压与高温物理学': {},
            '等离子体物理学': {},
            '热学与物质分子运动论': {},
            '分子物理学、原子物理学': {},
            '原子核物理学、高能物理学': {},
            '应用物理学': {},
        },
        '生物学': {
            '生物科学总论': {},
            '普通生物学': {},
            '细胞学': {},
            '遗传学': {},
            '生理学': {},
            '生物化学': {},
            '生物物理学': {},
            '分子生物学及遗传工程': {},
            '生物工程学': {},
            '环境生物学': {},
            '古生物学': {},
            '微生物学理论及其应用': {},
            '植物学': {},
            '动物学': {},
            '昆虫学': {},
            '人类学': {},
        },
        '天文学': {},
        '自然地理学和测绘学': {},
        '气象学': {},
        '海洋学': {},
        '地质学': {},
        '地球物理学': {},
        '资源科学': {},
    },
    '工程科技Ⅰ辑': {
        '化学': {
            '总论': {},
            '无机化学': {},
            '有机化学': {},
            '高分子化学(高聚物)': {},
            '物理化学(理论化学)、化学物理学': {},
            '分析化学': {},
            '应用化学': {},
            '晶体学': {},
        },
        '无机化工': {},
        '有机化工': {},
        '燃料化工': {},
        '一般化学工业': {},
        '石油天然气工业': {},
        '材料科学': {},
        '矿业工程': {},
        '金属学及金属工艺': {},
        '冶金工业': {},
        '轻工业手工业': {},
        '一般服务业': {},
        '安全科学与灾害防治': {},
        '环境科学与资源利用': {},
    },
    '工程科技Ⅱ辑': {
        '工业通用技术及设备': {},
        '机械工业': {},
        '仪器仪表工业': {},
        '航空航天科学与工程': {},
        '武器工业与军事技术': {},
        '铁路运输': {},
        '公路与水路运输': {},
        '汽车工业': {},
        '船舶工业': {},
        '水利水电工程': {},
        '建筑科学与工程': {},
        '动力工程': {},
        '核科学技术': {},
        '新能源': {},
        '电力工业': {},
    },
    '农业科技': {
        '农业基础科学': {},
        '农业工程': {},
        '农艺学': {},
        '植物保护': {},
        '农作物': {},
        '园艺': {},
        '林业': {},
        '畜牧与动物医学': {},
        '蚕蜂与野生动物保护': {},
        '水产和渔业': {},
    },
    '医药卫生科技': {
        '医药卫生方针政策与法律法规研究': {
            '卫生事业总论': {},
            '医药、卫生政策法规': {},
            '医疗保险': {},
            '医务人员及卫生宣传教育': {},
            '卫生标准与检查、医药管理': {},
            '卫生调查统计': {},
            '医疗服务、制度及改革': {},
            '医疗与法': {},
            '医疗卫生机构及管理': {},
            '医疗器械': {},
        },
        '医学教育与医学边缘学科': {},
        '预防医学与卫生学': {},
        '中医学': {},
        '中药学': {},
        '中西医结合': {},
        '基础医学': {
            '实验医学': {},
            '医用一般科学（医用数学、物理学、化学等）': {},
            '人体形态学': {},
            '人体生理学': {},
            '医学生物化学与分子生物学': {},
            '医学生物物理学': {},
            '病理学': {},
            '医学微生物学': {},
            '医学寄生虫学': {},
            '医学免疫学': {},
            '医学遗传学': {},
            '医学心理学、病理心理学': {},
        },
        '临床医学': {
            '临床医学总论': {},
            '诊断与治疗技术': {},
            '护理学': {},
            '吸毒与戒毒': {},
            '疼痛医学': {},
            '睡眠医学': {},
            '康复医学': {},
            '临终关怀学': {},
            '其他诊疗技术': {},
        },
        '感染性疾病及传染病': {},
        '心血管系统疾病': {
            '总论': {},
            '诊断学': {},
            '心脏疾病': {},
            '血管疾病': {},
            '血压异常': {},
            '血液及淋巴系疾病': {},
        },
        '呼吸系统疾病': {},
        '消化系统疾病': {},
        '内分泌腺及全身性疾病': {
            '内分泌腺疾病及代谢病': {
                '总论': {},
                '甲状腺疾病': {},
                '甲状旁腺疾病': {},
                '胸腺疾病': {},
                '垂体及间脑－垂体系统疾病': {},
                '松果体疾病及性早熟': {},
                '肾上腺疾病': {},
                '胰岛疾病': {},
                '性腺疾病': {},
                '骨质疏松': {},
                '代谢病': {},
            },
            '全身性疾病': {},
        },
        '外科学': {
            '外科学总论': {},
            '外科学基础': {},
            '外科手术学': {},
            '外科感染': {},
            '创伤外科学': {},
            '外科学各论': {
                '头部及神经外科学': {},
                '颈部外科学': {},
                '心脏血管及淋巴系外科学': {},
                '胸部外科学': {},
                '乳腺外科学': {},
                '腹部外科学': {},
                '四肢外科学': {},
                '小儿外科学': {},
            },
            '骨科学': {},
            '整形外科学': {},
        },
        '泌尿科学': {
            '一般疾病': {},
            '肾疾病': {},
            '输尿管疾病': {},
            '膀胱疾病': {},
            '尿的成分异常': {},
            '尿道疾病': {},
            '男性生殖器疾病': {},
            '男子性功能障碍': {},
            '泌尿及男性生殖系外科学': {},
        },
        '妇产科学': {
            '基础与临床': {},
            '妇科学': {},
            '产科学': {},
            '计划生育': {},
            '临床优生学': {},
            '辅助生殖技术': {},
        },
        '儿科学': {
            '新生儿、早产儿疾病': {},
            '儿童的营养障碍及保健与卫生': {},
            '小儿内科学': {},
            '小儿外科学': {},
            '小儿泌尿科学': {},
            '儿童神经、精神病': {},
            '小儿其他疾病': {},
            '中医儿科': {},
            '其他': {},
        },
        '神经病学': {
            '基础与临床': {},
            '脑部疾病': {},
            '脑血管疾病': {},
            '脊髓疾病': {},
            '周围神经及神经节疾病': {},
            '神经肌肉疾病': {},
            '植物性神经系疾病、血管舒缩及营养神经系疾病': {},
            '并发神经病': {},
            '神经系肿瘤': {},
            '儿童神经病': {},
        },
        '精神病学': {
            '基础与临床': {},
            '器质性疾病伴发的精神障碍': {},
            '精神分裂症': {},
            '情感性精神障碍': {},
            '反应性及中毒所致精神障碍': {},
            '神经官能症': {},
            '其他': {},
        },
        '肿瘤学': {
            '肿瘤学基础': {},
            '心血管肿瘤': {},
            '造血器及淋巴系肿瘤': {},
            '呼吸系肿瘤': {},
            '消化系肿瘤': {
                '总论': {},
                '食管肿瘤': {},
                '胃肿瘤': {},
                '肠肿瘤': {},
                '肠系膜及腹膜后肿瘤': {},
                '腹膜及膜壁肿瘤': {},
                '肝肿瘤': {},
                '胆囊、胆道肿瘤': {},
                '胰腺肿瘤': {},
                '壶腹癌': {},
            },
            '内分泌腺肿瘤': {},
            '泌尿生殖器肿瘤': {},
            '乳腺肿瘤': {},
            '运动系肿瘤': {},
            '神经系肿瘤': {},
            '皮肤肿瘤': {},
            '五官肿瘤': {},
            '其他部位肿瘤': {},
        },
        '眼科与耳鼻咽喉科': {
            '总论': {},
            '耳科学、耳疾病': {},
            '鼻科学、鼻疾病': {},
            '咽科学、咽疾病': {},
            '喉科学、喉疾病': {},
            '气管与食管镜学': {},
            '眼科学': {},
        },
        '口腔科学': {},
        '皮肤病与性病': {},
        '特种医学': {},
        '急救医学': {},
        '军事医学与卫生': {},
        '药学': {
            '总论': {},
            '药物基础科学': {},
            '药典、药方集及药物鉴定': {},
            '药剂学': {},
            '药事组织': {},
            '药理学': {},
            '药品': {},
            '毒物学(毒理学)': {},
        },
        '生物医学工程': {},
    },
    '哲学与人文科学': {},
    '社会科学Ⅰ辑': {},
    '社会科学Ⅱ辑': {},
    '信息科技': {
        '无线电电子学': {},
        '电信技术': {},
        '计算机硬件技术': {},
        '计算机软件及计算机应用': {
            '计算机理论与方法': {},
            '安全保密': {},
            '计算机软件概况': {},
            '程序设计、软件工程': {},
            '程序语言、算法语言': {},
            '编译程序、解释程序': {},
            '管理程序、管理系统': {},
            '操作系统': {},
            '数据库理论及系统': {},
            '程序包(应用软件)': {},
            '专用应用程序': {},
            '计算机的应用': {
                '信息处理(信息加工)': {
                    '模式识别与装置': {},
                    '机器辅助技术': {},
                    '计算机仿真': {},
                    '翻译、检索、诊断、教学及其他': {},
                },
                '计算机在其他方面的应用': {},
            },
        },
        '互联网技术': {},
        '自动化技术': {},
        '新闻与传媒': {},
        '出版': {},
        '图书情报与数字图书馆': {},
        '档案及博物馆': {},
    },
    '经济与管理科学': {},
}

# 兼容旧名：仅一级分类列表
LITERATURE_CATEGORIES = list(LITERATURE_CATEGORY_TREE.keys())

# 专利默认「国际专利分类」#RJUTCQD5，不用学科导航 #CNKI168
PATENT_IPC_NAV_ID = 'RJUTCQD5'
JOURNAL_NAV_ID = 'CNKI168'

PATENT_IPC_LABELS = {
    'A': '人类生活必需', 'B': '作业；运输', 'C': '化学；冶金', 'D': '纺织；造纸',
    'E': '固定建筑物', 'F': '机械工程；照明；加热；武器；爆破', 'G': '物理', 'H': '电学',
    'A01': '农业；林业；畜牧业；狩猎；诱捕；捕鱼',
    'A01B': '农业或林业的整地；一般农业机械或农具的部件、零件或附件',
    'A01C': '种植；播种；施肥', 'A01D': '收获；割草',
    'A01F': '脱粒；禾秆、干草或类似物的打捆等', 'A01G': '园艺；栽培；林业；浇水',
    'A01H': '新植物或获得新植物的方法', 'A01J': '乳制品的加工',
    'A01K': '畜牧业；养鸟业；养蜂业；养鱼业等', 'A01L': '动物钉蹄铁',
    'A01M': '动物的捕捉、诱捕或惊吓；消灭有害动植物用的装置',
    'A01N': '人体、动植物体保存；杀生剂；害虫驱避剂等',
    'A01P': '化学化合物或制剂的杀生、害虫驱避等活性',
    'A21': '焙烤；制作或处理面团的设备', 'A22': '屠宰；肉品处理；家禽或鱼的加工',
    'A23': '其他类不包含的食品或食料；及其处理',
    'A24': '烟草；雪茄烟；纸烟；模拟吸烟装置', 'A41': '服装', 'A42': '帽类制品',
    'A43': '鞋类', 'A44': '服饰缝纫用品；珠宝', 'A45': '手携物品或旅行品',
    'A46': '刷类制品', 'A47': '家具；家庭用物品或设备等',
    'A61': '医学或兽医学；卫生学', 'A62': '救生；消防',
    'A63': '运动；游戏；娱乐活动', 'A99': '本部其他类目中不包括的技术主题',
    'B01': '一般的物理或化学的方法或装置', 'B02': '破碎、磨粉或粉碎',
    'B03': '分离固体物料', 'B04': '离心装置或离心机', 'B05': '喷射或雾化；涂覆',
    'B06': '一般机械振动的发生或传递', 'B07': '将固体从固体中分离；分选',
    'B08': '清洁', 'B09': '固体废物的处理；被污染土壤的再生',
    'B21': '基本上无切削的金属机械加工；金属冲压', 'B22': '铸造；粉末冶金',
    'B23': '机床；其他类目中不包括的金属加工', 'B24': '磨削；抛光',
    'B25': '手动工具；轻便机动工具；机械手', 'B26': '手动切割工具；切割；切断',
    'B27': '木材或类似材料的加工或保存', 'B28': '加工水泥、黏土或石料',
    'B29': '塑料的加工', 'B30': '压力机', 'B31': '纸品制作与加工',
    'B32': '层状产品', 'B33': '增材制造技术', 'B41': '印刷；排版机；打字机',
    'B42': '装订；图册；文件夹', 'B43': '书写或绘图器具；办公用品',
    'B44': '装饰艺术', 'B60': '一般车辆', 'B61': '铁路', 'B62': '无轨陆用车辆',
    'B63': '船舶或其他水上船只', 'B64': '飞行器；航空；宇宙航行',
    'B65': '输送；包装；贮存', 'B66': '卷扬；提升；牵引',
    'B67': '开启或封闭瓶子、罐等', 'B68': '鞍具；家具罩面',
    'B81': '微观结构技术', 'B82': '超微技术', 'B99': '本部其他类目中不包括的技术主题',
    'C01': '无机化学', 'C02': '水、废水、污水或污泥的处理', 'C03': '玻璃；矿棉或渣棉',
    'C04': '水泥；混凝土；人造石；陶瓷', 'C05': '肥料；肥料制造', 'C06': '炸药；火柴',
    'C07': '有机化学', 'C08': '有机高分子化合物', 'C09': '染料；涂料；黏合剂等',
    'C10': '石油、煤气及炼焦工业', 'C11': '动物或植物油、脂等',
    'C12': '生物化学；啤酒；微生物学；酶学等', 'C13': '糖工业',
    'C14': '小原皮；大原皮；毛皮或皮革', 'C21': '铁的冶金',
    'C22': '冶金；合金', 'C23': '金属材料的镀覆；表面化学处理等',
    'C25': '电解或电泳工艺', 'C30': '晶体生长', 'C40': '组合技术',
    'C99': '本部其他类目不包括的技术主题',
    'D01': '天然或化学的线或纤维；纺纱或纺丝', 'D02': '纱线；机械整理',
    'D03': '织造', 'D04': '编织；花边；针织；非织造布', 'D05': '缝纫；绣花；簇绒',
    'D06': '织物等的处理；洗涤', 'D07': '绳；除电缆以外的缆索',
    'D21': '造纸；纤维素的生产', 'D99': '本部其他类目不包括的技术主题',
    'E01': '道路、铁路或桥梁的建筑', 'E02': '水利工程；基础；疏浚',
    'E03': '给水；排水', 'E04': '建筑物', 'E05': '锁；钥匙；门窗零件；保险箱',
    'E06': '一般门、窗、百叶窗或卷辊遮帘；梯子', 'E21': '土层或岩石的钻进；采矿',
    'E99': '本部其他类目不包括的技术主题',
    'F01': '一般机器或发动机；蒸汽机', 'F02': '燃烧发动机',
    'F03': '液力机械或液力发动机；风力等发动机', 'F04': '液体变容式机械；泵',
    'F15': '流体压力执行机构；液压气动技术', 'F16': '工程元件或部件；绝热',
    'F17': '气体或液体的贮存或分配', 'F21': '照明', 'F22': '蒸汽的发生',
    'F23': '燃烧设备；燃烧方法', 'F24': '供热；炉灶；通风',
    'F25': '制冷或冷却；热泵系统等', 'F26': '干燥', 'F27': '炉；窑；烘烤炉',
    'F28': '一般热交换', 'F41': '武器', 'F42': '弹药；爆破',
    'F99': '本部其他类目不包括的技术主题',
    'G01': '测量；测试', 'G02': '光学', 'G03': '摄影术；电影术等',
    'G04': '测时学', 'G05': '控制；调节', 'G06': '计算；推算或计数',
    'G07': '核算装置', 'G08': '信号装置', 'G09': '教育；密码术；显示；广告',
    'G10': '乐器；声学', 'G11': '信息存储', 'G12': '仪器的结构零部件',
    'G16': '特别适用于特定应用领域的信息通信技术[ICT]',
    'G21': '核物理；核工程', 'G99': '不包含在本部其他类目中的技术主题',
    'H01': '基本电气元件', 'H02': '发电、变电或配电', 'H03': '基本电子电路',
    'H04': '电通信技术', 'H05': '其他类目不包含的电技术',
    'H99': '本部中其他类目不包括的技术主题',
}


def _ipc_leaf(*codes):
    return {c: {} for c in codes}


PATENT_IPC_CATEGORY_TREE = {
    'A': {
        'A01': _ipc_leaf(
            'A01B', 'A01C', 'A01D', 'A01F', 'A01G', 'A01H',
            'A01J', 'A01K', 'A01L', 'A01M', 'A01N', 'A01P'),
        **_ipc_leaf(
            'A21', 'A22', 'A23', 'A24', 'A41', 'A42', 'A43', 'A44',
            'A45', 'A46', 'A47', 'A61', 'A62', 'A63', 'A99'),
    },
    'B': _ipc_leaf(
        'B01', 'B02', 'B03', 'B04', 'B05', 'B06', 'B07', 'B08', 'B09',
        'B21', 'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28', 'B29',
        'B30', 'B31', 'B32', 'B33', 'B41', 'B42', 'B43', 'B44',
        'B60', 'B61', 'B62', 'B63', 'B64', 'B65', 'B66', 'B67', 'B68',
        'B81', 'B82', 'B99'),
    'C': _ipc_leaf(
        'C01', 'C02', 'C03', 'C04', 'C05', 'C06', 'C07', 'C08', 'C09',
        'C10', 'C11', 'C12', 'C13', 'C14', 'C21', 'C22', 'C23', 'C25',
        'C30', 'C40', 'C99'),
    'D': _ipc_leaf(
        'D01', 'D02', 'D03', 'D04', 'D05', 'D06', 'D07', 'D21', 'D99'),
    'E': _ipc_leaf(
        'E01', 'E02', 'E03', 'E04', 'E05', 'E06', 'E21', 'E99'),
    'F': _ipc_leaf(
        'F01', 'F02', 'F03', 'F04', 'F15', 'F16', 'F17',
        'F21', 'F22', 'F23', 'F24', 'F25', 'F26', 'F27', 'F28',
        'F41', 'F42', 'F99'),
    'G': _ipc_leaf(
        'G01', 'G02', 'G03', 'G04', 'G05', 'G06', 'G07', 'G08', 'G09',
        'G10', 'G11', 'G12', 'G16', 'G21', 'G99'),
    'H': _ipc_leaf('H01', 'H02', 'H03', 'H04', 'H05', 'H99'),
}


def get_category_tree(doc_type=None):
    if normalize_doc_type(doc_type) == 'patent':
        return PATENT_IPC_CATEGORY_TREE
    return LITERATURE_CATEGORY_TREE


def category_nav_root_id(doc_type=None):
    if normalize_doc_type(doc_type) == 'patent':
        return PATENT_IPC_NAV_ID
    return JOURNAL_NAV_ID


def ipc_display_name(code):
    code = (code or '').strip()
    label = PATENT_IPC_LABELS.get(code, '')
    return f'{code} {label}'.strip() if label else code


def normalize_patent_date_field(value):
    """专利时间字段：AD=申请日，OD=公开日。"""
    v = (value or DEFAULT_PATENT_DATE_FIELD or 'AD').strip().upper()
    if v in ('OD', '公开日', 'GK', 'PUB', 'PUBLICATION', 'OPEN'):
        return 'OD'
    return 'AD'


def patent_date_field_label(value):
    return '公开日' if normalize_patent_date_field(value) == 'OD' else '申请日'


def has_captcha_anywhere(driver):
    """检查当前窗口及所有标签页是否存在验证码。"""
    current = driver.current_window_handle
    found = False
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if has_block_puzzle(driver):
            found = True
            break
    if current in driver.window_handles:
        driver.switch_to.window(current)
    return found


def _captcha_max_retries_for_type(captcha_type):
    """点选允许多试；滑块连败后应尽快重启浏览器。"""
    if captcha_type == 'text_click':
        return 10
    return 3


def handle_captcha(
    driver, save_dir, main_handle=None, max_retries=3, return_handle=None,
    download_dir=None, before_files=None,
):
    """检测并处理验证码（滑块 / 文字点选）。

    返回值：
        None — 未检测到验证码
        True — 检测到验证码且已通过
        False — 检测到验证码但未通过

    腾讯滑块通过后弹窗不关，以下载目录出现新文件作为通过依据。
    """
    original_handle = driver.current_window_handle
    captcha_handle = None

    if has_block_puzzle(driver):
        captcha_handle = original_handle
    else:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if has_block_puzzle(driver):
                captcha_handle = handle
                break

    if not captcha_handle:
        restore_handle = return_handle or original_handle
        if restore_handle in driver.window_handles:
            driver.switch_to.window(restore_handle)
        elif main_handle and main_handle in driver.window_handles:
            driver.switch_to.window(main_handle)
        return None

    try:
        cur_u = driver.current_url or ''
    except Exception:
        cur_u = ''
    if is_cnki_verify_url(cur_u):
        print('验证码页已打开，等待控件加载...')
        wait_for_any_captcha_widget(driver, timeout=12)

    captcha_type = get_captcha_type(driver)
    # 切到验证码窗后再定次数：点选 10，滑块 3（调用方传入更小值时抬到类型默认）
    type_retries = _captcha_max_retries_for_type(captcha_type)
    if captcha_type == 'text_click':
        max_retries = max(int(max_retries or 0), type_retries)
    elif max_retries is None or max_retries < 1:
        max_retries = type_retries
    type_label = {
        'text_click': '文字点选',
        'tencent_slider': '腾讯滑块 tencent-captcha-dy',
        'slider': 'blockPuzzle 滑块',
    }.get(captcha_type, captcha_type or '未知')
    print(f'检测到 {type_label} 验证码，开始自动处理（最多 {max_retries} 次）...')
    debug_dir = os.path.join(save_dir, 'captcha_debug')
    expect_download = before_files is not None
    dl_dir = download_dir or save_dir
    before = set(before_files) if expect_download else snapshot_download_names(dl_dir)
    ok = False
    stale_puzzle = False
    last_fp = None

    for attempt in range(1, max_retries + 1):
        print(f'验证码处理第 {attempt}/{max_retries} 次')
        if is_tencent_rate_limited(driver):
            print('腾讯提示过于频繁，本次不滑动')
            break
        if expect_download and has_new_download_file(dl_dir, before):
            print('验证码弹窗仍在，但已检测到下载文件，判定通过')
            ok = True
            break
        fp_before = captcha_puzzle_fingerprint(driver)
        # 上一轮未真正换题时，再拖同一滑块会一直失败 → 直接重启浏览器
        if (
            last_fp is not None
            and fp_before is not None
            and fp_before == last_fp
            and get_captcha_type(driver) != 'tencent_slider'
        ):
            print('题面与上次相同（未刷新），停止重复验证，改为重启浏览器')
            stale_puzzle = True
            break
        attempt_dir = os.path.join(debug_dir, f'attempt_{attempt}')
        ok = solve_captcha(
            driver, save_dir=attempt_dir,
            download_dir=dl_dir if expect_download else None,
            before_files=before if expect_download else None)
        if ok:
            print('验证码处理成功')
            break
        if expect_download and has_new_download_file(dl_dir, before):
            print('验证码弹窗仍在，但已检测到下载文件，判定通过')
            ok = True
            break
        # 滑块实际已通过但成功提示未识别时，控件会消失；不要把它当成失败
        time.sleep(0.6)
        if not has_visible_captcha_widget(driver):
            print('验证码控件已消失，按已通过处理')
            ok = True
            break
        last_fp = fp_before
        if attempt < max_retries:
            if tencent_last_fail() == 'rate_limit' or is_tencent_rate_limited(driver):
                print('腾讯提示过于频繁，本轮不再滑动')
                break
            ctype = get_captcha_type(driver)
            if not ctype:
                if expect_download and has_new_download_file(dl_dir, before):
                    print('验证码弹窗仍在，但已检测到下载文件，判定通过')
                    ok = True
                else:
                    print('验证码控件已不可交互，停止重试')
                break
            if ctype == 'tencent_slider':
                time.sleep(random.uniform(2.2, 3.6))
            else:
                # blockPuzzle / 点选：失败后必须强制换图；未换图则立即重启
                print('验证码失败，强制刷新换题后再试...')
            refreshed = refresh_captcha(
                driver,
                force=True if ctype != 'tencent_slider'
                else (tencent_last_fail() == 'no_gap'))
            fp_after = captcha_puzzle_fingerprint(driver)
            same_puzzle = (
                ctype != 'tencent_slider'
                and (
                    refreshed is False
                    or (fp_before is not None and fp_after == fp_before)
                    or (fp_before is None and fp_after is None)
                )
            )
            if same_puzzle:
                print('验证码未刷新（题面相同），停止重复验证，改为重启浏览器')
                stale_puzzle = True
                break
            if ctype != 'tencent_slider' and fp_after is not None:
                last_fp = None  # 已换题，下一轮允许再试
            try:
                if ctype == 'text_click':
                    wait_for_text_click_captcha(driver)
                elif ctype == 'tencent_slider':
                    wait_for_tencent_slider(driver)
                else:
                    wait_for_block_puzzle(driver)
            except Exception:
                if expect_download and has_new_download_file(dl_dir, before):
                    print('验证码弹窗仍在，但已检测到下载文件，判定通过')
                    ok = True
                break

    if not ok:
        if stale_puzzle:
            print('验证码题面未刷新，准备关闭浏览器后重试')
        else:
            print(f'验证码连续失败 {max_retries} 次，准备关闭浏览器后重试')
        time.sleep(0.6)

    restore_handle = return_handle or captcha_handle or original_handle
    if (
        ok
        and captcha_handle != restore_handle
        and captcha_handle != main_handle
        and captcha_handle in driver.window_handles
        and len(driver.window_handles) > 1
    ):
        try:
            driver.switch_to.window(captcha_handle)
            cur = driver.current_url or ''
            if (
                'verify/home' in cur
                or '/bar/verify' in cur
                or is_cnki_verify_done_page(driver)
            ):
                click_verify_done_close(driver)
                if captcha_handle in driver.window_handles and len(driver.window_handles) > 1:
                    driver.close()
        except Exception:
            pass

    if restore_handle in driver.window_handles:
        driver.switch_to.window(restore_handle)
    elif main_handle and main_handle in driver.window_handles:
        driver.switch_to.window(main_handle)
    elif original_handle in driver.window_handles:
        driver.switch_to.window(original_handle)

    return ok


def captcha_needs_browser_restart(result):
    """ensure_captcha_passed / ensure_cnki_captcha_after_nav 是否要求关浏览器。"""
    return result == 'restart_browser'


def ensure_captcha_passed(
    driver, save_dir, main_handle=None, return_handle=None, max_rounds=None,
    download_dir=None, before_files=None,
):
    """处理验证码；按类型重试，失败则要求重启浏览器。

    返回:
      True              — 已通过（或无验证码）
      False             — 失败（非重启场景）
      'restart_browser' — 连续失败或题面未刷新，需关浏览器重建

    说明：
      - 文字点选 text_click：最多 10 次，失败后强制换题，满 10 次前不进入等待重启
      - blockPuzzle / 腾讯滑块：最多 3 次，未换题或连败即 restart_browser
    max_rounds 已废弃，保留参数仅为兼容旧调用。
    """
    if max_rounds is not None and max_rounds != 1:
        print(f'忽略 max_rounds={max_rounds}，验证码失败后按类型重试或重启浏览器')
    if not has_captcha_anywhere(driver):
        return True
    if before_files is not None and has_new_download_file(
            download_dir or save_dir, before_files):
        print('验证码弹窗仍在，但已检测到下载文件，判定通过')
        return True
    if is_tencent_rate_limited(driver):
        print('腾讯仍提示过于频繁，停止滑动，准备关闭浏览器后重新认证')
        return 'restart_browser'
    ctype = get_captcha_type(driver)
    max_retries = _captcha_max_retries_for_type(ctype)
    if ctype == 'text_click':
        print(f'验证码处理：文字点选最多 {max_retries} 次，失败后强制换题')
    else:
        print(
            f'验证码处理：失败后强制换题，最多 {max_retries} 次；'
            '未换题或连败则重启浏览器')
    result = handle_captcha(
        driver, save_dir, main_handle,
        max_retries=max_retries, return_handle=return_handle,
        download_dir=download_dir, before_files=before_files)
    if result is True or result is None:
        return True
    if tencent_last_fail() == 'rate_limit' or is_tencent_rate_limited(driver):
        print('腾讯提示过于频繁，停止滑动，准备关闭浏览器后重新认证')
        return 'restart_browser'
    if before_files is not None and has_new_download_file(
            download_dir or save_dir, before_files):
        print('验证码弹窗仍在，但已检测到下载文件，判定通过')
        return True
    print(
        f'验证码失败（连败 {max_retries} 次或题面未刷新），'
        '改为关闭浏览器后重试')
    return 'restart_browser'


def handle_block_puzzle_captcha(driver, save_dir, main_handle=None, max_retries=3, return_handle=None):
    """兼容旧接口：检测到验证码时返回 True（含未通过），未检测到返回 False。"""
    result = handle_captcha(driver, save_dir, main_handle, max_retries, return_handle)
    return result is not None


def snapshot_excel_files(save_dir):
    """记录导出目录中已有的 Excel 文件。"""
    if not os.path.isdir(save_dir):
        return set()
    return {
        f for f in os.listdir(save_dir)
        if f.lower().endswith(('.xls', '.xlsx')) and not f.endswith('.crdownload')
    }


def get_new_excel_file(save_dir, before_files):
    """返回导出后新生成的 Excel 文件名，无则返回 None。"""
    if not os.path.isdir(save_dir):
        return None
    for name in os.listdir(save_dir):
        lower = name.lower()
        if not lower.endswith(('.xls', '.xlsx')):
            continue
        if name.endswith('.crdownload') or name.endswith('.tmp'):
            continue
        if name not in before_files:
            full = os.path.join(save_dir, name)
            if os.path.getsize(full) > 0:
                return name
    return None


def _is_element_display_block(el):
    """判断元素 style/computed display 是否为 block（可见登录框）。"""
    try:
        if not el.is_displayed():
            # is_displayed 对部分遮罩不准，再看 style
            pass
    except Exception:
        return False
    style = (el.get_attribute('style') or '').replace(' ', '').lower()
    if 'display:none' in style:
        return False
    if 'display:block' in style:
        return True
    try:
        display = (el.value_of_css_property('display') or '').lower()
        if display == 'none':
            return False
        if display == 'block':
            return True
    except Exception:
        pass
    try:
        return el.is_displayed()
    except Exception:
        return False


def _css_display(el):
    """读取元素 display（优先 inline style，再 computed）。"""
    style = (el.get_attribute('style') or '').replace(' ', '').lower()
    m = re.search(r'display:([^;]+)', style)
    if m:
        return m.group(1).strip()
    try:
        return (el.value_of_css_property('display') or '').lower().strip()
    except Exception:
        return ''


def is_export_page_url(url):
    """是否为知网导出页。"""
    u = effective_page_url(url) or ''
    return 'export.html' in u or '/dm8/manage/export' in u


def is_unit_header_login_visible(driver):
    """页头「机构登录」按钮可见（display 非 none）= 机构未登录。

    未登录: .ecp_header_unit_loginbg { display: inline-block }
    已登录: .ecp_header_unit_loginbg { display: none }
    """
    els = driver.find_elements(By.CSS_SELECTOR, '.ecp_header_unit_loginbg')
    for el in els:
        display = _css_display(el)
        if display == 'none':
            continue
        # 明确的可见 display
        if display in ('inline-block', 'block', 'inline', 'flex'):
            return True
        # style 未写死 none 时再看 is_displayed
        try:
            if el.is_displayed():
                return True
        except Exception:
            pass
    return False


def is_institution_logged_in(driver):
    """当前窗口机构是否已登录（页头「机构登录」为 display:none）。"""
    els = driver.find_elements(By.CSS_SELECTOR, '.ecp_header_unit_loginbg')
    if not els:
        # 无页头按钮时（如纯导出页）不据此判断，返回 None 语义用 False 弱判
        return find_visible_unit_login_box(driver) is None
    for el in els:
        if _css_display(el) == 'none':
            return True
        try:
            style = (el.get_attribute('style') or '').replace(' ', '').lower()
            if 'display:none' in style:
                return True
        except Exception:
            pass
    return False


def find_visible_unit_login_box(driver):
    """在当前窗口查找可见的机构登录框 .ecp_unit_LoginBox。"""
    boxes = driver.find_elements(By.CSS_SELECTOR, '.ecp_unit_LoginBox')
    for box in boxes:
        if _is_element_display_block(box):
            return box
    return None


def open_unit_login_dialog(driver):
    """打开机构登录弹窗（点击页头「机构登录」或调用 showUnitLogin）。"""
    try:
        opened = driver.execute_script("""
            var el = document.querySelector('.ecp_header_unit_loginbg');
            if (!el) return false;
            var style = (el.getAttribute('style') || '').replace(/\\s/g, '').toLowerCase();
            var disp = window.getComputedStyle(el).display || '';
            if (style.indexOf('display:none') >= 0 || disp === 'none') return false;
            try {
                if (typeof showUnitLogin === 'function') {
                    showUnitLogin(window.event || new MouseEvent('click'));
                    return true;
                }
            } catch (e) {}
            el.click();
            return true;
        """)
        if opened:
            print('已打开机构登录入口（页头「机构登录」）')
            time.sleep(1)
            return True
    except Exception as e:
        print('打开机构登录入口失败:', e)
    try:
        els = driver.find_elements(By.CSS_SELECTOR, '.ecp_header_unit_loginbg')
        for el in els:
            if _css_display(el) == 'none':
                continue
            driver.execute_script('arguments[0].click();', el)
            print('已点击页头「机构登录」')
            time.sleep(1)
            return True
    except Exception as e:
        print('点击页头「机构登录」失败:', e)
    return False


def has_unit_login_anywhere(driver):
    """检查所有窗口是否出现机构登录弹窗（display:block）。"""
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass
    found = False
    for handle in list(driver.window_handles):
        try:
            driver.switch_to.window(handle)
            if find_visible_unit_login_box(driver) is not None:
                found = True
                break
        except Exception:
            continue
    if current and current in driver.window_handles:
        try:
            driver.switch_to.window(current)
        except Exception:
            pass
    return found


def needs_unit_ip_login(driver, check_handle=None):
    """是否需要 IP 登录。

    - 登录弹窗 .ecp_unit_LoginBox：任意窗口（含导出页）均算未登录
    - 页头「机构登录」：只检查检索主页（check_handle），忽略导出页
      （导出页常自带可见的页头按钮，会造成误判并关掉已打开的导出窗）
    """
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass

    need = False
    try:
        # 1) 真实断线弹窗：任意窗口
        if has_unit_login_anywhere(driver):
            return True

        # 2) 页头状态：只看主检索页，跳过 export.html
        handles = []
        if check_handle and check_handle in driver.window_handles:
            handles = [check_handle]
        else:
            for handle in list(driver.window_handles):
                try:
                    driver.switch_to.window(handle)
                    if is_export_page_url(driver.current_url):
                        continue
                    handles.append(handle)
                except Exception:
                    continue

        for handle in handles:
            try:
                driver.switch_to.window(handle)
                if is_export_page_url(driver.current_url):
                    continue
                if is_unit_header_login_visible(driver):
                    need = True
                    break
            except Exception:
                continue
    finally:
        if current and current in driver.window_handles:
            try:
                driver.switch_to.window(current)
            except Exception:
                pass
    return need


def _webvpn_aes_cfb(data: bytes, encrypt=True):
    try:
        from Crypto.Cipher import AES
        cipher = AES.new(
            WEBVPN_AES_KEY, AES.MODE_CFB, iv=WEBVPN_AES_KEY, segment_size=128)
        return cipher.encrypt(data) if encrypt else cipher.decrypt(data)
    except Exception:
        return None


def webvpn_encrypt_host(host):
    """网瑞达 WebVPN 主机名编码。"""
    host = (host or '').strip().lower()
    if host in WEBVPN_HOST_ENC:
        return WEBVPN_HOST_ENC[host]
    out = _webvpn_aes_cfb(host.encode('utf-8'), encrypt=True)
    if out is None:
        raise ValueError(f'无法编码 WebVPN 主机: {host}')
    return out.hex()


def webvpn_decrypt_host(enc_hex):
    enc_hex = (enc_hex or '').split('-', 1)[0].strip().lower()
    for host, code in WEBVPN_HOST_ENC.items():
        if code == enc_hex:
            return host
    try:
        raw = bytes.fromhex(enc_hex)
    except ValueError:
        return ''
    out = _webvpn_aes_cfb(raw, encrypt=False)
    if out is None:
        return ''
    try:
        return out.decode('utf-8')
    except Exception:
        return ''


def is_webvpn_url(url):
    return 'webvpn.uibe.edu.cn' in (url or '').lower()


def is_cnki_404_url(url):
    """知网详情失效页（含 WebVPN 包装后的 /kcms2/article/404）。"""
    raw = (url or '').strip().lower()
    if not raw:
        return False
    inner = (unwrap_webvpn_url(raw) or raw).lower()
    for u in (raw, inner):
        path = urlparse(u).path if '://' in u else u
        if re.search(r'(?:^|/)404(?:/|$|\.html?)', path):
            return True
        if '/article/404' in u:
            return True
    return False


def unwrap_webvpn_url(url):
    """把 WebVPN 包装地址还原为原始 https://host/path，非 VPN 地址原样返回。"""
    raw = (url or '').strip()
    if not raw or not is_webvpn_url(raw):
        return raw
    m = re.search(
        r'https?://webvpn\.uibe\.edu\.cn/(https?)/([0-9a-f]+)(/[^?#]*)?(\?[^#]*)?(#.*)?',
        raw, re.I)
    if not m:
        return raw
    scheme = m.group(1).lower()
    code = m.group(2).lower()
    path = m.group(3) or '/'
    query = m.group(4) or ''
    frag = m.group(5) or ''
    if code.startswith(WEBVPN_MAGIC):
        host = webvpn_decrypt_host(code[len(WEBVPN_MAGIC):])
    else:
        host = webvpn_decrypt_host(code)
    if not host:
        return raw
    return f'{scheme}://{host}{path}{query}{frag}'


def effective_page_url(url):
    """用于主机/路径判断：VPN 包装时还原为原始知网地址。"""
    return unwrap_webvpn_url(url or '')


def to_webvpn_url(url):
    """将普通网址改写为 UIBE WebVPN 访问地址。已是 VPN 地址则不变。"""
    raw = (url or '').strip()
    if not raw:
        return raw
    if is_webvpn_url(raw):
        return raw
    parsed = urlparse(raw)
    host = (parsed.hostname or '').lower()
    if not host:
        return raw
    scheme = (parsed.scheme or 'https').lower()
    if scheme not in ('http', 'https'):
        scheme = 'https'
    try:
        enc = webvpn_encrypt_host(host)
    except ValueError:
        return raw
    code = WEBVPN_MAGIC + enc
    port = parsed.port
    if port and not (
            (scheme == 'https' and port == 443)
            or (scheme == 'http' and port == 80)):
        code = f'{code}-{port}'
    path = parsed.path or '/'
    query = ('?' + parsed.query) if parsed.query else ''
    frag = ('#' + parsed.fragment) if parsed.fragment else ''
    return f'{WEBVPN_ORIGIN}/{scheme}/{code}{path}{query}{frag}'


def is_webvpn_portal_login_url(url):
    u = (url or '').lower()
    if 'webvpn.uibe.edu.cn' not in u:
        return False
    if '/cas/login' in u:
        return True
    path = (urlparse(url).path or '').lower()
    return path.startswith('/login')


def is_login_page_url(url):
    """VPN / CAS / 知网统一认证登录页（含新标签里停留的登录页）。"""
    u = (url or '').strip().lower()
    if not u:
        return False
    if is_webvpn_portal_login_url(u):
        return True
    return (
        '/cas/login' in u
        or 'login.cnki.net' in u
        or 'fsso.cnki.net' in u
    )


def is_cnki_host_url(url):
    return 'cnki.net' in (effective_page_url(url) or '').lower()


def load_unit_credentials():
    """从 cnki_unit_credentials.json 读取机构登录配置。"""
    defaults = {
        'method': UNIT_LOGIN_METHOD,
        'org_name': '对外经济贸易大学',
        'username': '',
        'password': '',
    }
    path = UNIT_CREDENTIALS_FILE
    if not os.path.isfile(path):
        print(f'未找到机构凭据文件: {path}，使用默认 method={defaults["method"]}')
        return defaults
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return defaults
        out = dict(defaults)
        out.update({k: data[k] for k in (
            'method', 'org_name', 'username', 'password') if k in data})
        return out
    except Exception as e:
        print('读取机构凭据失败:', e)
        return defaults


def get_unit_login_method():
    """返回 vpn / ip / school_visit。"""
    method = (load_unit_credentials().get('method') or UNIT_LOGIN_METHOD or 'vpn')
    method = str(method).strip().lower()
    if method in ('vpn', 'webvpn', 'uibe_vpn', '校vpn', 'vpn网站', 'vpn登录'):
        return 'vpn'
    if method in ('school_visit', 'school', '校外', '校外访问', 'fsso'):
        return 'school_visit'
    if method in ('ip', 'ip登录', 'iplogin'):
        return 'ip'
    return 'vpn'


def to_cnki_access_url(url):
    """按当前登录方式给出实际访问地址（VPN 模式下改写为 WebVPN 域名）。"""
    raw = (url or '').strip()
    if not raw:
        return raw
    if get_unit_login_method() != 'vpn':
        return raw
    inner = unwrap_webvpn_url(raw)
    if not is_cnki_host_url(inner) and not is_webvpn_url(raw):
        return raw
    return to_webvpn_url(inner)


def normalize_doc_type(doc_type):
    """journal / patent / none（不选，跳过资源类型筛选）。"""
    value = (doc_type or DOWNLOAD_DOC_TYPE or 'journal').strip().lower()
    if value in ('none', 'any', 'all', '不选', '不限', '全部'):
        return 'none'
    if value in (
            'patent', 'pat', '专利权', '专利', 'zl',
            'vudi', PATENT_CLASSID.lower()):
        return 'patent'
    return 'journal'


def doc_type_label(doc_type):
    normalized = normalize_doc_type(doc_type)
    if normalized == 'patent':
        return '专利'
    if normalized == 'none':
        return '不选'
    return '期刊'


def get_adv_search_url(doc_type=None):
    """高级检索页。VPN 模式改写为学校 WebVPN 地址。

    doc_type=patent 时打开 classid=VUDIXAIY 专利检索；journal/none 用总库高级检索。
    """
    doc_type = normalize_doc_type(doc_type)
    if doc_type == 'patent':
        raw = ADV_SEARCH_URL_PATENT
    elif get_unit_login_method() == 'vpn':
        raw = 'https://kns.cnki.net' + WEBVPN_ADV_SEARCH_PATH
    else:
        raw = ADV_SEARCH_URL
    if get_unit_login_method() == 'vpn':
        return to_webvpn_url(raw)
    return raw


def get_login_home_url():
    """会话起始页：VPN 先打开 WebVPN 门户，其余直接打开知网高级检索。"""
    if get_unit_login_method() == 'vpn':
        return WEBVPN_ORIGIN + '/'
    return ADV_SEARCH_URL


def _short_url(url, maxlen=180):
    u = url or ''
    return u if len(u) <= maxlen else (u[:maxlen] + '...')


def wait_url_settle(driver, timeout=8, quiet=1.0):
    """等待地址栏停止跳转。"""
    deadline = time.time() + timeout
    last = ''
    stable_since = time.time()
    while time.time() < deadline:
        try:
            cur = driver.current_url or ''
        except Exception:
            time.sleep(0.3)
            continue
        if cur != last:
            last = cur
            stable_since = time.time()
        elif cur and (time.time() - stable_since) >= quiet:
            return cur
        time.sleep(0.3)
    return last


def open_cnki_url(driver, url):
    """打开知网地址。

    VPN 模式：先直连原始网址，等站点自己跳转拿到真实域名/路径，
    再按落地页拼接 WebVPN 地址并打开（避免一上来就拼错主机编码）。
    """
    raw = (url or '').strip()
    if not raw:
        return ''
    if get_unit_login_method() != 'vpn':
        driver.get(raw)
        wait_url_settle(driver)
        try:
            return driver.current_url or raw
        except Exception:
            return raw

    inner = unwrap_webvpn_url(raw)
    inner_l = inner.lower()
    if '/kcms2/article/abstract' in inner_l or '/kcms/detail' in inner_l:
        # 文章详情页已知真实地址，直连通常被校园网/VPN 拦截，直接走 WebVPN
        vpn_url = to_webvpn_url(inner)
        print(f'文章详情页直接走 WebVPN: {_short_url(vpn_url)}')
        try:
            driver.get(vpn_url)
        except Exception as e:
            print('WebVPN 文章页打开未完成:', e)
        wait_url_settle(driver)
        try:
            return driver.current_url or vpn_url
        except Exception:
            return vpn_url

    print(f'先直接访问以获取落地域名: {_short_url(inner)}')
    try:
        driver.get(inner)
    except Exception as e:
        print('直接访问未完成（超时或跳转中），继续读取当前地址:', e)
    wait_url_settle(driver)
    try:
        landed = driver.current_url or inner
    except Exception:
        landed = inner

    if is_login_page_url(landed):
        print('直接访问落到 WebVPN/CAS 登录页，先完成登录')
        try_unit_vpn_login(driver)
        try:
            landed = driver.current_url or landed
        except Exception:
            pass

    effective = unwrap_webvpn_url(landed) or inner
    if not is_cnki_host_url(effective) and is_cnki_host_url(inner):
        effective = inner
    if is_cnki_404_url(landed) or is_cnki_404_url(effective):
        print(f'落地页为 404，不再拼接跳转: {_short_url(landed or effective)}')
        return landed or effective
    vpn_url = to_webvpn_url(effective)
    already = (
        is_webvpn_url(landed)
        and unwrap_webvpn_url(landed).split('#')[0].rstrip('/')
        == effective.split('#')[0].rstrip('/')
    )
    if already:
        print(f'已在对应 WebVPN 地址: {_short_url(landed)}')
        return landed
    print(f'按落地页拼接 WebVPN 跳转: {_short_url(vpn_url)}')
    try:
        driver.get(vpn_url)
    except Exception as e:
        print('WebVPN 跳转未完成:', e)
    wait_url_settle(driver)
    try:
        return driver.current_url or vpn_url
    except Exception:
        return vpn_url


def open_adv_search_in_session(driver, doc_type=None):
    """同会话内回到高级检索页。

    直接打开 get_adv_search_url（VPN 模式下已是 WebVPN 地址），
    不做「先直连再拼 VPN」探测，避免任务内每段再出滑块验证码。
    """
    url = get_adv_search_url(doc_type)
    print(f'同会话打开高级检索（跳过直连探测）: {_short_url(url)}')
    try:
        driver.get(url)
    except Exception as e:
        print('同会话打开检索页未完成:', e)
    wait_url_settle(driver)
    try:
        return driver.current_url or url
    except Exception:
        return url


def ensure_cnki_captcha_after_nav(driver, save_dir, main_handle=None):
    """导航后若落到验证码页：等控件加载并处理；通过后若掉回直连再拼回 WebVPN。

    返回 True / False / 'restart_browser'（验证码连败需关浏览器）。
    """
    wait_url_settle(driver, timeout=6)
    try:
        url = driver.current_url or ''
    except Exception:
        url = ''
    if not (is_cnki_verify_url(url) or has_captcha_anywhere(driver)):
        return True
    if is_cnki_verify_url(url) and not has_visible_captcha_widget(driver):
        print('验证页暂未出现可交互控件，等待跳转...')
        deadline = time.time() + 6
        while time.time() < deadline:
            if has_visible_captcha_widget(driver):
                break
            time.sleep(0.4)
        else:
            print('验证页未出现可交互控件，按已通过处理')
            return True
    print(f'跳转后出现验证码，开始处理: {_short_url(url)}')
    wait_for_any_captcha_widget(driver, timeout=12)
    cap = ensure_captcha_passed(driver, save_dir, main_handle)
    if cap == 'restart_browser':
        return 'restart_browser'
    if not cap:
        return False
    wait_url_settle(driver, timeout=8)
    try:
        after = driver.current_url or ''
    except Exception:
        after = ''
    if is_cnki_verify_url(after):
        print('验证码处理后仍在校验页，再等跳转...')
        time.sleep(2)
        wait_url_settle(driver, timeout=8)
        try:
            after = driver.current_url or ''
        except Exception:
            after = ''
    if (
        get_unit_login_method() == 'vpn'
        and after
        and not is_webvpn_url(after)
        and is_cnki_host_url(after)
    ):
        vpn_url = to_webvpn_url(after)
        print(f'验证码后落到直连知网，按域名拼接 WebVPN: {_short_url(vpn_url)}')
        try:
            driver.get(vpn_url)
        except Exception as e:
            print('验证码后 WebVPN 跳转失败:', e)
            return False
        wait_url_settle(driver)
        try:
            url2 = driver.current_url or ''
        except Exception:
            url2 = ''
        if is_cnki_verify_url(url2) or has_captcha_anywhere(driver):
            print('拼接跳转后再次出现验证码，继续处理')
            wait_for_any_captcha_widget(driver, timeout=12)
            return ensure_captcha_passed(driver, save_dir, main_handle)
    return True


def _click_unit_ip_login(driver):
    """在当前窗口点击 IP 登录。"""
    btns = driver.find_elements(By.CSS_SELECTOR, 'button.ecp_IPLogin')
    if btns:
        try:
            driver.execute_script("arguments[0].click();", btns[0])
            print('已点击 button.ecp_IPLogin（IP登录）')
            return True
        except Exception as e:
            print('点击 IP登录 按钮失败:', e)
    try:
        ok = driver.execute_script(
            'if (typeof Ecp_AutoIPLogin === "function") { Ecp_AutoIPLogin(); return true; }'
            'return false;')
        if ok:
            print('已调用 Ecp_AutoIPLogin()')
            return True
    except Exception as e:
        print('调用 Ecp_AutoIPLogin 失败:', e)
    return False


def _click_unit_school_visit(driver):
    """在当前窗口点击「校外访问」。"""
    btns = driver.find_elements(By.CSS_SELECTOR, 'button.ecp_schoolVisit')
    if btns:
        try:
            driver.execute_script("arguments[0].click();", btns[0])
            print('已点击 button.ecp_schoolVisit（校外访问）')
            return True
        except Exception as e:
            print('点击 校外访问 按钮失败:', e)
    try:
        ok = driver.execute_script(
            'if (typeof Ecp_To_Fsso === "function") { Ecp_To_Fsso(); return true; }'
            'return false;')
        if ok:
            print('已调用 Ecp_To_Fsso()')
            return True
    except Exception as e:
        print('调用 Ecp_To_Fsso 失败:', e)
    return False


def _focus_sso_related_window(driver, main_handle=None):
    """优先切到 SSO/CAS/FSSO 相关窗口，否则切到最新非主窗。"""
    keywords = (
        'fsso.cnki.net', 'idp-lib.uibe.edu.cn', 'cas.uibe.edu.cn',
        '/cas/login', 'shibboleth', 'idp/profile',
    )
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass
    for handle in list(driver.window_handles):
        if main_handle and handle == main_handle:
            continue
        try:
            driver.switch_to.window(handle)
            url = (driver.current_url or '').lower()
            if any(k in url for k in keywords):
                return handle
        except Exception:
            continue
    # 最新非主窗
    for handle in reversed(list(driver.window_handles)):
        if main_handle and handle == main_handle:
            continue
        try:
            driver.switch_to.window(handle)
            return handle
        except Exception:
            continue
    if current and current in driver.window_handles:
        driver.switch_to.window(current)
    return current


def _try_fsso_select_org(driver, org_name):
    """fsso.cnki.net：输入机构名并点「前往」。已跳转则返回 False。"""
    url = (effective_page_url(driver.current_url or '') or '').lower()
    if 'fsso.cnki.net' not in url:
        return False
    inputs = driver.find_elements(By.CSS_SELECTOR, 'input#o')
    if not inputs:
        return False
    inp = inputs[0]
    try:
        cur = (inp.get_attribute('value') or '').strip()
        if org_name not in cur:
            inp.clear()
            inp.send_keys(org_name)
            time.sleep(0.6)
        submits = driver.find_elements(By.CSS_SELECTOR, 'div.submit_button')
        if not submits:
            submits = driver.find_elements(
                By.XPATH, "//*[contains(normalize-space(.),'前往')]")
        if submits:
            driver.execute_script("arguments[0].click();", submits[0])
            print(f'校外访问: 已输入「{org_name}」并点击前往')
            return True
    except Exception as e:
        print('校外访问: 选择机构失败:', e)
    return False


def _try_idp_consent_or_accept(driver):
    """idp-lib：勾选同意并点提交/接受。"""
    url = (driver.current_url or '').lower()
    if 'idp-lib.uibe.edu.cn' not in url and '/idp/' not in url:
        return False
    acted = False
    try:
        boxes = driver.find_elements(
            By.CSS_SELECTOR, 'input#accept[type="checkbox"]')
        for box in boxes:
            try:
                if not box.is_selected():
                    driver.execute_script("arguments[0].click();", box)
                    print('校外访问: 已勾选「同意」')
                    acted = True
            except Exception:
                pass
        submits = driver.find_elements(
            By.CSS_SELECTOR, 'input[type="submit"][name="_eventId_proceed"]')
        for btn in submits:
            val = (btn.get_attribute('value') or '').strip()
            if val in ('提交', '接受'):
                driver.execute_script("arguments[0].click();", btn)
                print(f'校外访问: 已点击「{val}」')
                return True
    except Exception as e:
        print('校外访问: 同意/提交失败:', e)
    return acted


def _try_cas_login(driver, username, password):
    """CAS（含 WebVPN 包装的 /cas/login）：填写账号密码并点立刻登录。"""
    url = (driver.current_url or '').lower()
    inner = (effective_page_url(url) or '').lower()
    if ('cas.uibe.edu.cn' not in url and '/cas/login' not in url
            and 'cas.uibe.edu.cn' not in inner):
        return False
    uns = driver.find_elements(By.CSS_SELECTOR, 'input#un')
    pds = driver.find_elements(By.CSS_SELECTOR, 'input#pd')
    if not uns or not pds:
        return False
    if not username or not password:
        print('CAS: 凭据文件缺少 username/password')
        return False
    try:
        un, pd = uns[0], pds[0]
        if not (un.get_attribute('value') or '').strip():
            un.clear()
            un.send_keys(username)
        if not (pd.get_attribute('value') or '').strip():
            pd.clear()
            pd.send_keys(password)
        rem = driver.find_elements(By.CSS_SELECTOR, 'input#rememberName')
        if rem and not rem[0].is_selected():
            driver.execute_script("arguments[0].click();", rem[0])
            print('CAS: 已勾选「记住用户名」')
        btns = driver.find_elements(
            By.CSS_SELECTOR, 'input.login_box_landing_btn')
        if not btns:
            btns = driver.find_elements(
                By.CSS_SELECTOR, 'input[type="button"].login_box_landing_btn')
        if btns:
            driver.execute_script("arguments[0].click();", btns[0])
            print('CAS: 已点击「立刻登录」')
            return True
        landing = driver.find_elements(By.CSS_SELECTOR, '#index_login_btn')
        if landing:
            driver.execute_script("arguments[0].click();", landing[0])
            print('CAS: 已点击「立刻登录」')
            return True
    except Exception as e:
        print('CAS 登录失败:', e)
    return False


def _click_webvpn_cas_entry(driver):
    """WebVPN 登录页点击「CAS统一身份认证登录」。"""
    url = (driver.current_url or '').lower()
    if '/cas/login' in url:
        return False
    if 'webvpn.uibe.edu.cn' not in url:
        return False
    els = driver.find_elements(By.CSS_SELECTOR, 'a#cas-login')
    if not els:
        els = driver.find_elements(
            By.XPATH,
            "//a[contains(normalize-space(.),'CAS统一身份认证登录')]")
    if not els:
        return False
    try:
        driver.execute_script("arguments[0].click();", els[0])
        print('WebVPN: 已点击「CAS统一身份认证登录」')
        return True
    except Exception as e:
        print('WebVPN: 点击 CAS 入口失败:', e)
        return False


def _click_webvpn_cnki_tile(driver):
    """在 WebVPN 主页点击「CNKI知网」资源块；新标签打开时切过去并关掉门户页。"""
    titles = driver.find_elements(
        By.CSS_SELECTOR, 'h2.block-group__item__name')
    target = None
    for h2 in titles:
        text = ((h2.get_attribute('title') or '') + ' ' + (h2.text or '')).strip()
        if 'CNKI' in text or '知网' in text:
            target = h2
            break
    if target is None:
        return False
    before = list(driver.window_handles)
    try:
        link = target.find_element(By.XPATH, './ancestor::a[1]')
    except Exception:
        link = target
    try:
        driver.execute_script("arguments[0].click();", link)
        print('WebVPN: 已点击「CNKI知网」')
    except Exception as e:
        print('WebVPN: 点击 CNKI知网 失败:', e)
        return False
    deadline = time.time() + 8
    new_handles = []
    while time.time() < deadline:
        after = list(driver.window_handles)
        new_handles = [h for h in after if h not in before]
        if new_handles:
            break
        time.sleep(0.3)
    cnki_handle = None
    if new_handles:
        try:
            driver.switch_to.window(new_handles[-1])
            cnki_handle = new_handles[-1]
            print('WebVPN: CNKI 在新标签打开，已切换')
        except Exception:
            pass
    found = _find_cnki_webvpn_handle(driver)
    if found:
        cnki_handle = found
    _close_webvpn_extra_windows(driver, keep_handle=cnki_handle)
    return True


def _find_cnki_webvpn_handle(driver):
    """找到已通过 WebVPN 打开的知网标签。"""
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass
    found = None
    for handle in list(driver.window_handles):
        try:
            driver.switch_to.window(handle)
            url = driver.current_url or ''
            if _is_cnki_via_webvpn(url):
                found = handle
                break
        except Exception:
            continue
    if found:
        try:
            driver.switch_to.window(found)
        except Exception:
            pass
        return found
    if current and current in driver.window_handles:
        try:
            driver.switch_to.window(current)
        except Exception:
            pass
    return None


def _close_webvpn_extra_windows(driver, keep_handle=None):
    """关掉 WebVPN 门户、CAS、空白页等多余标签，保留知网页。"""
    keep = keep_handle
    if not keep or keep not in getattr(driver, 'window_handles', []):
        keep = _find_cnki_webvpn_handle(driver)
    if (not keep) and driver.window_handles:
        keep = driver.window_handles[-1]
    for handle in list(driver.window_handles):
        if handle == keep:
            continue
        if len(driver.window_handles) <= 1:
            break
        try:
            driver.switch_to.window(handle)
            url = (driver.current_url or '').lower()
            close_it = (
                is_webvpn_portal_login_url(url)
                or '/cas/login' in url
                or 'about:blank' in url
                or url.startswith('chrome://')
                or _is_webvpn_resource_home(driver)
                or (is_webvpn_url(url) and not _is_cnki_via_webvpn(url)
                    and not is_cnki_verify_url(url))
            )
            if close_it:
                print(f'WebVPN: 关闭多余标签 {_short_url(driver.current_url or url)}')
                driver.close()
        except Exception:
            continue
    if keep and keep in driver.window_handles:
        try:
            driver.switch_to.window(keep)
        except Exception:
            pass
    elif driver.window_handles:
        try:
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass
    return keep if keep in getattr(driver, 'window_handles', []) else None


def _is_cnki_via_webvpn(url):
    u = url or ''
    if not is_webvpn_url(u):
        return False
    if is_webvpn_portal_login_url(u) or '/cas/login' in u.lower():
        return False
    return is_cnki_host_url(u)


def _webvpn_cnki_ready(driver):
    try:
        url = driver.current_url or ''
    except Exception:
        return False
    if not _is_cnki_via_webvpn(url):
        return False
    try:
        if is_unit_header_login_visible(driver):
            return False
        if find_visible_unit_login_box(driver) is not None:
            return False
    except Exception:
        pass
    return True


def _is_webvpn_resource_home(driver):
    try:
        url = driver.current_url or ''
    except Exception:
        return False
    if 'webvpn.uibe.edu.cn' not in url.lower():
        return False
    if is_webvpn_portal_login_url(url) or '/cas/login' in url.lower():
        return False
    if _is_cnki_via_webvpn(url):
        return False
    return bool(driver.find_elements(
        By.CSS_SELECTOR, 'h2.block-group__item__name'))


def needs_webvpn_login(driver, check_handle=None):
    """VPN 模式下：未走 WebVPN、停在登录页、或尚未进入知网。"""
    if get_unit_login_method() != 'vpn':
        return False
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass
    try:
        # 任一标签停在 VPN/CAS/知网登录页：即使还残留旧知网标签，也必须重新登录
        for handle in list(driver.window_handles):
            try:
                driver.switch_to.window(handle)
                if is_login_page_url(driver.current_url or ''):
                    return True
            except Exception:
                continue
        if _find_cnki_webvpn_handle(driver):
            return False
        url = ''
        if check_handle and check_handle in driver.window_handles:
            driver.switch_to.window(check_handle)
            url = driver.current_url or ''
        elif current and current in driver.window_handles:
            driver.switch_to.window(current)
            url = driver.current_url or ''
        else:
            return True
        if is_webvpn_portal_login_url(url) or '/cas/login' in url.lower():
            return True
        if _is_cnki_via_webvpn(url):
            return False
        return True
    except Exception:
        return True
    finally:
        if current and current in driver.window_handles:
            try:
                driver.switch_to.window(current)
            except Exception:
                pass


def try_unit_vpn_login(driver, restore_handle=None, timeout=180):
    """经 UIBE WebVPN + CAS 登录，再进入 CNKI 知网。"""
    creds = load_unit_credentials()
    user = (creds.get('username') or '').strip()
    pwd = creds.get('password') or ''
    if not user or not pwd:
        print('VPN 登录需要 cnki_unit_credentials.json 中的 username/password')
        return False

    wanted = None
    try:
        if restore_handle and restore_handle in driver.window_handles:
            driver.switch_to.window(restore_handle)
        cur = driver.current_url or ''
        inner = effective_page_url(cur)
        if (is_cnki_host_url(inner) and not is_webvpn_portal_login_url(cur)
                and not is_cnki_verify_url(cur) and not is_cnki_verify_url(inner)):
            wanted = inner
    except Exception:
        pass

    print('开始 WebVPN 网站登录...')
    try:
        driver.get(WEBVPN_ORIGIN + '/')
    except Exception as e:
        print('打开 WebVPN 失败:', e)
        return False
    time.sleep(2)

    deadline = time.time() + timeout
    cas_entry_clicked = False
    cas_submitted = False
    cnki_clicked = False
    reached_cnki_at = None

    while time.time() < deadline:
        try:
            url = driver.current_url or ''
        except Exception:
            time.sleep(1)
            continue

        if _webvpn_cnki_ready(driver):
            print('WebVPN: 知网访问已就绪')
            break
        if _is_cnki_via_webvpn(url):
            if reached_cnki_at is None:
                reached_cnki_at = time.time()
                print('WebVPN: 已进入知网，等待机构身份生效...')
            elif time.time() - reached_cnki_at >= 12:
                print('WebVPN: 已在知网页面，继续后续操作')
                break
            time.sleep(1.2)
            continue

        if not cas_entry_clicked and _click_webvpn_cas_entry(driver):
            cas_entry_clicked = True
            time.sleep(2)
            continue

        if not cas_submitted and _try_cas_login(driver, user, pwd):
            cas_submitted = True
            time.sleep(3)
            continue

        if not cnki_clicked and _click_webvpn_cnki_tile(driver):
            cnki_clicked = True
            time.sleep(2)
            continue

        if (_is_webvpn_resource_home(driver)
                and cas_submitted and not cnki_clicked):
            print('WebVPN: 已登录门户，直接打开知网')
            try:
                driver.get(to_webvpn_url('https://www.cnki.net/'))
                cnki_clicked = True
                time.sleep(3)
                continue
            except Exception as e:
                print('WebVPN: 打开知网失败:', e)

        time.sleep(1.2)
    else:
        print('WebVPN 登录超时')
        return False

    keep = _close_webvpn_extra_windows(driver)
    if keep:
        try:
            driver.switch_to.window(keep)
        except Exception:
            pass

    target = get_adv_search_url()
    if wanted:
        target = to_cnki_access_url(wanted)
    try:
        cur = driver.current_url or ''
        same = (
            unwrap_webvpn_url(cur).split('#')[0].rstrip('/')
            == unwrap_webvpn_url(target).split('#')[0].rstrip('/')
        )
        if not same:
            driver.get(target)
            time.sleep(2)
            _close_webvpn_extra_windows(driver)
    except Exception as e:
        print('WebVPN: 打开目标页失败:', e)
        return False
    return True


def _school_visit_step_on_current(driver, creds):
    """对当前页执行一步校外访问流程；有动作返回 True。"""
    org = creds.get('org_name') or ''
    user = creds.get('username') or ''
    pwd = creds.get('password') or ''
    if _try_fsso_select_org(driver, org):
        return True
    if _try_cas_login(driver, user, pwd):
        return True
    if _try_idp_consent_or_accept(driver):
        return True
    return False


def _sso_login_keywords():
    return (
        'fsso.cnki.net', 'idp-lib.uibe.edu.cn', 'cas.uibe.edu.cn',
        '/cas/login', '/idp/',
    )


def _classify_extra_windows(driver, main_handle=None):
    """区分仍在进行的认证窗 与 登录完成后的门户首页。"""
    sso_handles = []
    portal_handles = []
    for handle in list(driver.window_handles):
        if main_handle and handle == main_handle:
            continue
        try:
            driver.switch_to.window(handle)
            url = driver.current_url or ''
            if is_cnki_portal_home_url(url):
                portal_handles.append(handle)
            elif any(k in url.lower() for k in _sso_login_keywords()):
                sso_handles.append(handle)
        except Exception:
            continue
    return sso_handles, portal_handles


def _sso_or_portal_handles(driver, main_handle=None):
    """仍存在的 SSO / 门户首页窗口。"""
    sso_handles, portal_handles = _classify_extra_windows(driver, main_handle)
    return sso_handles + portal_handles


def _close_portal_windows(driver, main_handle=None):
    """只关 www.cnki.net 门户页，不关正在认证的 CAS/IdP。"""
    _, portal_handles = _classify_extra_windows(driver, main_handle)
    for handle in portal_handles:
        if len(driver.window_handles) <= 1:
            break
        if main_handle and handle == main_handle:
            continue
        try:
            driver.switch_to.window(handle)
            print(f'关闭多余标签: {driver.current_url}')
            driver.close()
        except Exception as e:
            print(f'关闭门户页失败: {e}')
    if main_handle and main_handle in driver.window_handles:
        try:
            driver.switch_to.window(main_handle)
        except Exception:
            pass


def _has_search_results_ui(driver):
    """当前页是否已有检索结果（SPA 下 URL 仍可能是 AdvSearch）。"""
    selectors = (
        (By.CLASS_NAME, 'checkAll'),
        (By.ID, 'Page_next_top'),
        (By.ID, 'briefBox'),
        (By.CLASS_NAME, 'result-table-list'),
        (By.CSS_SELECTOR, '#gridTable'),
    )
    for by, val in selectors:
        try:
            if driver.find_elements(by, val):
                return True
        except Exception:
            continue
    return False


def _search_page_logged_in(driver, handle):
    """检索/结果页是否已显示机构登录成功。"""
    if not handle or handle not in driver.window_handles:
        return False
    try:
        driver.switch_to.window(handle)
        if not is_cnki_work_page_url(driver.current_url):
            return False
        return (
            is_institution_logged_in(driver)
            and not is_unit_header_login_visible(driver)
            and find_visible_unit_login_box(driver) is None
        )
    except Exception:
        return False


def _capture_work_page(driver, handle=None):
    """记录当前工作页句柄、URL，以及是否已有检索结果。"""
    preserve_results = False
    try:
        if handle and handle in driver.window_handles:
            driver.switch_to.window(handle)
        handle = driver.current_window_handle
        url = driver.current_url or ''
        preserve_results = (
            _has_search_results_ui(driver) or is_cnki_results_page_url(url))
        if is_cnki_work_page_url(url):
            return handle, url, preserve_results
    except Exception:
        pass
    return handle, None, preserve_results


def _refresh_search_login_state(
        driver, handle, resume_url=None, preserve_results=False):
    """刷新工作页以更新登录态。已有检索结果时禁止刷新，避免回到初始页。"""
    if handle and handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
        except Exception:
            handle = None
    if (not handle or handle not in driver.window_handles) and driver.window_handles:
        handle = driver.window_handles[0]
        driver.switch_to.window(handle)

    url = ''
    try:
        url = driver.current_url or ''
    except Exception:
        pass

    try:
        if is_cnki_work_page_url(url):
            kind = '结果页' if is_cnki_results_page_url(url) else '高级检索页'
            print(f'校外访问: 刷新{kind}以更新登录状态')
            if preserve_results or _has_search_results_ui(driver):
                print('校外访问: 当前已有检索结果，跳过刷新以免回到初始检索页')
                return driver.current_window_handle
            driver.refresh()
            time.sleep(3)
            return driver.current_window_handle

        restore = resume_url if is_cnki_work_page_url(resume_url) else None
        if restore:
            kind = '结果页' if is_cnki_results_page_url(restore) else '高级检索页'
            print(f'校外访问: 工作页已离开，恢复{kind}')
            
            driver.get(to_cnki_access_url(restore))
            time.sleep(3)
            return driver.current_window_handle

        print('校外访问: 工作页丢失，打开高级检索页')
        driver.get(get_adv_search_url())
        time.sleep(3)
        return driver.current_window_handle
    except Exception as e:
        print('校外访问: 刷新检索页失败:', e)
        return handle


def complete_school_visit_login(
        driver, creds, main_handle=None, resume_url=None,
        preserve_results=False, timeout=180):
    """完成校外访问 SSO。认证过程中不关闭 CAS/IdP；有结果时不刷新工作页。"""
    deadline = time.time() + timeout
    idle_rounds = 0
    refresh_tries = 0
    max_refresh = 2
    sso_seen = False
    while time.time() < deadline:
        try:
            if main_handle and main_handle in driver.window_handles:
                if _search_page_logged_in(driver, main_handle):
                    print('校外访问: 原检索/结果页已显示机构登录成功')
                    _close_portal_windows(driver, main_handle=main_handle)
                    ensure_single_search_window(
                        driver, prefer_handle=main_handle, resume_url=resume_url)
                    return True
        except Exception:
            pass

        _focus_sso_related_window(driver, main_handle=main_handle)
        acted = False
        for handle in list(driver.window_handles):
            if main_handle and handle == main_handle:
                continue
            try:
                driver.switch_to.window(handle)
                url = (driver.current_url or '').lower()
                if any(k in url for k in _sso_login_keywords()):
                    if _school_visit_step_on_current(driver, creds):
                        acted = True
                        sso_seen = True
                        time.sleep(1.5)
                        break
            except Exception:
                continue

        if not acted:
            try:
                if _school_visit_step_on_current(driver, creds):
                    acted = True
                    sso_seen = True
                    time.sleep(1.5)
            except Exception:
                pass

        if acted:
            idle_rounds = 0
            continue

        sso_handles, portal_handles = _classify_extra_windows(
            driver, main_handle=main_handle)
        if sso_handles or portal_handles:
            sso_seen = True
        if portal_handles:
            print('校外访问: 检测到门户首页，关闭后回到工作页')
            _close_portal_windows(driver, main_handle=main_handle)

        # 认证窗还在：继续等跳转，绝不关掉 CAS/IdP
        if sso_handles:
            idle_rounds += 1
            if idle_rounds >= 3 and idle_rounds % 3 == 0:
                print('校外访问: 等待认证页跳转...')
            time.sleep(1.2)
            continue

        if sso_seen:
            kept = ensure_single_search_window(
                driver, prefer_handle=main_handle, resume_url=resume_url)
            if kept:
                main_handle = kept
            on_results = preserve_results or _has_search_results_ui(driver)
            if on_results:
                print('校外访问: 已回到检索结果页（不刷新，避免丢失结果）')
                if not _search_page_logged_in(driver, main_handle):
                    print('校外访问: 结果页页头可能尚未刷新，继续用当前结果导出')
                return True
            if _search_page_logged_in(driver, main_handle):
                print('校外访问: 已清理多余标签并确认机构登录成功')
                return True
            if refresh_tries < max_refresh:
                refresh_tries += 1
                main_handle = _refresh_search_login_state(
                    driver, main_handle, resume_url=resume_url,
                    preserve_results=False) or main_handle
                ensure_single_search_window(
                    driver, prefer_handle=main_handle, resume_url=resume_url)
                if _search_page_logged_in(driver, main_handle):
                    print('校外访问: 刷新后机构登录成功')
                    return True
                continue
            print('校外访问: 已返回检索页，但刷新后仍未登录')
            return False

        idle_rounds += 1
        if idle_rounds >= 3 and idle_rounds % 3 == 0:
            print('校外访问: 等待页面跳转...')
        time.sleep(1.2)

    print('校外访问: 登录超时')
    return False


def is_cnki_portal_home_url(url):
    """是否为知网门户首页（登录后常跳转 www.cnki.net，非检索页）。"""
    raw = (effective_page_url(url) or '').strip()
    u = raw.lower().split('#')[0].split('?')[0].rstrip('/')
    if not u:
        return False
    if 'fsso.cnki.net' in u or 'kns.cnki.net' in u:
        return False
    if re.match(r'^https?://(www\.)?cnki\.net$', u):
        return True
    if re.match(r'^https?://(www\.)?cnki\.net/(index(\.html?)?)?$', u):
        return True
    return False


def is_cnki_search_page_url(url):
    """是否为知网检索相关页（高级检索/结果页）。"""
    u = (effective_page_url(url) or '').lower()
    if 'kns.cnki.net' not in u:
        return False
    return any(k in u for k in (
        'advsearch', 'defaultresult', 'kns8', 'brief/', 'search'))


def is_cnki_results_page_url(url):
    """是否为检索结果页（已开始翻页/导出）。"""
    u = (effective_page_url(url) or '').lower()
    if 'kns.cnki.net' not in u:
        return False
    if 'advsearch' in u:
        return False
    return any(k in u for k in ('defaultresult', 'brief/', '/result'))


def is_cnki_work_page_url(url):
    """高级检索页或结果页，均可作为登录后应回到的工作页。"""
    return is_cnki_search_page_url(url) or is_cnki_results_page_url(url)


class SessionLostError(RuntimeError):
    """登录 / VPN 会话中断，需重建浏览器重新登录。"""


LOGIN_LOST_TEXT_MARKERS = (
    '登录状态可能中断',
    '重新进行登录',
    '重新登录',
    '登录已失效',
    '登录已过期',
    '登录状态已失效',
    '登录状态已过期',
    '会话已过期',
    '请先登录',
    '请登录后',
    '需要登录后',
    'session expired',
    'session has expired',
    'not logged in',
)


def text_suggests_login_lost(text):
    """异常信息 / 页面文案是否表示登录失效。"""
    raw = str(text or '').strip().lower()
    if not raw:
        return False
    compact = raw.replace(' ', '').replace('\u3000', '')
    for marker in LOGIN_LOST_TEXT_MARKERS:
        key = marker.lower().replace(' ', '')
        if key and key in compact:
            return True
    return False


def page_has_no_results(driver):
    """结果页正文是否明确显示「无结果」。"""
    try:
        body = driver.find_element(By.TAG_NAME, 'body').text or ''
    except Exception:
        return False
    text = body.replace(' ', '')
    return any(k in text for k in (
        '没有找到', '暂无数据', '无相关数据', '0条结果',
        '检索结果：0', '检索结果:0', '未检索到',
        'no results', 'no records found', '0 results',
    ))


def page_suggests_login_lost(driver, check_handle=None):
    """当前页是否已掉登录。

    判定顺序：VPN/CAS 登录页 → 结果页骨架缺失时的提示文案 / 登录弹窗。
    结果页骨架（secondUl/checkAll）存在时视为正常，避免把「未登录页头」误判。
    """
    original = None
    try:
        original = driver.current_window_handle
    except Exception:
        original = None
    try:
        if check_handle and check_handle in driver.window_handles:
            try:
                driver.switch_to.window(check_handle)
            except Exception:
                pass
        try:
            url = driver.current_url or ''
        except Exception:
            return False
        if is_cnki_verify_url(url):
            # 验证码页可能页头仍显示「机构登录」，不能据此判定掉登录
            return False
        if is_login_page_url(url):
            return True

        try:
            page_ready = bool(
                driver.find_elements(By.CLASS_NAME, 'secondUl')
                or driver.find_elements(By.CLASS_NAME, 'checkAll')
                or driver.find_elements(By.ID, 'gradetxt'))
        except Exception:
            page_ready = False

        try:
            body = driver.find_element(By.TAG_NAME, 'body').text or ''
        except Exception:
            body = ''
        if text_suggests_login_lost(body):
            return True
        if page_ready:
            # 页面骨架正常时，后台残留的登录标签不作为掉线依据；
            # 但页面上真实弹出机构登录框仍算断线
            try:
                if has_unit_login_anywhere(driver):
                    return True
            except Exception:
                pass
            return False

        # 工作页已无结果骨架：任一标签停在 VPN/CAS/知网登录页都算断线
        for handle in list(driver.window_handles):
            try:
                driver.switch_to.window(handle)
                if is_login_page_url(driver.current_url or ''):
                    return True
            except Exception:
                continue

        # 再看登录弹窗 / 页头是否提示需要机构登录
        try:
            if has_unit_login_anywhere(driver) or needs_unit_ip_login(
                    driver, check_handle=check_handle):
                return True
        except Exception:
            pass
        return False
    finally:
        if original and original in driver.window_handles:
            try:
                driver.switch_to.window(original)
            except Exception:
                pass


def exception_suggests_login_lost(exc, driver=None, check_handle=None):
    """异常本身或当前页是否表示登录中断。"""
    if text_suggests_login_lost(exc):
        return True
    if driver is None:
        return False
    try:
        return page_suggests_login_lost(driver, check_handle=check_handle)
    except Exception:
        return False


def _close_sso_extra_windows(driver, main_handle=None):
    """关闭 FSSO/CAS/IdP 及知网门户首页，保留检索窗。"""
    ensure_single_search_window(driver, prefer_handle=main_handle)


def ensure_single_search_window(driver, prefer_handle=None, resume_url=None):
    """登录后清理多余标签：关 SSO、关 www.cnki.net，保留原工作页。

    检索进行中优先保留结果页，避免退回空白高级检索。
    返回保留的工作页 window handle。
    """
    sso_keywords = (
        'fsso.cnki.net', 'idp-lib.uibe.edu.cn', 'cas.uibe.edu.cn',
        '/cas/login', '/idp/',
    )

    search_handle = None
    if prefer_handle and prefer_handle in driver.window_handles:
        try:
            driver.switch_to.window(prefer_handle)
            if is_cnki_work_page_url(driver.current_url):
                search_handle = prefer_handle
        except Exception:
            pass
    if search_handle is None:
        # 优先结果页，其次高级检索页
        result_handle = None
        adv_handle = None
        for handle in list(driver.window_handles):
            try:
                driver.switch_to.window(handle)
                url = driver.current_url or ''
                if is_cnki_results_page_url(url) and result_handle is None:
                    result_handle = handle
                elif is_cnki_search_page_url(url) and adv_handle is None:
                    adv_handle = handle
            except Exception:
                continue
        search_handle = result_handle or adv_handle

    for handle in list(driver.window_handles):
        if search_handle and handle == search_handle:
            continue
        if len(driver.window_handles) <= 1:
            break
        try:
            driver.switch_to.window(handle)
            url = driver.current_url or ''
            ul = (effective_page_url(url) or url).lower()
            close_it = (
                any(k in ul or k in url.lower() for k in sso_keywords)
                or is_cnki_portal_home_url(url)
                or (
                    search_handle is not None
                    and 'kns.cnki.net' in ul
                    and handle != search_handle
                )
            )
            if close_it:
                print(f'关闭多余标签: {url}')
                driver.close()
        except Exception as e:
            print(f'关闭多余标签失败: {e}')

    if search_handle and search_handle in driver.window_handles:
        try:
            driver.switch_to.window(search_handle)
            return search_handle
        except Exception:
            pass

    if not driver.window_handles:
        return None

    driver.switch_to.window(driver.window_handles[0])
    cur = driver.current_url or ''
    if is_cnki_portal_home_url(cur) or not is_cnki_work_page_url(cur):
        restore = resume_url if is_cnki_work_page_url(resume_url) else None
        if restore:
            kind = '结果页' if is_cnki_results_page_url(restore) else '高级检索页'
            print(f'当前非工作页，恢复{kind}...')
            driver.get(to_cnki_access_url(restore))
        else:
            print('当前非检索页，重新打开高级检索页...')
            driver.get(get_adv_search_url())
        time.sleep(3)
    kept = driver.current_window_handle
    for handle in list(driver.window_handles):
        if handle == kept or len(driver.window_handles) <= 1:
            continue
        try:
            driver.switch_to.window(handle)
            if is_cnki_portal_home_url(driver.current_url) or any(
                    k in (driver.current_url or '').lower()
                    for k in sso_keywords):
                print(f'关闭多余标签: {driver.current_url}')
                driver.close()
        except Exception:
            continue
    if kept in driver.window_handles:
        driver.switch_to.window(kept)
        return kept
    if driver.window_handles:
        driver.switch_to.window(driver.window_handles[0])
        return driver.current_window_handle
    return None


def _prepare_unit_login_dialog(driver, restore_handle=None, original=None):
    """定位并打开机构登录弹窗，返回 login_handle 或 None。"""
    login_handle = None
    header_handle = None
    for handle in list(driver.window_handles):
        try:
            driver.switch_to.window(handle)
        except Exception:
            continue
        url = ''
        try:
            url = driver.current_url or ''
        except Exception:
            pass
        if find_visible_unit_login_box(driver) is not None:
            login_handle = handle
            break
        if is_export_page_url(url):
            continue
        if header_handle is None and is_unit_header_login_visible(driver):
            header_handle = handle

    if not login_handle:
        prefer = None
        if restore_handle and restore_handle in driver.window_handles:
            prefer = restore_handle
        elif header_handle:
            prefer = header_handle
        if prefer:
            driver.switch_to.window(prefer)
            if is_unit_header_login_visible(driver) or prefer == restore_handle:
                print('检测到页头「机构登录」可见（未登录），打开登录框...')
                open_unit_login_dialog(driver)
                deadline_open = time.time() + 8
                while time.time() < deadline_open:
                    if find_visible_unit_login_box(driver) is not None:
                        login_handle = prefer
                        break
                    time.sleep(0.4)
                if not login_handle:
                    login_handle = prefer

    if not login_handle:
        if restore_handle and restore_handle in driver.window_handles:
            driver.switch_to.window(restore_handle)
        elif original and original in driver.window_handles:
            driver.switch_to.window(original)
    return login_handle


def try_unit_ip_login_reconnect(driver, restore_handle=None, timeout=None):
    """机构未登录重连：按配置使用 VPN / IP登录 / 校外访问。"""
    method = get_unit_login_method()
    if method == 'vpn':
        return try_unit_vpn_login(
            driver, restore_handle=restore_handle,
            timeout=timeout if timeout is not None else 180)
    if method == 'school_visit':
        return try_unit_school_visit_reconnect(
            driver, restore_handle=restore_handle,
            timeout=timeout if timeout is not None else 180)

    if timeout is None:
        timeout = 45
    original = None
    try:
        original = driver.current_window_handle
    except Exception:
        pass

    login_handle = _prepare_unit_login_dialog(
        driver, restore_handle=restore_handle, original=original)
    if not login_handle:
        return False

    driver.switch_to.window(login_handle)
    if find_visible_unit_login_box(driver) is not None:
        print('检测到机构登录弹窗，尝试 IP 登录...')
    else:
        print('在检索页尝试 IP 登录...')

    if not _click_unit_ip_login(driver):
        if open_unit_login_dialog(driver):
            time.sleep(1)
            if not _click_unit_ip_login(driver):
                print('未能触发 IP 登录')
                if restore_handle and restore_handle in driver.window_handles:
                    driver.switch_to.window(restore_handle)
                elif original and original in driver.window_handles:
                    driver.switch_to.window(original)
                return False
        else:
            print('未能触发 IP 登录')
            if restore_handle and restore_handle in driver.window_handles:
                driver.switch_to.window(restore_handle)
            elif original and original in driver.window_handles:
                driver.switch_to.window(original)
            return False

    verify_handle = restore_handle or login_handle
    deadline = time.time() + timeout
    ok = False
    while time.time() < deadline:
        time.sleep(1)
        try:
            if verify_handle and verify_handle in driver.window_handles:
                driver.switch_to.window(verify_handle)
        except Exception:
            pass
        box_gone = not has_unit_login_anywhere(driver)
        header_ok = False
        try:
            if verify_handle and verify_handle in driver.window_handles:
                driver.switch_to.window(verify_handle)
            if is_export_page_url(driver.current_url or ''):
                if restore_handle and restore_handle in driver.window_handles:
                    driver.switch_to.window(restore_handle)
            header_ok = is_institution_logged_in(driver)
            els = driver.find_elements(By.CSS_SELECTOR, '.ecp_header_unit_loginbg')
            if not els:
                header_ok = True
        except Exception:
            header_ok = box_gone

        if header_ok and box_gone:
            ok = True
            break

    if ok:
        print('机构 IP 登录成功（检索页页头「机构登录」已隐藏）')
    else:
        print('机构 IP 登录超时：页头仍显示未登录或登录框仍在')

    target = restore_handle or original
    if target and target in driver.window_handles:
        try:
            driver.switch_to.window(target)
        except Exception:
            pass
    return ok


def try_unit_school_visit_reconnect(driver, restore_handle=None, timeout=180):
    """校外访问重连：点「校外访问」并完成 FSSO/IdP/CAS 流程。"""
    creds = load_unit_credentials()
    if not creds.get('username') or not creds.get('password'):
        print('校外访问需要 cnki_unit_credentials.json 中的 username/password')
        return False

    original = None
    try:
        original = driver.current_window_handle
    except Exception:
        pass

    work_handle, resume_url, preserve_results = _capture_work_page(
        driver, restore_handle or original)
    if resume_url:
        print(f'校外访问: 登录后将回到工作页 {resume_url}')
    if preserve_results:
        print('校外访问: 检测到已有检索结果，登录后将保留结果页')

    login_handle = _prepare_unit_login_dialog(
        driver, restore_handle=restore_handle, original=original)
    if not login_handle:
        return False

    driver.switch_to.window(login_handle)
    print('检测到机构未登录，尝试「校外访问」登录...')
    handles_before = set(driver.window_handles)

    if not _click_unit_school_visit(driver):
        if open_unit_login_dialog(driver):
            time.sleep(1)
            if not _click_unit_school_visit(driver):
                print('未能触发校外访问')
                return False
        else:
            print('未能触发校外访问')
            return False

    time.sleep(2)
    new_handles = [h for h in driver.window_handles if h not in handles_before]
    if new_handles:
        driver.switch_to.window(new_handles[-1])
        print('校外访问: 已切换到新窗口', driver.current_url)
    else:
        _focus_sso_related_window(driver, main_handle=restore_handle or login_handle)

    verify_handle = work_handle or restore_handle or login_handle
    ok = complete_school_visit_login(
        driver, creds, main_handle=verify_handle,
        resume_url=resume_url, preserve_results=preserve_results,
        timeout=timeout)

    target_handle = ensure_single_search_window(
        driver, prefer_handle=verify_handle, resume_url=resume_url)
    if ok:
        try:
            if not target_handle:
                print('校外访问: 无可用窗口')
                ok = False
            else:
                if preserve_results or _has_search_results_ui(driver):
                    print('机构校外访问登录成功，继续原检索/导出流程')
                else:
                    url = driver.current_url or ''
                    if not is_cnki_work_page_url(url):
                        target_handle = _refresh_search_login_state(
                            driver, target_handle, resume_url=resume_url,
                            preserve_results=False)
                    if not needs_unit_ip_login(driver, check_handle=target_handle):
                        print('机构校外访问登录成功，继续原检索/导出流程')
                    else:
                        print('校外访问流程完成，但页头仍显示未登录')
                        ok = False
        except Exception as e:
            print('校外访问登录后确认状态失败:', e)
    else:
        print('机构校外访问登录失败或超时')

    kept = ensure_single_search_window(
        driver,
        prefer_handle=target_handle or restore_handle or original,
        resume_url=resume_url)
    if kept:
        try:
            driver.switch_to.window(kept)
        except Exception:
            pass
    return ok


def ensure_unit_login_if_needed(driver, restore_handle=None):
    """检查登录状态；未登录则按配置执行 VPN / IP登录 / 校外访问。"""
    if restore_handle and restore_handle in driver.window_handles:
        try:
            driver.switch_to.window(restore_handle)
        except Exception:
            pass

    method = get_unit_login_method()
    if method == 'vpn':
        need = needs_webvpn_login(driver, check_handle=restore_handle)
        if not need:
            need = needs_unit_ip_login(driver, check_handle=restore_handle)
        if not need:
            return True
        try:
            cur = driver.current_url or ''
        except Exception:
            cur = ''
        # 已直连知网：按当前落地页域名拼接 WebVPN，避免整段重登把页面冲掉
        if (not is_webvpn_url(cur) and is_cnki_host_url(cur)
                and not is_webvpn_portal_login_url(cur)
                and '/cas/login' not in cur.lower()):
            vpn_url = to_webvpn_url(unwrap_webvpn_url(cur))
            print(f'当前为直连知网，按落地域名拼接 WebVPN: {_short_url(vpn_url)}')
            try:
                driver.get(vpn_url)
                time.sleep(2)
                wait_url_settle(driver)
            except Exception as e:
                print('拼接 WebVPN 跳转失败:', e)
            if not needs_webvpn_login(driver, check_handle=restore_handle):
                if not needs_unit_ip_login(driver, check_handle=restore_handle):
                    return True
            try:
                cur = driver.current_url or ''
            except Exception:
                cur = ''
        if _is_cnki_via_webvpn(cur):
            print('WebVPN 已进入知网，尝试 IP 登录补全机构身份...')
            if open_unit_login_dialog(driver):
                time.sleep(0.8)
            if _click_unit_ip_login(driver):
                time.sleep(2)
                if not needs_unit_ip_login(driver, check_handle=restore_handle):
                    return True
        print('WebVPN 未认证或知网机构未登录，开始 VPN 网站登录...')
        return try_unit_vpn_login(driver, restore_handle=restore_handle)

    if not needs_unit_ip_login(driver, check_handle=restore_handle):
        return True

    label = '校外访问' if method == 'school_visit' else 'IP登录'
    print(f'机构未登录（页头「机构登录」可见或登录弹窗出现），开始{label}...')
    return try_unit_ip_login_reconnect(driver, restore_handle=restore_handle)


def wait_for_export_dialog(driver, export_handle, save_dir, main_handle, timeout=90):
    """等待导出弹窗出现 row-btns；仅在出现登录弹窗时重连，不因导出页页头误判。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if export_handle not in driver.window_handles:
            print('导出窗口已关闭')
            return False
        driver.switch_to.window(export_handle)
        if driver.find_elements(By.CLASS_NAME, 'row-btns'):
            return True
        alert_text = dismiss_alerts(driver, timeout=0.2)
        if text_suggests_login_lost(alert_text):
            print('导出弹窗加载中提示登录状态中断，需重建浏览器重新登录')
            return 'restart_browser'
        # 仅登录弹窗才算断线；导出页页头可见不算
        if has_unit_login_anywhere(driver):
            print('导出窗口加载中出现机构登录弹窗，正在重连...')
            close_export_windows(driver, main_handle=main_handle)
            ensure_unit_login_if_needed(driver, restore_handle=main_handle)
            return False  # 需外层重新打开导出页
        if has_captcha_anywhere(driver):
            print('导出窗口加载中遇到验证码，正在处理...')
            cap = ensure_captcha_passed(
                driver, save_dir, main_handle, return_handle=export_handle)
            if captcha_needs_browser_restart(cap):
                return 'restart_browser'
            if cap is False:
                time.sleep(2)
                continue
            continue
        time.sleep(1)
    print('等待导出弹窗超时')
    return False


def is_request_too_frequent_text(text):
    """WebVPN/知网「请求过于频繁，请稍后重试」类提示。"""
    if not text:
        return False
    t = str(text).replace(' ', '').replace('\u3000', '')
    return ('过于频繁' in t) or ('请求频繁' in t) or ('频繁' in t and '稍后' in t)


EXPORT_RATE_LIMIT_WAIT_SEC = 300  # 频繁提示后等待 5 分钟


def sleep_after_export_rate_limit(seconds=None):
    """频繁提示后等待（关闭浏览器之后调用）。"""
    sec = EXPORT_RATE_LIMIT_WAIT_SEC if seconds is None else int(seconds)
    mins = sec / 60.0
    print(f'请求过于频繁：等待 {mins:g} 分钟后重新打开浏览器...')
    time.sleep(sec)
    print('等待结束，重新启动浏览器')


def recreate_browser_after_rate_limit(driver, download_dir, doc_type=None):
    """关闭整个浏览器会话，等待后新建实例并打开检索页。返回新 driver。

    用于导出「请求过于频繁」、验证码连败或登录中断：仅关导出页无效，
    必须整浏览器重建。登录中断时跳过频繁等待，直接重新登录。
    """
    import gc
    login_lost = False
    try:
        if driver is not None and is_driver_alive(driver):
            login_lost = page_suggests_login_lost(driver)
    except Exception:
        login_lost = False
    print('关闭整个浏览器（仅关导出页无法解除频繁限制）...')
    try:
        if driver is not None:
            driver.quit()
    except Exception as e:
        print('关闭浏览器失败:', e)
    cleanup_chrome_profile()
    gc.collect()
    if login_lost:
        print('检测到登录状态中断：跳过频繁等待，直接重建并重新登录...')
    else:
        sleep_after_export_rate_limit()
    new_driver = create_chrome_driver(download_dir)
    open_cnki_url(new_driver, get_adv_search_url(doc_type))
    try:
        new_driver.maximize_window()
    except Exception:
        pass
    login_ok = ensure_unit_login_if_needed(
        new_driver, restore_handle=new_driver.current_window_handle)
    cap = ensure_cnki_captcha_after_nav(
        new_driver, download_dir, new_driver.current_window_handle)
    still_login = page_suggests_login_lost(
        new_driver, check_handle=new_driver.current_window_handle)
    if captcha_needs_browser_restart(cap):
        print('浏览器重建后验证码仍连败，将由上层再次重启或记失败')
    elif not cap:
        print('浏览器重建后验证码未通过，将由上层重试')
    elif still_login or not login_ok:
        print('浏览器重建后仍停在 WebVPN/CAS 登录页，'
              '请检查 VPN 凭据/账号状态或会话超时策略；将由上层决定是否记失败')
    else:
        print('浏览器已重建并完成登录/验证码，即将继续导出题录')
    return new_driver


def wait_for_export_file(driver, export_handle, save_dir, main_handle, before_files, timeout=120):
    """点击导出后等待文件下载；仅登录弹窗触发重连提示。

    返回:
      True            — 已下到新 Excel
      False           — 超时/失败
      'rate_limited'  — 出现请求过于频繁（调用方应关窗等待后重试）
    """
    deadline = time.time() + timeout
    retried_download = False
    captcha_seen = False
    login_retried = False
    while time.time() < deadline:
        alert_text = dismiss_alerts(driver, timeout=0.3)
        if text_suggests_login_lost(alert_text):
            print('导出等待中提示登录状态中断，需重建浏览器重新登录')
            return 'restart_browser'
        if is_request_too_frequent_text(alert_text):
            print('导出等待中出现「请求过于频繁」')
            return 'rate_limited'

        if has_unit_login_anywhere(driver):
            print('导出过程中出现机构登录弹窗，正在重连...')
            if ensure_unit_login_if_needed(driver, restore_handle=main_handle):
                login_retried = True
            time.sleep(2)
            continue

        if has_captcha_anywhere(driver):
            captcha_seen = True
            print('导出过程中遇到验证码，正在处理...')
            cap = ensure_captcha_passed(
                driver, save_dir, main_handle, return_handle=export_handle)
            if captcha_needs_browser_restart(cap):
                return 'restart_browser'
            if cap is False:
                time.sleep(2)
                continue

        new_file = get_new_excel_file(save_dir, before_files)
        if new_file:
            print(f'导出成功，新文件: {new_file}')
            return True

        need_retry_click = (
            (not retried_download)
            and (captcha_seen or login_retried)
            and export_handle in driver.window_handles
        )
        if need_retry_click:
            driver.switch_to.window(export_handle)
            if driver.find_elements(By.CLASS_NAME, 'row-btns'):
                try:
                    print('登录/验证码已处理，重新触发导出下载...')
                    cur_selall = driver.find_element(
                        By.CLASS_NAME, 'row-btns').find_elements(By.CSS_SELECTOR, 'a')[0]
                    driver.execute_script("arguments[0].click();", cur_selall)
                    download_excel = driver.find_element(
                        By.ID, 'litoexcel').find_element(By.CSS_SELECTOR, 'a')
                    driver.execute_script("arguments[0].click();", download_excel)
                    alert_text = dismiss_alerts(driver, timeout=2.0)
                    if is_request_too_frequent_text(alert_text):
                        print('重试导出时出现「请求过于频繁」')
                        return 'rate_limited'
                    retried_download = True
                    login_retried = False
                except Exception as e:
                    print('重试导出失败:', e)

        if export_handle in driver.window_handles:
            driver.switch_to.window(export_handle)
        time.sleep(2)

    print('等待导出文件超时，未检测到新 Excel')
    return False


def wait_for_result_page(driver, save_dir=None, main_handle=None, timeout=60):
    """等待检索结果页就绪（出现 perPageDiv / 无结果提示），期间处理验证码。

    返回：
        True              — 结果页已就绪（含无结果）
        False             — 超时仍未就绪
        'restart_browser' — 验证码连败，需关浏览器
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        if has_captcha_anywhere(driver):
            print('等待结果页时遇到验证码，正在处理...')
            if save_dir:
                cap = ensure_captcha_passed(driver, save_dir, main_handle)
                if captcha_needs_browser_restart(cap):
                    return 'restart_browser'
            else:
                time.sleep(2)
            continue

        if driver.find_elements(By.ID, 'perPageDiv'):
            return True

        # 无检索结果等提示页
        page_text = ''
        try:
            page_text = driver.find_element(By.TAG_NAME, 'body').text
        except Exception:
            pass
        if text_suggests_login_lost(page_text):
            print('等待结果页时检测到登录状态中断')
            return 'restart_browser'
        if any(kw in page_text for kw in ('找不到', '没有找到', '暂无数据', '0条结果', '检索结果：0')):
            print('检索结果为空或不存在 perPageDiv')
            return True

        # 结果列表或全选控件也视为已进入结果页
        if driver.find_elements(By.CLASS_NAME, 'checkAll'):
            return True

        if page_suggests_login_lost(driver, check_handle=main_handle):
            print('等待结果页时检测到登录中断（VPN/机构登录失效）')
            return 'restart_browser'

        time.sleep(1)

    print('等待检索结果页超时')
    return False


def set_result_page_size(driver, page_size=50, save_dir=None, main_handle=None):
    """将检索结果每页显示数量设置为指定值。

    返回 True / False / 'restart_browser'。
    """
    print(f'设置每页显示 {page_size} 条')
    ready = wait_for_result_page(
        driver, save_dir=save_dir, main_handle=main_handle)
    if captcha_needs_browser_restart(ready):
        return 'restart_browser'
    if not ready:
        return False

    per_page_divs = driver.find_elements(By.ID, 'perPageDiv')
    if not per_page_divs:
        print('未找到每页条数控件，跳过设置')
        return True

    try:
        per_page_div = per_page_divs[0]
        current = per_page_div.find_element(
            By.CSS_SELECTOR, '.sort-default span').text.strip()
        if current == str(page_size):
            print(f'当前已为每页 {page_size} 条，跳过')
            return True
        sort_default = per_page_div.find_element(By.CLASS_NAME, 'sort-default')
        driver.execute_script("arguments[0].click();", sort_default)
        time.sleep(random.randint(1, 2))
        option = per_page_div.find_element(
            By.CSS_SELECTOR, f'.sort-list li[data-val="{page_size}"] a')
        driver.execute_script("arguments[0].click();", option)
        time.sleep(3)
        # 验证每页条数是否真正生效，避免专利页仍按大页数累计已选
        try:
            per_page_divs = driver.find_elements(By.ID, 'perPageDiv')
            if per_page_divs:
                current = per_page_divs[0].find_element(
                    By.CSS_SELECTOR, '.sort-default span').text.strip()
                if current != str(page_size):
                    print(f'每页条数未生效（当前 {current}），重试一次')
                    sort_default = per_page_divs[0].find_element(
                        By.CLASS_NAME, 'sort-default')
                    driver.execute_script("arguments[0].click();", sort_default)
                    time.sleep(1)
                    option = per_page_divs[0].find_element(
                        By.CSS_SELECTOR, f'.sort-list li[data-val="{page_size}"] a')
                    driver.execute_script("arguments[0].click();", option)
                    time.sleep(3)
        except Exception as e:
            print('验证每页条数失败:', e)
        return True
    except Exception as e:
        print('设置每页条数失败:', e)
        return False

def close_export_windows(driver, main_handle=None):
    """关闭所有知网导出页（export.html / dm8/manage/export），并切回主窗口。

    重连后多次点击「自定义」可能留下多个导出页，必须全部关掉，否则后续检索会失败。
    """
    closed = 0
    for handle in list(driver.window_handles):
        if main_handle and handle == main_handle:
            continue
        try:
            driver.switch_to.window(handle)
            url = driver.current_url or ''
            if 'export.html' in url or '/dm8/manage/export' in url:
                print(f'关闭导出页: {url}')
                driver.close()
                closed += 1
        except Exception as e:
            print(f'关闭导出页失败: {e}')
    if main_handle and main_handle in driver.window_handles:
        try:
            driver.switch_to.window(main_handle)
        except Exception:
            pass
    elif driver.window_handles:
        try:
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass
    if closed:
        print(f'已关闭 {closed} 个导出页窗口')
    return closed


def downloadexcel(driver, save_dir, main_handle=None):
    """打开导出窗口并下载 Excel，失败时返回 False 且不中断主流程。

    流程：
      1) 导出前：在检索主页检查页头「机构登录」，未登录则 IP 登录
      2) 点击「自定义」打开导出页
      3) 导出页打开后：仅当出现登录弹窗才重连并重开；不以导出页页头误判
    """
    export_handle = None
    before_files = snapshot_excel_files(save_dir)
    handles_before = set(driver.window_handles)
    max_login_retries = 2

    def click_custom_export():
        if main_handle and main_handle in driver.window_handles:
            driver.switch_to.window(main_handle)
        close_export_windows(driver, main_handle=main_handle)
        handles_before.clear()
        handles_before.update(driver.window_handles)
        # 翻页/登录中断后可能已不在结果页，先等工具栏再判断，避免直接抛 no such element
        try:
            WebDriverWait(driver, 10, 0.5).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'secondUl')))
        except Exception:
            if page_suggests_login_lost(driver, check_handle=main_handle):
                raise SessionLostError(
                    '导出工具栏 .secondUl 缺失，登录状态已中断')
            raise
        li_uploadlist = driver.find_element(
            By.CLASS_NAME, 'secondUl').find_elements(By.CSS_SELECTOR, 'li')
        selfdef = li_uploadlist[-1].find_element(By.CSS_SELECTOR, 'a')
        print('自定义下载')
        driver.execute_script("arguments[0].click();", selfdef)

    def find_new_export_handle():
        """在已打开窗口中查找导出页。"""
        for handle in list(driver.window_handles):
            if handle in handles_before and handle != main_handle:
                # 也可能是新开但仍在 handles_before 之外
                pass
            if handle in handles_before:
                continue
            try:
                driver.switch_to.window(handle)
                if is_export_page_url(driver.current_url):
                    return handle
            except Exception:
                continue
        # 兜底：任意 export 页
        for handle in list(driver.window_handles):
            if main_handle and handle == main_handle:
                continue
            try:
                driver.switch_to.window(handle)
                if is_export_page_url(driver.current_url):
                    return handle
            except Exception:
                continue
        return None

    def wait_new_export_window(timeout=25):
        """等待导出页出现。优先认窗；仅登录弹窗才打断重连。"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            # 先认导出窗，避免把已打开的窗当未登录关掉
            h = find_new_export_handle()
            if h:
                return h
            if len(driver.window_handles) > len(handles_before):
                # 新窗但 URL 可能尚未跳到 export
                for handle in driver.window_handles:
                    if handle not in handles_before:
                        try:
                            driver.switch_to.window(handle)
                            time.sleep(0.5)
                            if is_export_page_url(driver.current_url):
                                return handle
                            # 新窗里若是登录弹窗
                            if find_visible_unit_login_box(driver) is not None:
                                return None  # 交给外层重连
                        except Exception:
                            pass

            # 仅真实登录弹窗才算需要重连（不用页头，避免导出页误判）
            if has_unit_login_anywhere(driver):
                print('等待导出窗时出现机构登录弹窗')
                return None
            time.sleep(0.4)
        return find_new_export_handle()

    try:
        close_export_windows(driver, main_handle=main_handle)
        handles_before.clear()
        handles_before.update(driver.window_handles)

        # ----- 阶段1：导出前在主页检查登录 -----
        print('导出前检查机构登录状态（检索页页头）...')
        if main_handle and main_handle in driver.window_handles:
            driver.switch_to.window(main_handle)
        if not ensure_unit_login_if_needed(driver, restore_handle=main_handle):
            print('机构未登录且 IP 登录失败，跳过本次下载')
            return False

        for attempt in range(max_login_retries + 1):
            click_custom_export()
            time.sleep(1.2)

            # 点击后立刻出现登录弹窗 → 重连再开
            if has_unit_login_anywhere(driver):
                print('点击「自定义」后出现机构登录弹窗，IP 登录后重开导出...')
                close_export_windows(driver, main_handle=main_handle)
                if not ensure_unit_login_if_needed(driver, restore_handle=main_handle):
                    print('IP 登录失败，跳过本次下载')
                    return False
                time.sleep(1.5)
                continue

            export_handle = wait_new_export_window(timeout=25)

            # ----- 阶段2：导出页已开，再检查（仅弹窗） -----
            if export_handle:
                driver.switch_to.window(export_handle)
                print('导出窗口:', driver.current_url)
                if has_unit_login_anywhere(driver):
                    print('导出页打开后出现机构登录弹窗，关闭后重连...')
                    close_export_windows(driver, main_handle=main_handle)
                    if not ensure_unit_login_if_needed(
                            driver, restore_handle=main_handle):
                        print('导出页后 IP 登录失败，跳过本次下载')
                        return False
                    time.sleep(1.5)
                    export_handle = None
                    continue

                dialog_ok = wait_for_export_dialog(
                    driver, export_handle, save_dir, main_handle)
                if dialog_ok == 'restart_browser':
                    return 'restart_browser'
                if dialog_ok is True:
                    break

                # dialog 未就绪：若因登录弹窗返回 False，重试；否则失败
                if has_unit_login_anywhere(driver) or needs_unit_ip_login(
                        driver, check_handle=main_handle):
                    print('导出弹窗未就绪且需登录，重连后重试...')
                    close_export_windows(driver, main_handle=main_handle)
                    if not ensure_unit_login_if_needed(
                            driver, restore_handle=main_handle):
                        return False
                    time.sleep(1.5)
                    continue
                print('导出弹窗未就绪，跳过本次下载')
                return False

            # 未打开导出窗
            if attempt < max_login_retries and (
                    has_unit_login_anywhere(driver)
                    or needs_unit_ip_login(driver, check_handle=main_handle)):
                print('导出窗未打开，检查登录后重试...')
                close_export_windows(driver, main_handle=main_handle)
                if not ensure_unit_login_if_needed(
                        driver, restore_handle=main_handle):
                    return False
                time.sleep(1.5)
                continue
            print('导出窗口未打开，跳过本次下载')
            return False
        else:
            print('导出重试次数用尽，跳过本次下载')
            return False

        if not export_handle or export_handle not in driver.window_handles:
            print('导出窗口无效，跳过本次下载')
            return False

        driver.switch_to.window(export_handle)
        cur_selall = driver.find_element(
            By.CLASS_NAME, 'row-btns').find_elements(By.CSS_SELECTOR, 'a')[0]
        driver.execute_script("arguments[0].click();", cur_selall)
        download_excel = driver.find_element(
            By.ID, 'litoexcel').find_element(By.CSS_SELECTOR, 'a')
        driver.execute_script("arguments[0].click();", download_excel)
        print('已点击导出 Excel，等待下载与验证码...')

        # 点击后立刻检查频繁/登录中断提示
        alert_text = dismiss_alerts(driver, timeout=2.5)
        if text_suggests_login_lost(alert_text):
            print('导出时提示登录状态中断，需重建浏览器重新登录')
            return 'restart_browser'
        rate_hit = is_request_too_frequent_text(alert_text)
        if not rate_hit:
            file_result = wait_for_export_file(
                driver, export_handle, save_dir, main_handle, before_files)
            if file_result is True:
                return True
            if file_result == 'restart_browser':
                print('导出等待中验证码连败，需关闭浏览器后重试')
                try:
                    close_export_windows(driver, main_handle=main_handle)
                except Exception:
                    pass
                return 'restart_browser'
            rate_hit = (file_result == 'rate_limited')
            if not rate_hit:
                print('导出未成功（无新文件或验证码未通过）')
                return False

        # 仅关导出页无效：通知上层关闭整个浏览器再重试
        print('导出触发请求过于频繁，需关闭浏览器后重试')
        try:
            close_export_windows(driver, main_handle=main_handle)
        except Exception:
            pass
        return 'restart_browser'
    except SessionLostError as e:
        print('downloadexcel 中止：登录状态已中断，需重建浏览器重新登录:', e)
        return 'restart_browser'
    except Exception as e:
        print('downloadexcel 失败:', type(e).__name__, e)
        if exception_suggests_login_lost(e, driver, check_handle=main_handle):
            print('检测到登录状态中断，需重建浏览器重新登录')
            return 'restart_browser'
        if has_unit_login_anywhere(driver) or needs_unit_ip_login(
                driver, check_handle=main_handle):
            print('导出异常且检测到机构未登录，尝试 IP 登录...')
            ensure_unit_login_if_needed(driver, restore_handle=main_handle)
        return False
    finally:
        # restart_browser 时上层会 quit，此处仍尝试关导出页
        try:
            if driver is not None and getattr(driver, 'session_id', None):
                close_export_windows(driver, main_handle=main_handle)
        except Exception:
            pass


def dismiss_alerts(driver, timeout=1.0):
    """关闭浏览器原生 alert/confirm，返回提示文本（无则 None）。"""
    texts = []
    deadline = time.time() + max(timeout, 0.1)
    while True:
        try:
            alert = driver.switch_to.alert
            text = (alert.text or '').strip()
            print(f'关闭弹窗提示: {text}')
            alert.accept()
            texts.append(text)
            time.sleep(0.2)
        except Exception:
            if time.time() >= deadline:
                break
            time.sleep(0.1)
            if not texts:
                try:
                    WebDriverWait(driver, 0.2).until(EC.alert_is_present())
                except Exception:
                    break
    return ' | '.join(texts) if texts else None


def hide_datetime_pickers(driver):
    """隐藏 xdsoft 日期选择器，避免挡住「清空」等按钮。"""
    try:
        driver.execute_script("""
            document.querySelectorAll('.xdsoft_datetimepicker').forEach(function(el) {
                el.style.display = 'none';
                el.style.visibility = 'hidden';
                el.classList.remove('active');
            });
            document.querySelectorAll('#datebox0, #datebox1, input.datebox').forEach(
                function(el) { try { el.blur(); } catch (e) {} }
            );
            try { document.body.click(); } catch (e) {}
        """)
    except Exception:
        pass


def get_selected_count(driver):
    """读取结果页已选条数；失败返回 None。"""
    try:
        box = driver.find_element(By.CLASS_NAME, 'checkcount')
        text = (box.text or '') + ' ' + (box.get_attribute('innerText') or '')
        m = re.search(r'(\d+)', text.replace(',', ''))
        return int(m.group(1)) if m else None
    except Exception:
        return None


def clear_selected_results(driver, main_handle=None, retries=3):
    """清空已选文献。用 JS/jQuery 调用，避开日期控件遮挡。"""
    if main_handle and main_handle in driver.window_handles:
        try:
            driver.switch_to.window(main_handle)
        except Exception:
            pass
    dismiss_alerts(driver, timeout=0.5)

    for attempt in range(1, retries + 1):
        hide_datetime_pickers(driver)
        try:
            ok = driver.execute_script("""
                var link = document.querySelector('.checkcount a');
                if (!link) return false;
                try {
                    if (window.jQuery) {
                        var $el = jQuery(link);
                        if (typeof $el.filenameClear === 'function') {
                            $el.filenameClear();
                            return true;
                        }
                    }
                } catch (e) {}
                try {
                    if (typeof link.filenameClear === 'function') {
                        link.filenameClear();
                        return true;
                    }
                } catch (e) {}
                link.click();
                return true;
            """)
            if not ok:
                link = driver.find_element(By.CSS_SELECTOR, '.checkcount a')
                driver.execute_script('arguments[0].click();', link)
            time.sleep(1)
            dismiss_alerts(driver, timeout=0.5)
            count = get_selected_count(driver)
            if count is None or count == 0:
                print('已清空已选')
                return True
            print(f'清空后仍显示已选 {count}，重试 ({attempt}/{retries})')
            # 兜底：再次用原生 JS 点击
            try:
                link = driver.find_element(By.CSS_SELECTOR, '.checkcount a')
                driver.execute_script('arguments[0].click();', link)
                time.sleep(1)
            except Exception:
                pass
        except Exception as e:
            print(f'清空已选失败(第{attempt}次): {e}')
            dismiss_alerts(driver, timeout=0.5)
            time.sleep(0.5)
    count = get_selected_count(driver)
    print(f'清空已选最终失败，当前已选={count}')
    return False


def run_download_excel_step(driver, save_dir, main_handle):
    """执行 Excel 导出并在成功后清空已选。

    返回:
      True              — 导出成功且已选已清空
      False             — 普通失败
      'restart_browser' — 请求过于频繁 / 登录中断，需关闭浏览器重建会话
    """
    result = downloadexcel(driver, save_dir, main_handle)
    if result == 'restart_browser':
        print('Excel 导出中止，需关闭浏览器重建会话（频繁/验证码/登录中断）')
        return 'restart_browser'
    if result is True:
        if main_handle and main_handle in driver.window_handles:
            driver.switch_to.window(main_handle)
        if clear_selected_results(driver, main_handle):
            print('下载完成一次===========================\n')
            return True
        print('导出成功但清空已选失败，标记本批异常')
        return False
    print('Excel 导出未完成，跳过清空已选')
    try:
        close_export_windows(driver, main_handle=main_handle)
        if main_handle and main_handle in driver.window_handles:
            driver.switch_to.window(main_handle)
    except Exception:
        pass
    return False


EXPORT_FAILURE_LOG = 'export_failures.txt'
EXPORT_OVER_6000_LOG = 'export_over_6000.txt'
SEARCH_HIT_LIMIT = 6000


def record_export_failure(save_dir, date_str):
    """将导出失败的日期追加写入导出目录下的 txt（去重）。"""
    path = os.path.join(save_dir, EXPORT_FAILURE_LOG)
    existing = set()
    if os.path.isfile(path):
        with open(path, 'r', encoding='utf-8') as f:
            existing = {line.strip() for line in f if line.strip()}
    if date_str in existing:
        print(f'导出失败日期已记录过: {date_str}')
        return
    with open(path, 'a', encoding='utf-8') as f:
        f.write(date_str + '\n')
    print(f'已记录导出失败日期: {date_str} -> {path}')


def read_export_failure_labels(save_dir):
    """读取失败日志中的原始时段标签（去重、保序）。"""
    path = os.path.join(save_dir, EXPORT_FAILURE_LOG)
    if not os.path.isfile(path):
        return []
    out = []
    seen = set()
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            label = line.strip()
            if not label or label.startswith('#') or label in seen:
                continue
            seen.add(label)
            out.append(label)
    return out


def remove_export_failure(save_dir, date_str):
    """从失败日志中删除指定时段。"""
    path = os.path.join(save_dir, EXPORT_FAILURE_LOG)
    if not os.path.isfile(path):
        return False
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    kept = []
    removed = False
    for line in lines:
        if line.strip() == date_str:
            removed = True
            continue
        kept.append(line)
    if not removed:
        return False
    with open(path, 'w', encoding='utf-8') as f:
        f.writelines(kept)
    print(f'已从失败日志移除: {date_str}')
    return True


def record_over_6000(save_dir, label, category_path, hit_count):
    """标记检索结果超过 6000、且无法再拆子类的时段。"""
    path = os.path.join(save_dir, EXPORT_OVER_6000_LOG)
    cat = ' / '.join(category_path) if isinstance(category_path, list) else str(category_path)
    row = f'{label}\t{cat}\t{hit_count}\n'
    existing = set()
    if os.path.isfile(path):
        with open(path, 'r', encoding='utf-8') as f:
            existing = {line.strip() for line in f if line.strip()}
    if row.strip() in existing:
        print(f'超限已记录过: {row.strip()}')
        return
    with open(path, 'a', encoding='utf-8') as f:
        f.write(row)
    print(f'检索超过 {SEARCH_HIT_LIMIT} 条，已标记: {path} | {row.strip()}')


def category_children(category_path, doc_type=None):
    """当前分类在对应分类树中的直接子类名。"""
    node = get_category_tree(doc_type)
    path = list(category_path or [])
    for name in path:
        if not isinstance(node, dict) or name not in node:
            return []
        node = node[name]
    if not isinstance(node, dict) or not node:
        return []
    return list(node.keys())


def _parse_hit_number(text):
    text = (text or '').replace(',', '').replace('，', '').replace(' ', '')
    m = re.search(r'(\d+)', text)
    if not m:
        return None
    try:
        return int(m.group(1))
    except ValueError:
        return None


def read_search_hit_count(driver):
    """解析结果页总条数；无结果返回 0；无法解析返回 None。"""
    selectors = (
        '#countPage',
        '.pagerTitleCell',
        'em.count',
        '#briefBox .pagerTitleCell',
        'span#countPage',
        '.countPage',
    )
    for selector in selectors:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, selector)
        except Exception:
            continue
        for el in els:
            try:
                text = (el.text or el.get_attribute('innerText') or '').strip()
            except Exception:
                continue
            n = _parse_hit_number(text)
            if n is not None:
                print(f'检索结果条数: {n} ({selector})')
                return n
    page_text = ''
    try:
        page_text = driver.find_element(By.TAG_NAME, 'body').text or ''
    except Exception:
        pass
    for pat in (
            r'共找到\s*([\d,，]+)\s*条',
            r'找到\s*([\d,，]+)\s*条结果',
            r'检索结果[：:]\s*([\d,，]+)',
            r'共\s*([\d,，]+)\s*条',
    ):
        m = re.search(pat, page_text)
        if m:
            n = _parse_hit_number(m.group(1))
            if n is not None:
                print(f'检索结果条数: {n} (正文)')
                return n
    if any(kw in page_text for kw in (
            '找不到', '没有找到', '暂无数据', '0条结果', '检索结果：0', '检索结果:0')):
        print('检索结果条数: 0（无结果页）')
        return 0
    print('未能解析检索结果条数')
    return None


PATENT_URL_COLUMNS = ('URL-网址', 'URL', '网址', 'Link', '链接')
PATENT_TITLE_COLUMNS = (
    'Title-题名', 'Title', '题名', '篇名', '专利名称', '专利题名',
)
PATENT_AUTHOR_COLUMNS = (
    'Author-作者', 'Author', '作者', '申请人', '发明人',
)
PATENT_DATE_COLUMNS = {
    'AD': ('申请日', '申请日期', 'ApplicationDate', 'Application Date'),
    'OD': (
        'PubTime-公开日期', '公开日期', '公开(公告)日', '公开（公告）日',
        '公开日', '公告日', 'PublicationDate', 'Publication Date',
    ),
}
PATENT_COMMON_DATE_COLUMNS = (
    'PubTime-发表时间', '发表时间', '出版日期', '出版时间', '日期',
)
_LOCAL_RECORDS_CACHE = {}
_LOCAL_RECORDS_CACHE_MAX_FILES = 5000


def _read_record_table(path):
    """读取题录表头与数据行，兼容 HTML 伪 xls 与真 Excel。"""
    from cnki_summary import TR_RE, _decode_xls, _is_html_xls, _td_texts
    path = Path(path)
    if path.suffix.lower() == '.xls' and _is_html_xls(path):
        rows = TR_RE.findall(_decode_xls(path))
        if not rows:
            return [], []
        headers = [h.strip() for h in _td_texts(rows[0])]
        body = [
            [cell.strip() for cell in _td_texts(row_html)]
            for row_html in rows[1:]
        ]
        return headers, body
    import pandas as pd
    df = pd.read_excel(path)
    headers = [str(c).strip() for c in df.columns]
    body = []
    for row in df.itertuples(index=False, name=None):
        cells = []
        for value in row:
            text = '' if value is None else str(value).strip()
            if text.lower() in ('nan', 'none'):
                text = ''
            cells.append(text)
        body.append(cells)
    return headers, body


def _pick_column(headers, candidates):
    """按候选列名找列下标；先精确匹配，再退化到包含匹配。"""
    from cnki_summary import _pick_index
    idx = _pick_index(headers, candidates)
    if idx is not None:
        return idx
    lowered = [h.strip().lower() for h in headers]
    for name in candidates:
        key = name.strip().lower()
        if not key:
            continue
        for i, header in enumerate(lowered):
            if header == key:
                return i
        for i, header in enumerate(lowered):
            if len(key) >= 2 and key in header:
                return i
    return None


def _patent_date_column_candidates(patent_date_field):
    """专利日期列候选。

    知网专利题录通常只有「公开日期」列，没有「申请日」列，因此：
      - OD（公开日）检索：用公开日期核对本地文件；
      - AD（申请日）检索：本地文件日期无法反映检索区间，不参与核对。
    """
    field = normalize_patent_date_field(patent_date_field)
    candidates = list(PATENT_DATE_COLUMNS.get(field, ()))
    if field == 'OD':
        candidates.extend(PATENT_COMMON_DATE_COLUMNS)
    return tuple(candidates)


def iter_patent_file_records(path, patent_date_field=None):
    """专利题录：日期取申请日/公开日列（专利没有「发表时间」列）。"""
    try:
        headers, rows = _read_record_table(path)
    except Exception as e:
        print(f'读取专利题录失败 {path}: {e}')
        return
    i_url = _pick_column(headers, PATENT_URL_COLUMNS)
    i_title = _pick_column(headers, PATENT_TITLE_COLUMNS)
    i_author = _pick_column(headers, PATENT_AUTHOR_COLUMNS)
    i_date = _pick_column(
        headers, _patent_date_column_candidates(patent_date_field))
    for cells in rows:
        def cell(idx):
            if idx is None or idx >= len(cells):
                return ''
            return cells[idx]
        if i_title is None and i_url is None:
            continue
        yield cell(i_url), cell(i_title), cell(i_author), cell(i_date)


def _local_records_for_file(path, doc_type, patent_date_field):
    """解析单个题录文件为 (记录键, 日期) 列表，按修改时间缓存避免反复读盘。"""
    try:
        st = os.stat(str(path))
        stamp = (st.st_mtime_ns, st.st_size)
    except Exception:
        stamp = None
    doc_type = normalize_doc_type(doc_type)
    date_field = (
        normalize_patent_date_field(patent_date_field)
        if doc_type == 'patent' else '')
    cache_key = (str(path), doc_type, date_field, stamp)
    cached = _LOCAL_RECORDS_CACHE.get(cache_key)
    if cached is not None:
        return cached
    from cnki_summary import iter_file_records, parse_iso_date, record_key
    records = []
    try:
        source = (
            iter_patent_file_records(path, date_field)
            if doc_type == 'patent' else iter_file_records(path))
        for url, title, author, pub_text in source:
            key = record_key(url, title, author)
            if key in ('u:', 't:|'):
                continue
            pub = parse_iso_date(pub_text)
            if pub is None:
                continue
            records.append((key, pub))
    except Exception as e:
        print(f'读取本地题录失败 {Path(path).name}: {e}')
        records = []
    if len(_LOCAL_RECORDS_CACHE) >= _LOCAL_RECORDS_CACHE_MAX_FILES:
        _LOCAL_RECORDS_CACHE.clear()
    _LOCAL_RECORDS_CACHE[cache_key] = records
    return records


def local_period_stats(
        save_dir, start, end, doc_type=None, patent_date_field=None):
    """本地该时段去重条数，以及多数时间落在该时段的题录文件。

    专利题录没有「发表时间」列，按公开日期/申请日解析。
    """
    from cnki_summary import iter_excel_files
    folder = Path(save_dir)
    unique = set()
    majority = {}
    for path in iter_excel_files(folder):
        hits = Counter()
        for key, pub in _local_records_for_file(
                path, doc_type, patent_date_field):
            hits[pub] += 1
            if start <= pub <= end:
                unique.add(key)
        if hits:
            majority[path] = hits.most_common(1)[0][0]
    mapped = [
        path for path, pub in majority.items()
        if start <= pub <= end
    ]
    return len(unique), mapped


def delete_period_files(paths):
    n = 0
    for path in paths:
        try:
            path = Path(path)
            if path.is_file():
                path.unlink()
                print(f'已删除该时段题录: {path.name}')
                n += 1
        except Exception as e:
            print(f'删除失败 {path}: {e}')
    return n


EXPORT_META_FILE = 'export_meta.json'


def _normalize_meta_entry(item):
    """规范化 export_meta.json 中的一条任务参数。"""
    if not isinstance(item, dict):
        return None
    doc_type = normalize_doc_type(item.get('doc_type'))
    category = item.get('category') or item.get('category_path')
    if isinstance(category, str):
        category = [
            p.strip() for p in category.replace('>', '/').split('/')
            if p.strip()
        ]
    if category is None:
        category = []
    if not isinstance(category, list):
        return None
    if not category and doc_type != 'none':
        return None
    return {
        'doc_type': doc_type,
        'lang': normalize_lang(item.get('lang')),
        'category': list(category),
        'theme': (item.get('theme') or '').strip(),
        'patent_date': normalize_patent_date_field(item.get('patent_date')),
    }


def read_export_meta(save_dir):
    """读取导出目录中记录的任务参数（补全模式用）。"""
    path = os.path.join(str(save_dir), EXPORT_META_FILE)
    if not os.path.isfile(path):
        return []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f'读取任务参数失败 {path}: {e}')
        return []
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        return []
    entries = []
    for item in data:
        meta = _normalize_meta_entry(item)
        if meta:
            entries.append(meta)
    return entries


def write_export_meta(save_dir, task):
    """把任务参数写入导出目录，补全模式据此还原主题与专利时间字段。"""
    meta = _normalize_meta_entry({
        'doc_type': task.get('doc_type'),
        'lang': task.get('lang'),
        'category': task.get('category'),
        'theme': task.get('theme'),
        'patent_date': task.get('patent_date'),
    })
    if not meta:
        return False
    entries = [
        e for e in read_export_meta(save_dir)
        if not (
            e['doc_type'] == meta['doc_type']
            and e['lang'] == meta['lang']
            and e['category'] == meta['category']
            and e['theme'] == meta['theme']
            and e['patent_date'] == meta['patent_date']
        )
    ]
    entries.append(meta)
    path = os.path.join(str(save_dir), EXPORT_META_FILE)
    try:
        os.makedirs(str(save_dir), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(entries, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f'记录任务参数失败（补全将退回目录名推断）: {e}')
        return False
    theme_part = f' | 主题: {meta["theme"]}' if meta['theme'] else ''
    date_part = (
        f' | {patent_date_field_label(meta["patent_date"])}'
        if meta['doc_type'] == 'patent' else '')
    print(
        f'已记录任务参数: {path}'
        f'（{doc_type_label(meta["doc_type"])}{theme_part}{date_part}）')
    return True


def _infer_fill_task_from_name(folder, doc_type=None):
    """无 export_meta.json 时按目录名推断 lang / category / theme / doc_type。

    目录名为「分类路径_主题词」，用分类树匹配最长前缀，其余视为主题词。
    """
    folder = Path(folder)
    root = Path(DOWNLOAD_ROOT)
    try:
        rel = folder.resolve().relative_to(root.resolve())
    except ValueError:
        rel = Path(folder.name)
    parts = list(rel.parts)
    lang = 'zh'
    signal_doc = None
    if parts and parts[0] == '英文期刊':
        lang = 'en'
        signal_doc = 'journal'
    elif parts and parts[0] == '专利':
        signal_doc = 'patent'
    elif parts and parts[0] in LITERATURE_CATEGORY_TREE:
        signal_doc = 'journal'
    if signal_doc is None:
        # 路径无法识别时才用命令行 --doc-type 兜底
        signal_doc = normalize_doc_type(doc_type or 'journal')
    if signal_doc == 'patent':
        lang = 'zh'
    name = parts[-1] if parts else folder.name
    tokens = [p for p in name.split('_') if p] or [name]

    node = get_category_tree(signal_doc)
    depth = 0
    for token in tokens:
        if not isinstance(node, dict) or token not in node:
            break
        node = node[token]
        depth += 1
    if 0 < depth < len(tokens):
        category = tokens[:depth]
        theme = '_'.join(tokens[depth:])
    else:
        category = tokens
        theme = ''
    return {
        'lang': lang,
        'doc_type': signal_doc,
        'category': category,
        'theme': theme,
        'patent_date': DEFAULT_PATENT_DATE_FIELD,
    }


def folder_to_fill_task(folder, doc_type=None, patent_date=None, theme=None):
    """由导出目录还原补全任务参数。

    优先读取导出时写入的 export_meta.json；缺失时按目录名推断。
    专利的申请日/公开日无法从目录名区分，只能来自任务参数或 --patent-date。
    """
    folder = Path(folder)
    inferred = _infer_fill_task_from_name(folder, doc_type=doc_type)
    task = dict(inferred)
    metas = read_export_meta(folder)
    if metas:
        matched = [m for m in metas if m['doc_type'] == inferred['doc_type']]
        chosen = (matched or metas)[-1]
        task.update({
            'doc_type': chosen['doc_type'],
            'lang': chosen['lang'],
            'category': chosen['category'],
            'theme': chosen['theme'],
            'patent_date': chosen['patent_date'],
        })
        if len(matched) > 1:
            detail = '；'.join(
                f'{doc_type_label(m["doc_type"])}'
                f'{"/" + m["theme"] if m["theme"] else ""}'
                f'{"/" + patent_date_field_label(m["patent_date"]) if m["doc_type"] == "patent" else ""}'
                for m in matched)
            print(
                f'目录 {folder} 有 {len(matched)} 条任务参数（{detail}），'
                '默认用最新一条，可用 --patent-date 指定时间字段')
    if patent_date:
        task['patent_date'] = normalize_patent_date_field(patent_date)
    if theme and not task.get('theme'):
        # 目录名/任务参数都没有主题时才用命令行兜底，避免污染其他目录
        task['theme'] = theme.strip()
    return task


def _fill_datebox_pair(driver, start_id, end_id, start_date, end_date):
    """填写一对日期框。"""
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, start_id)))
    except Exception:
        time.sleep(2)
        if not driver.find_elements(By.ID, start_id):
            raise RuntimeError(
                f'找不到 {start_id}，当前URL={driver.current_url}')
    start_el = driver.find_element(By.ID, start_id)
    driver.execute_script("arguments[0].removeAttribute('readonly');", start_el)
    start_el.clear()
    start_el.send_keys(start_date)
    end_el = driver.find_element(By.ID, end_id)
    driver.execute_script("arguments[0].removeAttribute('readonly');", end_el)
    end_el.clear()
    end_el.send_keys(end_date)


def _clear_datebox_pair(driver, start_id, end_id):
    """清空一对日期框（专利申请日/公开日互斥时用）。"""
    for box_id in (start_id, end_id):
        els = driver.find_elements(By.ID, box_id)
        if not els:
            continue
        el = els[0]
        try:
            driver.execute_script(
                "arguments[0].removeAttribute('readonly');"
                "arguments[0].value='';", el)
            el.clear()
        except Exception:
            pass


def set_adv_date_range(
        driver, start_date, end_date, doc_type=None, patent_date_field=None):
    """设置高级检索日期范围。

    期刊：datebox0/1（发表时间）
    专利时间范围区（patent-extend-row）：
      申请日 AD → datebox0 / datebox1
      公开日 OD → datebox2 / datebox3
      写入一侧时清空另一侧，避免双区间同时生效。
    """
    dismiss_alerts(driver, timeout=0.5)
    hide_datetime_pickers(driver)
    doc_type = normalize_doc_type(doc_type)
    if doc_type == 'patent':
        field = normalize_patent_date_field(patent_date_field)
        if field == 'OD':
            start_id, end_id = 'datebox2', 'datebox3'
            other_start, other_end = 'datebox0', 'datebox1'
            label = '公开日'
        else:
            start_id, end_id = 'datebox0', 'datebox1'
            other_start, other_end = 'datebox2', 'datebox3'
            label = '申请日'
        _clear_datebox_pair(driver, other_start, other_end)
        _fill_datebox_pair(driver, start_id, end_id, start_date, end_date)
        print(f'设置专利{label}(清空另一时间轴): {start_date} ~ {end_date}')
        return
    _fill_datebox_pair(driver, 'datebox0', 'datebox1', start_date, end_date)
    print(f'设置日期范围: {start_date} ~ {end_date}')


def prepare_patent_search_page(driver):
    """专利高级检索页预处理（与期刊 prepare_adv_search_page 对应）。

    页面特征（classid=VUDIXAIY）：
      - 侧栏默认 is-off，分类 tab 默认「国际专利分类」
      - #gradetxt 第1行公开号(GKH)、第2行主题(SU)、第3行申请人(SQR)
      - 时间：申请日 datebox0/1 + 公开日 datebox2/3
      - 扩展项「中英文扩展」data-id=EN 与期刊语种复选框不是同一概念，勿改
    """
    print('专利检索页预处理：重置条件，避免公开号/申请人残留')
    try:
        resets = driver.find_elements(By.CSS_SELECTOR, 'a.btn-reset')
        for btn in resets:
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1.0)
                print('已点击「重置条件」')
                break
    except Exception as e:
        print('重置条件失败（可忽略）:', e)
    # 展开侧栏并确保 IPC tab（后续选分类还会再确认）
    try:
        expand_literature_category_panel(driver)
        switch_category_nav_tab(driver, 'patent')
    except Exception as e:
        print('专利分类面板预展开失败（选分类时将重试）:', e)
    time.sleep(0.5)


def _set_literature_lang_checkbox(driver, data_id, want_checked):
    """设置高级检索页中文/外文文献复选框状态。"""
    els = driver.find_elements(By.CSS_SELECTOR, f'input[data-id="{data_id}"]')
    if not els:
        print(f'未找到文献语言复选框 data-id={data_id}')
        return False
    el = els[0]
    is_on = el.is_selected()
    if is_on != want_checked:
        try:
            driver.execute_script("arguments[0].click();", el)
        except Exception:
            el.click()
        print(f'已将 data-id={data_id} 设为 {"勾选" if want_checked else "取消"}')
    else:
        print(f'data-id={data_id} 已是 {"勾选" if want_checked else "取消"}，跳过')
    return True


def normalize_lang(lang):
    value = (lang or DOWNLOAD_LANG or 'zh').strip().lower()
    if value in ('none', 'any', 'all', '不选', '不限', '全部'):
        return 'none'
    if value in ('en', 'english', '英文', 'foreign', '外文'):
        return 'en'
    return 'zh'


def lang_label(lang):
    normalized = normalize_lang(lang)
    if normalized == 'en':
        return '英文'
    if normalized == 'none':
        return '不选'
    return '中文'


def prepare_adv_search_page(driver, lang=None):
    """高级检索页预处理：按下载类型勾选中文或外文文献；不选则跳过。"""
    lang = normalize_lang(lang)
    if lang == 'none':
        print('下载类型：不选，跳过检索页中/外文筛选')
        return
    if lang == 'en':
        print('检索范围：仅外文文献')
        if not click_foreign_filter(driver, timeout=15, stage='检索前'):
            print('检索前未点到外文按钮，继续尝试复选框设置')
        _set_literature_lang_checkbox(driver, 'CH', False)
        _set_literature_lang_checkbox(driver, 'EN', True)
    else:
        print('检索范围：仅中文文献')
        if not click_chinese_filter(driver, timeout=3, stage='检索前'):
            print('检索前未点到中文按钮，继续尝试复选框设置')
        _set_literature_lang_checkbox(driver, 'CH', True)
        _set_literature_lang_checkbox(driver, 'EN', False)
    time.sleep(2)


def click_search(driver):
    """点击检索按钮。"""
    searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
    driver.execute_script("arguments[0].click();", searchbtn)
    time.sleep(5)


def _click_lang_tab(driver, selectors, label, timeout, stage):
    last_err = None
    for selector in selectors:
        try:
            el = WebDriverWait(driver, timeout, 0.5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
            driver.execute_script("arguments[0].click();", el)
            print(f'已点击{label}({stage}): {selector} text={el.text!r}')
            time.sleep(2)
            return True
        except Exception as e:
            last_err = e
            continue
    print(f'未找到或无法点击{label}按钮({stage}):', last_err)
    return False


def click_foreign_filter(driver, timeout=20, stage='结果页'):
    """点击「外文」：<a class="en" data-val="Foreign">外文</a>。"""
    print(f'点击{stage}「外文」')
    return _click_lang_tab(
        driver,
        ('a.en[data-val="Foreign"]', 'a[data-val="Foreign"]', 'a.en'),
        '外文', timeout, stage)


def click_chinese_filter(driver, timeout=20, stage='结果页'):
    """点击「中文」：<a class="zh" data-val="Chinese">中文</a>。"""
    print(f'点击{stage}「中文」')
    return _click_lang_tab(
        driver,
        ('a.zh[data-val="Chinese"]', 'a[data-val="Chinese"]', 'a.zh'),
        '中文', timeout, stage)


def apply_lang_filter(driver, lang=None):
    """结果页按下载类型点中文/外文；下载类型=不选时跳过。"""
    lang = normalize_lang(lang)
    if lang == 'none':
        print('下载类型：不选，跳过结果页中/外文筛选')
        return
    if lang == 'en':
        if not click_foreign_filter(driver, stage='结果页'):
            raise RuntimeError('结果页外文筛选失败')
        print('结果页只保留外文')
    else:
        if not click_chinese_filter(driver, timeout=3, stage='结果页'):
            print('结果页未点到中文按钮，继续后续筛选')
        print('结果页只保留中文')


def apply_journal_filter(driver, lang=None):
    """检索后按下载类型点中文/外文，再选择学术期刊；不选则跳过语言筛选。"""
    apply_lang_filter(driver, lang)
    print('只选择学术期刊')
    journal = WebDriverWait(driver, 20, 0.5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[resource="JOURNAL"]')))
    driver.execute_script("arguments[0].click();", journal)
    time.sleep(3)


def apply_patent_filter(driver):
    """检索结果页只保留专利。"""
    print('只选择专利')
    selectors = (
        'a[resource="PATENT"]',
        'a[resource="patent"]',
        'a[data-id="PATENT"]',
    )
    last_err = None
    for sel in selectors:
        try:
            el = WebDriverWait(driver, 12, 0.5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
            driver.execute_script("arguments[0].click();", el)
            time.sleep(3)
            print(f'已点击专利筛选: {sel}')
            return True
        except Exception as e:
            last_err = e
    # 文案兜底
    try:
        links = driver.find_elements(By.CSS_SELECTOR, 'a[resource], .resource a, .doctype a')
        for el in links:
            text = (el.text or '').strip()
            if '专利' in text:
                driver.execute_script("arguments[0].click();", el)
                time.sleep(3)
                print(f'已点击专利筛选(文案): {text!r}')
                return True
    except Exception as e:
        last_err = e
    raise RuntimeError(f'结果页专利筛选失败: {last_err}')


def apply_resource_filter(driver, doc_type=None, lang=None):
    """按导出类型筛选结果；导出类型=不选时跳过整个资源类型筛选。

    期刊：中/外文（下载类型不选时跳过）+ 学术期刊(JOURNAL)。
    专利：已在 classid=VUDIXAIY 专利库检索，结果页通常无 JOURNAL/PATENT 切换，默认跳过。
    """
    doc_type = normalize_doc_type(doc_type)
    if doc_type == 'none':
        print('导出类型：不选，跳过结果页资源类型筛选')
        apply_lang_filter(driver, lang)
        return True
    if doc_type == 'patent':
        print('专利库结果页：跳过期刊式资源类型筛选')
        return True
    apply_journal_filter(driver, lang=lang)
    return True


def _grade_row_field_value(dd):
    """读取检索行当前字段 data-value / 显示名。"""
    try:
        span = dd.find_element(By.CSS_SELECTOR, '.sort.reopt .sort-default span')
        return (
            (span.get_attribute('data-value')
             or span.get_attribute('value')
             or span.text
             or '').strip()
        )
    except Exception:
        return ''


def _set_grade_row_field(driver, dd, data_val):
    """将检索行字段切到 data-val（如 SU）。已是目标则跳过。"""
    cur = _grade_row_field_value(dd)
    if cur == data_val or (data_val == 'SU' and cur in ('SU', '主题')):
        return True
    reopt = dd.find_element(By.CSS_SELECTOR, '.sort.reopt')
    driver.execute_script("arguments[0].click();", reopt)
    time.sleep(0.35)
    items = dd.find_elements(
        By.CSS_SELECTOR, f'.sort-list li[data-val="{data_val}"]')
    if not items:
        return False
    el = items[0]
    try:
        clickable = el.find_element(By.CSS_SELECTOR, 'a')
    except Exception:
        clickable = el
    driver.execute_script("arguments[0].click();", clickable)
    time.sleep(0.3)
    return True


def _fill_grade_row_text(driver, dd, text):
    text_input = dd.find_element(By.CSS_SELECTOR, 'input[type="text"]')
    try:
        text_input.clear()
    except Exception:
        driver.execute_script("arguments[0].value='';", text_input)
    text_input.send_keys(text)


def _find_grade_row_by_field(dd_list, data_val, title_hint=None):
    """在 #gradetxt 的 dd 列表中找当前字段为 data_val 的行。"""
    title_hint = title_hint or ''
    for dd in dd_list:
        val = _grade_row_field_value(dd)
        if val == data_val or (title_hint and title_hint in val):
            return dd
        # 下拉里带 cur 的也算
        try:
            cur = dd.find_elements(
                By.CSS_SELECTOR,
                f'.sort.reopt .sort-list li.cur[data-val="{data_val}"]')
            if cur:
                return dd
        except Exception:
            pass
    return None


def set_adv_theme_keyword(driver, theme, doc_type=None):
    """高级检索可选「主题」条件。theme 为空则跳过。

    期刊：第 1 行切到 SU 后填写。
    专利：第1行默认公开号(GKH)、第2行主题(SU)、第3行申请人(SQR)；
          只写入主题行，并清空公开号/申请人，避免空条件以外的残留干扰。
    """
    theme = (theme or '').strip()
    doc_type = normalize_doc_type(doc_type)
    if not theme:
        print('未设置主题检索词，仅按分类+日期检索')
        if doc_type == 'patent':
            try:
                WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.ID, 'gradetxt')))
                dd_list = driver.find_element(By.ID, 'gradetxt').find_elements(
                    By.CSS_SELECTOR, 'dd')
                for code, hint in (('GKH', '公开号'), ('SQR', '申请人')):
                    row = _find_grade_row_by_field(dd_list, code, hint)
                    if row is not None:
                        _fill_grade_row_text(driver, row, '')
            except Exception:
                pass
        return False

    print(f'设置主题检索词: {theme}')
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, 'gradetxt')))
    dd_list = driver.find_element(By.ID, 'gradetxt').find_elements(
        By.CSS_SELECTOR, 'dd')
    if not dd_list:
        raise RuntimeError('找不到检索行 #gradetxt dd')

    if doc_type == 'patent':
        sel_dd = _find_grade_row_by_field(dd_list, 'SU', '主题')
        if sel_dd is None:
            # 专利页通常第 2 行已是主题
            sel_dd = dd_list[1] if len(dd_list) > 1 else dd_list[0]
            if not _set_grade_row_field(driver, sel_dd, 'SU'):
                raise RuntimeError('专利页无法切换到主题(SU)字段')
        print('专利主题写入第2类「主题」行（跳过公开号/申请人）')
        _fill_grade_row_text(driver, sel_dd, theme)
        # 清空公开号、申请人
        for code, hint in (('GKH', '公开号'), ('SQR', '申请人')):
            row = _find_grade_row_by_field(dd_list, code, hint)
            if row is not None and row != sel_dd:
                _fill_grade_row_text(driver, row, '')
    else:
        sel_dd = dd_list[0]
        if not _set_grade_row_field(driver, sel_dd, 'SU'):
            raise RuntimeError('找不到主题字段选项 data-val=SU')
        _fill_grade_row_text(driver, sel_dd, theme)

    time.sleep(0.5)
    print(f'已填入主题: {theme}')
    return True


def prompt_literature_category(doc_type=None):
    """控制台逐级选分类。期刊=学科；专利=国际专利分类 IPC；不选则跳过。"""
    doc_type = normalize_doc_type(doc_type)
    if doc_type == 'none':
        print('导出类型：不选，跳过文献分类(学科导航)选择')
        return []
    path = []
    node = get_category_tree(doc_type)
    title = '国际专利分类(IPC)' if doc_type == 'patent' else '文献分类(学科导航)'
    while True:
        if doc_type == 'patent':
            path_label = (
                ' / '.join(ipc_display_name(p) for p in path) if path else '(根)')
        else:
            path_label = ' / '.join(path) if path else '(根)'
        print(f'\n{title}（当前: {path_label}）')
        children = list(node.keys()) if isinstance(node, dict) else []
        if path:
            print('  0. 确认选择当前分类')
        for index, name in enumerate(children, 1):
            mark = ' →' if node.get(name) else ''
            shown = ipc_display_name(name) if doc_type == 'patent' else name
            if len(shown) > 60:
                shown = shown[:57] + '...'
            print(f'  {index}. {shown}{mark}')
        if not children and path:
            print('无更细分类，已确认当前路径')
            return path

        choice = input(
            '请选择编号（带 → 可继续细分；0=确认当前）: '
        ).strip()
        if choice == '0':
            if not path:
                print('请先选择一级分类')
                continue
            return path
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(children):
                name = children[idx - 1]
                path.append(name)
                child = node.get(name) or {}
                if not child:
                    return path
                node = child
                continue
        print('输入无效，请重新选择')


def prompt_patent_date_field(default=None):
    """专利：申请日 / 公开日。"""
    default = normalize_patent_date_field(default or DEFAULT_PATENT_DATE_FIELD)
    options = [('1', 'AD', '申请日'), ('2', 'OD', '公开日')]
    default_key = '1' if default == 'AD' else '2'
    print('\n专利时间字段：')
    for key, _v, label in options:
        mark = ' (默认)' if key == default_key else ''
        print(f'  {key}. {label}{mark}')
    while True:
        choice = input(
            f'请选择时间字段编号（直接回车用默认 {default_key}）: '
        ).strip()
        if not choice:
            choice = default_key
        for key, value, label in options:
            if choice == key or choice == value or choice == label:
                print(f'已选择专利时间字段: {label}')
                return value
        print('输入无效，请重新选择')


def category_path_to_dirname(category_path):
    """文献分类路径转为目录名。"""
    if isinstance(category_path, str):
        return category_path or '不限分类'
    if not category_path:
        return '不限分类'
    return '_'.join(category_path)


def category_path_label(category_path, doc_type=None):
    """分类路径显示名；导出类型不选或无路径时显示「不限分类」。"""
    doc_type = normalize_doc_type(doc_type)
    if isinstance(category_path, str):
        category_path = [category_path] if category_path.strip() else []
    category_path = list(category_path or [])
    if doc_type == 'none' or not category_path:
        return '不限分类'
    if doc_type == 'patent':
        return ' / '.join(ipc_display_name(p) for p in category_path)
    return ' / '.join(category_path)


def task_download_dir(task):
    """按下载类型确定保存目录。

    期刊中文：大类或分类路径；英文：英文期刊/分类；
    专利：专利/分类[_主题]；导出类型不选：不限分类[_主题]。
    """
    category_path = task.get('category') or []
    category_dir = category_path_to_dirname(category_path)
    theme = (task.get('theme') or '').strip()
    if theme:
        # 主题写入子目录名，避免不同主题互相覆盖
        safe_theme = re.sub(r'[\\/:*?"<>|]+', '_', theme)[:40]
        if category_dir == '不限分类':
            category_dir = safe_theme
        else:
            category_dir = f'{category_dir}_{safe_theme}'
    if normalize_doc_type(task.get('doc_type')) == 'patent':
        return os.path.join(DOWNLOAD_ROOT, '专利', category_dir)
    if normalize_lang(task.get('lang')) == 'en':
        return os.path.join(DOWNLOAD_ROOT, '英文期刊', category_dir)
    if isinstance(category_path, list) and len(category_path) == 1 and not theme:
        return os.path.join(DOWNLOAD_ROOT, category_path[0])
    return os.path.join(DOWNLOAD_ROOT, category_dir)


def expand_literature_category_panel(driver):
    """展开高级检索页文献分类面板（专利页默认 search-mainbox is-off）。"""
    WebDriverWait(driver, 20, 0.5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '.search-sidebar-a')))
    mainbox = driver.find_element(By.CSS_SELECTOR, '.search-mainbox')
    cls = (mainbox.get_attribute('class') or '')
    if 'is-on' in cls and 'is-off' not in cls:
        return
    toggle = driver.find_element(By.CSS_SELECTOR, '.search-sidebar-a a.icon-arrow')
    driver.execute_script("arguments[0].click();", toggle)
    WebDriverWait(driver, 10, 0.5).until(
        lambda d: 'is-on' in (
            d.find_element(By.CSS_SELECTOR, '.search-mainbox').get_attribute('class')
            or ''))
    time.sleep(1)
    print('已展开文献分类侧栏')


def switch_category_nav_tab(driver, doc_type=None):
    """专利→国际专利分类 #RJUTCQD5；期刊→学科导航 #CNKI168。

    期刊总库页默认即学科导航，通常没有「国际专利分类」双 tab，
    只要存在 #CNKI168 即直接使用，不强行点 tab。
    """
    doc_type = normalize_doc_type(doc_type)
    root_id = category_nav_root_id(doc_type)
    tab_title = '国际专利分类' if doc_type == 'patent' else '学科导航'

    def _root_usable(rid):
        try:
            root = driver.find_element(By.ID, rid)
        except Exception:
            return False
        style = (root.get_attribute('style') or '').replace(' ', '').lower()
        try:
            display = (root.value_of_css_property('display') or '').lower()
        except Exception:
            display = ''
        if 'display:none' in style or display == 'none':
            return False
        return True

    # 期刊：无 IPC tab → 默认学科；或 #CNKI168 已可见 → 直接用
    if doc_type != 'patent':
        ipc_tabs = driver.find_elements(
            By.CSS_SELECTOR, f'.top-tabs a[data-href="{PATENT_IPC_NAV_ID}"]')
        if not ipc_tabs:
            if driver.find_elements(By.ID, JOURNAL_NAV_ID):
                print('期刊页无专利分类 tab，默认使用学科导航 #CNKI168')
                return JOURNAL_NAV_ID
        elif _root_usable(JOURNAL_NAV_ID):
            print('分类面板已是学科导航（默认），无需切换')
            return JOURNAL_NAV_ID

    if _root_usable(root_id):
        tabs = driver.find_elements(
            By.CSS_SELECTOR, f'.top-tabs a[data-href="{root_id}"]')
        if not tabs or 'cur' in (tabs[0].get_attribute('class') or ''):
            print(f'分类面板已是「{tab_title}」')
            return root_id

    last_err = None
    for sel in (
            f'.top-tabs a[data-href="{root_id}"]',
            f'.top-tabs a[title="{tab_title}"]',
            f'a[data-href="{root_id}"]'):
        try:
            tab = WebDriverWait(driver, 8, 0.5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
            driver.execute_script("arguments[0].click();", tab)
            WebDriverWait(driver, 10, 0.5).until(
                EC.visibility_of_element_located((By.ID, root_id)))
            time.sleep(0.5)
            print(f'已切换分类面板: {tab_title} (#{root_id})')
            return root_id
        except Exception as e:
            last_err = e

    # 期刊/不选类型兜底：tab 点不到但 DOM 有 CNKI168 仍继续
    if doc_type != 'patent' and driver.find_elements(By.ID, JOURNAL_NAV_ID):
        print('未点到学科导航 tab，仍使用默认 #CNKI168')
        return JOURNAL_NAV_ID
    raise RuntimeError(f'无法切换到「{tab_title}」: {last_err}')


def clear_category_nav_selection(driver):
    """点击分类「清除」。"""
    try:
        for btn in driver.find_elements(
                By.CSS_SELECTOR,
                '.top-operate-btns a.clearall, .nav-content a.clearall'):
            try:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.4)
                    print('已清除原有分类勾选')
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _element_text(el):
    try:
        text = (el.text or '').strip()
    except Exception:
        text = ''
    if not text:
        try:
            text = (el.get_attribute('textContent') or '').strip()
        except Exception:
            text = ''
    return text


def _category_token_match(name, li=None, link=None):
    want = (name or '').strip()
    if not want:
        return False
    code = ''
    text = ''
    try:
        if li is not None:
            code = (li.get_attribute('name') or '').strip()
    except Exception:
        pass
    try:
        if link is not None:
            text = _element_text(link)
    except Exception:
        pass
    if want == code or want == text:
        return True
    label = PATENT_IPC_LABELS.get(code, '')
    if want == label:
        return True
    if code and want.upper() == code.upper():
        return True
    if code:
        # IPC 子节点可能没有 name 属性，只能靠链接文本前缀匹配
        if re.match(rf'^{re.escape(code)}(?![A-Za-z0-9])', text, re.I):
            return True
        if label and want.replace(' ', '') in (
                f'{code}{label}'.replace(' ', ''),
                f'{code} {label}'.replace(' ', '')):
            return True
    if re.match(rf'^{re.escape(want)}(?![A-Za-z0-9])', text, re.I):
        return True
    return False


def _find_direct_child_category(scope_el, name):
    """在 scope 直接子 li 中按中文名或 IPC 分类号查找。"""
    try:
        lis = scope_el.find_elements(By.XPATH, './li')
    except Exception:
        return None, None, None
    for li in lis:
        try:
            link = li.find_element(
                By.XPATH,
                './div[contains(@class,"item-nav")]/a[contains(@class,"navi-search")]')
            if _category_token_match(name, li=li, link=link):
                item_nav = link.find_element(By.XPATH, './..')
                return li, item_nav, link
        except Exception:
            continue
    return None, None, None


def _find_any_category(scope_el, name):
    """scope 内任意层级按名称/分类号查找。"""
    try:
        links = scope_el.find_elements(By.CSS_SELECTOR, 'a.navi-search')
    except Exception:
        return None, None, None
    for link in links:
        try:
            item_nav = link.find_element(By.XPATH, './..')
            li = item_nav.find_element(By.XPATH, './..')
            if _category_token_match(name, li=li, link=link):
                return li, item_nav, link
        except Exception:
            continue
    return None, None, None


def _child_container(li):
    """li 的直接子级容器；兼容 ul 外包一层 div 的结构。"""
    for xpath in (
            './ul',
            './div/ul',
            './div[contains(@class,"sub-nav")]',
            './div[contains(@class,"sub")]',
            './div[contains(@class,"child")]',
            './div[contains(@class,"list")]'):
        try:
            els = li.find_elements(By.XPATH, xpath)
        except Exception:
            continue
        if els:
            return els[0]
    return None


def _find_nav_item_along_path(driver, path, root_id=None):
    """从分类根节点按路径层层定位。"""
    root_id = root_id or JOURNAL_NAV_ID
    try:
        scope = driver.find_element(By.ID, root_id)
    except Exception:
        return None, None, None
    li = item_nav = link = None
    for name in path:
        li, item_nav, link = _find_direct_child_category(scope, name)
        if li is None:
            li, item_nav, link = _find_any_category(scope, name)
        if li is None:
            return None, None, None
        child = _child_container(li)
        scope = child if child is not None else li
    return li, item_nav, link


def _child_ul_visible(li):
    """判断 li 下子列表是否已展开可见。"""
    container = _child_container(li)
    if container is None:
        return False
    try:
        return bool(container.is_displayed())
    except Exception:
        pass
    try:
        display = (container.value_of_css_property('display') or '').lower()
        if display == 'none':
            return False
    except Exception:
        pass
    style = (container.get_attribute('style') or '').replace(' ', '').lower()
    if 'display:none' in style:
        return False
    cls = (container.get_attribute('class') or '').lower()
    if 'hide' in cls or 'hidden' in cls:
        return False
    return True


def _expand_nav_item(driver, li, item_nav, expect_child_name=None, timeout=10):
    """展开分类节点。兼容 IPC 树：icon 无响应时点父节点，最后强制显示子容器。"""
    try:
        link = item_nav.find_element(By.CSS_SELECTOR, 'a.navi-search')
        name = _element_text(link)
    except Exception:
        link = None
        name = ''

    def child_ready():
        container = _child_container(li)
        if container is None or not _child_ul_visible(li):
            return False
        if not expect_child_name:
            return bool(container.find_elements(By.CSS_SELECTOR, 'a.navi-search'))
        c_li, _, _ = _find_direct_child_category(container, expect_child_name)
        if c_li is None:
            c_li, _, _ = _find_any_category(container, expect_child_name)
        if c_li is not None:
            return True
        # 子列表已展开但名称未匹配上：交给调用方全树兜底，避免反复点击把树收起
        return bool(container.find_elements(By.CSS_SELECTOR, 'a.navi-search'))

    if child_ready():
        return True

    icons = []
    try:
        icons = item_nav.find_elements(By.CSS_SELECTOR, 'i.icon')
    except Exception:
        pass
    arrow = None
    for icon in icons:
        try:
            cls = (icon.get_attribute('class') or '').lower()
        except Exception:
            continue
        if any(k in cls for k in ('arrow', 'down', 'up', 'plus', 'minus', 'expand')):
            arrow = icon
            break
    targets = []
    if arrow is not None:
        targets.append(arrow)
    elif icons:
        targets.append(icons[0])
    targets.append(item_nav)
    if link is not None:
        targets.append(link)

    for attempt in range(1, 4):
        target = targets[(attempt - 1) % len(targets)]
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", target)
            if attempt == 1:
                driver.execute_script("arguments[0].click();", target)
            else:
                try:
                    target.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", target)
        except Exception as e:
            print(f'展开点击失败({name}) 第{attempt}次:', e)
        deadline = time.time() + timeout
        while time.time() < deadline:
            if child_ready():
                print(f'已展开分类节点: {name}')
                return True
            time.sleep(0.3)
        print(f'展开等待超时({name}) 第{attempt}次，重试点击')

    # 最后兜底：直接展开子容器（部分 IPC 节点的 icon 点击不触发）
    container = _child_container(li)
    if container is not None:
        try:
            driver.execute_script(
                "var e=arguments[0]; if(!e) return;"
                "e.style.display='block'; e.style.visibility='visible';"
                "e.classList.remove('hide','hidden');", container)
            time.sleep(0.6)
            if child_ready():
                print(f'已强制展开分类节点: {name}')
                return True
        except Exception:
            pass
    print(f'展开分类失败: {name}')
    return False


def select_literature_category(driver, category_path, doc_type=None):
    """按路径选择分类。专利用国际专利分类 IPC；期刊用学科导航；不选则跳过。"""
    doc_type = normalize_doc_type(doc_type)
    if doc_type == 'none':
        print('导出类型：不选，跳过文献分类(学科导航)选择')
        return
    if isinstance(category_path, str):
        category_path = [category_path]
    if not category_path:
        raise RuntimeError('分类路径为空')

    expand_literature_category_panel(driver)
    root_id = switch_category_nav_tab(driver, doc_type)
    WebDriverWait(driver, 20, 0.5).until(
        EC.presence_of_element_located((By.ID, root_id)))
    time.sleep(0.5)
    if doc_type == 'patent':
        clear_category_nav_selection(driver)

    path_label = (
        ' / '.join(ipc_display_name(p) for p in category_path)
        if doc_type == 'patent' else ' / '.join(category_path))

    for depth, name in enumerate(category_path):
        is_last = depth == len(category_path) - 1
        parent_path = category_path[:depth]
        current_path = category_path[:depth + 1]
        expect_child = None if is_last else (
            category_path[depth + 1] if depth + 1 < len(category_path) else None)

        for i in range(len(parent_path)):
            ancestor_path = category_path[:i + 1]
            a_li, a_nav, _ = _find_nav_item_along_path(
                driver, ancestor_path, root_id=root_id)
            if a_li is None:
                raise RuntimeError(
                    f'未找到祖先分类: {category_path[i]}（路径 {path_label}）')
            next_name = category_path[i + 1] if i + 1 < len(category_path) else None
            _expand_nav_item(driver, a_li, a_nav, expect_child_name=next_name)

        li, item_nav, _link = _find_nav_item_along_path(
            driver, current_path, root_id=root_id)
        if li is None and parent_path:
            p_li, p_nav, _ = _find_nav_item_along_path(
                driver, parent_path, root_id=root_id)
            if p_li is not None:
                print(f'未找到「{name}」，强制展开父级后重试...')
                _expand_nav_item(
                    driver, p_li, p_nav, expect_child_name=name, timeout=10)
                time.sleep(0.5)
                li, item_nav, _link = _find_nav_item_along_path(
                    driver, current_path, root_id=root_id)

        if li is None:
            root = driver.find_element(By.ID, root_id)
            if doc_type == 'patent':
                # IPC 树有时靠样式隐藏子级，强制展开所有 ul 后再全树查找
                try:
                    driver.execute_script(
                        "arguments[0].querySelectorAll("
                        "'ul, div[class*=\"sub\"], div[class*=\"list\"]')"
                        ".forEach(function(u){"
                        "u.style.display='block';"
                        "u.classList.remove('hide','hidden');});", root)
                    time.sleep(0.4)
                except Exception:
                    pass
            li, item_nav, _link = _find_any_category(root, name)
            if li is not None:
                print(f'警告: 以全树兜底方式找到「{name}」')

        if li is None or item_nav is None:
            raise RuntimeError(f'未找到分类: {name}（路径 {path_label}）')

        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", item_nav)
        except Exception:
            pass
        time.sleep(0.3)

        if not is_last:
            if not _expand_nav_item(
                    driver, li, item_nav, expect_child_name=expect_child):
                print(f'警告: 「{name}」展开可能未完成，继续尝试')
            continue

        icon_select = item_nav.find_element(By.CSS_SELECTOR, 'i.icon-select')
        driver.execute_script("arguments[0].click();", icon_select)
        time.sleep(2)
        kind = '国际专利分类' if doc_type == 'patent' else '学科导航'
        print(f'已选择{kind}: {path_label}')
        return

    raise RuntimeError(f'未能选择分类: {path_label}')


def run_one_search_export(
        driver, download_dir, category_path, start_date, end_date, label,
        lang=None, doc_type=None, theme=None, patent_date_field=None,
        reuse_session=False):
    """执行一次检索+导出。

    专利与期刊差异：
      - 专利页 classid=VUDIXAIY，分类用国际专利分类(IPC)
      - 主题写入「主题」行（非默认的公开号行）
      - 日期为申请日(AD)或公开日(OD)
      - 不做期刊中/外文勾选与 JOURNAL 强制筛选

    reuse_session=True：同任务内已通过滑块/登录，只回到高级检索改日期，
    不做直连探测跳转，避免每段再出验证码。
    """
    lang = normalize_lang(lang)
    doc_type = normalize_doc_type(doc_type)
    theme = (theme or '').strip()
    patent_date_field = normalize_patent_date_field(patent_date_field)
    cat_label = category_path_label(category_path, doc_type)
    theme_part = f' | 主题: {theme}' if theme else ''
    date_part = (
        f' | {patent_date_field_label(patent_date_field)}'
        if doc_type == 'patent' else '')
    if doc_type == 'none':
        cat_kind = '分类'
    elif doc_type == 'patent':
        cat_kind = '国际专利分类'
    else:
        cat_kind = '学科导航'
    print(
        f'\n===== 开始检索: {label} | {doc_type_label(doc_type)}'
        f'{"/" + lang_label(lang) if doc_type != "patent" else ""}'
        f' | {cat_kind}: {cat_label}{theme_part}{date_part} ====='
    )

    main_handle = None
    try:
        # 进入新检索前先关掉残留导出页，避免焦点/句柄错乱
        if driver.window_handles:
            close_export_windows(driver, main_handle=driver.current_window_handle)

        search_url = get_adv_search_url(doc_type)
        if reuse_session:
            print(f'{label} 复用本任务已验证会话，回到高级检索（不重新探测跳转）')
            open_adv_search_in_session(driver, doc_type)
            time.sleep(0.8)
            main_handle = driver.current_window_handle
            if page_suggests_login_lost(driver, check_handle=main_handle):
                print(f'{label} 复用会话时检测到登录中断，需重建浏览器重新登录')
                return 'restart_browser'
            # 仅当页面再次出现验证码时才处理，不主动触发探测链
            try:
                cur = driver.current_url or ''
            except Exception:
                cur = ''
            if is_cnki_verify_url(cur) or has_captcha_anywhere(driver):
                print(f'{label} 复用会话时仍出现验证码，开始处理')
                cap = ensure_cnki_captcha_after_nav(
                    driver, download_dir, main_handle)
                if captcha_needs_browser_restart(cap):
                    print(f'{label} 验证码连败，需关闭浏览器后重试')
                    return 'restart_browser'
                if not cap:
                    print(f'{label} 复用会话验证码未通过，记录并跳过')
                    record_export_failure(download_dir, label)
                    return 'failed'
            else:
                print(f'{label} 会话有效，跳过滑块验证')
        else:
            open_cnki_url(driver, search_url)
            time.sleep(1)
            # 新段开始前先确保机构登录（页头）
            ensure_unit_login_if_needed(
                driver, restore_handle=driver.current_window_handle)
            main_handle = driver.current_window_handle
            cap = ensure_cnki_captcha_after_nav(
                driver, download_dir, main_handle)
            if captcha_needs_browser_restart(cap):
                print(f'{label} 验证码连败，需关闭浏览器后重试')
                return 'restart_browser'
            if not cap:
                print(f'{label} 登录后验证码未通过，记录并跳过')
                record_export_failure(download_dir, label)
                return 'failed'
            if not is_cnki_search_page_url(driver.current_url or ''):
                print('验证码后不在高级检索页，按落地页重新打开检索')
                open_cnki_url(driver, search_url)
                ensure_unit_login_if_needed(
                    driver, restore_handle=driver.current_window_handle)
                main_handle = driver.current_window_handle
                cap = ensure_cnki_captcha_after_nav(
                    driver, download_dir, main_handle)
                if captcha_needs_browser_restart(cap):
                    print(f'{label} 验证码连败，需关闭浏览器后重试')
                    return 'restart_browser'
                if not cap:
                    print(f'{label} 打开检索页后验证码未通过，记录并跳过')
                    record_export_failure(download_dir, label)
                    return 'failed'
        if doc_type != 'patent':
            prepare_adv_search_page(driver, lang)
        else:
            prepare_patent_search_page(driver)
            print(
                '专利检索：IPC + 主题行(SU) + '
                f'{patent_date_field_label(patent_date_field)}'
                '（已跳过期刊中/外文与 JOURNAL 筛选）')
        # 重置条件/展开侧栏后验证码可能延迟弹出，选分类前再确认一次
        time.sleep(1.0)
        if has_captcha_anywhere(driver):
            print(f'{label} 选分类前检测到验证码，先处理')
            cap = ensure_cnki_captcha_after_nav(
                driver, download_dir, main_handle)
            if captcha_needs_browser_restart(cap):
                print(f'{label} 选分类前验证码连败，需关闭浏览器后重试')
                return 'restart_browser'
            if not cap:
                print(f'{label} 选分类前验证码未通过，记录并跳过')
                record_export_failure(download_dir, label)
                return 'failed'
        category_done = False
        for category_attempt in range(1, 3):
            try:
                select_literature_category(
                    driver, category_path, doc_type=doc_type)
                category_done = True
                break
            except Exception as e:
                if exception_suggests_login_lost(
                        e, driver, check_handle=main_handle):
                    print(f'{label} 分类选择时登录中断，需重建浏览器后重试')
                    return 'restart_browser'
                if category_attempt == 1 and has_captcha_anywhere(driver):
                    print(f'{label} 选分类时出现验证码，处理后重试分类')
                    cap = ensure_cnki_captcha_after_nav(
                        driver, download_dir, main_handle)
                    if captcha_needs_browser_restart(cap):
                        print(f'{label} 选分类时验证码连败，需关闭浏览器后重试')
                        return 'restart_browser'
                    if cap:
                        continue
                print(f'{label} 分类选择失败，记录并跳过:', e)
                record_export_failure(download_dir, label)
                return 'failed'
        if not category_done:
            print(f'{label} 分类选择未完成，记录并跳过')
            record_export_failure(download_dir, label)
            return 'failed'

        try:
            set_adv_theme_keyword(driver, theme, doc_type=doc_type)
        except Exception as e:
            if exception_suggests_login_lost(e, driver, check_handle=main_handle):
                print(f'{label} 主题设置时登录中断，需重建浏览器后重试')
                return 'restart_browser'
            print(f'{label} 主题设置失败，记录并跳过:', e)
            record_export_failure(download_dir, label)
            return 'failed'

        try:
            set_adv_date_range(
                driver, start_date, end_date,
                doc_type=doc_type, patent_date_field=patent_date_field)
        except Exception as e:
            if exception_suggests_login_lost(e, driver, check_handle=main_handle):
                print(f'{label} 设置日期时登录中断，需重建浏览器后重试')
                return 'restart_browser'
            print(f'{label} 设置日期失败，记录并跳过:', e)
            record_export_failure(download_dir, label)
            return 'failed'

        click_search(driver)
        print('主页:', driver.current_url)

        main_handle = driver.current_window_handle
        # 检索后若弹出验证码再处理；复用会话时多数情况无需再滑
        if has_captcha_anywhere(driver):
            cap = ensure_captcha_passed(driver, download_dir, main_handle)
            if captcha_needs_browser_restart(cap):
                print(f'{label} 检索后验证码连败，需关闭浏览器后重试')
                return 'restart_browser'
            if cap is False:
                print(f'{label} 检索后验证码未通过，记录并跳过')
                record_export_failure(download_dir, label)
                return 'failed'

        sps = set_result_page_size(
            driver, 50, save_dir=download_dir, main_handle=main_handle)
        if captcha_needs_browser_restart(sps):
            print(f'{label} 结果页验证码连败，需关闭浏览器后重试')
            return 'restart_browser'
        if not sps:
            if page_suggests_login_lost(driver, check_handle=main_handle):
                print(f'{label} 结果页未就绪且登录中断，需重建浏览器')
                return 'restart_browser'
            print(f'{label} 结果页未就绪，记录并跳过')
            record_export_failure(download_dir, label)
            return 'failed'

        try:
            apply_resource_filter(driver, doc_type=doc_type, lang=lang)
        except Exception as e:
            if exception_suggests_login_lost(e, driver, check_handle=main_handle):
                print(f'{label} {doc_type_label(doc_type)}筛选时登录中断，需重建浏览器')
                return 'restart_browser'
            print(f'{label} {doc_type_label(doc_type)}筛选失败，记录并跳过:', e)
            record_export_failure(download_dir, label)
            return 'failed'

        main_handle = driver.current_window_handle
        # 导出采集前再确认一次登录
        ensure_unit_login_if_needed(driver, restore_handle=main_handle)
        if page_suggests_login_lost(driver, check_handle=main_handle):
            print(f'{label} 导出前检测到登录中断，需重建浏览器')
            return 'restart_browser'
        status = collect_day_results(
            driver, download_dir, main_handle, day_str=label)
        if status == 'ok':
            print(f'{label} 检索导出完成')
            return 'ok'
        if status == 'restart_browser':
            print(f'{label} 需关闭浏览器后重试本时段（频繁/验证码/登录中断）')
            return 'restart_browser'
        if status == 'failed':
            print(f'{label} 导出失败，开始下一个')
            return 'failed'
        print(f'{label} 无检索结果，跳过导出')
        return 'empty'
    except Exception as e:
        if exception_suggests_login_lost(e, driver, check_handle=main_handle):
            print(f'{label} 登录状态中断，需重建浏览器后重试本时段:', e)
            return 'restart_browser'
        print(f'{label} 检索导出异常，记录并跳过:', type(e).__name__, e)
        try:
            record_export_failure(download_dir, label)
        except Exception:
            pass
        return 'failed'


def iter_crawl_date_ranges(start, end, granularity):
    """按粒度切分 [start, end] 区间。

    granularity: day / month / year / all
    不足整月/整年的尾段截到 end（最后一天）。
    生成 (range_start, range_end) 的 date 对。
    """
    import calendar

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
            period_end = min(
                current.replace(day=last_day),
                end,
            )
            yield current, period_end
            current = period_end + timedelta(days=1)
            continue

        if g in ('year', 'y', '年'):
            period_end = min(
                current.replace(month=12, day=31),
                end,
            )
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


def prompt_download_lang(default=None):
    """控制台选择下载类型：中文/英文/不选（不选中/外文筛选）。"""
    default = normalize_lang(default or DOWNLOAD_LANG)
    options = [
        ('1', 'zh', '中文'),
        ('2', 'en', '英文'),
        ('3', 'none', '不选'),
    ]
    default_key = '1'
    for key, value, _label in options:
        if value == default:
            default_key = key
            break

    print('\n下载类型：')
    for key, _value, label in options:
        mark = ' (默认)' if key == default_key else ''
        print(f'  {key}. {label}{mark}')
    print('  说明：选择「不选」将跳过检索页/结果页的中/外文筛选。')
    while True:
        choice = input(f'请选择下载类型编号（直接回车用默认 {default_key}）: ').strip()
        if not choice:
            choice = default_key
        for key, value, label in options:
            if choice == key or choice == value or choice == label:
                print(f'已选择下载类型: {label}')
                return value
        print('输入无效，请重新选择')


def prompt_crawl_granularity(default=None):
    """控制台选择检索粒度：天/月/年/总。"""
    default = default or CRAWL_GRANULARITY
    options = [
        ('1', 'day', '按天'),
        ('2', 'month', '按月'),
        ('3', 'year', '按年'),
        ('4', 'all', '总（整段一次）'),
    ]
    default_key = '4'
    for key, value, _label in options:
        if value == default:
            default_key = key
            break

    print('\n检索粒度：')
    for key, _value, label in options:
        mark = ' (默认)' if key == default_key else ''
        print(f'  {key}. {label}{mark}')
    while True:
        choice = input(f'请选择检索粒度编号（直接回车用默认 {default_key}）: ').strip()
        if not choice:
            choice = default_key
        for key, value, label in options:
            if choice == key or choice == value or choice == label:
                print(f'已选择检索粒度: {label}')
                return value
        print('输入无效，请重新选择')


def prompt_task_date_range(default_bg=None, default_end=None):
    """可选覆盖任务日期范围；回车使用默认。"""
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


def prompt_doc_type(default=None):
    """控制台选择导出文献类型：期刊 / 专利 / 不选（跳过资源类型筛选）。"""
    default = normalize_doc_type(default or DOWNLOAD_DOC_TYPE)
    options = [
        ('1', 'journal', '学术期刊题录'),
        ('2', 'patent', '专利题录'),
        ('3', 'none', '不选'),
    ]
    default_key = '1'
    for key, value, _label in options:
        if value == default:
            default_key = key
            break
    print('\n导出类型：')
    for key, _value, label in options:
        mark = ' (默认)' if key == default_key else ''
        print(f'  {key}. {label}{mark}')
    print('  说明：选择「不选」将跳过结果页资源类型筛选。')
    while True:
        choice = input(f'请选择导出类型编号（直接回车用默认 {default_key}）: ').strip()
        if not choice:
            choice = default_key
        for key, value, label in options:
            if choice == key or choice == value or choice == label:
                print(f'已选择导出类型: {label}')
                return value
        print('输入无效，请重新选择')


def prompt_theme_keyword(default=''):
    """可选主题检索词；回车表示不限制主题。"""
    default = (default or '').strip()
    hint = f'，默认「{default}」' if default else '，回车=不按主题'
    raw = input(f'主题检索词（可选{hint}）: ').strip()
    theme = raw if raw else default
    if theme:
        print(f'已设置主题: {theme}')
    else:
        print('不按主题限制（仅分类+日期）')
    return theme


def normalize_export_task(task, index=None):
    """规范化单个导出任务字典。"""
    if not isinstance(task, dict):
        raise ValueError(f'任务格式错误: {task!r}')
    doc_type = normalize_doc_type(
        task.get('doc_type') or task.get('type') or DOWNLOAD_DOC_TYPE)
    category = task.get('category') or task.get('category_path')
    if isinstance(category, str):
        category = [p.strip() for p in category.replace('>', '/').split('/') if p.strip()]
    if category is None:
        category = []
    if not isinstance(category, list):
        raise ValueError(f'任务 category 格式错误: {task!r}')
    if doc_type == 'none':
        category = []
    elif not category:
        raise ValueError(f'任务缺少 category 路径: {task!r}')
    granularity = task.get('granularity') or CRAWL_GRANULARITY
    lang = normalize_lang(task.get('lang') or DOWNLOAD_LANG)
    theme = (task.get('theme') or task.get('keyword') or task.get('su') or '').strip()
    patent_date = normalize_patent_date_field(
        task.get('patent_date') or task.get('patent_date_field')
        or DEFAULT_PATENT_DATE_FIELD)
    start_s = task.get('bgdate') or bgdate
    end_s = task.get('initdate') or initdate
    datetime.strptime(start_s, '%Y-%m-%d')
    datetime.strptime(end_s, '%Y-%m-%d')
    return {
        'name': task.get('name') or (
            f'任务{index}' if index is not None else '任务'),
        'category': list(category),
        'doc_type': doc_type,
        'theme': theme,
        'patent_date': patent_date,
        'lang': lang,
        'granularity': granularity,
        'bgdate': start_s,
        'initdate': end_s,
    }


def print_export_tasks(tasks):
    """打印任务列表摘要。"""
    print(f'\n======== 导出任务列表（共 {len(tasks)} 个）========')
    for i, task in enumerate(tasks, 1):
        doc_type = normalize_doc_type(task.get('doc_type'))
        cat = category_path_label(task.get('category'), doc_type)
        if doc_type == 'patent':
            date_kind = patent_date_field_label(task.get('patent_date'))
        else:
            date_kind = '日期' if doc_type == 'none' else '发表'
        theme = task.get('theme') or ''
        theme_part = f' | 主题:{theme}' if theme else ''
        print(
            f'  {i}. [{task.get("name", f"任务{i}")}] '
            f'{doc_type_label(task.get("doc_type"))}/'
            f'{lang_label(task.get("lang"))} | '
            f'{granularity_label(task["granularity"])} | '
            f'{date_kind} {task["bgdate"]} ~ {task["initdate"]} | '
            f'{cat}{theme_part}'
        )
    print('==========================================')


def prompt_export_tasks():
    """启动时交互配置多个导出任务。"""
    tasks = []
    print('\n开始配置导出任务列表（可添加多个分类/粒度组合）')
    doc_type = prompt_doc_type(DOWNLOAD_DOC_TYPE)
    download_lang = 'zh'
    if doc_type == 'patent':
        print('专利导出默认中文语种设置')
    else:
        download_lang = prompt_download_lang(DOWNLOAD_LANG)
    while True:
        print(f'\n----- 添加第 {len(tasks) + 1} 个任务 -----')
        granularity = prompt_crawl_granularity(CRAWL_GRANULARITY)
        category_path = prompt_literature_category(doc_type)
        theme = prompt_theme_keyword(DEFAULT_THEME)
        patent_date = DEFAULT_PATENT_DATE_FIELD
        if doc_type == 'patent':
            patent_date = prompt_patent_date_field(DEFAULT_PATENT_DATE_FIELD)
        start_s, end_s = prompt_task_date_range(bgdate, initdate)
        task = normalize_export_task({
            'name': f'任务{len(tasks) + 1}',
            'category': category_path,
            'doc_type': doc_type,
            'theme': theme,
            'patent_date': patent_date,
            'lang': download_lang,
            'granularity': granularity,
            'bgdate': start_s,
            'initdate': end_s,
        }, index=len(tasks) + 1)
        tasks.append(task)
        theme_part = f' | 主题:{task["theme"]}' if task.get('theme') else ''
        print(
            f'已加入: {doc_type_label(task["doc_type"])}/'
            f'{lang_label(task["lang"])} | '
            f'{granularity_label(task["granularity"])} | '
            f'{task["bgdate"]}~{task["initdate"]} | '
            f'{category_path_label(task["category"], task["doc_type"])}'
            f'{theme_part}'
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
                    print(
                        f'已删除: '
                        f'{category_path_label(removed.get("category"), removed.get("doc_type"))}')
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
    """加载任务：优先 EXPORT_TASKS 配置，否则交互配置。"""
    if EXPORT_TASKS:
        tasks = [
            normalize_export_task(t, index=i)
            for i, t in enumerate(EXPORT_TASKS, 1)
        ]
        print('使用配置文件中的 EXPORT_TASKS')
        print_export_tasks(tasks)
        return tasks
    return prompt_export_tasks()


def is_driver_alive(driver):
    if driver is None:
        return False
    try:
        _ = driver.window_handles
        return True
    except Exception:
        return False


def set_chrome_download_dir(driver, download_dir):
    """运行中切换 Chrome 下载目录，换任务时不必关闭浏览器。"""
    abs_dir = os.path.abspath(download_dir)
    os.makedirs(abs_dir, exist_ok=True)
    payload = {
        'behavior': 'allow',
        'downloadPath': abs_dir,
    }
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


def run_export_task(driver, task, task_index, task_total):
    """执行单个导出任务（同一浏览器会话内）。返回 (driver, download_dir)。"""
    category_path = task['category']
    lang = normalize_lang(task.get('lang'))
    doc_type = normalize_doc_type(task.get('doc_type'))
    theme = (task.get('theme') or '').strip()
    patent_date = normalize_patent_date_field(task.get('patent_date'))
    granularity = task['granularity']
    start = datetime.strptime(task['bgdate'], '%Y-%m-%d').date()
    end = datetime.strptime(task['initdate'], '%Y-%m-%d').date()
    ranges = list(iter_crawl_date_ranges(start, end, granularity))

    download_dir = task_download_dir(task)
    os.makedirs(download_dir, exist_ok=True)
    write_export_meta(download_dir, task)

    theme_part = f' | 主题: {theme}' if theme else ''
    cat_part = category_path_label(category_path, doc_type)
    if doc_type == 'patent':
        date_kind = patent_date_field_label(patent_date)
    else:
        date_kind = '日期' if doc_type == 'none' else '发表'
    print(
        f'\n######## 任务 {task_index}/{task_total}: {task.get("name", "")} ########\n'
        f'类型: {doc_type_label(doc_type)}/{lang_label(lang)} | '
        f'分类: {cat_part}{theme_part}\n'
        f'粒度: {granularity_label(granularity)} | {date_kind}: '
        f'{task["bgdate"]} ~ {task["initdate"]} | 共 {len(ranges)} 段\n'
        f'保存: {download_dir}'
    )

    # 仅在新任务开始时主动重建浏览器；任务内日/月/年时段成功后不主动重启
    if is_driver_alive(driver):
        print(
            f'任务 {task_index}/{task_total}：主动关闭旧浏览器并重建'
            '（任务内时段切换不重启）...')
        try:
            driver.quit()
        except Exception as e:
            print('关闭旧浏览器失败:', e)
        cleanup_chrome_profile()
        driver = None

    driver = create_chrome_driver(download_dir)
    open_cnki_url(driver, get_adv_search_url(doc_type))
    try:
        driver.maximize_window()
    except Exception:
        pass
    login_ok = ensure_unit_login_if_needed(
        driver, restore_handle=driver.current_window_handle)
    cap0 = ensure_cnki_captcha_after_nav(
        driver, download_dir, driver.current_window_handle)
    still_login = page_suggests_login_lost(
        driver, check_handle=driver.current_window_handle)
    # 任务启动已过验证码后，各时段复用会话，不再直连探测跳转
    session_ready = (
        bool(cap0)
        and not captcha_needs_browser_restart(cap0)
        and login_ok
        and not still_login
    )
    if captcha_needs_browser_restart(cap0) or not cap0:
        print('启动时验证码未通过，将在首段检索中关闭浏览器重试')
        session_ready = False
    elif still_login or not login_ok:
        print('启动时仍停在 WebVPN/CAS 登录页，将在首段检索中重建浏览器并重新登录')
        session_ready = False
    else:
        print('已通过验证码验证，本任务后续时段将复用页面（不再重复滑块）')

    max_browser_restarts = 2
    for idx, (rs, re_) in enumerate(ranges, 1):
        start_str = rs.strftime('%Y-%m-%d')
        end_str = re_.strftime('%Y-%m-%d')
        label = start_str if start_str == end_str else f'{start_str}_{end_str}'
        print(f'\n----- 任务{task_index} 第 {idx}/{len(ranges)} 段: {label} -----')
        restart_left = max_browser_restarts
        while True:
            try:
                status = run_one_search_export(
                    driver, download_dir, category_path, start_str, end_str,
                    label, lang=lang, doc_type=doc_type, theme=theme,
                    patent_date_field=patent_date,
                    reuse_session=session_ready)
            except Exception as e:
                if exception_suggests_login_lost(e, driver):
                    print(
                        f'任务段 {label} 登录状态中断，'
                        '准备关闭浏览器重新登录后重试本时段:', e)
                    status = 'restart_browser'
                else:
                    print(
                        f'任务段 {label} 未捕获异常，记录并继续下一段:',
                        type(e).__name__, e)
                    try:
                        record_export_failure(download_dir, label)
                    except Exception:
                        pass
                    status = 'failed'
                    try:
                        close_export_windows(driver, main_handle=None)
                        if driver.window_handles:
                            driver.switch_to.window(driver.window_handles[0])
                        if session_ready:
                            open_adv_search_in_session(driver, doc_type)
                        else:
                            open_cnki_url(driver, get_adv_search_url(doc_type))
                        time.sleep(2)
                    except Exception:
                        pass

            # 仅失败恢复（频繁/验证码）才重建；正常完成时段不重启
            if status == 'restart_browser' and restart_left > 0:
                restart_left -= 1
                print(
                    f'{label} 关闭浏览器并重建'
                    f'（频繁、验证码或登录中断，剩余重建次数 {restart_left}）...')
                driver = recreate_browser_after_rate_limit(
                    driver, download_dir, doc_type=doc_type)
                session_ready = not page_suggests_login_lost(driver)
                if not session_ready:
                    print(
                        f'{label} 重建后仍停在 WebVPN/CAS 登录页，'
                        '本段将重新走登录流程')
                print(
                    f'{label} 浏览器已重建，继续导出题录'
                    '（重新检索本时段并导出）...')
                continue

            if status == 'restart_browser':
                print(f'{label} 浏览器重建次数用尽，记入失败并开始下一个')
                try:
                    record_export_failure(download_dir, label)
                except Exception:
                    pass
            elif status in ('ok', 'empty'):
                session_ready = True
                print(
                    f'{label} 本时段题录导出完成，继续下一段'
                    '（复用页面，不重启、不重新跳转）')
            elif status == 'failed' and is_driver_alive(driver):
                # 浏览器仍在则后续段继续复用，避免无谓再出滑块
                session_ready = True
            break

    return driver, download_dir


def cnkipaperfordatav4():
    """期刊/专利多任务检索导出：可选主题，支持分类/粒度/日期组合。"""
    tasks = load_export_tasks()
    if not tasks:
        print('无导出任务，结束')
        return

    preload_click_captcha()
    driver = None
    try:
        for i, task in enumerate(tasks, 1):
            driver, _download_dir = run_export_task(
                driver, task, i, len(tasks))
            print(f'\n######## 任务 {i}/{len(tasks)} 完成 ########')
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        cleanup_chrome_profile()
    print('\n全部导出任务已结束')


def create_chrome_driver(download_dir):
    """创建独立 Chrome 实例，可与 SCI 脚本并行。"""
    global _chrome_profile_dir
    os.makedirs(CHROME_PROFILE_ROOT, exist_ok=True)
    _chrome_profile_dir = tempfile.mkdtemp(
        prefix=f'cnki_{os.getpid()}_', dir=CHROME_PROFILE_ROOT)
    debug_port = CHROME_DEBUG_PORT_BASE + (os.getpid() % 400)
    cache_dir = os.path.join(_chrome_profile_dir, 'Cache')
    os.makedirs(cache_dir, exist_ok=True)

    chrome_options = Options()
    prefs = {
        'profile.default_content_settings.popups': 0,
        'download.default_directory': download_dir,
        'directory_upgrade': True,
    }
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument(f'--user-data-dir={_chrome_profile_dir}')
    chrome_options.add_argument(f'--disk-cache-dir={cache_dir}')
    chrome_options.add_argument('--disk-cache-size=1')
    chrome_options.add_argument('--disable-gpu-shader-disk-cache')
    chrome_options.add_argument(f'--remote-debugging-port={debug_port}')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--no-default-browser-check')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    print(f'Chrome profile: {_chrome_profile_dir} | debug-port={debug_port}')
    driver = webdriver.Chrome(options=chrome_options)
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


def collect_day_results(driver, save_dir, main_handle, day_str=None):
    """当日结果：逐页全选，每500条导出一次，翻页至末页导出剩余。

    返回:
      'ok'              — 有结果且全部导出成功
      'empty'           — 无检索结果
      'failed'          — 导出失败已中止（已记失败日志）
      'restart_browser' — 请求过于频繁 / 验证码连败 / 登录中断，需关浏览器重试本时段
    """
    count = 1
    page_count = 0
    has_results = False
    export_failed = False
    page_error = False
    need_browser_restart = False

    def _do_export(fail_msg='本批导出失败，取消后续导出，开始下一个'):
        nonlocal export_failed, need_browser_restart
        result = run_download_excel_step(driver, save_dir, main_handle)
        if result == 'restart_browser':
            need_browser_restart = True
            export_failed = True
            print(fail_msg + '（需重启浏览器）')
            return False
        if result is True:
            return True
        print(fail_msg)
        export_failed = True
        return False

    def _export_remainder_and_finish(last_page):
        """翻到末页后导出未满 500 的剩余已选。"""
        nonlocal export_failed, page_error, need_browser_restart
        rem = last_page % 10
        if not has_results or rem == 0 or export_failed:
            return
        selected = get_selected_count(driver)
        if selected == 0:
            print(f'末页已选为 0，跳过导出（rem={rem}），记失败交给补全')
            page_error = True
            return
        try:
            has_toolbar = bool(driver.find_elements(By.CLASS_NAME, 'secondUl'))
        except Exception:
            has_toolbar = False
        if not has_toolbar:
            if page_suggests_login_lost(driver, check_handle=main_handle):
                print('末页导出前检测到登录中断，需重建浏览器')
                need_browser_restart = True
            else:
                print('末页导出前未找到导出工具栏，跳过本批并记失败')
                page_error = True
            return
        print(f'末页处理：导出剩余 {rem} 页（约至第{last_page}页）')
        if not _do_export('末批导出失败'):
            pass

    # 开始采集前先清理上一轮残留已选，避免第 1 页起就达到 500 条上限
    try:
        initial_selected = get_selected_count(driver)
    except Exception:
        initial_selected = None
    if initial_selected is not None and initial_selected > 0:
        print(f'采集开始前检测到残留已选 {initial_selected} 条，先导出并清空')
        _do_export('采集开始前残留批次导出失败')
        clear_selected_results(driver, main_handle)

    while True:
        if need_browser_restart:
            break
        alert_text = dismiss_alerts(driver, timeout=0.3)
        print(f'第{count}页')
        if text_suggests_login_lost(alert_text) or page_suggests_login_lost(
                driver, check_handle=main_handle):
            print('检测到登录状态中断（VPN/机构登录失效），需重建浏览器后重试本时段')
            need_browser_restart = True
            break
        if has_captcha_anywhere(driver):
            cap = ensure_captcha_passed(driver, save_dir, main_handle)
            if captcha_needs_browser_restart(cap):
                need_browser_restart = True
                break
            continue
        time.sleep(3)
        reloadele = driver.find_elements(By.ID, 'reload-button')
        if len(reloadele) > 0:
            print('需要重新加载')
            reloadele[0].click()
            time.sleep(3)
            continue

        # 若上次清空失败导致已选仍满 500：优先清空，避免把同一批再导出一遍
        selected = get_selected_count(driver)
        if selected is not None and selected >= 500:
            print(f'检测到已选 {selected} 条（达上限），先导出当前批次再清空，避免丢数据或重复导出')
            if not _do_export('已选达上限，先导出当前批次'):
                clear_selected_results(driver, main_handle)
                break
            if not clear_selected_results(driver, main_handle):
                print('导出成功但清空已选失败，本时段记失败')
                page_error = True
                break
            continue

        try:
            WebDriverWait(driver, 20, 0.5).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'checkAll')))
            selallclick = driver.find_element(By.CLASS_NAME, 'checkAll').find_element(
                By.CSS_SELECTOR, 'input')
            driver.execute_script("arguments[0].click();", selallclick)
            has_results = True
            alert_text = dismiss_alerts(driver, timeout=0.8)
            if text_suggests_login_lost(alert_text):
                print('勾选时提示登录状态中断，需重建浏览器后重试本时段')
                need_browser_restart = True
                break
            if alert_text and '不能超过500' in alert_text:
                print('勾选触发超过500上限，优先清空后再重选本页')
                if not clear_selected_results(driver, main_handle):
                    print('清空失败，尝试导出+清空')
                    if not _do_export():
                        clear_selected_results(driver, main_handle)
                        break
                # 清空后重选当前页
                try:
                    selallclick = driver.find_element(
                        By.CLASS_NAME, 'checkAll').find_element(By.CSS_SELECTOR, 'input')
                    driver.execute_script("arguments[0].click();", selallclick)
                    dismiss_alerts(driver, timeout=0.5)
                except Exception as e:
                    print('清空后重选本页失败:', e)
        except Exception as e:
            dismiss_alerts(driver, timeout=0.5)
            if has_captcha_anywhere(driver):
                cap = ensure_captcha_passed(driver, save_dir, main_handle)
                if captcha_needs_browser_restart(cap):
                    need_browser_restart = True
                    break
                continue
            if text_suggests_login_lost(e) or page_suggests_login_lost(
                    driver, check_handle=main_handle):
                print('提取失败且登录状态中断，需重建浏览器后重试本时段')
                need_browser_restart = True
                break
            print('当日无更多结果或提取失败:', type(e).__name__, e)
            # 当前页未勾选成功：先按上一页导出剩余，再判断是正常末页还是页面异常
            _export_remainder_and_finish(count - 1 if count > 1 else 0)
            if not page_has_no_results(driver):
                print('页面未正常加载（非明确无结果），本时段记失败，交给补全重导')
                page_error = True
            break

        time.sleep(2)
        if count % 10 == 0 and count > page_count:
            print(f'第{count // 10}个表格开始下载（500条）')
            if not _do_export():
                clear_selected_results(driver, main_handle)
                break

        try:
            next_ele = driver.find_element(By.ID, 'Page_next_top')
            driver.execute_script("arguments[0].click();", next_ele)
            alert_text = dismiss_alerts(driver, timeout=0.8)
            if text_suggests_login_lost(alert_text):
                print('翻页时提示登录状态中断，需重建浏览器后重试本时段')
                need_browser_restart = True
                break
            if alert_text and '不能超过500' in alert_text:
                print('翻页时提示超过500，优先清空再继续（已导出过则不再导出）')
                if not clear_selected_results(driver, main_handle):
                    if not _do_export():
                        clear_selected_results(driver, main_handle)
                        break
                try:
                    next_ele = driver.find_element(By.ID, 'Page_next_top')
                    driver.execute_script("arguments[0].click();", next_ele)
                    dismiss_alerts(driver, timeout=0.5)
                except Exception as e2:
                    print('清空后仍无法翻页，视为末页:', e2)
                    _export_remainder_and_finish(count)
                    break
            count += 1
            time.sleep(3)
        except Exception as e:
            alert_text = dismiss_alerts(driver, timeout=0.8)
            if alert_text and '不能超过500' in (alert_text or ''):
                print('翻页异常且因超过500：清空后重试翻页，不按末页重复导出')
                if not clear_selected_results(driver, main_handle):
                    if not _do_export():
                        clear_selected_results(driver, main_handle)
                        break
                continue
            if text_suggests_login_lost(e) or page_suggests_login_lost(
                    driver, check_handle=main_handle):
                print('翻页失败且登录状态中断，需重建浏览器后重试本时段')
                need_browser_restart = True
                break
            print('翻页结束:', e)
            _export_remainder_and_finish(count)
            break

    if need_browser_restart:
        # 不在此处记失败：上层关浏览器重试本时段，成功后再移除记录
        return 'restart_browser'
    if export_failed or page_error:
        if day_str:
            record_export_failure(save_dir, day_str)
        return 'failed'
    return 'ok' if has_results else 'empty'

global next_flag
next_flag=False

def cnkipaperfordata(keywords, keywords2):
    global initdate
    global next_flag
    download_dir = r'D:\2024年工作\知网数据库下载'  # 路径一定是斜杠，不能是反斜杠
    chrome_options = Options()

    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': download_dir,
             "directory_upgrade": True}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=chrome_options)

    #driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
    #driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
    driver.get(get_adv_search_url())
    driver.maximize_window()  # 最大化窗口
    time.sleep(3)
    ensure_unit_login_if_needed(driver, restore_handle=driver.current_window_handle)
    select_eleddlist=driver.find_element(By.ID, 'gradetxt').find_elements(By.CSS_SELECTOR, 'dd')
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

    #去掉英文文献
    input=driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
    try:
        input.click()
    except Exception as e:
        print(e)
    # input.click()
    time.sleep(3)
    #填写开始时间
    times=driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
    #start_time=times[0]#.send_keys('2011-01-01')
    #start_time.send_keys('2011-01-01')
    date_txt_bg = driver.find_element(By.ID, 'datebox0')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
    date_txt_bg.clear()  # 先清除原来的日期值
    date_txt_bg.send_keys(bgdate)
    #date_txt_bg.send_keys(initdate)
    #print('更新开始论文时间：',initdate)

    date_txt_end = driver.find_element(By.ID, 'datebox1')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
    date_txt_end.clear()  # 先清除原来的日期值
    # date_txt.send_keys('2010-01-01')
    date_txt_end.send_keys(initdate)
    print('更新结束论文时间：', initdate)

    print(times[0].get_attribute('outerHTML'))
    time.sleep(5)
    #start_time = time[0].send_keys('2011-01-01')
    searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
    #searchbtn.click()
    driver.execute_script("arguments[0].click();", searchbtn)
    time.sleep(5)
    now_url = driver.current_url #当前页面
    print('主页:',now_url)

    # 只采集学术期刊
    print('只选择学术期刊')
    journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
    journal.click()
    time.sleep(3)

    main_handle = driver.current_window_handle
    nextbutton=True
    count=1
    page_count=0
    while nextbutton == True:
        nextbutton = False
        print("第%s页" % count)
        if count>=page_count:
            if handle_block_puzzle_captcha(driver, download_dir, main_handle):
                nextbutton = True
                continue
            time.sleep(3)
            # 2重新加载、
            reloadele=driver.find_elements(By.ID, 'reload-button')
            if len(reloadele) > 0:
                print('需要重新加载')
                reloadbtn=reloadele[0]
                reloadbtn.click()
                time.sleep(3)
                nextbutton = True
                continue


        try:
            #1、题录全选
            WebDriverWait(driver, 20, 0.5).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'checkAll')))

            #倒叙排列，点击排序按钮

            #selallclick.click()
            # 开始点击全选
            selallclick = driver.find_element(By.CLASS_NAME, 'checkAll').find_element(By.CSS_SELECTOR, 'input')
            driver.execute_script("arguments[0].click();", selallclick)
            datelist=driver.find_element(By.CLASS_NAME, 'result-table-list').find_elements(By.CLASS_NAME, 'date')
            if (len(datelist)>0) and  count%10==1:
                initdates=datelist[0].get_attribute('innerHTML')
                #print(initdates)
                #initdate=initdates.split(' ')[1]
                initdate = re.findall(r'\d{4}-\d{2}-\d{2}', initdates)[0]
                print('更新initdate:', initdate)
        except Exception as e:
            if handle_block_puzzle_captcha(driver, download_dir, main_handle):
                nextbutton = True
                continue
            driver.quit()
            print('提取题录失败，关闭网页:',e)
            nextbutton = False
            break

        time.sleep(2)

        if ((count%10)==0 and count>page_count):#500下载excel
            print('第%s个表格开始下载'%(count/10))
            run_download_excel_step(driver, download_dir, main_handle)
        #2、翻页
        try:
            next_ele=driver.find_element(By.ID, 'Page_next_top')
            #next_ele.click()
            driver.execute_script("arguments[0].click();", next_ele)
            nextbutton = True
            time.sleep(3)
        except Exception as e:
            print("翻页失败",e)
            next_flag=True
            nextbutton = False
            if (count%25)!=0:
               run_download_excel_step(driver, download_dir, main_handle)
        count = count + 1

    driver.quit()

def spiderfordate(text):#增加日期
        global next_flag
        global initdate
        while ('2010-06-' in initdate)==False and next_flag==False:
            print("开始：initdate", initdate)
            try:
              cnkipaperfordata(text)
            except Exception as e:
                print(e)


from datetime import datetime
def rename_folder(old_path, new_name):
    # 构造新的文件夹路径
    new_path = os.path.join(os.path.dirname(old_path), new_name)
    if  os.path.exists(new_path):
        #c# 获取当前时间，精确到秒
        now = datetime.now().replace(microsecond=0)
        print(now)
        formatted_time = now.strftime('%Y-%m-%d %H-%M-%S')
        os.rename(new_path, new_path+str(formatted_time))
    os.rename(old_path, new_path)
#新增：将页面处理函数全都放到以下函数
def setexcelcon(driver,pdfpath, subject):
    main_handle = driver.current_window_handle
    nextbutton = True
    count = 1
    page_count = 0
    #翻页采集
    while nextbutton == True:
        nextbutton = False
        print("第%s页" % count)
        if count >= page_count:
            if handle_block_puzzle_captcha(driver, pdfpath, main_handle):
                nextbutton = True
                continue
            time.sleep(3)
            # 2重新加载、
            reloadele = driver.find_elements(By.ID, 'reload-button')
            if len(reloadele) > 0:
                print('需要重新加载')
                reloadbtn = reloadele[0]
                reloadbtn.click()
                time.sleep(3)
                nextbutton = True
                continue

        try:
            # 1、题录全选
            WebDriverWait(driver, 20, 0.5).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'checkAll')))

            # 倒叙排列，点击排序按钮

            # selallclick.click()
            # 开始点击全选
            selallclick = driver.find_element(By.CLASS_NAME, 'checkAll').find_element(By.CSS_SELECTOR, 'input')
            driver.execute_script("arguments[0].click();", selallclick)
            datelist = driver.find_element(By.CLASS_NAME, 'result-table-list').find_elements(By.CLASS_NAME, 'date')
            if (len(datelist) > 0) and count % 10 == 1:
                initdates = datelist[0].get_attribute('innerHTML')
                # print(initdates)
                # initdate=initdates.split(' ')[1]
                initdate = re.findall(r'\d{4}-\d{2}-\d{2}', initdates)[0]
                print('更新initdate:', initdate)
        except Exception as e:
            if handle_block_puzzle_captcha(driver, pdfpath, main_handle):
                nextbutton = True
                continue
            driver.quit()
            print('提取题录失败，关闭网页:', e)
            nextbutton = False
            break

        time.sleep(2)

        if ((count % 10) == 0 and count > page_count):  # 500下载excel
            print('第%s个表格开始下载' % (count / 10))
            run_download_excel_step(driver, pdfpath, main_handle)
        # 2、翻页
        try:
            next_ele = driver.find_element(By.ID, 'Page_next_top')
            # next_ele.click()
            driver.execute_script("arguments[0].click();", next_ele)
            nextbutton = True
            time.sleep(3)
        except Exception as e:
            print("翻页失败", e)
            next_flag = True
            nextbutton = False
            if (count % 10) != 0:
                run_download_excel_step(driver, pdfpath, main_handle)
                print('该学科下载完成===========================\n')
        count = count + 1

    #进行修改名称，然后再重建一个文件夹
    rename_folder(pdfpath, subject)
    if not os.path.exists(pdfpath):
        # 如果文件夹不存在，则创建文件夹
        os.makedirs(pdfpath)


#新增接口，增加学术期刊，
def cnkipaperfordatav2(text):
    global initdate
    global next_flag
    download_dir = os.path.join(r'D:\2024年工作\知网数据库下载', text, 'excel下载')
    chrome_options = Options()
    print(download_dir)
    if not os.path.exists(download_dir):
        # 如果文件夹不存在，则创建文件夹
        os.makedirs(download_dir)

    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': download_dir,
             "directory_upgrade": True}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=chrome_options)

    #driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
    #driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
    driver.get(get_adv_search_url())
    driver.maximize_window()  # 最大化窗口
    time.sleep(3)
    ensure_unit_login_if_needed(driver, restore_handle=driver.current_window_handle)
    select_eleddlist=driver.find_element(By.ID, 'gradetxt').find_elements(By.CSS_SELECTOR, 'dd')
    #获取第一个选择项
    sel_dd=select_eleddlist[0]
    dd_cli=sel_dd.find_element(By.CLASS_NAME, 'sort.reopt')
    dd_cli.click()
    time.sleep(3)
    #获取下拉选项:文献来源
    #sel_li=dd_cli.find_element(By.CLASS_NAME, 'sort-list').find_element(By.CSS_SELECTOR, 'li[data-val="LY"]')
    sel_li=dd_cli.find_element(By.CLASS_NAME, 'sort-list').find_element(By.CSS_SELECTOR, 'li[data-val="SU"]')
    print(sel_li.get_attribute('outerHTML'))
    sel_li.click()
    time.sleep(3)
    text_element = sel_dd.find_element(By.CSS_SELECTOR, 'input[type="text"]')
    print('text_element:',text_element.get_attribute('outerHTML'))
    time.sleep(3)
    text_element.send_keys(text)
    time.sleep(3)

    #去掉英文文献
    input=driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
    input.click()
    time.sleep(3)
    #填写开始时间
    times=driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
    #start_time=times[0]#.send_keys('2011-01-01')
    #start_time.send_keys('2011-01-01')
    date_txt_bg = driver.find_element(By.ID, 'datebox0')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
    date_txt_bg.clear()  # 先清除原来的日期值
    date_txt_bg.send_keys('2010-08-24')
    #date_txt_bg.send_keys(initdate)
    #print('更新开始论文时间：',initdate)

    date_txt_end = driver.find_element(By.ID, 'datebox1')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
    date_txt_end.clear()  # 先清除原来的日期值
    # date_txt.send_keys('2010-01-01')
    date_txt_end.send_keys(initdate)
    print('更新结束论文时间：', initdate)

    print(times[0].get_attribute('outerHTML'))
    time.sleep(5)
    #start_time = time[0].send_keys('2011-01-01')
    searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
    #searchbtn.click()
    driver.execute_script("arguments[0].click();", searchbtn)
    time.sleep(5)
    now_url = driver.current_url #当前页面
    print('主页:',now_url)

    # 只采集学术期刊
    print('只选择学术期刊')
    journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
    journal.click()
    time.sleep(3)

    # 只选择核心期刊
    print('只选择北大核心期刊')
    LYBSMele = driver.find_element(By.CSS_SELECTOR, 'dl[groupid="LYBSM"]')
    driver.execute_script("arguments[0].class= '  ';", LYBSMele)
    LYBSMele.click()
    time.sleep(3)
    # print(LYBSMele.get_attribute('outerHTML'))
    # print(LYBSMele.get_attribute('innerHTML'))
    central = LYBSMele.find_element(By.CSS_SELECTOR, 'input[text="北大核心"]')
    central.click()
    time.sleep(3)

    # 分学科进行采集
    print('每个选项选择学科属性')
    cclele = driver.find_element(By.CSS_SELECTOR, 'dl[groupid="CCL"]')
    driver.execute_script("arguments[0].class= '  ';", cclele)
    cclele.click()
    time.sleep(3)
    # cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
    ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
    print(len(ccllist))
    j = 0
    while j < len(ccllist):
        if j > 9:
            driver.execute_script("arguments[0].style.display = 'list-item';", ccllist[j])
        # print(ccllist[j].get_attribute('outerHTML'))
        checkbox = ccllist[j].find_element(By.CSS_SELECTOR, 'input')
        print("\n===========选择学科%s,名称：%s" % (j, checkbox.get_attribute('title')))
        subject = checkbox.get_attribute('title')  # 获取学科类别
        # print(checkbox.get_attribute('outerHTML'))
        checkbox.click()  # 点击选中
        time.sleep(3)

        # 下载pdf文件
        setexcelcon(driver, download_dir, subject)

        windows = driver.window_handles
        print("当前有%s个窗口" % len(windows))
        # 下载完成之后需要再取消一下,需要重新获取一下当前的节点
        cclelev2 = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        ccllistv2 = cclelev2.find_elements(By.CSS_SELECTOR, 'input[checked="checked"]')
        # checkbox = ccllistv2[0].find_element(By.CSS_SELECTOR, 'input')
        ccllistv2[0].click()
        time.sleep(2)

        # 重新获取一下自选框,取消完之后重新获取
        cclele = driver.find_element(By.CSS_SELECTOR, 'dd[field="CCL"]')
        ccllist = cclele.find_elements(By.CSS_SELECTOR, 'li')
        j = j + 1
    driver.quit()

def cnkipaperfordatav3(keywords):
    global initdate
    global next_flag
    download_dir = r'D:\2024年工作\知网数据库下载'  # 路径一定是斜杠，不能是反斜杠
    chrome_options = Options()

    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': download_dir,
             "directory_upgrade": True}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=chrome_options)

    #driver = webdriver.Chrome('F:\ChromDriver\chromedriver_win32\chromedriver.exe')
    #driver.get('https://kns.cnki.net/KNS8/AdvSearch?dbcode=SCOD&searchType=majorSearch')
    driver.get(get_adv_search_url())
    driver.maximize_window()  # 最大化窗口
    ensure_unit_login_if_needed(driver, restore_handle=driver.current_window_handle)

    # 等待手动输入检索条件
    search_condition = input("请输入检索条件后按回车继续: ")

    #去掉英文文献
    input_en = driver.find_element(By.CSS_SELECTOR, 'input[data-id="EN"]')
    try:
        input_en.click()
    except Exception as e:
        print(e)
    # input.click()
    time.sleep(3)
    #填写开始时间
    times=driver.find_element(By.CLASS_NAME, 'tit-date-box').find_elements(By.CSS_SELECTOR, 'input')
    #start_time=times[0]#.send_keys('2011-01-01')
    #start_time.send_keys('2011-01-01')
    date_txt_bg = driver.find_element(By.ID, 'datebox0')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_bg)
    date_txt_bg.clear()  # 先清除原来的日期值
    date_txt_bg.send_keys(bgdate)
    #date_txt_bg.send_keys(initdate)
    #print('更新开始论文时间：',initdate)

    date_txt_end = driver.find_element(By.ID, 'datebox1')
    driver.execute_script("arguments[0].removeAttribute('readonly');", date_txt_end)
    date_txt_end.clear()  # 先清除原来的日期值
    # date_txt.send_keys('2010-01-01')
    date_txt_end.send_keys(initdate)
    print('更新结束论文时间：', initdate)

    print(times[0].get_attribute('outerHTML'))
    time.sleep(5)
    #start_time = time[0].send_keys('2011-01-01')
    searchbtn = driver.find_element(By.CLASS_NAME, 'btn-search')
    #searchbtn.click()
    driver.execute_script("arguments[0].click();", searchbtn)
    time.sleep(5)
    now_url = driver.current_url #当前页面
    print('主页:',now_url)

    set_result_page_size(driver, 50)

    # 只采集学术期刊
    print('只选择学术期刊')
    journal = driver.find_element(By.CSS_SELECTOR, 'a[resource="JOURNAL"]')
    journal.click()
    time.sleep(3)

    main_handle = driver.current_window_handle
    nextbutton=True
    count=1
    page_count=0
    while nextbutton == True:
        nextbutton = False
        print("第%s页" % count)
        if count>=page_count:
            if handle_block_puzzle_captcha(driver, download_dir, main_handle):
                nextbutton = True
                continue
            time.sleep(3)
            # 2重新加载、
            reloadele=driver.find_elements(By.ID, 'reload-button')
            if len(reloadele) > 0:
                print('需要重新加载')
                reloadbtn=reloadele[0]
                reloadbtn.click()
                time.sleep(3)
                nextbutton = True
                continue


        try:
            #1、题录全选
            WebDriverWait(driver, 20, 0.5).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'checkAll')))

            #倒叙排列，点击排序按钮

            #selallclick.click()
            # 开始点击全选
            selallclick = driver.find_element(By.CLASS_NAME, 'checkAll').find_element(By.CSS_SELECTOR, 'input')
            driver.execute_script("arguments[0].click();", selallclick)
            datelist=driver.find_element(By.CLASS_NAME, 'result-table-list').find_elements(By.CLASS_NAME, 'date')
            if (len(datelist)>0) and  count%10==1:
                initdates=datelist[0].get_attribute('innerHTML')
                #print(initdates)
                #initdate=initdates.split(' ')[1]
                initdate = re.findall(r'\d{4}-\d{2}-\d{2}', initdates)[0]
                print('更新initdate:', initdate)
        except Exception as e:
            if handle_block_puzzle_captcha(driver, download_dir, main_handle):
                nextbutton = True
                continue
            driver.quit()
            print('提取题录失败，关闭网页:',e)
            nextbutton = False
            break

        time.sleep(2)

        if ((count%10)==0 and count>page_count):#500下载excel
            print('第%s个表格开始下载'%(count/10))
            run_download_excel_step(driver, download_dir, main_handle)
        #2、翻页
        try:
            next_ele=driver.find_element(By.ID, 'Page_next_top')
            #next_ele.click()
            driver.execute_script("arguments[0].click();", next_ele)
            nextbutton = True
            time.sleep(3)
        except Exception as e:
            print("翻页失败",e)
            next_flag=True
            nextbutton = False
            if (count%25)!=0:
               run_download_excel_step(driver, download_dir, main_handle)
        count = count + 1

    driver.quit()


    driver.quit()


def parse_fill_label(label):
    """失败日志标签 -> (start_date, end_date, start_str, end_str)。"""
    import re as _re
    from datetime import datetime as _dt
    text = (label or '').strip()
    m = _re.match(r'^(\d{4}-\d{2}-\d{2})(?:_(\d{4}-\d{2}-\d{2}))?$', text)
    if not m:
        return None
    s = _dt.strptime(m.group(1), '%Y-%m-%d').date()
    e = _dt.strptime(m.group(2) or m.group(1), '%Y-%m-%d').date()
    if s > e:
        s, e = e, s
    return s, e, s.isoformat(), e.isoformat()


def do_search_return_hit(
        driver, download_dir, category_path, start_str, end_str, lang,
        doc_type=None, theme=None, patent_date_field=None,
        reuse_session=False):
    """打开检索页、选类、可选主题、设日期、检索、筛期刊/专利，返回 hit_count。

    特殊返回：'restart_browser'（验证码连败需关浏览器）。
    reuse_session=True：同会话内回到高级检索，不做直连探测。
    """
    doc_type = normalize_doc_type(doc_type)
    theme = (theme or '').strip()
    patent_date_field = normalize_patent_date_field(patent_date_field)
    if reuse_session:
        print('补全：复用已验证会话打开高级检索')
        open_adv_search_in_session(driver, doc_type)
        time.sleep(0.8)
        main_handle = driver.current_window_handle
        if page_suggests_login_lost(driver, check_handle=main_handle):
            print('补全：复用会话时登录已中断，需重建浏览器后重试')
            return 'restart_browser'
        try:
            cur = driver.current_url or ''
        except Exception:
            cur = ''
        if is_cnki_verify_url(cur) or has_captcha_anywhere(driver):
            cap = ensure_cnki_captcha_after_nav(
                driver, download_dir, main_handle)
            if captcha_needs_browser_restart(cap):
                return 'restart_browser'
            if not cap:
                return None
    else:
        open_cnki_url(driver, get_adv_search_url(doc_type))
        time.sleep(1)
        ensure_unit_login_if_needed(
            driver, restore_handle=driver.current_window_handle)
        main_handle = driver.current_window_handle
        if page_suggests_login_lost(driver, check_handle=main_handle):
            print('补全：登录未恢复，需重建浏览器后重试')
            return 'restart_browser'
        cap = ensure_cnki_captcha_after_nav(driver, download_dir, main_handle)
        if captcha_needs_browser_restart(cap):
            return 'restart_browser'
        if not cap:
            return None
    if doc_type != 'patent':
        prepare_adv_search_page(driver, lang)
    else:
        prepare_patent_search_page(driver)
    select_literature_category(driver, category_path, doc_type=doc_type)
    set_adv_theme_keyword(driver, theme, doc_type=doc_type)
    set_adv_date_range(
        driver, start_str, end_str,
        doc_type=doc_type, patent_date_field=patent_date_field)
    click_search(driver)
    main_handle = driver.current_window_handle
    if has_captcha_anywhere(driver):
        cap = ensure_captcha_passed(driver, download_dir, main_handle)
        if captcha_needs_browser_restart(cap):
            return 'restart_browser'
        if cap is False:
            return None
    sps = set_result_page_size(
        driver, 50, save_dir=download_dir, main_handle=main_handle)
    if captcha_needs_browser_restart(sps):
        return 'restart_browser'
    if not sps:
        return None
    apply_resource_filter(driver, doc_type=doc_type, lang=lang)
    return read_search_hit_count(driver)


def export_with_split(
        driver, download_dir, category_path, label, start, end, hit_count, lang,
        main_handle, doc_type=None, theme=None, patent_date_field=None):
    """超6000则按子类拆分；叶子超限写入 export_over_6000.txt。返回 True=已处理。

    导出类型=不选时未勾选学科，无法按子类拆分，超限直接记入 over_6000。
    """
    doc_type = normalize_doc_type(doc_type)
    theme = (theme or '').strip()
    patent_date_field = normalize_patent_date_field(patent_date_field)
    if hit_count is not None and hit_count > SEARCH_HIT_LIMIT:
        # 不选类型未使用学科导航，不能按子类拆分
        children = (
            [] if doc_type == 'none'
            else category_children(category_path, doc_type=doc_type))
        if children:
            print(f'{label} 共 {hit_count} 条，拆 {len(children)} 个子类')
            ok_all = True
            for child in children:
                sub_path = list(category_path) + [child]
                try:
                    sub_hit = do_search_return_hit(
                        driver, download_dir, sub_path,
                        start.isoformat(), end.isoformat(), lang,
                        doc_type=doc_type, theme=theme,
                        patent_date_field=patent_date_field,
                        reuse_session=True)
                except Exception as e:
                    if exception_suggests_login_lost(e, driver):
                        print(f'{label} 子类 {child} 检索时登录中断，需重建浏览器')
                        return 'restart_browser'
                    print(f'{label} 子类 {child} 检索失败: {e}')
                    record_export_failure(download_dir, label)
                    ok_all = False
                    continue
                if captcha_needs_browser_restart(sub_hit):
                    print(f'{label} 子类 {child} 验证码连败，需重启浏览器')
                    return 'restart_browser'
                if sub_hit is not None and sub_hit > SEARCH_HIT_LIMIT:
                    split_ret = export_with_split(
                            driver, download_dir, sub_path, label,
                            start, end, sub_hit, lang, driver.current_window_handle,
                            doc_type=doc_type, theme=theme,
                            patent_date_field=patent_date_field)
                    if captcha_needs_browser_restart(split_ret):
                        return 'restart_browser'
                    if not split_ret:
                        ok_all = False
                    continue
                if sub_hit == 0:
                    continue
                try:
                    driver.switch_to.window(driver.current_window_handle)
                    sub_status = collect_day_results(
                        driver, download_dir, driver.current_window_handle,
                        day_str=None)
                    if sub_status == 'restart_browser':
                        print(
                            f'{label} 子类 {child} 需重启浏览器'
                            f'（频繁或验证码连败）')
                        return 'restart_browser'
                    if sub_status == 'failed':
                        print(f'{label} 子类 {child} 导出失败，取消后续')
                        ok_all = False
                        record_export_failure(download_dir, label)
                        break
                except Exception as e:
                    if exception_suggests_login_lost(e, driver):
                        print(f'{label} 子类 {child} 导出时登录中断，需重建浏览器')
                        return 'restart_browser'
                    print(f'{label} 子类 {child} 导出失败: {e}')
                    ok_all = False
                    record_export_failure(download_dir, label)
                    break
            if ok_all:
                remove_export_failure(download_dir, label)
            return ok_all
        record_over_6000(download_dir, label, category_path, hit_count)
        return False
    # 未超限：当前结果页直接导出
    try:
        status = collect_day_results(
            driver, download_dir, main_handle, day_str=label)
        if status == 'ok' or (status == 'empty' and hit_count == 0):
            remove_export_failure(download_dir, label)
            return True
        if status == 'restart_browser':
            print(f'{label} 需重启浏览器（频繁或验证码连败）')
            return 'restart_browser'
        if status == 'failed':
            print(f'{label} 导出失败，保留失败记录，开始下一个')
            return False
        # empty 但 hit_count>0：异常，保留/记入失败
        print(f'{label} 无勾选结果但命中 {hit_count}，记入失败')
        record_export_failure(download_dir, label)
        return False
    except Exception as e:
        print(f'{label} 导出异常: {e}')
        record_export_failure(download_dir, label)
        return False


def run_fill_one_label(
        driver, download_dir, category_path, lang, label,
        doc_type=None, theme=None, patent_date=None, reuse_session=False):
    """补全单个失败时段：先检索出条数，本地一致则跳过，否则删文件重导。"""
    parsed = parse_fill_label(label)
    if not parsed:
        print(f'跳过无法解析的失败标签: {label}')
        return False
    start, end, start_str, end_str = parsed
    doc_type = normalize_doc_type(doc_type)
    theme = (theme or '').strip()
    patent_date = normalize_patent_date_field(patent_date)
    if doc_type == 'patent' and patent_date != 'OD':
        # 知网专利题录只含公开日期，按申请日检索无法用本地文件核对条数
        local_n, mapped = 0, []
    else:
        local_n, mapped = local_period_stats(
            str(download_dir), start, end,
            doc_type=doc_type, patent_date_field=patent_date)
    try:
        hit = do_search_return_hit(
            driver, download_dir, category_path, start_str, end_str, lang,
            doc_type=doc_type, theme=theme,
            patent_date_field=patent_date, reuse_session=reuse_session)
    except Exception as e:
        if exception_suggests_login_lost(e, driver):
            print(f'{label} 检索时登录状态中断，需重建浏览器后重试: {e}')
            return 'restart_browser'
        print(f'{label} 检索失败，保留失败记录: {type(e).__name__}: {e}')
        record_export_failure(str(download_dir), label)
        return False
    if captcha_needs_browser_restart(hit):
        print(f'{label} 检索验证码连败，需关闭浏览器后重试')
        return 'restart_browser'
    if hit is None:
        print(f'{label} 条数解析失败，按普通导出处理')
        if mapped and local_n > 0:
            hit = local_n
        elif category_children(category_path, doc_type=doc_type):
            # 有子类可拆时保守按超限处理，避免整段漏导
            hit = SEARCH_HIT_LIMIT + 1
        else:
            # 叶子分类无法再拆，直接导出当前结果页
            hit = None
    # 类型2：本地已有且条数一致 -> 跳过
    if mapped and local_n > 0 and hit == local_n:
        print(f'{label} 知网 {hit} == 本地 {local_n}，跳过重导')
        remove_export_failure(str(download_dir), label)
        return True
    if mapped and local_n > 0 and hit != local_n:
        print(f'{label} 知网 {hit} != 本地 {local_n}，删 {len(mapped)} 个文件后重导')
        delete_period_files(mapped)
    if hit == 0:
        print(f'{label} 知网 0 条，移除失败记录')
        remove_export_failure(str(download_dir), label)
        return True
    main_handle = driver.current_window_handle
    return export_with_split(
        driver, str(download_dir), category_path, label,
        start, end, hit, lang, main_handle,
        doc_type=doc_type, theme=theme, patent_date_field=patent_date)


def run_fill_mode(
        roots=None, limit=0, dry_run=False, doc_type=None, patent_date=None,
        theme=None):
    """补全模式：只按各目录 export_failures.txt 检索（期刊/专利均支持）。"""
    from cnki_captcha import preload_click_captcha
    root_paths = [Path(p) for p in roots] if roots else [Path(DOWNLOAD_ROOT)]
    units = []
    for rp in root_paths:
        if not rp.is_dir():
            print(f'目录不存在: {rp}')
            continue
        # 复用 summary 的单元划分
        try:
            from cnki_summary import iter_export_units
            units.extend(iter_export_units(rp))
        except Exception:
            units.append(rp)
    print(f'补全单元 {len(units)} 个')
    if dry_run:
        for u in units:
            labels = read_export_failure_labels(str(u))
            info = folder_to_fill_task(
                u, doc_type=doc_type, patent_date=patent_date, theme=theme)
            kind = doc_type_label(info.get('doc_type'))
            theme_part = f' | 主题: {info["theme"]}' if info.get('theme') else ''
            date_part = (
                f' | {patent_date_field_label(info.get("patent_date"))}'
                if normalize_doc_type(info.get('doc_type')) == 'patent' else '')
            print(
                f'{u}: {len(labels)} 个失败时段 {labels[:5]}'
                f' [{kind}{theme_part}{date_part}]')
        return True
    preload_click_captcha()
    driver = None
    try:
        for u in units:
            info = folder_to_fill_task(
                u, doc_type=doc_type, patent_date=patent_date, theme=theme)
            labels = read_export_failure_labels(str(u))
            if limit and limit > 0:
                labels = labels[:limit]
            if not labels:
                continue
            download_dir = str(u)
            os.makedirs(download_dir, exist_ok=True)
            fill_doc_type = normalize_doc_type(info.get('doc_type'))
            fill_theme = (info.get('theme') or '').strip()
            fill_patent_date = normalize_patent_date_field(
                info.get('patent_date'))
            fill_date_part = (
                f' | {patent_date_field_label(fill_patent_date)}'
                if fill_doc_type == 'patent' else '')
            fill_theme_part = f' | 主题: {fill_theme}' if fill_theme else ''
            print(
                f'\n==== 补全 {u} 共 {len(labels)} 个时段'
                f'（{doc_type_label(fill_doc_type)}{fill_theme_part}'
                f'{fill_date_part}）====')
            if fill_doc_type == 'patent' and fill_patent_date != 'OD':
                print(
                    '提示：专利题录只含公开日期列，按申请日补全时'
                    '不做本地条数核对，直接重导失败时段')
            if fill_doc_type == 'patent' and not patent_date and not read_export_meta(download_dir):
                print(
                    '提示：该目录没有 export_meta.json（旧版导出），'
                    '专利时间字段按默认申请日；若原导出用公开日请加 --patent-date OD')
            fill_session_ready = False
            if not is_driver_alive(driver):
                driver = create_chrome_driver(download_dir)
                open_cnki_url(driver, get_adv_search_url(fill_doc_type))
                try:
                    driver.maximize_window()
                except Exception:
                    pass
                login_ok = ensure_unit_login_if_needed(
                    driver, restore_handle=driver.current_window_handle)
                cap0 = ensure_cnki_captcha_after_nav(
                    driver, download_dir, driver.current_window_handle)
                still_login = page_suggests_login_lost(
                    driver, check_handle=driver.current_window_handle)
                if captcha_needs_browser_restart(cap0) or not cap0:
                    print('补全启动时验证码未通过，将在时段重试中重建浏览器')
                    fill_session_ready = False
                elif still_login or not login_ok:
                    print('补全启动时仍停在 WebVPN/CAS 登录页，'
                          '将在时段重试中重建浏览器并重新登录')
                    fill_session_ready = False
                else:
                    fill_session_ready = True
                    print('补全：已通过验证码，后续时段复用会话')
            else:
                set_chrome_download_dir(driver, download_dir)
                fill_session_ready = True
            for label in labels:
                try:
                    fill_result = run_fill_one_label(
                        driver, download_dir, info['category'], info['lang'],
                        label, doc_type=fill_doc_type, theme=fill_theme,
                        patent_date=fill_patent_date,
                        reuse_session=fill_session_ready)
                    if fill_result == 'restart_browser':
                        print(
                            f'{label} 补全需关浏览器'
                            f'（频繁、验证码或登录中断），重建后重试...')
                        driver = recreate_browser_after_rate_limit(
                            driver, download_dir, doc_type=fill_doc_type)
                        fill_session_ready = not page_suggests_login_lost(driver)
                        if not fill_session_ready:
                            print(f'{label} 重建后仍停在登录页，'
                                  '本时段将重新走登录流程')
                        fill_result = run_fill_one_label(
                            driver, download_dir, info['category'],
                            info['lang'], label,
                            doc_type=fill_doc_type, theme=fill_theme,
                            patent_date=fill_patent_date,
                            reuse_session=True)
                        if fill_result == 'restart_browser':
                            print(
                                f'{label} 重建后仍需重启，保留失败记录')
                            record_export_failure(download_dir, label)
                        elif fill_result:
                            fill_session_ready = True
                    elif fill_result:
                        fill_session_ready = True
                except Exception as e:
                    print(f'{label} 未捕获异常，保留记录: {e}')
                    record_export_failure(download_dir, label)
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
            cleanup_chrome_profile()
    print('补全模式结束')
    return True


if __name__ == "__main__":
    import argparse
    _parser = argparse.ArgumentParser(
        description='CNKI 期刊/专利题录导出（含补全、可选主题）')
    _parser.add_argument('--fill', action='store_true', help='按失败日期补全')
    _parser.add_argument('--root', action='append', default=None, help='只补指定大类目录')
    _parser.add_argument('--limit', type=int, default=0, help='每个目录最多补N个时段，0=全部')
    _parser.add_argument('--dry-run', action='store_true', help='只列队不检索')
    _parser.add_argument(
        '--doc-type', choices=('journal', 'patent', 'none'), default=None,
        help='默认导出类型（journal/patent/none=不选跳过资源类型筛选；'
             'EXPORT_TASKS 为空且交互未选时作为默认；'
             '补全模式作为目录无法识别时的兜底）')
    _parser.add_argument(
        '--lang', choices=('zh', 'en', 'none'), default=None,
        help='默认下载类型（zh/en/none=不选跳过中外文筛选；'
             '交互配置时作为默认；专利任务仍固定中文设置）')
    _parser.add_argument(
        '--patent-date', choices=('AD', 'OD'), default=None,
        help='补全模式专利时间字段兜底：AD=申请日（默认），OD=公开日')
    _parser.add_argument(
        '--theme', default=None,
        help='可选默认主题词（写入全局提示默认；任务仍可单独覆盖）')
    _args, _ = _parser.parse_known_args()
    if _args.doc_type:
        DOWNLOAD_DOC_TYPE = _args.doc_type
    if _args.lang:
        DOWNLOAD_LANG = _args.lang
    if _args.theme is not None:
        DEFAULT_THEME = _args.theme.strip()
    if _args.fill:
        run_fill_mode(
            roots=_args.root, limit=_args.limit, dry_run=_args.dry_run,
            doc_type=_args.doc_type, patent_date=_args.patent_date,
            theme=_args.theme)
    else:
        cnkipaperfordatav4()
