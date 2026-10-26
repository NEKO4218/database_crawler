#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sci-Hub 全文下载工具 v5 (带路径输出与汇总)
- 支持 .xls / .xlsx / .xlsm
- 支持按类别、期刊子文件夹筛选（--journal / --folder / --list-journals）
- 自动过滤 masthead/editorial 等非论文条目
- 下载成功立即打印完整绝对路径
- 运行结束打印汇总报告
"""

import os
import sys
import re
import time
import argparse
from pathlib import Path
from urllib.parse import urljoin

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("[错误] 缺少依赖，请运行: pip install requests beautifulsoup4")
    sys.exit(1)

# 尝试导入 Excel 读取库
try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import xlrd
except ImportError:
    xlrd = None

# ================= 配置区 =================
DEFAULT_INPUT_ROOT = r"D:\2026年工作\SCI数据库下载\九类期刊题录"
DEFAULT_OUTPUT_ROOT = r"D:\2026年工作\SCI数据库下载\九类期刊全文"

MIRRORS = [
    "https://sci-hub.ren/",
    "https://sci-hub.ee/",
    "https://sci-hub.vg/",
    "https://sci-hub.ru/",
    "https://sci-hub.shop/",
]

# 非论文 DOI 关键词过滤
SKIP_KEYWORDS = ["masthead", "editorial", "publication-info", "cover", "toc"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# ================= 核心逻辑 =================

def read_dois_from_excel(filepath):
    """从 Excel 文件中读取 DOI 列表"""
    dois = []
    suffix = Path(filepath).suffix.lower()
    
    if suffix == ".xls":
        if not xlrd:
            print(f"  [警告] 缺少 xlrd，无法读取 .xls 文件: {filepath}")
            return []
        try:
            wb = xlrd.open_workbook(filepath)
            sheet = wb.sheet_by_index(0)
            header = [str(cell.value).strip().lower() for cell in sheet.row(0)]
            doi_col = next((i for i, h in enumerate(header) if "doi" in h), None)
            if doi_col is None: return []
            for row_idx in range(1, sheet.nrows):
                val = str(sheet.cell_value(row_idx, doi_col)).strip()
                if val and val.lower() != "doi": dois.append(val)
        except Exception as e:
            print(f"  [错误] 读取 xls 失败: {e}")
            
    elif suffix in [".xlsx", ".xlsm"]:
        if not openpyxl:
            print(f"  [警告] 缺少 openpyxl，无法读取 .xlsx 文件: {filepath}")
            return []
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows: return []
            header = [str(c).strip().lower() if c else "" for c in rows[0]]
            doi_col = next((i for i, h in enumerate(header) if "doi" in h), None)
            if doi_col is None: return []
            for row in rows[1:]:
                if doi_col < len(row):
                    val = str(row[doi_col]).strip()
                    if val and val.lower() != "doi": dois.append(val)
            wb.close()
        except Exception as e:
            print(f"  [错误] 读取 xlsx 失败: {e}")
    return dois

def get_pdf_url(session, doi, mirror, timeout):
    """从 Sci-Hub 页面提取 PDF 链接"""
    url = urljoin(mirror, doi)
    try:
        resp = session.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        if resp.status_code != 200: return None, f"HTTP {resp.status_code}"
        
        # 检查是否为验证页
        if "cf-turnstile" in resp.text or "check-your-browser" in resp.text:
            return None, "Cloudflare验证页"
            
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # 1. 尝试 embed 标签
        embed = soup.find("embed", type="application/pdf")
        if embed and embed.get("src"):
            return embed["src"], "embed"
            
        # 2. 尝试 iframe
        iframe = soup.find("iframe", src=re.compile(r"\.pdf"))
        if iframe and iframe.get("src"):
            return iframe["src"], "iframe"
            
        # 3. 尝试按钮
        btn = soup.find("button", onclick=re.compile(r"location\.href.*pdf"))
        if btn:
            match = re.search(r"location\.href='([^']+)'", btn.get("onclick", ""))
            if match: return urljoin(mirror, match.group(1)), "button"
            
        return None, "未找到PDF链接"
    except Exception as e:
        return None, str(e)

def download_pdf(session, pdf_url, save_path, timeout):
    """下载 PDF 文件"""
    try:
        if pdf_url.startswith("//"): pdf_url = "https:" + pdf_url
        resp = session.get(pdf_url, headers=HEADERS, timeout=timeout, stream=True)
        if resp.status_code == 200 and "pdf" in resp.headers.get("Content-Type", ""):
            with open(save_path, "wb") as f:
                for chunk in resp.iter_content(8192): f.write(chunk)
            return True, os.path.getsize(save_path)
        return False, f"下载失败: {resp.status_code}"
    except Exception as e:
        return False, str(e)

def is_valid_paper_doi(doi):
    """检查是否为有效论文 DOI (过滤 masthead 等)"""
    lower_doi = doi.lower()
    return not any(kw in lower_doi for kw in SKIP_KEYWORDS)


def is_excel_file(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
        and not path.name.startswith(".")
    )


def list_journal_dirs(root: Path) -> list[Path]:
    """类别目录下的期刊子文件夹（含 Excel 的文件夹）。"""
    journals = []
    for item in sorted(root.iterdir(), key=lambda p: p.name.lower()):
        if not item.is_dir() or item.name.startswith("."):
            continue
        if any(is_excel_file(p) for p in item.iterdir() if p.is_file()):
            journals.append(item)
    return journals


def parse_journal_filters(raw_items) -> list[str]:
    """把 -j 的参数收成刊名。支持：多次 -j、带空格未加引号、逗号分隔、引号包裹。"""
    filters = []
    for item in raw_items or []:
        if isinstance(item, (list, tuple)):
            text = " ".join(str(x) for x in item).strip()
        else:
            text = str(item).strip()
        if not text:
            continue
        parts = [p.strip() for p in text.split(",")] if "," in text else [text]
        filters.extend(p for p in parts if p)
    return filters


def journal_matches(journal_name: str, filters: list[str]) -> bool:
    if not filters:
        return True
    name = journal_name.casefold()
    return any(f.casefold() in name or name == f.casefold() for f in filters)


def resolve_scan_root(input_root: Path, category: str, folder: str) -> Path:
    """确定要扫描的目录：类别文件夹，或 --folder 指定的路径。"""
    cat_dir = input_root / category
    if not folder:
        return cat_dir
    specified = Path(folder)
    if specified.is_absolute():
        return specified
    under_cat = cat_dir / folder
    if under_cat.exists():
        return under_cat
    under_root = input_root / folder
    if under_root.exists():
        return under_root
    return specified


def collect_excel_files(scan_root: Path, filters: list[str]) -> tuple[list[Path], list[str]]:
    """
    收集待处理 Excel。
    - 若 scan_root 本身是某刊文件夹（内含表格），只处理该刊
    - 否则按子文件夹当作刊名筛选
    """
    excel_files: list[Path] = []
    matched_journals: list[str] = []

    direct = [p for p in scan_root.iterdir() if is_excel_file(p)] if scan_root.is_dir() else []
    if direct:
        if journal_matches(scan_root.name, filters):
            excel_files.extend(sorted(direct, key=lambda p: p.name.lower()))
            matched_journals.append(scan_root.name)
        return excel_files, matched_journals

    for journal_dir in list_journal_dirs(scan_root):
        if not journal_matches(journal_dir.name, filters):
            continue
        matched_journals.append(journal_dir.name)
        files = [p for p in journal_dir.iterdir() if is_excel_file(p)]
        excel_files.extend(sorted(files, key=lambda p: p.name.lower()))
    return excel_files, matched_journals


# ================= 主流程 =================

def main():
    parser = argparse.ArgumentParser(description="Sci-Hub 下载工具 v5")
    parser.add_argument("--category", required=True, help="类别名称，如 '集成电路'")
    parser.add_argument(
        "--journal", "-j", action="append", nargs="+", default=[],
        metavar="刊名",
        help="只处理刊名包含这些词的子文件夹。带空格可直接写在 -j 后，也可加引号；可多次指定",
    )
    parser.add_argument(
        "--folder", default="",
        help="指定要扫描的文件夹：绝对路径，或相对类别目录的期刊名/子路径",
    )
    parser.add_argument(
        "--list-journals", action="store_true",
        help="列出指定文件夹下的期刊子目录后退出",
    )
    parser.add_argument("--input-root", default=DEFAULT_INPUT_ROOT, help="题录根目录")
    parser.add_argument("--output-root", default=DEFAULT_OUTPUT_ROOT, help="全文输出根目录")
    parser.add_argument("--timeout", type=float, default=30.0, help="请求超时时间")
    parser.add_argument("--proxy", help="代理地址，如 http://127.0.0.1:7890")
    parser.add_argument("--dry-run", action="store_true", help="仅扫描不下载")
    args = parser.parse_args()

    session = requests.Session()
    if args.proxy: session.proxies = {"http": args.proxy, "https": args.proxy}

    input_root = Path(args.input_root)
    scan_root = resolve_scan_root(input_root, args.category, args.folder)
    if not scan_root.exists():
        print(f"[错误] 目录不存在: {scan_root}")
        return

    filters = parse_journal_filters(args.journal)
    all_journals = [p.name for p in list_journal_dirs(scan_root)]
    if not all_journals and any(is_excel_file(p) for p in scan_root.iterdir() if p.is_file()):
        all_journals = [scan_root.name]

    if args.list_journals:
        print(f"[期刊] {scan_root} 共 {len(all_journals)} 种")
        for i, name in enumerate(all_journals, 1):
            mark = " *" if journal_matches(name, filters) and filters else ""
            print(f"  {i}. {name}{mark}")
        if filters:
            print("带 * 的为当前 --journal 能匹配到的刊")
        return

    excel_files, matched = collect_excel_files(scan_root, filters)
    print(f"[扫描] 目录: {scan_root}")
    if filters:
        print(f"[筛选] 期刊关键词: {', '.join(filters)}")
        print(f"[筛选] 匹配到 {len(matched)} 种刊: {', '.join(matched) if matched else '无'}")
        unmatched = [f for f in filters if not any(journal_matches(n, [f]) for n in all_journals)]
        if unmatched:
            print(f"[提示] 未匹配: {', '.join(unmatched)}")
            print("可用 --list-journals 查看全部刊名")
    print(f"[扫描] 找到 {len(excel_files)} 个 Excel 文件")
    if not excel_files:
        return

    stats = {"success": 0, "fail": 0, "skip": 0}
    downloaded_files = []

    for ef in excel_files:
        journal_name = ef.parent.name
        try:
            rel = ef.relative_to(input_root)
        except ValueError:
            rel = Path(args.category) / journal_name / ef.name
        print(f"\n[处理] {rel}")
        dois = read_dois_from_excel(ef)
        journal_name = ef.parent.name
        out_dir = Path(args.output_root) / args.category / journal_name
        out_dir.mkdir(parents=True, exist_ok=True)

        for doi in dois:
            if not is_valid_paper_doi(doi):
                print(f"  [跳过] 非论文条目: {doi}")
                stats["skip"] += 1
                continue

            safe_name = re.sub(r'[\\/*?:"<>|]', "_", doi.replace("/", "_")) + ".pdf"
            save_path = out_dir / safe_name

            if save_path.exists():
                print(f"  [已存在] {save_path.name}")
                stats["success"] += 1
                downloaded_files.append(str(save_path.resolve()))
                continue

            if args.dry_run:
                print(f"  [预览] {doi}")
                continue

            # 尝试下载
            found = False
            for mirror in MIRRORS:
                pdf_url, msg = get_pdf_url(session, doi, mirror, args.timeout)
                if pdf_url:
                    ok, info = download_pdf(session, pdf_url, save_path, args.timeout)
                    if ok:
                        size_mb = info / 1024 / 1024
                        print(f"  [下载完成] {save_path.name} ({size_mb:.1f} MB)")
                        print(f"    -> {save_path.resolve()}")  # 关键：打印完整路径
                        stats["success"] += 1
                        downloaded_files.append(str(save_path.resolve()))
                        found = True
                        break
                else:
                    print(f"  [尝试失败] {mirror} -> {msg}")
            
            if not found:
                print(f"  [全部失败] {doi}")
                stats["fail"] += 1
            time.sleep(1)  # 礼貌延迟

    # 汇总报告
    print("\n" + "="*60)
    print("下载汇总")
    print("="*60)
    print(f"  成功下载: {stats['success']} 篇")
    print(f"  失败: {stats['fail']} 篇")
    print(f"  跳过(非论文): {stats['skip']} 篇")
    
    if downloaded_files:
        print(f"\n已下载文件列表 ({len(downloaded_files)} 个):")
        for i, p in enumerate(downloaded_files, 1):
            print(f"  [{i}] {p}")
    print("="*60)

if __name__ == "__main__":
    main()