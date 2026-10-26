# -*- coding: utf-8 -*-
"""按期刊来源（Publication/Source Titles）检索并导出 WoS 题录。

复用 sciforexcel.py 的登录、合集勾选、导出 Excel 流程。
检索字段改为 Publication/Source Titles，刊名来自 Q1九类期刊.xlsx，
结果按九类分目录存放。

示例：
  python sci_journals.py --dry-run
  python sci_journals.py --domain 人工智能
  python sci_journals.py --limit 1
  python sci_journals.py --journal CELL
  python sci_journals.py --no-resume
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import sciforexcel as sci

JOURNALS_XLSX = Path(r"D:\AI科技类内容图文匹配\题录\Q1九类期刊.xlsx")
DOWNLOAD_ROOT = os.path.join(sci.DOWNLOAD_ROOT, "九类期刊题录")

DOMAINS = [
    "集成电路",
    "高端科学仪器",
    "智能制造",
    "高端机器人",
    "人工智能",
    "智能网联新能源车",
    "商业航天",
    "生物医药",
    "生物育种",
]

SOURCE_FIELD_ALIASES = (
    "Publication/Source Titles",
    "Publication Titles",
    "Source Titles",
    "Source Title",
    "出版物名称",
    "来源出版物",
    "出版物/来源出版物",
)


def load_domain_journals(
        path: Path, only_domain: str = "", only_journal: str = ""
        ) -> list[tuple[str, str]]:
    """读取九类期刊表，返回 [(领域, 刊名), ...]。"""
    xl = pd.ExcelFile(path)
    rows: list[tuple[str, str]] = []
    wanted = {only_domain} if only_domain else set(DOMAINS)
    needle = only_journal.strip().casefold()
    for domain in DOMAINS:
        if domain not in wanted:
            continue
        if domain not in xl.sheet_names:
            print(f"跳过缺失工作表: {domain}")
            continue
        df = pd.read_excel(xl, sheet_name=domain)
        if "Journal" not in df.columns:
            print(f"{domain} 无 Journal 列，跳过")
            continue
        seen = set()
        for name in df["Journal"].astype(str):
            journal = " ".join(name.split()).strip()
            if not journal or journal.lower() == "nan":
                continue
            if needle and needle not in journal.casefold():
                continue
            key = journal.casefold()
            if key in seen:
                continue
            seen.add(key)
            rows.append((domain, journal))
    return rows


def journal_query(name: str) -> str:
    text = str(name).strip()
    if text.startswith('"') and text.endswith('"'):
        return text
    return f'"{text}"'


def journal_dir(out_root: str, domain: str, journal: str) -> str:
    path = os.path.join(out_root, sci._safe_dirname(domain), sci._safe_dirname(journal))
    os.makedirs(path, exist_ok=True)
    return path


def label_prefix(journal: str, database, collections) -> str:
    db_path = sci.resolve_db_path(database, collections)
    return (
        f"SCI_{sci._safe_dirname(journal)}_"
        f"{sci._safe_dirname(db_path[0])}_"
        f"{sci._safe_dirname(db_path[1] if len(db_path) > 1 else collections)}"
    )


def _source_title_field_active(driver) -> bool:
    """当前检索字段是否已是 Publication/Source Titles。"""
    if sci._date_range_inputs_present(driver):
        return False
    return bool(driver.execute_script(
        """
        const aliases = arguments[0].map(s => s.toLowerCase());
        const btns = Array.from(document.querySelectorAll(
          'button[aria-label*="Select search field"], button[aria-haspopup="dialog"]'));
        for (const b of btns) {
          const r = b.getBoundingClientRect();
          if (r.width < 2 || r.height < 2) continue;
          const ta = (b.getAttribute('data-ta') || '').trim();
          const aria = (b.getAttribute('aria-label') || '');
          const lowAria = aria.toLowerCase();
          if (lowAria.includes('database') || lowAria.includes('edition')) continue;
          const t = (b.innerText || '').trim();
          const blob = (ta + ' ' + aria + ' ' + t).toLowerCase();
          if (ta === 'Publication/Source Titles' || ta === 'SO') return true;
          if (aliases.some(a => blob.includes(a.toLowerCase()))) return true;
        }
        return false;
        """,
        list(SOURCE_FIELD_ALIASES)))


def _open_source_field_dropdown(driver) -> bool:
    """打开检索字段下拉；当前字段已是 SO 时也能点开。"""
    if sci._open_search_field_dropdown(driver):
        return True
    sci._dismiss_open_overlays(driver)
    for sel in (
        'button[aria-label*="Select search field"]',
        'button[aria-haspopup="dialog"][data-ta="Publication/Source Titles"]',
        'button[aria-haspopup="dialog"][data-ta="SO"]',
        'button[aria-haspopup="dialog"][data-ta="Topic"]',
        'button[aria-haspopup="dialog"][data-ta="Publication Date"]',
    ):
        els = driver.find_elements(By.CSS_SELECTOR, sel)
        for el in els:
            try:
                if not el.is_displayed():
                    continue
                aria = (el.get_attribute("aria-label") or "").lower()
                if "database" in aria or "edition" in aria:
                    continue
                if sci.safe_click(driver, el):
                    time.sleep(0.8)
                    return True
            except Exception:
                continue
    return bool(driver.execute_script(
        """
        const aliases = arguments[0];
        const btns = Array.from(document.querySelectorAll(
          'button[role="combobox"], button.dropdown, button[aria-haspopup="dialog"]'));
        for (const b of btns) {
          const aria = (b.getAttribute('aria-label') || '').toLowerCase();
          const ta = (b.getAttribute('data-ta') || '');
          const t = (b.innerText || '').trim();
          if (aria.includes('database') || aria.includes('edition')) continue;
          if (aria.includes('select search field') || aria.includes('检索字段')
              || t === '主题' || t === 'Topic' || t === '出版日期'
              || t === 'Publication Date' || t === 'Publication/Source Titles'
              || aliases.includes(t) || aliases.includes(ta)) {
            const r = b.getBoundingClientRect();
            if (r.width < 2 || r.height < 2) continue;
            b.click();
            return true;
          }
          if (ta === 'Topic' || ta === 'Publication Date' || ta === 'Author'
              || ta === 'Title' || ta === 'Publication/Source Titles' || ta === 'SO') {
            const r = b.getBoundingClientRect();
            if (r.width < 2 || r.height < 2) continue;
            b.click();
            return true;
          }
        }
        return false;
        """,
        list(SOURCE_FIELD_ALIASES)))


def _click_source_title_option(driver) -> bool:
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
        const aliases = arguments[0];
        const panes = Array.from(document.querySelectorAll(
          '.cdk-overlay-pane, [role="dialog"], [role="listbox"]'));
        const roots = panes.length ? panes : [document.body];
        const selectors = [
          '#option-SO',
          'button[data-ta="Publication/Source Titles"]',
          '[data-ta="Publication/Source Titles"]',
          '[role="option"]',
          '[role="menuitem"]',
          'button', 'li', 'div', 'span', 'a'
        ];
        for (const root of roots) {
          for (const sel of selectors) {
            let nodes;
            try { nodes = Array.from(root.querySelectorAll(sel)); }
            catch (e) { continue; }
            for (const n of nodes) {
              if (!visible(n) || isTrigger(n)) continue;
              const id = n.id || '';
              const ta = (n.getAttribute('data-ta') || '').trim();
              const label = (n.getAttribute('aria-label') || '').trim();
              const t = (n.innerText || '').replace(/\\s+/g, ' ').trim();
              if (id === 'option-SO' || ta === 'Publication/Source Titles'
                  || ta === 'SO' || label === 'Publication/Source Titles') {
                n.click();
                return t || ta || id;
              }
              if (aliases.some(a => t === a || label === a || ta === a)) {
                n.click();
                return t || ta;
              }
            }
          }
        }
        return null;
        """,
        list(SOURCE_FIELD_ALIASES))
    if clicked:
        print(f"已点击筛选项: {clicked}")
        time.sleep(0.6)
        return True
    return False


def select_source_title_field(driver, timeout=18) -> bool:
    """打开检索字段下拉并选择 Publication/Source Titles。"""
    if _source_title_field_active(driver):
        print("检索字段已是 Publication/Source Titles")
        return True
    if sci.is_page_unusable(driver):
        print("选择检索字段时页面已失效")
        return False
    if not sci.is_search_form_ready(driver):
        if not sci.wait_search_form_ready(driver, timeout=8):
            print("检索表单未就绪，无法选择 Publication/Source Titles")
            return False

    miss = 0
    deadline = time.time() + timeout
    while time.time() < deadline:
        if sci.is_page_unusable(driver):
            return False
        try:
            url = (driver.current_url or "").lower()
        except Exception:
            url = ""
        if "basic-search" not in url and "advanced-search" not in url:
            print("当前不是检索页，停止点检索字段")
            return False
        if not sci.is_search_form_ready(driver) and not _source_title_field_active(driver):
            miss += 1
            if miss >= 4:
                print("检索表单未出现，停止空点检索字段")
                return False
            print("检索表单未就绪，等待...")
            sci.wait_search_form_ready(driver, timeout=3)
            continue
        if not _open_source_field_dropdown(driver):
            miss += 1
            print(f"未找到检索字段下拉，重试 ({miss}/6)...")
            if miss >= 6:
                break
            time.sleep(0.6)
            continue
        miss = 0
        if not _click_source_title_option(driver):
            print("下拉已开但未点到 Publication/Source Titles，重试...")
            sci._dismiss_open_overlays(driver)
            time.sleep(0.4)
            continue
        wait_deadline = time.time() + 6
        while time.time() < wait_deadline:
            if _source_title_field_active(driver):
                print("已选择检索字段: Publication/Source Titles")
                time.sleep(0.3)
                return True
            time.sleep(0.25)
        print("已点击 Publication/Source Titles，但字段尚未生效，重试...")
        sci._dismiss_open_overlays(driver)
        time.sleep(0.4)
    print("未能切换到 Publication/Source Titles")
    return False


def fill_source_title(driver, journal: str) -> bool:
    """在检索框填入刊名（加引号精确匹配）。"""
    query = journal_query(journal)
    els = driver.find_elements(By.CSS_SELECTOR, 'input[name="search-main-box"]')
    if not els:
        els = driver.find_elements(By.CSS_SELECTOR, "#search-option-0")
    if not els:
        els = driver.find_elements(
            By.CSS_SELECTOR, 'input[data-ta="search-criteria-input"]')
    visible = []
    for el in els:
        try:
            if el.is_displayed():
                visible.append(el)
        except Exception:
            continue
    if not visible:
        print("找不到刊名检索框")
        return False
    sci.fill_input(driver, visible[0], query)
    try:
        visible[0].send_keys(Keys.ESCAPE)
    except Exception:
        pass
    print(f"已输入刊名: {query}")
    time.sleep(0.4)
    return True


def prepare_source_search_form(driver, database, collections) -> bool:
    target = sci.search_url_for(database, collections)
    last_err = None
    for attempt in range(1, sci.PAGE_LOAD_RETRIES + 1):
        if (sci.is_page_unusable(driver)
                or not sci.is_on_target_search_page(driver, target)
                or not sci.is_search_form_ready(driver)):
            print(f"检索表单未就绪，重新打开 ({attempt}/{sci.PAGE_LOAD_RETRIES})")
            if not sci.goto_search_url(driver, target):
                last_err = RuntimeError("检索页恢复失败")
                sci.recover_wos_page(driver, target_url=target)
                time.sleep(1.5)
                continue
        if not sci.ensure_database_and_collections(driver, database, collections):
            last_err = RuntimeError("未能打开目标检索库")
            sci.recover_wos_page(driver, target_url=target)
            continue
        if sci.is_page_unusable(driver):
            last_err = RuntimeError("打开检索库后页面失效")
            continue
        sci._dismiss_open_overlays(driver)
        time.sleep(0.5)
        if select_source_title_field(driver):
            return True
        last_err = RuntimeError("无法选择 Publication/Source Titles")
        print(f"未能选择刊名字段，重开检索页 ({attempt}/{sci.PAGE_LOAD_RETRIES})")
        sci.recover_wos_page(driver, target_url=target)
        time.sleep(1)
    raise last_err or RuntimeError("无法选择 Publication/Source Titles")


def run_journal_search_once(driver, journal, database, collections, download_dir=None):
    """按刊名检索一次，返回命中数；失败返回 None。"""
    if download_dir:
        sci.bind_browser_ctx(driver, download_dir, database, collections)
        sci.set_chrome_download_dir(driver, download_dir)
    else:
        sci.bind_browser_ctx(driver, database=database, collections=collections)

    target = sci.search_url_for(database, collections)
    last_err = None
    server_round = 0
    timeout_round = 0

    while True:
        driver = sci.ctx_driver(driver)
        if not sci.is_driver_alive(driver):
            print("检索前浏览器不可用")
            last_err = "driver_dead"
            break

        hit_server_error = False
        hit_search_timeout = False
        form_ok = False

        for attempt in range(1, sci.PAGE_LOAD_RETRIES + 1):
            driver = sci.ctx_driver(driver)
            if not sci.goto_search_url(driver, target):
                print(f"无法打开检索页 ({attempt}/{sci.PAGE_LOAD_RETRIES})")
                sci.recover_wos_page(driver, target_url=target)
                last_err = "open_search_page"
                continue
            try:
                prepare_source_search_form(driver, database, collections)
                if not fill_source_title(driver, journal):
                    raise RuntimeError("未能填写刊名")
                sci.click_run_search(driver)
                sci.accept_cookies(driver, timeout=3)
                form_ok = True

                status = sci.wait_search_result_or_server_error(
                    driver, timeout=sci.SEARCH_RESULT_WAIT)
                if status == "no_results":
                    print(f"检索无结果，跳过: {journal}")
                    sci.bind_browser_ctx(driver)
                    return 0
                if status == "server_error":
                    hit_server_error = True
                    last_err = "server_unexpected_error"
                    break
                if status == "timeout":
                    if sci.has_wos_server_unexpected_error(driver):
                        hit_server_error = True
                        last_err = "server_unexpected_error"
                        break
                    hit_search_timeout = True
                    last_err = "no_results_page_timeout"
                    print("检索后仍未进入结果页，将关闭浏览器重试")
                    break
                if status == "unusable":
                    print("检索后页面失效，恢复后重试")
                    sci.recover_wos_page(driver, target_url=target)
                    last_err = "page_after_search"
                    continue

                hits = sci.parse_hit_count(driver, timeout=sci.HIT_COUNT_TIMEOUT)
                if hits is None:
                    if sci.has_wos_server_unexpected_error(driver):
                        hit_server_error = True
                        last_err = "server_unexpected_error"
                        break
                    if sci.is_page_unusable(driver):
                        sci.recover_wos_page(driver, target_url=target)
                        last_err = "hit_count_page"
                        continue
                    print(f"解析命中条数失败 ({attempt}/{sci.PAGE_LOAD_RETRIES})")
                    last_err = "hit_count"
                    time.sleep(2)
                    continue
                if hits > 0:
                    sci.set_sort_date_ascending(driver)
                sci.bind_browser_ctx(driver)
                return hits
            except Exception as e:
                last_err = e
                print(f"检索异常 ({attempt}/{sci.PAGE_LOAD_RETRIES}):", e)
                if sci.has_wos_server_unexpected_error(driver):
                    hit_server_error = True
                    last_err = "server_unexpected_error"
                    break
                if sci.is_page_unusable(driver):
                    sci.recover_wos_page(driver, target_url=target)
                time.sleep(2)

            if hit_server_error or hit_search_timeout:
                break

        if hit_server_error:
            server_round += 1
            if server_round > sci.SERVER_UNEXPECTED_ERROR_RETRIES:
                print("Server.unexpectedError 多次失败，标记本刊失败")
                break
            print(f"Server.unexpectedError：第 {server_round} 次冷却重试")
            driver = sci.restart_browser_after_server_error(
                driver, reason=f"刊名检索 Server.unexpectedError（{journal}）")
            if not sci.is_driver_alive(driver) or not sci.is_wos_logged_in(driver):
                last_err = "restart_after_server_error"
                break
            continue

        if hit_search_timeout:
            timeout_round += 1
            if timeout_round > sci.SEARCH_TIMEOUT_RETRIES:
                print("检索超时多次失败，标记本刊失败")
                break
            print(f"检索超时：第 {timeout_round} 次关闭浏览器并重试")
            driver = sci.restart_browser_after_search_timeout(
                driver, reason=f"刊名检索超时（{journal}）")
            if not sci.is_driver_alive(driver) or not sci.is_wos_logged_in(driver):
                last_err = "restart_after_search_timeout"
                break
            continue

        if not form_ok:
            print("检索表单准备多次失败，终止本刊")
        break

    print("刊名检索最终失败:", last_err)
    return None


def export_journal(driver, domain, journal, database, collections, out_root, resume=True):
    """检索一种期刊并按批次导出到 领域/刊名 目录。"""
    save_dir = journal_dir(out_root, domain, journal)
    done_flag = os.path.join(save_dir, "_done.txt")
    if resume and os.path.isfile(done_flag):
        print(f"已完成，跳过: {domain} / {journal}")
        return driver, "skipped"

    sci.set_chrome_download_dir(driver, save_dir)
    sci.bind_browser_ctx(driver, save_dir, database, collections)

    def restore():
        hits_again = run_journal_search_once(
            sci.ctx_driver(driver), journal, database, collections, save_dir)
        return hits_again is not None

    hits = run_journal_search_once(driver, journal, database, collections, save_dir)
    driver = sci.ctx_driver(driver)
    if hits is None:
        sci.record_export_failure(save_dir, sci._safe_dirname(journal))
        return driver, "failed"
    if hits == 0:
        Path(done_flag).write_text("0\n", encoding="utf-8")
        print(f"{journal}: 0 条，跳过")
        return driver, "empty"

    export_hits = hits
    if hits > sci.WOS_MAX_HITS:
        print(f"{journal}: 命中 {hits} 超过 {sci.WOS_MAX_HITS}，只导出前 {sci.WOS_MAX_HITS} 条")
        export_hits = sci.WOS_MAX_HITS

    prefix = label_prefix(journal, database, collections)
    print(f"{domain} / {journal}: {hits} 条，导出到 {save_dir}")
    fail_path = os.path.join(save_dir, sci.EXPORT_FAILURE_LOG)
    fail_mtime = os.path.getmtime(fail_path) if os.path.isfile(fail_path) else 0
    sci.export_all_hits(
        driver, save_dir, export_hits, prefix,
        restore_fn=restore, expected_hits=hits)
    new_fail = (
        os.path.isfile(fail_path)
        and os.path.getmtime(fail_path) > fail_mtime
    )
    if new_fail:
        print(f"{journal}: 有批次失败，不标记完成")
        return sci.ctx_driver(driver), "failed"
    Path(done_flag).write_text(f"{hits}\n", encoding="utf-8")
    return sci.ctx_driver(driver), "ok"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按九类期刊名检索并导出 WoS 题录")
    parser.add_argument("--journals", type=Path, default=JOURNALS_XLSX)
    parser.add_argument("--out", type=str, default=DOWNLOAD_ROOT)
    parser.add_argument("--domain", type=str, default="", help="只跑某一个领域")
    parser.add_argument("--journal", type=str, default="", help="刊名包含该字符串才检索")
    parser.add_argument("--limit", type=int, default=0, help="最多处理几种刊（调试）")
    parser.add_argument(
        "--resume", dest="resume", action="store_true", default=True,
        help="跳过已有 _done.txt 的刊（默认开启）")
    parser.add_argument(
        "--no-resume", dest="resume", action="store_false",
        help="不跳过已完成的刊，重新检索导出")
    parser.add_argument("--dry-run", action="store_true", help="只列出将检索的期刊")
    parser.add_argument(
        "--restart-every", type=int, default=8,
        help="每处理多少种刊重启一次浏览器")
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()
    jobs = load_domain_journals(
        args.journals, args.domain.strip(), args.journal.strip())
    if args.limit and args.limit > 0:
        jobs = jobs[: args.limit]
    if not jobs:
        print("没有可检索的期刊")
        return 1

    print(f"期刊表: {args.journals}")
    print(f"保存目录: {args.out}")
    already_done = 0
    if args.resume:
        already_done = sum(
            1 for domain, journal in jobs
            if os.path.isfile(os.path.join(
                journal_dir(args.out, domain, journal), "_done.txt"))
        )
    print(f"待检索: {len(jobs)} 种")
    if already_done:
        print(f"其中已完成将跳过: {already_done} 种，实际待处理: {len(jobs) - already_done} 种")
    by_domain: dict[str, int] = {}
    for domain, _journal in jobs:
        by_domain[domain] = by_domain.get(domain, 0) + 1
    for domain in DOMAINS:
        if domain in by_domain:
            print(f"  {domain}: {by_domain[domain]}")
    if args.dry_run:
        for i, (domain, journal) in enumerate(jobs, 1):
            print(f"  {i}. [{domain}] {journal}")
        return 0

    database = sci.SEARCH_DATABASE
    collections = sci.SEARCH_COLLECTIONS
    os.makedirs(args.out, exist_ok=True)

    first_pending = None
    if args.resume:
        for domain, journal in jobs:
            if not os.path.isfile(os.path.join(
                    journal_dir(args.out, domain, journal), "_done.txt")):
                first_pending = (domain, journal)
                break
        if first_pending is None:
            print("全部期刊均已完成，无需启动浏览器")
            return 0
    else:
        first_pending = jobs[0]

    first_dir = journal_dir(args.out, first_pending[0], first_pending[1])
    driver = sci.start_browser_session(first_dir, database, collections)
    sci.bind_browser_ctx(driver, first_dir, database, collections)
    if not sci.is_driver_alive(driver) or not sci.is_wos_logged_in(driver):
        print("首次登录失败")
        sci.quit_chrome_driver(driver)
        return 1

    stats = {"ok": 0, "empty": 0, "failed": 0, "skipped": 0}
    processed = 0
    try:
        for i, (domain, journal) in enumerate(jobs, 1):
            save_dir = journal_dir(args.out, domain, journal)
            done_flag = os.path.join(save_dir, "_done.txt")
            if args.resume and os.path.isfile(done_flag):
                print(f"\n######## {i}/{len(jobs)} [{domain}] {journal} ########")
                print(f"已完成，跳过: {domain} / {journal}")
                stats["skipped"] += 1
                continue

            print(f"\n######## {i}/{len(jobs)} [{domain}] {journal} ########")
            if args.restart_every and processed > 0 and processed % args.restart_every == 0:
                driver = sci.restart_browser_session(
                    driver, save_dir, database, collections,
                    reason=f"已实际处理 {processed} 种刊，重启浏览器")
            elif not sci.is_driver_alive(driver) or not sci.is_wos_logged_in(driver):
                driver = sci.restart_browser_session(
                    driver, save_dir, database, collections,
                    reason="掉线或浏览器失效，重启")
            if not sci.is_driver_alive(driver) or not sci.ensure_logged_in(driver):
                print("登录失效，记录失败并跳过")
                sci.record_export_failure(save_dir, sci._safe_dirname(journal))
                stats["failed"] += 1
                processed += 1
                continue
            try:
                driver, status = export_journal(
                    driver, domain, journal, database, collections,
                    args.out, resume=args.resume)
            except Exception as e:
                print(f"{journal} 异常:", e)
                sci.record_export_failure(save_dir, sci._safe_dirname(journal))
                status = "failed"
                driver = sci.ctx_driver(driver)
            stats[status] = stats.get(status, 0) + 1
            if status != "skipped":
                processed += 1
    finally:
        sci.quit_chrome_driver(driver)

    print("\n全部刊名检索结束:", stats)
    return 0 if stats["failed"] == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
