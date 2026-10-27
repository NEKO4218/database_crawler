# -*- coding: utf-8 -*-
"""通过 NCBI E-utilities 按日期批量获取 PubMed 题录，并导出 Excel。

不重不漏：
  - 默认按入库日期 EDAT 逐日检索。每条记录只有一个 EDAT，日期之间不重叠。
  - 每一天先用 ESearch（usehistory）取全量 PMID，并核对条数与 Count 一致后才进入下载。
  - 全日期范围再核对一次 Count；若逐日之和小于总命中，按月/年补齐未落到具体日期的记录。
  - 以 PMID 为主键去重，同一 PMID 只写入一次。

断点续跑：
  - 进度默认保存在输出目录的 pubmed_checkpoint.sqlite，可用 --checkpoint 改到其他路径。
  - 已完成的日期直接跳过；中断后只补未取到的 PMID，不重复请求已入库记录。
  - Excel 在该日记录全部落库后再导出；导出中断则只重写 Excel，不重新下载。

示例：
  python PubMed.py --start 2024-01-01 --end 2024-01-03 --out D:\\PubMed题录
  python PubMed.py --datetype pdat --start 2024-06-01 --end 2024-06-30 --out D:\\PubMed题录
  python PubMed.py --out D:\\PubMed题录 --checkpoint E:\\cache\\pubmed_checkpoint.sqlite
"""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path

import requests

try:
    from openpyxl import Workbook
except ImportError:
    print("缺少 openpyxl，请先运行: pip install openpyxl")
    raise

# ---------------------------------------------------------------------------
# 基本配置（命令行参数会覆盖这里）
# ---------------------------------------------------------------------------
_SCRIPT_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = _SCRIPT_DIR / "pubmed_credentials.json"
TOOL = "PubMed.py"


def load_credentials() -> dict:
    """从 pubmed_credentials.json 读取 API Key。"""
    defaults = {"api_key": "", "email": ""}
    if not CREDENTIALS_FILE.is_file():
        print(f"未找到凭据文件: {CREDENTIALS_FILE}，请复制 pubmed_credentials.example.json")
        return defaults
    try:
        data = json.loads(CREDENTIALS_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"读取 PubMed 凭据失败: {exc}")
        return defaults
    if not isinstance(data, dict):
        return defaults
    out = dict(defaults)
    if "api_key" in data:
        out["api_key"] = str(data["api_key"] or "").strip()
    if "email" in data:
        out["email"] = str(data["email"] or "").strip()
    return out


_CREDS = load_credentials()
API_KEY = _CREDS["api_key"]
EMAIL = _CREDS["email"]

# 检索起止日期（含首尾），格式 YYYY-MM-DD
BGDATE = "2024-01-01"
ENDDATE = "2024-01-01"
# edat=入库日期（推荐，不重不漏） | pdat=出版日期 | mdat=修订日期
DATETYPE = "edat"
# 额外检索式，会与日期条件 AND。留空表示该日期内的全部记录。
TERM = ""
OUT_DIR = r"D:\2026年工作\PubMed题录"
# 断点库文件路径。留空则使用 OUT_DIR 下的 pubmed_checkpoint.sqlite
CHECKPOINT = ""

ESEARCH_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
EFETCH_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"

# 带 api_key 时 NCBI 上限为 10 次/秒，留出余量
MIN_INTERVAL = 0.12
ESEARCH_RETMAX = 9999
UID_MIN = 1
UID_MAX = 80_000_000
EFETCH_BATCH = 200
MAX_RETRIES = 6
# NCBI 检索临时故障（HTTP 500、SOLR 不可用）时持续等待，避免整次任务退出
TRANSIENT_MAX_WAITS = 48
EXCEL_MAX_DATA_ROWS = 1_000_000
REQUEST_TIMEOUT = 120

DATE_TYPES = {
    "edat": "EDAT",
    "pdat": "PDAT",
    "mdat": "MDAT",
}

COLUMNS = [
    "PMID",
    "标题",
    "摘要",
    "作者",
    "作者单位",
    "期刊",
    "期刊缩写",
    "ISSN",
    "卷",
    "期",
    "页码",
    "出版日期",
    "电子出版日期",
    "入库日期",
    "DOI",
    "PMCID",
    "文献类型",
    "MeSH主题词",
    "关键词",
    "语种",
    "出版国",
    "NLM唯一号",
    "版权",
    "基金",
    "利益冲突",
    "PubMed链接",
]

ILLEGAL_XLSX = re.compile(r"[\000-\010\013\014\016-\037]")


def log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def parse_date(text: str) -> datetime:
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d")
    except ValueError as exc:
        raise SystemExit(f"日期格式应为 YYYY-MM-DD，收到: {text}") from exc


def daterange(start: datetime, end: datetime):
    cur = start
    while cur <= end:
        yield cur
        cur += timedelta(days=1)


def month_spans(start: datetime, end: datetime):
    cur = start.replace(day=1)
    while cur <= end:
        if cur.month == 12:
            nxt = cur.replace(year=cur.year + 1, month=1, day=1)
        else:
            nxt = cur.replace(month=cur.month + 1, day=1)
        last = min(end, nxt - timedelta(days=1))
        first = max(start, cur)
        yield first, last
        cur = nxt


def year_spans(start: datetime, end: datetime):
    year = start.year
    while year <= end.year:
        first = max(start, datetime(year, 1, 1))
        last = min(end, datetime(year, 12, 31))
        yield first, last
        year += 1


def ymd(day: datetime) -> str:
    return day.strftime("%Y/%m/%d")


def day_key(day: datetime) -> str:
    return day.strftime("%Y-%m-%d")


def brief_error(text: str) -> str:
    """把 NCBI 的 HTML 诊断页收成一行说明。"""
    raw = text or ""
    lowered = raw.lower()
    if "solr" in lowered:
        return "NCBI 检索暂时不可用（无法连接 SOLR）"
    if "temporarily unavailable" in lowered or "search backend failed" in lowered:
        return "NCBI 检索暂时不可用"
    if "<html" in lowered or "error 500" in lowered:
        return "NCBI 服务暂时不可用（HTTP 500）"
    return clean_text(raw)[:180]


def is_transient(text: str) -> bool:
    lowered = (text or "").lower()
    return any(key in lowered for key in (
        "solr",
        "temporarily unavailable",
        "search backend failed",
        "http 500",
        "http 502",
        "http 503",
        "http 504",
        "服务暂时不可用",
        "检索暂时不可用",
    ))


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    text = " ".join(value.replace("\xa0", " ").split())
    return ILLEGAL_XLSX.sub("", text)


def clip_cell(value: str) -> str:
    if len(value) <= 32767:
        return value
    return value[:32760] + "…[截断]"


def node_text(node: ET.Element | None, path: str = "") -> str:
    if node is None:
        return ""
    target = node.find(path) if path else node
    if target is None or target.text is None:
        return ""
    return clean_text(target.text)


def join_parts(parts: list[str], sep: str = "; ") -> str:
    return sep.join(p for p in parts if p)


class RateLimiter:
    def __init__(self, interval: float) -> None:
        self.interval = interval
        self._next = 0.0

    def wait(self) -> None:
        now = time.monotonic()
        if now < self._next:
            time.sleep(self._next - now)
        self._next = time.monotonic() + self.interval


class EUtils:
    def __init__(self, api_key: str, email: str, tool: str) -> None:
        self.api_key = api_key.strip()
        self.email = email.strip()
        self.tool = tool
        self.session = requests.Session()
        self.limiter = RateLimiter(MIN_INTERVAL)

    def _common(self) -> dict:
        params = {"tool": self.tool, "api_key": self.api_key}
        if self.email:
            params["email"] = self.email
        return params

    def _ncbi_error(self, resp: requests.Response) -> str:
        match = re.search(r"<ERROR>(.*?)</ERROR>", resp.text, flags=re.I | re.S)
        if match:
            return brief_error(match.group(1))
        return brief_error(resp.text[:800])

    def _request(self, url: str, params: dict, label: str, soft: bool = False) -> requests.Response:
        last_error = ""
        attempt = 0
        transient_waits = 0
        while True:
            attempt += 1
            self.limiter.wait()
            try:
                resp = self.session.post(url, data=params, timeout=REQUEST_TIMEOUT)
            except requests.RequestException as exc:
                last_error = str(exc)
                if is_transient(last_error):
                    transient_waits = self._wait_transient(transient_waits, label, last_error)
                    continue
                self._backoff(attempt, label, last_error)
                if attempt >= MAX_RETRIES:
                    break
                continue
            if resp.status_code == 200 and resp.content:
                lowered = resp.text[:500].lower()
                if "error" in lowered and ("<error" in lowered or "api key" in lowered):
                    last_error = self._ncbi_error(resp)
                    if "api key" in lowered:
                        raise SystemExit(f"NCBI 拒绝了 API 密钥: {last_error}")
                    if is_transient(last_error):
                        transient_waits = self._wait_transient(transient_waits, label, last_error)
                        continue
                    self._backoff(attempt, label, last_error)
                    if attempt >= MAX_RETRIES:
                        break
                    continue
                return resp
            last_error = f"HTTP {resp.status_code} {self._ncbi_error(resp)}"
            transient = resp.status_code in (429, 500, 502, 503, 504) or not resp.content or is_transient(last_error)
            if transient:
                transient_waits = self._wait_transient(transient_waits, label, last_error)
                continue
            if soft:
                if attempt < 3:
                    log(f"{label} 第 {attempt} 次失败，2s 后重试: {last_error}")
                    time.sleep(2 * attempt)
                    continue
                raise RuntimeError(f"{label} 失败: {last_error}")
            raise SystemExit(f"{label} 失败: {last_error}")
        if soft or is_transient(last_error):
            raise RuntimeError(f"{label} 失败: {last_error}")
        raise SystemExit(f"{label} 连续失败 {MAX_RETRIES} 次: {last_error}")

    def _backoff(self, attempt: int, label: str, reason: str) -> None:
        delay = min(60, 2 ** attempt)
        log(f"{label} 第 {attempt} 次失败，{delay}s 后重试: {brief_error(reason)}")
        time.sleep(delay)

    def _wait_transient(self, waits: int, label: str, reason: str) -> int:
        waits += 1
        if waits > TRANSIENT_MAX_WAITS:
            raise SystemExit(
                f"{label} 持续不可用，已等待 {waits - 1} 次。重新运行同一命令会从断点继续。"
                f"{brief_error(reason)}"
            )
        delay = min(300, 30 * waits)
        log(f"{label} 暂时不可用，{delay}s 后重试（{waits}/{TRANSIENT_MAX_WAITS}）: {brief_error(reason)}")
        time.sleep(delay)
        return waits

    def _date_params(self, start: datetime, end: datetime, datetype: str, extra: str) -> dict:
        return {
            "db": "pubmed",
            "term": extra.strip() or "all[sb]",
            "mindate": start.strftime("%Y/%m/%d"),
            "maxdate": end.strftime("%Y/%m/%d"),
            "datetype": datetype,
        }

    def _parse_esearch(self, resp: requests.Response) -> dict:
        """NCBI 的 JSON 偶尔夹带未转义控制字符，先清掉再解析。"""
        text = resp.content.decode("utf-8", errors="replace")
        cleaned = re.sub(r"[\x00-\x1f]", " ", text)
        try:
            payload = json.loads(cleaned)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"ESearch 返回无法解析: {cleaned[:240]}") from exc
        data = payload.get("esearchresult")
        if not isinstance(data, dict):
            raise RuntimeError(f"ESearch 响应异常: {cleaned[:240]}")
        if data.get("ERROR"):
            raise RuntimeError(f"ESearch 错误: {brief_error(str(data['ERROR']))}")
        return data

    def _esearch(self, params: dict, label: str) -> dict:
        """检索遇到 SOLR / HTTP 500 时长时间等待，不把整次任务打断。"""
        waits = 0
        while True:
            try:
                resp = self._request(ESEARCH_URL, params, label)
                return self._parse_esearch(resp)
            except RuntimeError as exc:
                message = brief_error(str(exc))
                if not is_transient(message):
                    raise RuntimeError(message) from exc
                waits = self._wait_transient(waits, label, message)

    def search_count(self, start: datetime, end: datetime, datetype: str, extra: str) -> int:
        params = {
            **self._common(),
            **self._date_params(start, end, datetype, extra),
            "retmode": "json",
            "retmax": 0,
        }
        data = self._esearch(params, "ESearch计数")
        return int(data.get("count") or 0)

    def _uid_term(self, extra: str, lo: int, hi: int) -> str:
        base = extra.strip() or "all[sb]"
        return f"({base}) AND {lo}:{hi}[uid]"

    def _search_one_page(self, start: datetime, end: datetime, datetype: str, extra: str) -> tuple[int, list[str]]:
        """PubMed 的 ESearch 最多返回 9999 条，超过时只带回 Count。"""
        params = {
            **self._common(),
            **self._date_params(start, end, datetype, extra),
            "retmode": "json",
            "retmax": ESEARCH_RETMAX,
            "retstart": 0,
        }
        data = self._esearch(params, "ESearch")
        count = int(data.get("count") or 0)
        ids = list(dict.fromkeys(data.get("idlist") or []))
        if count <= ESEARCH_RETMAX and len(ids) != count:
            raise RuntimeError(f"期望 {count} 条 PMID，实际返回 {len(ids)}")
        return count, ids

    def _collect_by_uid(self, start: datetime, end: datetime, datetype: str,
                        extra: str, lo: int, hi: int) -> list[str]:
        count, ids = self._search_one_page(start, end, datetype, self._uid_term(extra, lo, hi))
        if count == 0:
            return []
        if count <= ESEARCH_RETMAX:
            return ids
        if lo >= hi:
            raise RuntimeError(f"PMID {lo} 命中 {count}，无法继续切分")
        mid = (lo + hi) // 2
        log(f"命中 {count} 超过 {ESEARCH_RETMAX}，按 PMID {lo}-{mid} / {mid + 1}-{hi} 拆开")
        left = self._collect_by_uid(start, end, datetype, extra, lo, mid)
        right = self._collect_by_uid(start, end, datetype, extra, mid + 1, hi)
        return left + right

    def search_ids(self, start: datetime, end: datetime, datetype: str, extra: str) -> list[str]:
        """命中不超过 9999 时直接取 PMID；超过则按 PMID 区间对半切分，直到每段取全。"""
        label = query_label(start, end, datetype, extra)
        last_gap = ""
        attempt = 0
        transient_waits = 0
        while attempt < MAX_RETRIES:
            try:
                total = self.search_count(start, end, datetype, extra)
                if total == 0:
                    return []
                if total <= ESEARCH_RETMAX:
                    count, ids = self._search_one_page(start, end, datetype, extra)
                    if len(ids) == count == total:
                        return ids
                    last_gap = f"期望 {total}，实际 {len(ids)}"
                else:
                    log(f"{label} 命中 {total}，超过单次 9999 条，开始按 PMID 分段")
                    ids = self._collect_by_uid(start, end, datetype, extra, UID_MIN, UID_MAX)
                    unique = list(dict.fromkeys(ids))
                    if len(unique) == total:
                        return unique
                    last_gap = f"期望 {total}，分段合计 {len(unique)}"
            except RuntimeError as exc:
                last_gap = brief_error(str(exc))
                if is_transient(last_gap):
                    transient_waits = self._wait_transient(transient_waits, "ESearch", last_gap)
                    continue
            attempt += 1
            log(f"PMID 列表不完整（{last_gap}），重新检索（{attempt}/{MAX_RETRIES}）")
            time.sleep(min(30, 2 ** attempt))
        raise SystemExit(f"无法取全 PMID: {label}（{last_gap}）")

    def fetch_xml(self, pmids: list[str]) -> ET.Element:
        params = {
            **self._common(),
            "db": "pubmed",
            "id": ",".join(pmids),
            "retmode": "xml",
            "rettype": "xml",
        }
        resp = self._request(EFETCH_URL, params, f"EFetch({len(pmids)})", soft=True)
        payload = resp.content
        start = payload.find(b"<")
        if start > 0:
            payload = payload[start:]
        try:
            return ET.fromstring(payload)
        except ET.ParseError:
            text = re.sub(
                r"[\x00-\x08\x0b\x0c\x0e-\x1f]",
                "",
                payload.decode("utf-8", errors="replace"),
            )
            try:
                return ET.fromstring(text)
            except ET.ParseError as exc:
                recovered = recover_xml(text)
                if recovered is not None:
                    return recovered
                raise RuntimeError(f"EFetch XML 解析失败: {exc}") from exc


def recover_xml(text: str) -> ET.Element | None:
    """标准解析失败时，用 lxml 的容错模式尽量收回这一批。未安装 lxml 则跳过。"""
    try:
        from lxml import etree
    except ImportError:
        return None
    parser = etree.XMLParser(recover=True, huge_tree=True)
    root = etree.fromstring(text.encode("utf-8"), parser)
    if root is None:
        return None
    try:
        return ET.fromstring(etree.tostring(root))
    except ET.ParseError:
        return None


def query_label(start: datetime, end: datetime, datetype: str, extra: str) -> str:
    clause = f"{ymd(start)}:{ymd(end)}[{DATE_TYPES[datetype]}]"
    extra = extra.strip()
    if extra:
        return f"({extra}) AND ({clause})"
    return clause


def format_pubdate(node: ET.Element | None) -> str:
    if node is None:
        return ""
    medline = node_text(node, "MedlineDate")
    if medline:
        return medline
    year = node_text(node, "Year")
    month = node_text(node, "Month")
    day = node_text(node, "Day")
    season = node_text(node, "Season")
    return join_parts([year, month or season, day], " ")


def format_author(author: ET.Element) -> str:
    collective = node_text(author, "CollectiveName")
    if collective:
        return collective
    last = node_text(author, "LastName")
    fore = node_text(author, "ForeName") or node_text(author, "Initials")
    return clean_text(f"{last} {fore}")


def parse_article(article: ET.Element) -> dict | None:
    citation = article.find("MedlineCitation")
    book = article.find("BookDocument")
    if citation is None and book is not None:
        return parse_book(article)
    if citation is None:
        return None
    pmid = node_text(citation, "PMID")
    if not pmid:
        return None
    art = citation.find("Article")
    if art is None:
        art = ET.Element("Article")
    journal = art.find("Journal")
    issue = journal.find("JournalIssue") if journal is not None else None
    authors: list[str] = []
    affiliations: list[str] = []
    for author in art.findall("./AuthorList/Author"):
        name = format_author(author)
        if not name:
            continue
        authors.append(name)
        affs = [node_text(a, "Affiliation") for a in author.findall("AffiliationInfo")]
        affs = [a for a in affs if a]
        if affs:
            affiliations.append(f"{name}: {join_parts(affs)}")
    abstract_parts = []
    for block in art.findall("./Abstract/AbstractText"):
        label = clean_text(block.attrib.get("Label") or "")
        body = clean_text("".join(block.itertext()))
        if label and body:
            abstract_parts.append(f"{label}: {body}")
        elif body:
            abstract_parts.append(body)
    mesh = []
    for heading in citation.findall("./MeshHeadingList/MeshHeading"):
        desc = heading.find("DescriptorName")
        if desc is None or not desc.text:
            continue
        major = "*" if desc.attrib.get("MajorTopicYN") == "Y" else ""
        quals = []
        for qual in heading.findall("QualifierName"):
            q = clean_text(qual.text)
            if not q:
                continue
            star = "*" if qual.attrib.get("MajorTopicYN") == "Y" else ""
            quals.append(f"{q}{star}")
        item = clean_text(desc.text) + major
        if quals:
            item += "/" + "/".join(quals)
        mesh.append(item)
    keywords = [
        clean_text("".join(k.itertext()))
        for k in citation.findall("./KeywordList/Keyword")
    ]
    id_map = {}
    for aid in article.findall("./PubmedData/ArticleIdList/ArticleId"):
        id_map[(aid.attrib.get("IdType") or "").lower()] = clean_text(aid.text)
    doi = id_map.get("doi") or ""
    if not doi:
        for loc in art.findall("ELocationID"):
            if (loc.attrib.get("EIdType") or "").lower() == "doi" and loc.text:
                doi = clean_text(loc.text)
                break
    entrez = ""
    epub = ""
    history = {}
    for hist in article.findall("./PubmedData/History/PubMedPubDate"):
        status = (hist.attrib.get("PubStatus") or "").lower()
        history[status] = format_pubdate(hist)
    entrez = history.get("pubmed") or history.get("medline") or history.get("entrez") or ""
    epub = history.get("epublish") or ""
    grants = []
    for grant in art.findall("./GrantList/Grant"):
        grants.append(join_parts([
            node_text(grant, "Agency"),
            node_text(grant, "GrantID"),
            node_text(grant, "Country"),
        ], " "))
    medline_info = citation.find("MedlineJournalInfo")
    row = {
        "PMID": pmid,
        "标题": clean_text("".join(art.find("ArticleTitle").itertext())) if art.find("ArticleTitle") is not None else "",
        "摘要": join_parts(abstract_parts, "\n"),
        "作者": join_parts(authors),
        "作者单位": join_parts(affiliations, " | "),
        "期刊": node_text(journal, "Title") if journal is not None else "",
        "期刊缩写": node_text(journal, "ISOAbbreviation") if journal is not None else "",
        "ISSN": node_text(journal, "ISSN") if journal is not None else "",
        "卷": node_text(issue, "Volume") if issue is not None else "",
        "期": node_text(issue, "Issue") if issue is not None else "",
        "页码": node_text(art, "Pagination/MedlinePgn"),
        "出版日期": format_pubdate(issue.find("PubDate") if issue is not None else None),
        "电子出版日期": epub or format_pubdate(art.find("ArticleDate")),
        "入库日期": entrez,
        "DOI": doi,
        "PMCID": id_map.get("pmc", ""),
        "文献类型": join_parts([
            clean_text(p.text) for p in art.findall("./PublicationTypeList/PublicationType")
        ]),
        "MeSH主题词": join_parts(mesh, " | "),
        "关键词": join_parts([k for k in keywords if k]),
        "语种": join_parts([node_text(lang) for lang in art.findall("Language")]),
        "出版国": node_text(medline_info, "Country") if medline_info is not None else "",
        "NLM唯一号": node_text(medline_info, "NlmUniqueID") if medline_info is not None else "",
        "版权": node_text(art, "Abstract/CopyrightInformation"),
        "基金": join_parts(grants, " | "),
        "利益冲突": node_text(art, "CoiStatement"),
        "PubMed链接": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
    }
    return {key: clip_cell(row.get(key, "")) for key in COLUMNS}


def parse_book(article: ET.Element) -> dict | None:
    book = article.find("BookDocument")
    if book is None:
        return None
    pmid = node_text(book, "PMID")
    if not pmid:
        return None
    book_node = book.find("Book")
    authors = []
    if book_node is not None:
        for author in book_node.findall("./AuthorList/Author"):
            name = format_author(author)
            if name:
                authors.append(name)
    title = ""
    if book_node is not None:
        title = clean_text("".join(book_node.findtext("BookTitle", default="") or ""))
    row = {key: "" for key in COLUMNS}
    row.update({
        "PMID": pmid,
        "标题": title or node_text(book, "ArticleTitle"),
        "作者": join_parts(authors),
        "期刊": title,
        "出版日期": format_pubdate(book_node.find("PubDate") if book_node is not None else None),
        "文献类型": "Book",
        "PubMed链接": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
    })
    return {key: clip_cell(row.get(key, "")) for key in COLUMNS}


def parse_pubmed_xml(root: ET.Element) -> tuple[list[dict], list[str]]:
    rows = []
    errors = []
    for err in root.findall(".//ERROR"):
        if err.text:
            errors.append(clean_text(err.text))
    for tag in ("PubmedArticle", "PubmedBookArticle"):
        for node in root.findall(tag):
            row = parse_article(node)
            if row and row["PMID"]:
                rows.append(row)
    return rows, errors


class Store:
    def __init__(self, db_path: Path) -> None:
        db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(db_path))
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA synchronous=NORMAL")
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS jobs (
                job_key TEXT PRIMARY KEY,
                label TEXT NOT NULL,
                expected_count INTEGER,
                status TEXT NOT NULL,
                updated_at TEXT
            );
            CREATE TABLE IF NOT EXISTS job_pmids (
                job_key TEXT NOT NULL,
                pmid TEXT NOT NULL,
                PRIMARY KEY (job_key, pmid)
            );
            CREATE TABLE IF NOT EXISTS records (
                pmid TEXT PRIMARY KEY,
                payload TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS missing (
                pmid TEXT PRIMARY KEY,
                reason TEXT,
                tries INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS exported (
                pmid TEXT PRIMARY KEY
            );
            """
        )
        self.conn.commit()

    def close(self) -> None:
        self.conn.close()

    def job_status(self, job_key: str) -> str | None:
        row = self.conn.execute(
            "SELECT status FROM jobs WHERE job_key = ?", (job_key,)
        ).fetchone()
        return row[0] if row else None

    def mark_job(self, job_key: str, label: str, status: str, expected: int | None = None) -> None:
        self.conn.execute(
            """
            INSERT INTO jobs (job_key, label, expected_count, status, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(job_key) DO UPDATE SET
                label = excluded.label,
                expected_count = COALESCE(excluded.expected_count, jobs.expected_count),
                status = excluded.status,
                updated_at = excluded.updated_at
            """,
            (job_key, label, expected, status, datetime.now().isoformat(timespec="seconds")),
        )
        self.conn.commit()

    def reset_ids(self, job_key: str) -> None:
        self.conn.execute("DELETE FROM job_pmids WHERE job_key = ?", (job_key,))
        self.conn.commit()

    def add_pmids(self, job_key: str, pmids: list[str]) -> None:
        self.conn.executemany(
            "INSERT OR IGNORE INTO job_pmids (job_key, pmid) VALUES (?, ?)",
            ((job_key, pmid) for pmid in pmids),
        )
        self.conn.commit()

    def pmid_count(self, job_key: str) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM job_pmids WHERE job_key = ?", (job_key,)
        ).fetchone()[0]

    def missing_pmids(self, job_key: str, limit: int) -> list[str]:
        rows = self.conn.execute(
            """
            SELECT j.pmid
            FROM job_pmids j
            LEFT JOIN records r ON r.pmid = j.pmid
            LEFT JOIN missing m ON m.pmid = j.pmid
            WHERE j.job_key = ?
              AND r.pmid IS NULL
              AND (m.pmid IS NULL OR m.tries < ?)
            LIMIT ?
            """,
            (job_key, MAX_RETRIES, limit),
        ).fetchall()
        return [row[0] for row in rows]

    def pending_count(self, job_key: str) -> int:
        return self.conn.execute(
            """
            SELECT COUNT(*)
            FROM job_pmids j
            LEFT JOIN records r ON r.pmid = j.pmid
            LEFT JOIN missing m ON m.pmid = j.pmid AND m.tries >= ?
            WHERE j.job_key = ? AND r.pmid IS NULL AND m.pmid IS NULL
            """,
            (MAX_RETRIES, job_key),
        ).fetchone()[0]

    def save_records(self, rows: list[dict]) -> int:
        new_rows = 0
        for row in rows:
            cur = self.conn.execute(
                "INSERT OR IGNORE INTO records (pmid, payload) VALUES (?, ?)",
                (row["PMID"], json.dumps(row, ensure_ascii=False)),
            )
            new_rows += cur.rowcount
        self.conn.commit()
        return new_rows

    def mark_missing(self, pmids: list[str], reason: str) -> None:
        self.conn.executemany(
            """
            INSERT INTO missing (pmid, reason, tries) VALUES (?, ?, 1)
            ON CONFLICT(pmid) DO UPDATE SET
                tries = tries + 1,
                reason = excluded.reason
            """,
            ((pmid, reason) for pmid in pmids),
        )
        self.conn.commit()

    def unexported_rows(self, job_key: str) -> list[dict]:
        """同一 PMID 只导出一次；已写入 Excel 的记录不再重复写出。"""
        rows = self.conn.execute(
            """
            SELECT r.pmid, r.payload
            FROM job_pmids j
            JOIN records r ON r.pmid = j.pmid
            LEFT JOIN exported e ON e.pmid = r.pmid
            WHERE j.job_key = ? AND e.pmid IS NULL
            ORDER BY CAST(j.pmid AS INTEGER)
            """,
            (job_key,),
        ).fetchall()
        return [json.loads(payload) for _pmid, payload in rows]

    def mark_exported(self, pmids: list[str]) -> None:
        self.conn.executemany(
            "INSERT OR IGNORE INTO exported (pmid) VALUES (?)",
            ((pmid,) for pmid in pmids),
        )
        self.conn.commit()

    def saved_count(self, job_key: str) -> int:
        return self.conn.execute(
            """
            SELECT COUNT(*)
            FROM job_pmids j
            LEFT JOIN records r ON r.pmid = j.pmid
            WHERE j.job_key = ? AND r.pmid IS NOT NULL
            """,
            (job_key,),
        ).fetchone()[0]

    def distinct_saved(self, job_keys: list[str]) -> int:
        if not job_keys:
            return 0
        marks = ",".join("?" for _ in job_keys)
        return self.conn.execute(
            f"""
            SELECT COUNT(DISTINCT j.pmid)
            FROM job_pmids j
            JOIN records r ON r.pmid = j.pmid
            WHERE j.job_key IN ({marks})
            """,
            job_keys,
        ).fetchone()[0]


def job_key_for(kind: str, start: datetime, end: datetime, datetype: str, term: str) -> str:
    extra = term.strip() or "-"
    return f"{datetype}|{kind}|{day_key(start)}|{day_key(end)}|{extra}"


def collect_job(api: EUtils, store: Store, key: str, label: str,
                start: datetime, end: datetime, datetype: str, term: str) -> None:
    status = store.job_status(key)
    if status in ("ready", "done"):
        log(f"跳过已列出的 PMID: {label}（{store.pmid_count(key)} 条）")
        return
    log(f"检索 PMID: {label}")
    pmids = api.search_ids(start, end, datetype, term)
    store.reset_ids(key)
    store.add_pmids(key, pmids)
    saved = store.pmid_count(key)
    if saved != len(pmids):
        raise SystemExit(f"{label} PMID 入库数量不一致: {saved} != {len(pmids)}")
    store.mark_job(key, label, "ready", expected=len(pmids))
    log(f"PMID 已冻结: {label}，{len(pmids)} 条")


def fetch_batch(api: EUtils, store: Store, batch: list[str]) -> None:
    """请求或解析失败先重试；整批仍失败则对半拆开，避免一条坏记录中断全部下载。"""
    last_error = ""
    root = None
    for attempt in range(1, 4):
        try:
            root = api.fetch_xml(batch)
            break
        except RuntimeError as exc:
            last_error = str(exc)
            if attempt < 3:
                log(f"  本批 {len(batch)} 条下载失败，重试 {attempt}/3: {exc}")
                time.sleep(min(30, 2 ** attempt))
                continue
    if root is None:
        if len(batch) == 1:
            store.mark_missing(batch, last_error)
            log(f"  PMID {batch[0]} 下载失败，将稍后重试: {last_error}")
            return
        mid = max(1, len(batch) // 2)
        log(f"  本批 {len(batch)} 条下载失败，拆成 {mid} 与 {len(batch) - mid} 条")
        fetch_batch(api, store, batch[:mid])
        fetch_batch(api, store, batch[mid:])
        return
    rows, errors = parse_pubmed_xml(root)
    got = {row["PMID"] for row in rows}
    store.save_records(rows)
    missed = [pmid for pmid in batch if pmid not in got]
    if missed:
        reason = join_parts(errors) or "EFetch 未返回该 PMID"
        store.mark_missing(missed, reason)
        log(f"  本批 {len(batch)} 条，返回 {len(got)}，缺失 {len(missed)}（将重试）")


def fetch_job(api: EUtils, store: Store, key: str, label: str) -> None:
    store.conn.execute(
        """
        UPDATE missing SET tries = 0
        WHERE pmid IN (SELECT pmid FROM job_pmids WHERE job_key = ?)
        """,
        (key,),
    )
    store.conn.commit()
    total = store.pmid_count(key)
    pending = store.pending_count(key)
    log(f"下载题录: {label}，共 {total}，待下载 {pending}")
    done = total - pending
    while True:
        batch = store.missing_pmids(key, EFETCH_BATCH)
        if not batch:
            break
        fetch_batch(api, store, batch)
        done = total - store.pending_count(key)
        log(f"  进度 {done}/{total}")
    left = store.pending_count(key)
    if left:
        raise SystemExit(f"{label} 仍有 {left} 条 PMID 未取到，可直接重新运行本脚本继续")
    unresolved = store.pmid_count(key) - store.saved_count(key)
    if unresolved:
        raise SystemExit(
            f"{label} 有 {unresolved} 条 PMID 多次请求仍未返回。"
            "这些编号已记在断点库，重新运行会继续重试，不会重下已成功的记录。"
        )


def excel_paths(out_dir: Path, label: str, parts: int) -> list[Path]:
    safe = re.sub(r'[<>:"/\\|?*]+', "_", label).strip(" ._")
    paths = []
    for i in range(parts):
        suffix = f"_{i + 1}" if parts > 1 else ""
        paths.append(out_dir / f"PubMed_{safe}{suffix}.xlsx")
    return paths


def write_excel(store: Store, key: str, out_dir: Path, label: str) -> int:
    out_dir.mkdir(parents=True, exist_ok=True)
    rows = store.unexported_rows(key)
    if not rows:
        log(f"无新记录需要导出: {label}")
        return 0
    parts = (len(rows) + EXCEL_MAX_DATA_ROWS - 1) // EXCEL_MAX_DATA_ROWS
    paths = excel_paths(out_dir, label, parts)
    for index, path in enumerate(paths):
        chunk = rows[index * EXCEL_MAX_DATA_ROWS:(index + 1) * EXCEL_MAX_DATA_ROWS]
        tmp = path.with_suffix(".xlsx.part")
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("题录")
        ws.append(COLUMNS)
        for row in chunk:
            ws.append([row.get(col, "") for col in COLUMNS])
        wb.save(tmp)
        tmp.replace(path)
        log(f"已写入 {path}（{len(chunk)} 行）")
    store.mark_exported([row["PMID"] for row in rows])
    return len(rows)


def run_unit(api: EUtils, store: Store, out_dir: Path, kind: str, start: datetime,
             end: datetime, datetype: str, term: str) -> str:
    label = f"{day_key(start)}_{day_key(end)}_{datetype}"
    if kind != "day":
        label = f"{kind}_{label}"
    key = job_key_for(kind, start, end, datetype, term)
    if store.job_status(key) == "done":
        log(f"已完成，跳过: {label}")
        return key
    collect_job(api, store, key, label, start, end, datetype, term)
    fetch_job(api, store, key, label)
    written = write_excel(store, key, out_dir, label)
    expected = store.pmid_count(key)
    saved = store.saved_count(key)
    if saved != expected:
        raise SystemExit(f"{label} 核对失败: 应有 {expected}，已保存 {saved}")
    store.mark_job(key, label, "done", expected=expected)
    log(f"完成: {label}，PMID {expected}，本文件新写入 {written}")
    return key


def fill_gaps(api: EUtils, store: Store, out_dir: Path, start: datetime, end: datetime,
              datetype: str, term: str, day_keys: list[str]) -> list[str]:
    """逐日条数小于整段命中时，按月、再按年补齐没有具体日的记录。"""
    keys = list(day_keys)
    daily_sum = sum(store.pmid_count(key) for key in day_keys)
    total = api.search_count(start, end, datetype, term)
    log(f"整段命中 {total}，逐日合计 {daily_sum}")
    if daily_sum == total:
        log("逐日结果与整段命中一致，无需补漏")
        return keys
    if daily_sum > total:
        log(f"逐日合计比整段多 {daily_sum - total}，导出时已按 PMID 去重")
        return keys
    log(f"相差 {total - daily_sum}，开始按月补齐未落到单日的记录")
    for first, last in month_spans(start, end):
        if first == last:
            continue
        keys.append(run_unit(api, store, out_dir, "month", first, last, datetype, term))
    covered = store.distinct_saved(keys)
    again = api.search_count(start, end, datetype, term)
    if covered >= again:
        log(f"补齐后已覆盖 {covered} 条，整段命中 {again}")
        return keys
    log(f"按月补齐后覆盖 {covered}，整段命中 {again}，继续按年补齐")
    for first, last in year_spans(start, end):
        keys.append(run_unit(api, store, out_dir, "year", first, last, datetype, term))
    return keys


def resolve_checkpoint(out_dir: Path, checkpoint: str) -> Path:
    """未指定时使用 Excel 目录下的 pubmed_checkpoint.sqlite。"""
    text = checkpoint.strip().strip('"')
    if not text:
        return out_dir / "pubmed_checkpoint.sqlite"
    path = Path(text)
    if text.endswith(("\\", "/")) or (path.exists() and path.is_dir()):
        return path / "pubmed_checkpoint.sqlite"
    return path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="按日期从 PubMed 下载题录并导出 Excel")
    parser.add_argument("--start", default=BGDATE, help="开始日期 YYYY-MM-DD，含当天")
    parser.add_argument("--end", default=ENDDATE, help="结束日期 YYYY-MM-DD，含当天")
    parser.add_argument("--datetype", default=DATETYPE, choices=sorted(DATE_TYPES),
                        help="edat 入库日期（默认，不重不漏）；pdat 出版日期；mdat 修订日期")
    parser.add_argument("--term", default=TERM, help="附加检索式，与日期条件 AND")
    parser.add_argument("--out", default=OUT_DIR, help="Excel 保存目录")
    parser.add_argument(
        "--checkpoint",
        default=CHECKPOINT,
        help="pubmed_checkpoint.sqlite 的路径。留空则放在 Excel 目录下；"
             "也可指定目录，文件名仍为 pubmed_checkpoint.sqlite",
    )
    parser.add_argument("--api-key", default=API_KEY, help="NCBI API Key")
    parser.add_argument("--email", default=EMAIL, help="联系邮箱，NCBI 建议填写")
    return parser


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    args = build_parser().parse_args()
    start = parse_date(args.start)
    end = parse_date(args.end)
    if end < start:
        raise SystemExit("结束日期不能早于开始日期")
    if not args.api_key.strip():
        raise SystemExit("缺少 NCBI API Key")
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    db_path = resolve_checkpoint(out_dir, args.checkpoint)
    log(f"日期 {day_key(start)} ~ {day_key(end)}，类型 {args.datetype}")
    log(f"输出目录: {out_dir}")
    log(f"断点库: {db_path}")
    if args.datetype != "edat":
        log("当前不是 EDAT。出版日期/修订日期里只有年月的记录不会落在某一天，脚本会在逐日结束后按月、按年补齐并按 PMID 去重。")
    api = EUtils(args.api_key, args.email, TOOL)
    store = Store(db_path)
    try:
        day_keys = []
        for day in daterange(start, end):
            day_keys.append(run_unit(
                api, store, out_dir, "day", day, day, args.datetype, args.term
            ))
        if start != end or args.datetype != "edat":
            job_keys = fill_gaps(
                api, store, out_dir, start, end, args.datetype, args.term, day_keys
            )
        else:
            job_keys = day_keys
        whole = api.search_count(start, end, args.datetype, args.term)
        saved = store.distinct_saved(job_keys)
        log(f"核对: 整段命中 {whole}，已保存唯一 PMID {saved}")
        if saved < whole:
            raise SystemExit(f"结果不完整: 命中 {whole}，已保存 {saved}。重新运行将只补缺失部分。")
        exported = store.conn.execute("SELECT COUNT(*) FROM exported").fetchone()[0]
        log(f"全部完成。输出目录内已导出不重复 PMID {exported} 条。断点库: {db_path}")
    finally:
        store.close()


if __name__ == "__main__":
    main()
