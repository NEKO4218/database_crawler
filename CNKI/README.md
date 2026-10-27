# CNKI（中国知网）

本目录为知网题录导出、全文 PDF、验证码处理与结果统计相关脚本。

## 文件说明

### 核心流程

| 文件 | 作用 |
|------|------|
| `cnkiforexcel.py` | 主流程：高级检索 → 导出题录 Excel（期刊 / 专利、文献分类、可选主题、日期粒度、单位登录与验证码） |
| `cnki.py` | 较早的 PDF 下载工具（按时间段 / 学科目录下载，依赖 `cnkiforexcel` 登录与验证码能力） |
| `cnki_excel.py` | 根据已导出的 `CNKI-*.xls/xlsx` 中的 URL 批量下载 PDF，并抓取作者单位等 |
| `cnki_author.py` | 从详情页补充作者与单位信息到 Excel |

### 验证码与登录

| 文件 | 作用 |
|------|------|
| `cnki_captcha.py` | 验证码公共模块（腾讯验证码、滑块 blockPuzzle、文字点选等） |
| `cnki_unit_credentials.example.json` | 单位登录凭证模板（复制为下方正式文件后填写） |
| `cnki_unit_credentials.json` | 单位登录凭证（本地使用，勿外传） |
| `test_block_puzzle_captcha.py` | 滑块验证码联调 |
| `test_text_click.py` | 文字点选验证码联调 |
| `test_anticap.py` | AntiCAP 相关试验 |
| `Text_select_captcha/` | 文字点选验证码服务与模型相关代码包 |

### 统计与质检

| 文件 | 作用 |
|------|------|
| `cnki_summary.py` | 统计题录覆盖、缺失时段、重复记录，生成汇总表 |
| `cnki_fill_failures.py` | 根据汇总结果整理/补写失败时段日志 |
| `cnki_dedup_files.py` | 按汇总中的重复信息删除冗余题录文件 |

## 建议运行顺序

1. 配置 `cnki_unit_credentials.json`（可参考 example）。
2. 运行 `cnkiforexcel.py` 导出题录。
3. 需要全文时运行 `cnki_excel.py`（或历史流程 `cnki.py`）。
4. 用 `cnki_summary.py` 做覆盖与重复检查，必要时去重或补失败时段。
5. 需要作者-单位对应时运行 `cnki_author.py`：无参数运行会进入交互模式，按提示输入源目录、输出目录、并发数、文件/论文上限等；也可用 `--folder`、`--out` 等参数非交互运行。

## 期刊 / 专利题录导出

在 `cnkiforexcel.py` 中通过任务字段或交互配置：

| 字段 | 说明 |
|------|------|
| `doc_type` | `journal`（学术期刊）、`patent`（专利库 `classid=VUDIXAIY`）或 `none`（不选，跳过结果页资源类型筛选） |
| `lang` | `zh`=中文、`en`=英文、`none`=不选（跳过检索页/结果页中/外文筛选） |
| `category` | 期刊=学科导航；专利=国际专利分类 IPC（如 `['G','G06']`） |
| `theme` | 可选主题；专利写入第2行「主题」(SU)，不填「公开号」 |
| `patent_date` | 专利：`AD`=申请日 / `OD`=公开日（默认申请日） |
| `bgdate` / `initdate` / `granularity` | 日期范围与切分粒度 |

专利检索流程与期刊不同：重置条件 → 展开 IPC 侧栏 → 勾选分类号 → 主题行 → 申请日或公开日（互斥）→ 检索；结果页不再点 JOURNAL。

```bash
python cnkiforexcel.py --doc-type patent
python cnkiforexcel.py --doc-type journal
python cnkiforexcel.py --doc-type none   # 不选导出类型，跳过资源类型筛选
```

选择「不选」时：`doc_type=none` 不点 JOURNAL/PATENT；`lang=none` 不勾选中/外文，也不点结果页语言 tab。

专利结果保存到 `DOWNLOAD_ROOT/专利/...`；带主题时目录名会追加主题后缀。

## 注意

- 请在本目录（`code/CNKI`）下执行脚本，以保证同目录模块互相导入正常。
- 下载根目录等路径多在脚本内配置或通过命令行参数指定，与代码目录相互独立。
