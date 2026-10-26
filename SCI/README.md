# SCI（Web of Science）

本目录为 Web of Science（SCI）题录自动导出、Q1/学科筛选、期刊合并与全文获取相关脚本。

## 文件说明

### 题录导出与登录

| 文件 | 作用 |
|------|------|
| `sciforexcel.py` | 主流程：登录 WoS → 检索 → 自定义 Excel 导出（结构对齐知网 `cnkiforexcel.py`） |
| `sci_journals.py` | 按「期刊 / Source Titles」检索并导出（可配合 Q1 期刊清单） |
| `sci_credentials.example.json` | 登录凭证模板 |
| `sci_credentials.json` | 登录凭证（本地使用，勿外传） |
| `SCI自动化登录流程.md` | 登录流程说明文档 |
| `SCI自动化导出题录流程.md` | 导出题录流程说明文档 |

### Q1 / 学科后处理

| 文件 | 作用 |
|------|------|
| `extract_q1.py` | 从中科院期刊分区表提取 1 区期刊为 `Q1.xlsx` |
| `classify_q1_journals.py` | 将 Q1 期刊按领域关键词分到若干专业类 |
| `extract_wos_q1.py` | 按 Q1 清单从 WoS 题录合集中筛出 1 区记录 |
| `extract_wos_q1_domains.py` | 在 Q1 结果上再按领域筛选 |
| `extract_wos_q1_journals.py` | 从 Q1 结果提取期刊名（Source Title）去重列表 |
| `extract_wos_categories.py` | 提取并汇总 WoS Categories |
| `merge_journal_records.py` | 合并各领域/学科下的期刊题录 Excel |
| `check_ra_export_counts.py` | 按 Research Areas 统计表核对导出条数，可补缺/删多余 |

### 全文 PDF

| 文件 | 作用 |
|------|------|
| `download_doi_pdfs.py` | 按 DOI 通过官方跳转 / Unpaywall 等合法 OA 渠道下载 PDF |
| `sci_hub.py` | Sci-Hub 全链路下载工具（使用前请自行确认合规性） |

## 建议运行顺序

1. 配置 `sci_credentials.json`。
2. 运行 `sciforexcel.py`（或 `sci_journals.py`）导出题录。
3. 需要分区筛选时：`extract_q1.py` → `extract_wos_q1.py` →（可选）领域/期刊提取与合并。
4. 需要全文时再用 DOI/OA 或其它下载脚本。

## 注意

- 请在本目录（`code/SCI`）下执行脚本，以保证 `sci_journals` 等对 `sciforexcel` 的导入正常。
- 数据输出目录通常指向「SCI数据库下载」等业务路径，与代码目录分离。
