# PubMed

本目录为通过 NCBI E-utilities 批量获取 PubMed 题录并导出 Excel 的脚本。

## 文件说明

| 文件 | 作用 |
|------|------|
| `PubMed.py` | 按日期范围检索 PubMed，导出题录 Excel；支持断点续跑与 PMID 去重 |
| `download_pmc_fulltext.py` | 扫描指定文件夹下题录 Excel，按 PMCID 下载开放获取 PDF/XML |

## 能力要点

- **不重不漏**：默认按入库日（EDAT）逐日检索；以 PMID 为主键去重；支持与总命中数核对。
- **断点续跑**：进度默认保存在输出目录的 `pubmed_checkpoint.sqlite`，中断后可续跑。
- **日期类型**：可用参数切换为出版日（PDAT）等。

## 示例

```bash
cd code/PubMed
python PubMed.py --start 2024-01-01 --end 2024-01-03 --out D:\PubMed题录
python PubMed.py --datetype pdat --start 2024-06-01 --end 2024-06-30 --out D:\PubMed题录

# 按题录中的 PMCID 下载开放获取全文（按文件流式读，边读边下）
python download_pmc_fulltext.py --excel D:\PubMed题录\PubMed_2023-01-02_2023-01-02_pdat.xlsx --out D:\PubMed全文 --format pdf --resume
python download_pmc_fulltext.py --excel-dir D:\PubMed题录 --out D:\PubMed全文 --format pdf --resume
python download_pmc_fulltext.py --excel-dir D:\PubMed题录 --format both --limit 100
```

## 依赖

- `requests`、`openpyxl`、`pandas`（缺失时脚本会提示安装）

## 注意

- 遵守 NCBI 使用规范与请求频率限制；大批量任务建议错峰、控制并发。
- 输出目录、断点库路径均通过命令行参数配置，不依赖本目录写数据。
