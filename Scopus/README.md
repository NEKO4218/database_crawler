# Scopus

本目录为通过 Scopus Search API 批量获取题录并导出 Excel 的脚本。

## 文件说明

| 文件 | 作用 |
|------|------|
| `Scopus.py` | 按日期（及可选学科）检索 Scopus，导出题录 Excel；支持断点续跑与 EID 去重 |

## 能力要点

- **检索拆分**：默认按首次入库日逐日检索；单日超过 API 上限时按文献类型、出版年继续拆分。
- **断点续跑**：进度保存在输出目录的 `scopus_checkpoint.sqlite`；配额用尽或中断后可同命令续跑。
- **学科筛选**：可通过 `--subject` 等参数限制学科组；导出按 EID 去重。
- **摘要**：可选 Abstract Retrieval；取决于当前 API Key 权限视图。

## 示例

```bash
cd code/Scopus
python Scopus.py
python Scopus.py --year-from 2023 --year-to 2026 --start 2023-01-01 --end 2026-06-30 --out D:\Scopus题录
python Scopus.py --subject COMP,MATH --out D:\Scopus题录
```

## 注意

- API Key、周配额与每页条数上限以当前密钥实际权限为准（脚本启动时会探测）。
- 请在本目录运行；输出路径通过参数指定，勿把大批量数据写进代码目录。
