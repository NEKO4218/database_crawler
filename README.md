# 题录自动化脚本（code）

本目录集中存放「数据库获取」项目中，面向四大文献数据库的 Python 脚本与相关配置（JSON）。脚本已按数据库分类，便于维护与按库运行。

## 目录结构

```
code/
├── CNKI/      # 中国知网（题录导出、PDF、验证码、统计去重等）
├── PubMed/    # PubMed / NCBI E-utilities 题录批量导出
├── SCI/       # Web of Science（SCI）题录导出、Q1 筛选、PDF 等
├── Scopus/    # Scopus Search API 题录批量导出
├── other/     # 与具体库无关的通用工具（压缩、Excel/PDF 统计）
└── README.md  # 本说明
```

## 四大数据库一览

| 目录        | 数据库         | 典型入口脚本                     | 说明                                               |
| ----------- | -------------- | -------------------------------- | -------------------------------------------------- |
| `CNKI/`   | 中国知网       | `cnkiforexcel.py`、`cnki.py` | Selenium 导出题录 / 下载 PDF；单位登录与验证码模块 |
| `PubMed/` | PubMed         | `PubMed.py`                    | 官方 API 按日检索，断点续跑，导出 Excel            |
| `SCI/`    | Web of Science | `sciforexcel.py`               | Selenium 导出题录；Q1/学科后处理与 OA PDF          |
| `Scopus/` | Scopus         | `Scopus.py`                    | Scopus API 按日检索，断点续跑，导出 Excel          |

更细的文件清单与用法见各子目录中的 `README.md`。

## 使用约定

1. **工作目录**：进入对应子目录再运行脚本（同目录互相 `import` 的模块会一并找到）。
2. **凭证**：知网见 `CNKI/cnki_unit_credentials*.json`；SCI 见 `SCI/sci_credentials*.json`。请先复制 `*.example.json` 再填写真实账号，勿把含密码的 JSON 提交到公开仓库。
3. **输出路径**：多数脚本仍使用绝对路径或命令行参数指定下载/导出目录（如「知网数据库下载」「SCI数据库下载」「PubMed题录」「Scopus题录」），与代码目录分离。
4. **依赖**：在仓库根目录执行 `python -m pip install -r requirements.txt`（Python 3.10-3.12，另需本机安装 Chrome）。知网 GPU 验证码、`Text_select_captcha` 服务及 AntiCAP 试验所需的可选依赖见文件内注释。

## 其他

- `other/`：跨库通用的批处理小工具，不属于某一数据库流水线。
- 题录根目录下的数据、Excel、Chrome 配置等**不会**放入本目录，仅整理脚本与配置类 `.py` / `.json`。
