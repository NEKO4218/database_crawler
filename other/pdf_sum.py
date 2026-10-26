import pandas as pd
from pathlib import Path

def count_pdf_files(root_folder, output_filename="PDF统计结果.xlsx"):
    """
    递归统计指定文件夹下各子文件夹的PDF文件数量。
    """
    root_path = Path(root_folder)
    
    if not root_path.exists():
        print(f"❌ 错误：未找到文件夹 '{root_folder}'")
        return

    results = []
    # 定义PDF文件后缀
    pdf_extensions = ('.pdf',)

    print(f"🔍 正在递归扫描文件夹: {root_path.resolve()} ...")

    # 遍历根目录下的直接子文件夹（作为一级分类）
    for subfolder in sorted(root_path.iterdir()):
        if not subfolder.is_dir():
            continue
            
        file_count = 0
        
        # 使用 rglob 递归查找该子文件夹下所有层级的 PDF 文件
        for file in subfolder.rglob('*'):
            if file.is_file() and file.suffix.lower() in pdf_extensions:
                file_count += 1

        results.append({
            "一级子文件夹": subfolder.name,
            "PDF文件总数": file_count
        })

    if not results:
        print("ℹ️ 该目录下没有找到子文件夹。")
        return

    # 创建 DataFrame 并保存为 Excel
    df_result = pd.DataFrame(results)
    output_path = root_path / output_filename
    
    try:
        df_result.to_excel(output_path, index=False)
        print(f"✅ 统计完成！结果已保存至: {output_path}")
        print("\n📊 统计概览:")
        print(df_result.to_string(index=False))
    except PermissionError:
        print(f"❌ 保存失败：文件 '{output_filename}' 可能正在被打开，请关闭后重试。")

if __name__ == "__main__":
    # 👇 在这里修改你需要统计的文件夹路径
    TARGET_FOLDER = r"D:\2026年工作\SCI数据库下载\九类期刊全文" 
    
    count_pdf_files(TARGET_FOLDER)