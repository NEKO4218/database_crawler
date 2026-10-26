import os
import pandas as pd
from pathlib import Path

def count_excel_data_recursive(root_folder, output_filename="统计结果.xlsx"):
    root_path = Path(root_folder)
    
    if not root_path.exists():
        print(f"❌ 错误：未找到文件夹 '{root_folder}'")
        return

    results = []
    excel_extensions = ('.xlsx', '.xls')

    print(f"🔍 正在递归扫描文件夹: {root_path.resolve()} ...")

    # 遍历根目录下的直接子文件夹（作为一级分类）
    for subfolder in sorted(root_path.iterdir()):
        if not subfolder.is_dir():
            continue
            
        file_count = 0
        total_rows = 0
        
        # 使用 rglob 递归查找该子文件夹下所有层级的 Excel 文件
        for file in subfolder.rglob('*'):
            if file.is_file() and file.suffix.lower() in excel_extensions:
                file_count += 1
                try:
                    df = pd.read_excel(file, header=0)
                    total_rows += len(df)
                except Exception as e:
                    print(f"⚠️ 读取文件失败: {file.name}，错误信息: {e}")

        results.append({
            "一级子文件夹": subfolder.name,
            "Excel文件总数(含深层目录)": file_count,
            "数据总行数(不含表头)": total_rows
        })

    if not results:
        print("ℹ️ 该目录下没有找到子文件夹。")
        return

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
    TARGET_FOLDER = r"D:\2026年工作\SCI数据库下载" 
    count_excel_data_recursive(TARGET_FOLDER)