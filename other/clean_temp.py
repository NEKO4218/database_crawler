import os
import shutil

def clean_directory(target_path):
    # 检查路径是否存在
    if not os.path.exists(target_path):
        print(f"错误：路径 '{target_path}' 不存在。")
        return

    print(f"开始扫描目录：{target_path}")
    
    # 用于存储不为空的 export_failures.txt 的路径
    non_empty_failures = []

    # 使用 os.walk 遍历目录树
    # topdown=False 表示先处理子目录，再处理父目录
    for root, dirs, files in os.walk(target_path, topdown=False):
        
        # --- 文件处理逻辑 ---
        for filename in files:
            file_path = os.path.join(root, filename)
            
            try:
                # 1. 名为 export_failures.txt 的文件
                if filename == "export_failures.txt":
                    if os.path.getsize(file_path) == 0:
                        os.remove(file_path)
                        print(f"[已删除] 空文件: {file_path}")
                    else:
                        non_empty_failures.append(file_path)

                # 2. 名为 downloads.html 的文件
                elif filename == "downloads.html":
                    os.remove(file_path)
                    print(f"[已删除] HTML文件: {file_path}")

                # 3. 以 download_ 开头的 .txt 空文件
                elif filename.startswith("download_") and filename.endswith(".txt"):
                    if os.path.getsize(file_path) == 0:
                        os.remove(file_path)
                        print(f"[已删除] 空下载日志: {file_path}")

            except Exception as e:
                print(f"处理文件 {file_path} 时出错: {e}")

        # --- 文件夹处理逻辑 ---
        # 4. 名为 captcha_debug 的文件夹（包括其内容）
        if "captcha_debug" in dirs:
            dir_path = os.path.join(root, "captcha_debug")
            try:
                shutil.rmtree(dir_path)
                print(f"[已删除] 文件夹: {dir_path}")
            except Exception as e:
                print(f"处理文件夹 {dir_path} 时出错: {e}")

    # 最后打印不为空的 export_failures.txt 路径
    if non_empty_failures:
        print("\n" + "="*40)
        print("以下 export_failures.txt 文件不为空，已保留：")
        for path in non_empty_failures:
            print(path)
        print("="*40)
    else:
        print("\n未发现不为空的 export_failures.txt 文件。")

if __name__ == "__main__":
    # 在这里输入你要清理的文件夹路径
    path_to_clean = input("请输入要清理的文件夹路径: ").strip().strip('"')
    
    clean_directory(path_to_clean)