import os
import zipfile
import sys

def compress_subfolders(root_path):
    """
    遍历指定路径下的所有子文件夹，并将每个子文件夹单独压缩为zip文件。
    """
    # 检查路径是否存在
    if not os.path.exists(root_path):
        print(f"错误：路径 '{root_path}' 不存在！")
        return

    # 获取所有子文件夹
    subfolders = [f.path for f in os.scandir(root_path) if f.is_dir()]
    
    if not subfolders:
        print(f"在 '{root_path}' 下未找到任何子文件夹。")
        return

    print(f"找到 {len(subfolders)} 个子文件夹，开始压缩...")

    for folder_path in subfolders:
        folder_name = os.path.basename(folder_path)
        zip_filename = os.path.join(root_path, f"{folder_name}.zip")
        
        try:
            # 创建zip文件
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # 遍历文件夹中的所有文件
                for root, dirs, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        # 计算文件在zip中的相对路径，保持目录结构
                        arcname = os.path.relpath(file_path, root_path)
                        zipf.write(file_path, arcname)
            
            print(f"成功压缩: {folder_name} -> {zip_filename}")
            
        except Exception as e:
            print(f"压缩失败 {folder_name}: {e}")

if __name__ == "__main__":
    
    # 如果运行脚本时带了参数，则使用参数作为路径
    if len(sys.argv) > 1:
        target_folder = sys.argv[1]
        compress_subfolders(target_folder)
    else:
        # 否则提示用户输入路径
        target_folder = input("请输入要压缩的文件夹路径: ")
        compress_subfolders(target_folder)