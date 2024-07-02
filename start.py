import subprocess
import time  
import os
  
# 等待3秒  
#time.sleep(3) 

#获取文件夹和路径
def get_folder_info(path):  
    folder_info = []  
    for item in os.listdir(path):  
        full_path = os.path.join(path, item)  
        if os.path.isdir(full_path):  
            folder_info.append((item, full_path))  
    return folder_info  
  
# 指定要遍历的路径  
directory_path = 'C:\\AAA\\工作室\\backup\\整理'  
folder_info = get_folder_info(directory_path)  

for folder_name, folder_path in folder_info:  
    print(f"Folder Name: {folder_name}, Folder Path: {folder_path}")
    # 使用列表的形式组织命令行参数，第一个元素是要执行的脚本名  
    args = ['python', 'copyworld.py', folder_path, folder_name] 
    # 启动script2.py并等待其完成  
    subprocess.run(args)