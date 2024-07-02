import os
import sys #允许程序传入参数
from glob import glob
from win32com import client as wc
from docx import Document
from docxcompose.composer import Composer

# cwd = os.getcwd()
cwd = sys.argv[1]

#获取doc 
doc_files = glob(cwd+"\\*.doc")
doc_files.sort()  #对doc_files列表中的文件路径进行排序

# 将doc文件转化为docx
word = wc.Dispatch("Word.Application") # 打开word应用程序
for file in doc_files:    #遍历doc_files列表中的每个文件路径
    doc = word.Documents.Open(file) # 打开word文件
    doc.SaveAs("{}x".format(file), 12) # 另存为后缀为".docx"的文件
    doc.Close() # 关闭原来word文件
    # if '.docx' in file:
    #     os.remove(file)
    #     os.rename(file + 'x', file)
word.Quit()

#列表排序函数
# 定义关键词及其优先级  
keywords = [('开题报告', 1), ('中期检查', 2), ('指导过程', 3), ('指导教师', 4), ('评阅教师', 5), ('答辩委员', 6),('答辩记录', 7),('答辩意见', 8)]  
# 创建一个排序关键字函数  
def sort_key(file_name):  
    # 初始化一个包含所有优先级的元组，初始值为一个很大的数（表示不存在该关键词）  
    priority_tuple = tuple([float('inf')] * len(keywords))  
      
    # 遍历关键词，更新优先级元组  
    for keyword, priority in keywords:  
        if keyword in file_name:  
            # 如果关键词在文件名中，更新对应位置的优先级  
            priority_tuple = tuple([priority if i == keywords.index((keyword, priority)) else x for i, x in enumerate(priority_tuple)])  
            break  # 如果已经找到一个关键词，就不需要再继续查找了（因为你想让含有某个关键词的元素排在前面）  
      
    # 返回优先级元组作为排序关键字  
    return priority_tuple  


# 合并docx

result=[]  # 建立一个空列表

#创建函数search|搜索：将路径下的name文件的 路径 放到result列表中
def search(path=".", name=""):
    for item in os.listdir(path):  # 遍历path下所有的文件和目录，赋值给item
        item_path = os.path.join(path, item) # 将文件与路径合成一个可执行文件路径

                    #检查变量 item_path 所表示的路径是否是一个目录/文件
        if os.path.isdir(item_path): #是目录就再来一遍
            search(item_path, name) 
        elif os.path.isfile(item_path): #是文件就判断是不是要用的文档
            if name in item:
                global result#声明访问了全局变量
                result.append(item_path) #将 item_path 添加到名为 result 的列表中。
                print (item_path)
    #排序
    result = sorted(result, key=lambda x: sort_key(x)) 



# search(path=cwd+'\\Word', name=".docx")
search(path = cwd, name=".docx")


files = result
#创建函数：combine_all_docx|合并整个文档
def combine_all_docx(filename_master,files_list):
    number_of_files = len(files_list) #获取列表中文件数量
    master = Document(filename_master) #创建一个新的文档，内容是filename_master路径下的文件
    master.add_page_break() # 在文末添加一个分页符
    composer = Composer(master) # 使用 composer 操作 master 文档

     #遍历文档，并添加
    for i in range(1, number_of_files):
        doc_temp = Document(files_list[i])
        doc_temp.add_page_break() # 在文末添加一个分页符
        composer.append(doc_temp)
    composer.save(sys.argv[2] + "-附件材料.docx")

combine_all_docx(result[0],result)
print("...合并完成...")
