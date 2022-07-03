from docx import Document
import os
from win32com import client
import time

class Change:
    #输入基本信息
    def __init__(self):
        self.jiance_path=os.getcwd()
        self.path=self.jiance_path+'\\jiance'
        self.path1=self.jiance_path+'\\tanshang'
        #评估文件夹路径
        self.ping_gu=self.jiance_path+'\\pinggu'
        self.file_list=[]#目标文件绝对路径列表
        self.file_list_tanshang=[]
        self.file_list_pinggu=[]

    #获取上传文档绝对路径列表
    def get_file_list(self):
        self.file_list0=os.listdir(self.path)
        for j in self.file_list0:
            self.file_list.append(os.path.join(self.path,j))

        self.file_list1 = os.listdir(self.path1)
        for i in self.file_list1:
            self.file_list_tanshang.append(os.path.join(self.path1, i))
        #评估统计文件路径列表
        self.file_list2=os.listdir(self.ping_gu)
        for k in self.file_list2:
            self.file_list_pinggu.append(os.path.join(self.ping_gu,k))
    def change(self,i,ix):
        word = client.Dispatch("Word.Application")
        doc = word.Documents.Open(i)
        # 使用参数16表示将doc转换成docx
        doc.SaveAs(ix, 16)
        doc.Close()
        word.Quit()
    #doc转化为docx格式
    def doc2docx(self):
        self.get_file_list()
        for i in self.file_list:
            if i[-1]=='c':
                ix=i+'x'
                self.change(i,ix)
                os.remove(i)

        for j in self.file_list_tanshang:
            if j[-1]=='c':
                jx=j+'x'
                self.change(j,jx)
                os.remove(j)

        for k in self.file_list_pinggu:
            if k[-1]=='c':
                kx=k+'x'
                self.change(k,kx)
                os.remove(k)

    def output1(self):
        self.doc2docx()
        file_list=[]
        file_list1=os.listdir(self.path)
        for j in file_list1:
            file_list.append(os.path.join(self.path,j))
        return file_list

    def output2(self):
        self.doc2docx()
        file_list0 = []
        file_list2 = os.listdir(self.path1)
        for j in file_list2:
            file_list0.append(os.path.join(self.path1, j))
        return file_list0

    def output3(self):
        self.doc2docx()
        file_list_pinggu = []
        file_list_pinggu1 = os.listdir(self.ping_gu)
        for j in file_list_pinggu1:
            file_list_pinggu.append(os.path.join(self.ping_gu, j))
        return file_list_pinggu
