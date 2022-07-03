from first import Change
from jiance import JianCe
from tanshang import TanShang
from ping_gu import PingGu
import os
#调用井架检验类
def jing_jia_jian_yan(shen_he,bian_zhi):
    a = Change().output1()
    ji_jia = JianCe(shen_he, bian_zhi)
    ji_jia.copy_excel()
    ji_jia.open_excel()
    for i in range(len(a)):
        try:
            ab = ji_jia.open_docx(a[i])
            ji_jia.write_excel(ab, i)
            os.remove(a[i])
        except:
            print(a[i],'无法统计')
            continue
    ji_jia.save_excel()

#调用探伤类
def jian_ce_tan_shang(bian_zhi):
    b = Change().output2()
    tan_shang = TanShang(bian_zhi)
    tan_shang.copy_excel()
    tan_shang.open_excel()
    for i in range(len(b)):
        try:
            ab = tan_shang.open_docx(b[i])
            tan_shang.write_excel(ab, i)
            os.remove(b[i])
        except:
            print(b[i],'无法统计')
            continue
    tan_shang.save_excel()

#调用评估统计类
def zi_zhi_ping_gu(shen_he,bian_zhi):
    a = Change().output3()
    ping_gu = PingGu(shen_he, bian_zhi)
    ping_gu.copy_excel()
    ping_gu.open_excel()
    for i in range(len(a)):
        try:
            ab = ping_gu.open_docx(a[i])
            ping_gu.write_excel(ab, i)
            os.remove(a[i])
        except:
            print(a[i],'无法统计')
            continue
    ping_gu.save_excel()

if __name__ == '__main__':
    print('欢迎进入上传报告统计系统')
    choice=input('输入‘1’选择井架检验上传统计，输入‘2’选择探伤上传统计，输入‘3’选择评估上传统计:')
    shen_he=input('请输入审核人：')
    bian_zhi=input('请输入编制人：')
    if choice == '1':
        jing_jia_jian_yan(shen_he,bian_zhi)
    elif choice == '2':
        jian_ce_tan_shang(bian_zhi)
    elif choice == '3':
        zi_zhi_ping_gu(shen_he,bian_zhi)


