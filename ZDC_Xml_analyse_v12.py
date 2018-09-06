from openpyxl import Workbook
from openpyxl import load_workbook
import operator
from array import array
import re
import os,os.path
import shutil
import gzip, io
import xml.dom.minidom
from xml.dom.minidom import parse
from functools import reduce
#from Excel import *
from bs4 import BeautifulSoup
import bs4
import os,os.path
import shutil
import sys
import PyQt5.QtCore
from PyQt5.QtCore import pyqtSignal, QObject
from PyQt5.QtWidgets import (QMainWindow, QTextEdit,
                             QAction, QFileDialog, QApplication, QLabel, QLineEdit, QGridLayout, QWidget, QPushButton,QMessageBox)
from PyQt5.QtGui import QIcon

class ZDC_file_Calculate_class(QObject):
    trigger_output_Z = PyQt5.QtCore.pyqtSignal(str)

    def __init__(self):
        print('ZDC_file_Calculate_class chushihua ')

        super().__init__()

        self.initUI()

    def initUI(self):
        self.PR_Num_list=[]
        self.ZDC_file_list=[]

    def example(self,text):
        print('example')
        self.trigger_output_Z.emit(text)


    def Xml_Analyse(self,PR_Num_List_io,filename_io,ZDC_file_path):
        temp="开始运算："+filename_io
        self.trigger_output_Z.emit(temp)
        print(("Xml_Analyse已启动"))
        #PR_Num_List=['A8F','J0T','0XF','9HA','6LH','8GU','TBD','TBD','3P1','KD0','KK1','4ZB','9ZX','1ER','B36','ONA','TBD','3ZU','0SM','EL0','EA0','K8G','D60','7L6','G1D','1X1','U10','1G9','I06','G01','N6K','Q2J','3KP','3L4','7P1','3H0','4A3','4D0','QN0','3G0','QE0','KL0','C00','EF1','9AK','2ZU','4LC','8YZ','7Q0','7QA','UF0','QV0','ER6','8RM','8IX','8SK','8G0','9I5','8WG','8TB','8X0','8N6','8Q5','8S7','6T2','9TC','9E4','QQ8','8T3','6K3','8W0','9S5','7X3','7W2','EM1','KA1','QK0','7Y8','4UE','4X3','5ZF','3T2','4L6','4H3','9P3','3C7','3QT','4B3','VL0','7M3','5A2','5MB','6FF','6XQ','5SL','5RQ','QR0','UG1','UK3','0Y5','1E4','0TA','ND0','1D0','UH2','1KU','1LX','1AS','1N3','1T2','1U8','3B4','3CA','3D6','3FU','4AA','3S1','3GE','3U2','6M0','6SC','4E7','5XH','4I3','4GF','4KC','KS0','QH0','QJ1','5D1','6A0','6E3','6KD','6N0','6Q2','7AA','7B2','7E0','7K9','9T0','7N1','1NL','1PA','1SA','1S1','1Z0','2A0','2H5','3M0','2J2','2WA','1TH','8Y1','8ZH','8Z6','9JC','9M0','3X0','ES0','E0A','FB0','AV1','7R7','L0L','U5A','TBD','TBD','F0A','VF0','V0A','0B4','0K0','T9R','1A7','7CM','0EX','TBD','1JA','QG1','0AG','0KA','QI3','2W0','8FA','9Z0','9WC','3B4','B1P','9U0','EV0','0P0']
        #PR_Num_List_io=['A8F','T9D','DP2','G1A','0IC','7L6','7CR','0P2','4ZG','C3R','H3Y','1G8','8IU','8G0','8Q3','8K9','8X0','8WA','8TL','8VM','8N6','9T0','6JE','2JW','3FU','8S7','4KC','4GF','8ZH','QQ9','6T2','5MQ','9S0','N7D','Q4H','3L4','7P1','4V1','5ZC','3Q7','4A3','4D3','3NU','VT3','6DQ','9TL','3LJ','6E3','6D2','9JA','QJ9','2ZH','6Q2','3Y4','1U8','4R3','4I3','1N3','6XQ','5SL','5RQ','4L2','8UW','7T0','77N','QH1','U9F','9ZX','9WT','9Z0','8RM','KH7','4E2','2H0','KS0','EL0','7I4','7AA','1AP','1KU','1ZD','9U1','7X2','KA2','4UE','4X3','9P9','3QW','3ZU','3B4','7K9','8T3','6K3','VL3','QK0','EM0','7Y8','7W0','0MP','E0A','1PA','1D0','1E4','GM2','0K3','J1D','B1P','1S3','C00','G01','2G7','EF0','0FS','UH2','3W4','8GZ','TBD','TBD','3U0','U5J','K8B','0M5','4H1','KD0','QA0','0Y5','QR8','1A7','KK0','FB0','ES4','L0L','3H0','EV6','AV1','FM0','8Z6','1NL','ER6','V0A','0B6','NZ0','0SM','QI3','3TF','3SF','TBD','7U0','1Q2','B36','5D1','3CA','QV0','1SA','1FP','QG1','7E0','8FC','0TD','9M0','1JA']
        #filename_io="C:/Users/SVW/Desktop/PyTemp/v0631J083GB___BCM.xml"
        PR_Num_List =PR_Num_List_io
        filename=filename_io
        #ZDC_file_path = "C:\\users\\SVW\\Desktop\\nms NF\\" # 文件路径

        desktop_path=os.path.join(os.path.expanduser("~"), 'Desktop')
        Programming_file_path = desktop_path+"\\ProjektZ\\Programming_Files\\"
        if(os.path.exists(Programming_file_path)):
            pass
        else:
            os.makedirs(Programming_file_path)

        ZDC_data_ALL_Address=[]#所有地址的ZDC数据
        print(("Xml_Analyse已启动1"))

        ZDC_filename=filename_io+"_ZDC.txt"
        f = open(ZDC_filename, 'w')  # 打开文件
        print('已打开文件')
        print(filename)
        if(os.path.exists(filename)):
            print("exist")
            pass
        else:
            print("▇▇ 错误 ▇▇：ZDC文件不存在")
            Error_info = '▇▇ 错误 ▇▇：ZDC文件:'+filename+'不存在!'
            return (Error_info,'', [], [], [])

        dom = xml.dom.minidom.parse(filename)
        print('已成功解析文件')
        ZDC_root = dom.documentElement
        #print(("Xml_Analyse已启动2"))

        IDENT = ZDC_root.getElementsByTagName("IDENT")
        INFO= ZDC_root.getElementsByTagName("INFO")
        VORSCHRIFT= ZDC_root.getElementsByTagName("VORSCHRIFT")
        ###############################IDENT#################################
        ###############################INFO##################################
        ############################VORSCHRIFT###############################
        DIREKT = VORSCHRIFT[0].getElementsByTagName("DIREKT")
        DIAGN = DIREKT[0].getElementsByTagName("DIAGN")
        ADR =DIAGN[0].getElementsByTagName("ADR")
        Controller_id= str(ADR[0].childNodes[0].data)


        TABELLEN = DIREKT[0].getElementsByTagName("TABELLEN")#所有地址对应的数值块节点
        REFERENZ = TABELLEN[0].getElementsByTagName("REFERENZ")
        TABELLE = TABELLEN[0].getElementsByTagName("TABELLE")#所有地址对应的数值表的列表

        Programming_list_name=[0]
        Programming_list_content=[0]
        Programming_list_content_filename=[]
        Programming_Num = 0  # 参数文件数量

        Programming_list_content_gen2_list=[]
        Programming_list_gen2_content=[]
        na_position_list=[]
        cannot_caculate_position_list=[]
        #print('Xml_Ana启动_3')

        for TABELLE_n in TABELLE:#遍历所有地址
            na_position_list = []
            cannot_caculate_position_list=[]
            MODUS=TABELLE_n.getElementsByTagName("MODUS")[0]
            #print('Xml_Ana_3')
            if str(MODUS.childNodes[0].data)=='K':
                RDIDENTIFIER=TABELLE_n.getElementsByTagName("RDIDENTIFIER")[0]
                #print("Address:",str(RDIDENTIFIER.childNodes[0].data))
                #temp="Address:"+str(RDIDENTIFIER.childNodes[0].data)
                #self.trigger_output_Z.emit(temp)
                #self.trigger_output_Z.emit(str(RDIDENTIFIER.childNodes[0].data))
                KOPF = TABELLE_n.getElementsByTagName("KOPF")
                KOPF_ZDE=TABELLE_n.getElementsByTagName("ZDE")#ZDE按照顺序定义每个Byte怎么分割
                TAB = TABELLE_n.getElementsByTagName("TAB")
                FAM = TAB[0].getElementsByTagName("FAM")#FAM存储的是每个族下面的多个PR号组合
                ZDC_LIST= [0]#初始化ZDC数值的列表
                ZDC_flag=[0]#初始化ZDC标志位的列表
                for n in range(400):
                    ZDC_LIST.append(0)
                    ZDC_flag.append(0)
                data_length=0#ZDC数据的最大长度
                for FAM_n in FAM:#遍历FAM结点
                    FAMNR = FAM_n.getElementsByTagName("FAMNR")#FAMNR定义FAM的名称，3位字母
                    FAMBEZ = FAM_n.getElementsByTagName("FAMBEZ")
                    TEGUE = FAM_n.getElementsByTagName("TEGUE")
                    TEGUE_flag=0#标志位，定义该PR号组合是否能对应本车kopp
                    FAM_flag=0#标志位，定义该FAM是否有对应的PR号组合
                    for TEGUE_n in TEGUE:#遍历PR号组合列表
                        PRNR = TEGUE_n.getElementsByTagName("PRNR")[0]#PRNR定义的是PR号组合的具体内容
                        PRNR_list=(str(PRNR.childNodes[0].data).strip( '+' )).split('+')#按照+分割PR号组合
                        exist=0
                        #print(PRNR_list)
                        for prs in PRNR_list:#遍历分割后PR号
                            if len(prs)>3:#长度大于3说明是多个PR号可选的组合
                                pr_list=prs.split('/')#按照/分割多个PR号可选的组合
                                pr_flag=0
                                for pr in pr_list:
                                    if pr in PR_Num_List:
                                        pr_flag+=1
                                if pr_flag>0:#只要有一个在kopp中，标志位置1
                                    exist=1
                                else:
                                    exist=0
                                    break#跳出循环
                            elif len(prs)==3:
                                if prs in PR_Num_List:
                                    exist=1
                                else:
                                    exist=0
                                    break
                            else:#长度小于3，跳过
                                pass
                        #print('exist:',exist)
                        if exist==1:#如果PR号组合在kopp中，需要计算后面的所有KONTEN中的数据
                            KNOTEN_LIST = TEGUE_n.getElementsByTagName("KNOTEN")
                            FAM_flag=1#将FAM标志位置1，说明该FAM已经有对应的PR号组合
                            TEGUE_flag+=1#PR号组合标志位加1
                            if TEGUE_flag>1:
                                print('▇▇ 错误 ▇▇：一个FAM对应多个PR号组合！','FAM:',FAMNR[0].childNodes[0].data,'  PR:',PRNR_list,'  次数:',TEGUE_flag)
                                #temp="▇▇ 错误 ▇▇：一个FAM对应多个PR号组合！FAM:"+str(FAMNR[0].childNodes[0].data)+" PR:"+str(PRNR_list)+" 次数："+str(TEGUE_flag)
                                temp="▇▇ 错误 ▇▇：一个FAM对应多个PR号组合！FAM:"+str(FAMNR[0].childNodes[0].data)+"次数："+str(TEGUE_flag)
                                self.trigger_output_Z.emit(temp)
                            else:
                                for kn in KNOTEN_LIST:#遍历KONTEN列表
                                    STELLE = kn.getElementsByTagName("STELLE")[0]
                                    if data_length < int(str(STELLE.childNodes[0].data), 16):
                                        data_length = int(str(STELLE.childNodes[0].data), 16)
                                    LSB = kn.getElementsByTagName("LSB")[0]
                                    #print('----------------------Wert:',('WERT' in kn.childNodes))#后还需要判断wert是否在node节点里面
                                    WERT_exist=0
                                    DATEN_exist=0
                                    for i in kn.childNodes:
                                        if('WERT'==str(i.nodeName)):
                                            WERT_exist=1
                                        elif('DATEINAME'==str(i.nodeName)):
                                            DATEN_exist=1
                                    #print(WERT_exist,DATEN_exist)
                                    if(WERT_exist==1):
                                        WERT = kn.getElementsByTagName("WERT")[0]#WERT记录的是该位的数据
                                        if str(STELLE.childNodes[0].data) == 'n/a':#排除n/a行
                                            self.trigger_output_Z.emit("n/a:passed!")
                                            print(("n/a:passed!"))
                                        elif(len(str(STELLE.childNodes[0].data))<4):
                                            #print('计算中')
                                            ZDC_LIST[int(str(STELLE.childNodes[0].data),16)]=ZDC_LIST[int(str(STELLE.childNodes[0].data),16)]+(int(str(WERT.childNodes[0].data), 16))*2**(int(str(LSB.childNodes[0].data)))#将该行数据加入到对应的ZDC_LIST中，转换成十进制
                                            for KOPF_ZDBYTE_n in KOPF_ZDE[int(str(STELLE.childNodes[0].data),16)-1].getElementsByTagName("ZDBYTE"):
                                                if int(KOPF_ZDBYTE_n.getElementsByTagName("ZLSB")[0].childNodes[0].data)==int(LSB.childNodes[0].data):#LSB数值相等说明找到了对应的ZDBYTE列表
                                                    LSB_n=int(LSB.childNodes[0].data)#数值转换成int
                                                    while LSB_n<=(int(KOPF_ZDBYTE_n.getElementsByTagName("ZMSB")[0].childNodes[0].data)):#从LSB循环到MSB
                                                        ZDC_flag[int(str(STELLE.childNodes[0].data),16)]=ZDC_flag[int(str(STELLE.childNodes[0].data),16)]+2**(LSB_n)
                                                        LSB_n=LSB_n+1
                                        else:
                                            pass
                                    elif(DATEN_exist==1):
                                        DATEINAME = kn.getElementsByTagName("DATEINAME")[0]
                                        print("二代参数P：", str(DATEINAME.childNodes[0].data))
                                        temp="\t包含二代参数："+ str(DATEINAME.childNodes[0].data)
                                        self.trigger_output_Z.emit(temp)

                                        Programming_list_content_gen2_list.append(str(DATEINAME.childNodes[0].data))
                                        for KOPF_ZDBYTE_n in KOPF_ZDE[
                                            int(str(STELLE.childNodes[0].data), 16) - 1].getElementsByTagName(
                                                "ZDBYTE"):  # 遍历KOPF下面对应的ZDE列表的ZDBYTE，int(str(STELLE.childNodes[0].data),16)-1就是ZDE列表的对应位，寻找当前位置对应的MSB数值，也就是高位从几开始
                                            if int(KOPF_ZDBYTE_n.getElementsByTagName("ZLSB")[0].childNodes[0].data) == int(
                                                    LSB.childNodes[0].data):  # LSB数值相等说明找到了对应的ZDBYTE列表
                                                LSB_n = int(LSB.childNodes[0].data)  # 数值转换成int
                                                while LSB_n <= (int(KOPF_ZDBYTE_n.getElementsByTagName("ZMSB")[0].childNodes[
                                                                        0].data)):  # 从LSB循环到MSB
                                                    ZDC_flag[int(str(STELLE.childNodes[0].data), 16)] = ZDC_flag[int(
                                                        str(STELLE.childNodes[0].data), 16)] + 2 ** (LSB_n)  # 标志位加上2的幂级数，用来确定该Byte是否完整
                                                    LSB_n = LSB_n + 1
                                    else:
                                        #ZDC_LIST[int(str(STELLE.childNodes[0].data), 16)]='NA'
                                        na_position_list.append(int(str(STELLE.childNodes[0].data), 16))
                                        temp=str(RDIDENTIFIER.childNodes[0].data)+"有不包含WERT/DATEN结点的Byte，请检查文件格式."
                                        #self.trigger_output_Z.emit(temp)
                                        #print(temp)
                    if FAM_flag==0:
                        print('该FAM：',FAMNR[0].childNodes[0].data,' 没有对应的PR号组合, 影响下列位置的数据：')
                        temp="▇▇ 错误 ▇▇： 地址： "+str(RDIDENTIFIER.childNodes[0].data)+"  的PR Family："+str(FAMNR[0].childNodes[0].data)+" 没有对应的PR号组合, 影响下列位置的数据："
                        self.trigger_output_Z.emit(temp)
                        TEGUE_already_outputed_flag=0
                        for TEGUE_n in TEGUE:
                            if TEGUE_already_outputed_flag==0:
                                PRNR = TEGUE_n.getElementsByTagName("PRNR")[0]#由于每个组合对应的位置都是一样的，这里直接选取第一组
                                KNOTEN_LIST = TEGUE_n.getElementsByTagName("KNOTEN")#获取KNOTEN列表
                                for kn in KNOTEN_LIST:  # 遍历KONTEN列表
                                    WERT_exist = 0  # 是否存在WERT结点的标志位
                                    for i in kn.childNodes:
                                        if ('WERT' == str(i.nodeName)):  # 是否有结点名称是WERT
                                            WERT_exist = 1
                                    if (WERT_exist == 1):  # 判断是否包含WERT结点
                                        STELLE = kn.getElementsByTagName("STELLE")[0]  # STELLE存储的是Byte的位置
                                        if data_length < int(str(STELLE.childNodes[0].data), 16):  # 如果Byte位置大于当前记录的数据长度，更改数据长度到该Byte位置
                                            data_length = int(str(STELLE.childNodes[0].data), 16)
                                        LSB = kn.getElementsByTagName("LSB")[0]  # LSB记录的是低位从几开始
                                        WERT = kn.getElementsByTagName("WERT")[0]  # WERT记录的是该位的数据
                                        if str(STELLE.childNodes[0].data) == 'n/a':  # 排除n/a行
                                            self.trigger_output_Z.emit("n/a:passed!")
                                            print("n/a:passed!")
                                        else:
                                            for KOPF_ZDBYTE_n in KOPF_ZDE[int(str(STELLE.childNodes[0].data), 16) - 1].getElementsByTagName(
                                                    "ZDBYTE"):  # 遍历KOPF下面对应的ZDE列表的ZDBYTE，int(str(STELLE.childNodes[0].data),16)-1就是ZDE列表的对应位，寻找当前位置对应的MSB数值，也就是高位从几开始
                                                if int(KOPF_ZDBYTE_n.getElementsByTagName("ZLSB")[0].childNodes[0].data) == int(LSB.childNodes[0].data):  # LSB数值相等说明找到了对应的ZDBYTE列表
                                                    #print('Byte:', str(STELLE.childNodes[0].data))#输出Byte位置
                                                    temp="Byte"+ str(int(str(STELLE.childNodes[0].data),16))+"的 "#输出Byte位置
                                                    #temp="Byte:"+ int(str(STELLE.childNodes[0].data),16)+"的 "#输出Byte位置
                                                    cannot_caculate_position_list.append(int(str(STELLE.childNodes[0].data), 16))
                                                    LSB_n = int(LSB.childNodes[0].data)#数值转换成int
                                                    while LSB_n <= (int(KOPF_ZDBYTE_n.getElementsByTagName("ZMSB")[0].childNodes[0].data)):#从LSB循环到MSB
                                                        ZDC_flag[int(str(STELLE.childNodes[0].data), 16)] = ZDC_flag[int(
                                                            str(STELLE.childNodes[0].data), 16)] + 2 ** (LSB_n)
                                                        LSB_n = LSB_n + 1
                                                    if (int(KOPF_ZDBYTE_n.getElementsByTagName("ZLSB")[0].childNodes[0].data))==(int(KOPF_ZDBYTE_n.getElementsByTagName("ZMSB")[0].childNodes[0].data)):
                                                        temp=temp+"bit"+str(int(KOPF_ZDBYTE_n.getElementsByTagName("ZMSB")[0].childNodes[0].data))
                                                    else:# 输出从低位到高位
                                                        temp = temp +"bit"+str(int(KOPF_ZDBYTE_n.getElementsByTagName("ZLSB")[0].childNodes[0].data))+"到 bit"+str(int(KOPF_ZDBYTE_n.getElementsByTagName("ZMSB")[0].childNodes[0].data))
                                            TEGUE_already_outputed_flag=1
                                            self.trigger_output_Z.emit(temp)
                                            print(temp)

                #print('Adress:',str(RDIDENTIFIER.childNodes[0].data),list(map(hex,ZDC_LIST[1:data_length+1])))#输出计算出来的该地址的ZDC数据
                ZDC_index=1
                f.write(str(RDIDENTIFIER.childNodes[0].data))#写数据地址到文件中
                f.write(':\n')
                while ZDC_index<data_length+1:#遍历ZDC_LIST
                    HEX_ZDC =str(hex(int(ZDC_LIST[ZDC_index])))#转换成str
                    ZDC_LIST[ZDC_index] = HEX_ZDC[2:].zfill(2).upper()#转换成两位大写的格式
                    f.write(str(ZDC_LIST[ZDC_index]).zfill(2).upper())  # 数据写入到文件中
                    f.write(' ')
                    ZDC_index+=1
                f.write('\n')#换行
                ZDC_LIST[0]=str(RDIDENTIFIER.childNodes[0].data)
                if len(na_position_list)>0:
                    for na_position in na_position_list:
                        ZDC_LIST[int(na_position)]='NA'

                if len(cannot_caculate_position_list)>0:
                    for cannot_caculate_position in cannot_caculate_position_list:
                        ZDC_LIST[int(cannot_caculate_position)]='**'

                ZDC_data_ALL_Address.append(ZDC_LIST[0:data_length+1])
                print('Adress:',str(RDIDENTIFIER.childNodes[0].data),ZDC_LIST[1:data_length+1])#输出地址数值
                #print('Adress:',str(RDIDENTIFIER.childNodes[0].data),ZDC_flag[1:data_length+1])#输出该地址的标志位
                ZDC_flag_n=1#起始位
                while ZDC_flag_n < (data_length+1):#循环到数据最长
                    if ZDC_flag[ZDC_flag_n]>255:#大于255说明数据有重复定义，后续有必要的话加入自动查询
                        print('▇▇ 错误 ▇▇： 地址：',str(RDIDENTIFIER.childNodes[0].data),'的 Byte',ZDC_flag_n,'有重复定义,请检查！')
                        temp="▇▇ 错误 ▇▇： 地址："+str(RDIDENTIFIER.childNodes[0].data)+"的 Byte"+str(ZDC_flag_n)+"有重复定义,请检查！"
                        self.trigger_output_Z.emit(temp)
                        ZDC_flag_n=ZDC_flag_n+1
                        continue
                    elif ZDC_flag[ZDC_flag_n]<255:
                        if(ZDC_LIST[ZDC_flag_n]=='NA'):
                            ZDC_flag_n = ZDC_flag_n + 1
                        else:
                            bin_str = str(bin(ZDC_flag[ZDC_flag_n])).replace('ob', '')#转换成str
                            bin_str_std8=bin_str[2:].zfill(8)#去出头两位ob后自动补0填充到8位
                            #print(bin_str_std8,':')#输出二进制，bit几是0就是该bit没有数据
                            for i in range(0,len(bin_str_std8)):#遍历这8位
                                if bin_str_std8[i]=='0':#如果是0，输出错误
                                    print('▇▇ 错误 ▇▇：地址:',str(RDIDENTIFIER.childNodes[0].data),'的 Byte',ZDC_flag_n,', bit',str(7-i),'没有数据定义,请检查！')
                                    temp="▇▇ 错误 ▇▇：地址:"+ str(RDIDENTIFIER.childNodes[0].data)+"的 Byte"+ str(ZDC_flag_n)+" bit"+ str(7-i)+"没有数据定义,请检查！"
                                    self.trigger_output_Z.emit(temp)
                            ZDC_flag_n = ZDC_flag_n + 1
                    else:
                        ZDC_flag_n = ZDC_flag_n + 1
                        pass
            elif str(MODUS.childNodes[0].data)=='P':
                MODUSTEIL = TABELLE_n.getElementsByTagName("MODUSTEIL")[0]
                #print('一代参数:',str(MODUSTEIL.childNodes[0].data))
                TAB = TABELLE_n.getElementsByTagName("TAB")[0]
                FAM = TAB.getElementsByTagName("FAM")
                for FAM_n in FAM:
                    FAMNR = FAM_n.getElementsByTagName("FAMNR")  # FAMNR定义FAM的名称，3位字母
                    TEGUE = FAM_n.getElementsByTagName("TEGUE")
                    TEGUE_flag = 0  # 标志位，定义该PR号组合是否能对应本车kopp
                    FAM_flag = 0  # 标志位，定义该FAM是否有对应的PR号组合
                    for TEGUE_n in TEGUE:  # 遍历PR号组合列表
                        PRNR = TEGUE_n.getElementsByTagName("PRNR")[0] # PRNR定义的是PR号组合的具体内容
                        PRNR_list = (str(PRNR.childNodes[0].data).strip( '+' )).split('+')  # 按照+分割PR号组合
                        exist = 0  # 标志位，判断PR号组合是否在本车的kopp定义中
                        # print(PRNR_list)
                        for prs in PRNR_list:  # 遍历分割后PR号
                            if len(prs) > 3:  # 长度大于3说明是多个PR号可选的组合
                                pr_list = prs.split('/')  # 按照/分割多个PR号可选的组合
                                pr_flag = 0
                                for pr in pr_list:
                                    if pr in PR_Num_List:
                                        pr_flag += 1
                                if pr_flag > 0:  # 只要有一个在kopp中，标志位置1
                                    exist = 1
                                else:
                                    exist = 0  # PR号都不在kopp中，标志位置0
                                    break  # 跳出循环
                            else:  # 长度等于3 说明只有一个PR号
                                if prs in PR_Num_List:  # 该PR号在kopp中，标志位置1
                                    exist = 1
                                else:
                                    exist = 0  # 该PR号不在kopp中，标志位置0
                                    break
                        #print('exist:',exist)
                        if exist == 1:  # 如果PR号组合在kopp中，需要记录后面的所有KONTEN中的参数
                            KNOTEN_LIST = TEGUE_n.getElementsByTagName("KNOTEN")  # 获取KNOTEN，该列表定义的是有效数据的：位置和具体数值
                            FAM_flag = 1  # 将FAM标志位置1，说明该FAM已经有对应的PR号组合
                            TEGUE_flag += 1  # PR号组合标志位加1
                            if TEGUE_flag > 1:
                                print('▇▇ 错误 ▇▇：一个FAM对应多个PR号组合！', 'FAM:', FAMNR[0].childNodes[0].data, '  PR:', PRNR_list, '  次数:', TEGUE_flag)
                                temp="▇▇ 错误 ▇▇：一个FAM对应多个PR号组合！ FAM:"+str( FAMNR[0].childNodes[0].data)+ " PR:"
                                for PRNR_list_n in PRNR_list:
                                    temp=temp+str(PRNR_list_n)
                                temp=temp+ " 次数:" + str(TEGUE_flag)
                                self.trigger_output_Z.emit(temp)
                            else:
                                for kn in KNOTEN_LIST:  # 遍历KONTEN列表
                                    DATEN_NAME_exist = 0
                                    for i in kn.childNodes:
                                        if ('DATEN-NAME' == str(i.nodeName)):  # 是否有结点名称是DATEN_NAME
                                            DATEN_NAME_exist = 1
                                            #print(str(i.nodeName))
                                        else:  # 是否有结点名称是DATEINAME
                                            pass
                                    #print('DATEN_NAME_exist:',DATEN_NAME_exist)
                                    if (DATEN_NAME_exist == 1):  # 判断是否包含DATEN_NAME结点
                                        DATEN_NAME = kn.getElementsByTagName("DATEN-NAME")[0]
                                        if str(DATEN_NAME.childNodes[0].data) == 'n/a':  # 排除n/a行
                                            self.trigger_output_Z.emit("n/a:passed!")
                                            print("n/a:passed!")
                                        elif (len(str(DATEN_NAME.childNodes[0].data)) > 6):
                                            #print('一代参数计算中')
                                            Programming_list_name.append(str(MODUSTEIL.childNodes[0].data))
                                            Programming_list_content.append(str(DATEN_NAME.childNodes[0].data))
                                            Programming_Num=Programming_Num+1
                                        else:
                                            pass

                                    else:
                                        #print('该地址：',str(MODUSTEIL.childNodes[0].data),'的FAM：',str(FAMNR[0].childNodes[0].data),'不包含当前配置的PR号组合，请检查.')
                                        temp="▇ 注意 ▇ 该地址："+str(MODUSTEIL.childNodes[0].data)+" 的FAM："+str(FAMNR[0].childNodes[0].data)+"不包含当前配置的PR号组合，请检查."
                                        self.trigger_output_Z.emit(temp)
                                        print(temp)
                #print('list',Programming_list_content,Programming_list_name)
            else:
                pass
        if(len(Programming_list_content) > 1):
            temp="\t共"+str(Programming_Num)+"个一代参数: "
            for Programming_list_content_n in Programming_list_content[1:]:
                temp = temp +str(Programming_list_content_n)+'; '
            self.trigger_output_Z.emit(temp.strip(' ').strip(';'))
            temp="\t参数名称: "
            for Programming_list_name_n in Programming_list_name[1:]:
                temp = temp +str(Programming_list_name_n)+ '; '
            self.trigger_output_Z.emit(temp.strip(' ').strip(';'))
            DATENBEREICHE = DIREKT[0].getElementsByTagName("DATENBEREICHE")[0]#所有的DATENBEREICHE节点列表
            DATENBEREICH = DATENBEREICHE.getElementsByTagName("DATENBEREICH")
            for DATENBEREICH_n in DATENBEREICH:
                DATEN_NAME = DATENBEREICH_n.getElementsByTagName("DATEN-NAME")[0]  # 所有的DATEN-NAME节点列表
                if str(DATEN_NAME.childNodes[0].data) in Programming_list_content:
                    START_ADR = DATENBEREICH_n.getElementsByTagName("START-ADR")[0]
                    GROESSE_DEKOMPRIMIERT = DATENBEREICH_n.getElementsByTagName("GROESSE-DEKOMPRIMIERT")[0]
                    DATEN = DATENBEREICH_n.getElementsByTagName("DATEN")[0]  # 所有的DATEN节点列表
                    filename=str(DATEN_NAME.childNodes[0].data)
                    #print(str(DATEN.childNodes[0].data))
                    #Programming_file_path="C:/Users/SVW/Desktop/ZDC_Files/"#文件路径
                    name = Programming_file_path +Controller_id+ '-gen1-' + str(DATEN_NAME.childNodes[0].data) + '.xml'
                    Programming_list_content_filename.append(name)
                    p = open(name, 'w')
                    p.writelines(['<?xml version="1.0" encoding="ISO-8859-1"?>\n', \
                                  '<MESSAGE DTD="XMLMSG" VERSION="1.1">\n', \
                                  '  <RESULT>\n', \
                                  '    <RESPONSE NAME="GetParametrizeData" DTD="RepairHints" VERSION="1.4.0.0" ID="0">\n', \
                                  '      <DATA>\n', \
                                  '        <PARAMETER_DATA DIAGNOSTIC_ADDRESS="0x'])
                    p.write(Controller_id)
                    p.write('" START_ADDRESS="')
                    p.write(str(START_ADR.childNodes[0].data).split('/')[0].lstrip('0'))
                    p.write('" PR_IDX="0" ZDC_NAME="-----------" ZDC_VERSION="----" LOGIN="')
                    if operator.eq(str(Controller_id), '09'):
                        p.write('31347')
                    else:
                        p.write('20103')
                    p.write('" LOGIN_IND=""> \n')
                    p.write(str(DATEN.childNodes[0].data))
                    p.writelines(['\n</PARAMETER_DATA>\n', \
                                  '      </DATA>\n', \
                                  '    </RESPONSE>\n', \
                                  '  </RESULT>\n', \
                                  '</MESSAGE>\n'])
                    print(name, ' generated, total', (len(str(DATEN.childNodes[0].data))) / 3, ' Bytes.')
        print('开始处理二代参数列表',Programming_list_content_gen2_list)
        if(len(Programming_list_content_gen2_list) > 0):  # 处理二代参数
            print('len(Programming_list_content_gen2_list)',len(Programming_list_content_gen2_list))
            if (len(ZDC_file_path) < 4):
                Error_info = 'ZDC文件所在路径错误'
                return (Error_info,'',[], [], [])
            print('开始遍历')
            for root, subdirs, files in os.walk(ZDC_file_path):
                for filepath in files:  # 遍历
                    if filepath in Programming_list_content_gen2_list:  # 如果ZDC列表中存在
                        # print('已找到:',filepath,'.')
                        Programming_list_content_gen2_list[
                            Programming_list_content_gen2_list.index(filepath)] = '0'
                        oldfile = os.path.join(root, filepath)  # 老文件地址
                        # newfile = desktop_path+"\\PyQt5" + filepath#新文件地址
                        newfile = Programming_file_path + filepath  # 新文件地址
                        shutil.copyfile(oldfile, newfile)
                        name = Programming_file_path + Controller_id + '-gen2-' + filepath + '.xml'
                        p = open(name, 'w')
                        if Controller_id=='75':
                            p.writelines(['<?xml version="1.0" encoding="UTF-8"?>\n', \
                                          '<MESSAGE DTD="XMLMSG" VERSION="1.1">\n', \
                                          '  <RESULT>\n', \
                                          '    <RESPONSE NAME="GetParametrizeData" DTD="RepairHints" VERSION="1.4.7.0" ID="0">\n', \
                                          '      <DATA>\n', \
                                          '        <REQUEST_ID>\n', \
                                          '        <INFORMATION>\n', \
                                          '          <CODE/>\n', \
                                          '        </INFORMATION>\n', \
                                          '        <PARAMETER_DATA DIAGNOSTIC_ADDRESS="0x00'])
                            p.write(Controller_id)
                            p.write('" DSD_TYPE="2" FILENAME="')
                            p.write(filepath)  # 写入起始地址并去除前面的0
                            p.write('" SESSIONNAME="SESD')
                            p.write(filepath[2:-4])
                            p.writelines([
                                             '" START_ADDRESS="" PR_IDX="" ZDC_NAME="V42000999AA" ZDC_VERSION="0001" LOGIN="20103" LOGIN_IND=""/>\n', \
                                             '      </DATA>\n', \
                                             '    </RESPONSE>\n', \
                                             '  </RESULT>\n', \
                                             '</MESSAGE>\n'])
                        else:
                            p.writelines(['<?xml version="1.0" encoding="UTF-8"?>\n', \
                                          '<MESSAGE DTD="XMLMSG" VERSION="1.1">\n', \
                                          '  <RESULT>\n', \
                                          '    <RESPONSE NAME="GetParametrizeData" DTD="RepairHints" VERSION="1.4.7.0" ID="0">\n', \
                                          '      <DATA>\n', \
                                          '        <REQUEST_ID>47251711</REQUEST_ID>\n', \
                                          '        <PARAMETER_DATA DIAGNOSTIC_ADDRESS="0x'])
                            p.write(Controller_id)
                            p.write('" DSD_TYPE="2" FILENAME="')
                            p.write(filepath)
                            p.writelines([
                                '" SESSIONNAME="" START_ADDRESS="" PR_IDX="" ZDC_NAME="-----------" ZDC_VERSION="----" LOGIN="20103" LOGIN_IND="" /> \n', \
                                '        <INFORMATION>\n', \
                                '          <CODE />\n', \
                                '        </INFORMATION>\n', \
                                '      </DATA>\n', \
                                '    </RESPONSE>\n', \
                                '  </RESULT>\n', \
                                '</MESSAGE>\n'])
                        print(name, ' gen2 generated.')
                        Programming_list_gen2_content.append(name)
            for file in Programming_list_content_gen2_list:
                if (file != '0' and len(str(file)) > 8):
                    self.trigger_output_Z.emit("该二代参数文件在指定文件夹中未找到：" + file+", 请检查. ")

        f.close()



        Error_info = '0'
        return(Error_info,Controller_id,ZDC_data_ALL_Address,Programming_list_content_filename,Programming_list_gen2_content)


