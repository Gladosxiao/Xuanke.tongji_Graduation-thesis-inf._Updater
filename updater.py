# -*- coding: utf-8 -*-
"""
Created on Wed Mar 18 10:53:45 2020

@author: GlaDo
"""

from selenium import webdriver

from openpyxl import load_workbook
import random

NoListin = []
NameListin = []
NameList2in = []
LAIYUANListin = []
LEIXINGListin = []
XINGZHIListin = []
LAOSHIListin = []


path = 'D://V6.0.xlsx' #文件路径

#读文件
wb=load_workbook(path)#读整个文件
#获取sheet：
Sheet1 = wb.get_sheet_by_name('Sheet1')   #通过表名获取  sheet1这张表
NoList = Sheet1['A']# 获取第二列
NameList = Sheet1['C']# 获取第二列
NameList2 = Sheet1['B']# 获取第二列
LAIYUANList = Sheet1['F']# 获取第二列
LEIXINGList = Sheet1['E']# 获取第二列
XINGZHIList = Sheet1['G']# 获取第二列
LAOSHIList = Sheet1['D']

for i in range(len(NoList)):
   NoListin.append(NoList[i].value) 
for i in range(len(NameList)):
   NameListin.append(NameList[i].value) 
for i in range(len(NameList2)):
    NameList2in.append(NameList2[i].value) 
for i in range(len(LAIYUANList)):
   LAIYUANListin.append(LAIYUANList[i].value) 
for i in range(len(NameList)):
   LEIXINGListin.append(LEIXINGList[i].value) 
for i in range(len(XINGZHIList)):
   XINGZHIListin.append(XINGZHIList[i].value) 
for i in range(len(LAOSHIList)):
   LAOSHIListin.append(LAOSHIList[i].value) 
    
driver = webdriver.Ie()    # 创建Chrome对象.
# 操作这个对象.
driver.get('http://xuanke.tongji.edu.cn/tj_login/frame.jsp')     # get方式访问百度.
#driver.get('http://www.baidu.com')     # get方式访问百度.
start = input("inputID = ")

driver.switch_to_frame("detailfrm")
#driver.find_element_by_name("KCMC").send_keys("159")
startno = NoListin.index(int(start))
j=startno
while 1:
        driver.find_element_by_name("KCMC").send_keys(NameListin[j])
        driver.find_element_by_name("jsgh").send_keys(LAOSHIListin[j])   
        driver.find_element_by_name("BYSJDD").send_keys("校内")   
        driver.find_element_by_name("KTLX").find_elements_by_tag_name("option")[1].click()
        if LAIYUANListin[j] == "国家科研":
            driver.find_element_by_name("KTLY").find_elements_by_tag_name("option")[1].click()
        elif LAIYUANListin[j] == "省部级科研":
            driver.find_element_by_name("KTLY").find_elements_by_tag_name("option")[2].click()  
        elif LAIYUANListin[j] == "校级项目":
            driver.find_element_by_name("KTLY").find_elements_by_tag_name("option")[3].click()  
        elif LAIYUANListin[j] == "校外写作":
            driver.find_element_by_name("KTLY").find_elements_by_tag_name("option")[4].click()  
        elif LAIYUANListin[j] == "学生创新项目":
            driver.find_element_by_name("KTLY").find_elements_by_tag_name("option")[5].click()  
        elif LAIYUANListin[j] == "自拟题目":
            driver.find_element_by_name("KTLY").find_elements_by_tag_name("option")[6].click()  
        elif LAIYUANListin[j] == "其他":
            driver.find_element_by_name("KTLY").find_elements_by_tag_name("option")[7].click() 
           
              
        if XINGZHIListin[j] == "结合科研":
            driver.find_element_by_name("KTXZ2").find_elements_by_tag_name("option")[2].click()
        elif XINGZHIListin[j] == "结合实际":  
            driver.find_element_by_name("KTXZ2").find_elements_by_tag_name("option")[1].click()


        j=j+1
        input("下一个是："+NameList2in[j])

        
        
 
