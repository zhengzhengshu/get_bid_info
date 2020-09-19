# -*- coding: utf-8 -*-
import logging
import os
import time
import sys
import xlwt
from log import logger
from splinter import Browser
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import threading

options = Options()
executable_path = '/Users/aaronbrant/Downloads/chromedriver'
if sys.platform == 'win32':
    options.binary_path = './chrome/chrome.exe'
    options.add_argument("--user-data-dir=chrome/UserData")
    executable_path = './chrome/chromedriver.exe'

options.add_experimental_option(
    "excludeSwitches", ["enable-automation"]
)  # 绕过js检测

file_num = 0

def do_auto_task(begin_num, end_num, file_num):
    global browser 
    url_template = "https://ec.mcc.com.cn/common/modal.jsp?url=/bidAction.do@actionType=toBidB2bView!bidid={}"
    #open_excel
    f = xlwt.Workbook()
    sheet1 = f.add_sheet("投标数据",cell_overwrite_ok=True)
    row0 = ["数字编号","招标编码","招标编号","招标名称","投标开标时间","供应商名称"]
    for i in range(0,len(row0)):
        sheet1.write(0,i,row0[i])

    row_index = 1
    logger.info("正在写入excel...")
    while begin_num <= end_num:
        num_text = f"{begin_num :06d}"
        url =  url_template.format(num_text)
        logger.info(url)
        browser.visit(url)
        browser.driver.switch_to.frame(0)
        browser.is_element_present_by_css(".stdColumnLinkContent",wait_time= 10)
        #找到对应的元素
        html = browser.html
        row = get_data_from_source(num_text,html)
        logger.info(f"获取招标信息: {','.join(row)}")
        #write_to_excel
        for col_index,col in enumerate(row):
            sheet1.write(row_index,col_index,col)

        f.save(f"招标数据{time.strftime('%Y%m%d')}_{ file_num }.xls")
        begin_num += 1
        row_index += 1

    #close_excel
    sheet1.col(0).width = 10*256
    sheet1.col(1).width = 15*256
    sheet1.col(2).width = 50*256
    sheet1.col(3).width = 50*256
    sheet1.col(4).width = 30*256
    sheet1.col(5).width = 30*256
    f.save(f"招标数据{time.strftime('%Y%m%d')}_{ file_num }.xls")
    logger.info("写入excel完成")


def get_data_from_source(begin_num,html):

    zbbm,zbbh,zbmc,tpkbsj,gysmc = "","","","",""
    soup = BeautifulSoup(html,"html.parser")

    #招标编码
    try:
        zbbm = soup.body.find_all('td',attrs={'colspan':"1", 'rowspan':"1",'class':"stdColumnLinkContent"})[0].text
    except:
        zbbm = "未找到"
        pass
    #招标编号
    try:
        zbbh = soup.body.find_all('td',attrs={'colspan':"1", 'rowspan':"1",'class':"StdColumnContent"})[0].text
    except:
        zbbh = "未找到"
    #招标名称
    try:
        zbmc = soup.body.find_all('td',attrs={'colspan':"1", 'rowspan':"1",'class':"StdColumnContent"})[1].text
    except:
        zbmc = "未找到"
    #投标开标时间
    try:
        tpkbsj = soup.body.find_all('td',attrs={'colspan':"1", 'rowspan':"1",'class':"StdColumnContent"})[4].text
    except:
        tpkbsj = "未找到"
    #供应商名称，如果有。
    try:
        gysmc = soup.body.find_all('td',attrs={'colspan':"1", 'rowspan':"1","align":"left",'class':"stdFieldInput"})[0].text
    except:
        gysmc = ""
    return [begin_num,zbbm,zbbh,zbmc,tpkbsj,gysmc]

def main():
    global file_num,browser
    while True:
        try:
            file_num += 1  
            begin_num = int(input("请输入起始编号 "))
            end_num = int(input("请输入结束编号 "))
            logger.info(f"起始编号：{begin_num}, 结束编号：{end_num}, 第{ file_num }次爬取")
            do_auto_task(begin_num,end_num,file_num)
        except KeyboardInterrupt:
            print("本次爬取被终止")
            break;


if __name__ == "__main__":
    url1 = "https://ec.mcc.com.cn/logonAction.do?source=validateCode"
    print("请在即将打开的浏览器上手动登陆，登陆成功后，在此处按回车确认")
    #browser = Browser(driver_name="chrome", executable_path=executable_path, options=options)
    browser = Browser(driver_name="chrome", executable_path=executable_path, options=options)
    browser.visit(url1)
    time.sleep(5)
    ok = input("如果已经成功登陆，请回车 ")
    while True:
        try:
            file_num += 1  
            begin_num = int(input("请输入起始编号 "))
            end_num = int(input("请输入结束编号 "))
            logger.info(f"起始编号：{begin_num}, 结束编号：{end_num}, 第{ file_num }次爬取")
            do_auto_task(begin_num,end_num,file_num)
        except:
            print("浏览器已经退出，需要重新登录")
            ok = input("如果已经成功登陆，请回车 ")
            browser = Browser(driver_name="chrome", executable_path=executable_path, options=options)
            browser.visit(url1)
            time.sleep(5)
    browser.quit()
    logger.info("程序运行结束")
