#coding:utf-8

from qqQunManage import *
import json
from time import sleep
from qunDataBeautifulSoup import *
from bs4 import BeautifulStoneSoup
import pandas as pd
import math

#将群管理地址和QQ群号进行拼接,得到指定的群的网址
def getQunNumberUrl(qunUrl, qunNumberlist):
    res = []
    for i in qunNumberlist:
        tmp = qunUrl + i
        res.append(tmp)
    return res

def main():
    getWeb = qqQunManage()
    data = getJsonData('pathData.json')

    dictXpath, dictNum, dictUrl  = data[0], data[1], data[2]
    loginXpath, groupManageXpath = dictXpath["login"], dictXpath["groupManage"]
    qunNumList = dictNum["qqNumber"]
    QQqunURL, QQqunURLnumber = dictUrl["QQqunUrl"], dictUrl["QQqunUrlNumber"]

    getWeb.getUrl(QQqunURL)                                                 #进入首页
    getWeb.webVisitXpath(loginXpath)                                        #登录页面

    sleep(10)                                                               #等待扫码
    getWeb.webVisitXpath(groupManageXpath)                                  #进入QQ群

    url = getQunNumberUrl(QQqunURLnumber, qunNumList)                       #拼接网址,获取访问QQ群的URL网址，
    sleep(10)                                                                #休眠5秒，等待进入QQ群完成

    html_name_lst = []                                                      #存放下载好的网页路径
    #print(url)
    try:
        for i in range(len(url)):
            # getWeb.getUrl(url[i])
            sleep(5)
            # getWeb.closeWindow()
            js = 'window.open("'+ str(url[i]) +'");'                            #利用JS打开新的网页
            getWeb.executeScript(js)
            print(getWeb.windowHandle())
    except:
        pass
        
    all_handles = getWeb.windowHandle()[1:]
    print(all_handles)

    html_lst = []                                                                       #存放所有爬取的html页面

    try:
        for i in range(len(all_handles)):                                                   #在已经打开的页面上逐个跳转
            getWeb.switchWindow(all_handles[i])                                             #进行跳转
            #html_Name = qunNumList[i] + "_source_code.html"
            sleep(1)
            a = getWeb.webVisitXpathNoclick('//*[@id="groupMemberNum"]')                      #selenium 通过Xpath 获取QQ群 人数
            a.click()
            num = a.text
            count = math.ceil(int(num) / 10)                                                         #计算下滑次数
            print(count)
            if count != 0:
                for i in range(count + 1):
                    # js="var q=document.documentElement.scrollTop=10000"                              #通过JS 下滑页面
                    # getWeb.executeScript(js)
                    getWeb.sendKey()
                    sleep(1)
            #getWeb.getSourceCode(html_Name)
            html_lst.append(getWeb.getHtml())
    except:
        pass

    print("网页源代码已下载完成")
    #print(html_name_lst)

    output_cols = ["成员", "群昵称", "QQ号", "性别", "Q龄", "入群时间", "最后发言"]
    writer =pd.ExcelWriter('resExcel.xlsx')                  # pylint: disable=abstract-class-instantiated
    for i in range(len(html_lst)):
        two_arry = getQQdata(html_lst[i], "selector.json")
        #print(two_arry)
        try:
            resDict = dict(zip(output_cols, two_arry))
            resExcel = pd.DataFrame(resDict)
            resExcel.to_excel(writer, sheet_name=qunNumList[i],index=False)
        except ValueError as e:
            print(e)
    writer.save() 
    writer.close()

    getWeb.quitChrome()

if __name__ == "__main__":
    main()