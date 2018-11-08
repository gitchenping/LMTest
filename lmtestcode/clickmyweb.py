
#encoding=utf-8
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import sys
import mouse
sys.path.append("D:\\Program Files\\Python\\Python36\\lib\\site-packages")

import win32con
import win32api
import win32gui

from docx import Document
from docx.shared import Inches

# iedriver="C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"
iedriver="D:\chromedriver_win32\chromedriver.exe"
os.environ["webdriver.chrome.driver"] = iedriver
options = webdriver.ChromeOptions()
options.add_argument('disable-infobars')
driver=webdriver.Chrome(iedriver,chrome_options=options)

url="http://www.xinyunfuwu.com/firsttransfer.jsp?enc=79a2424ac5f85024f2f6d82ce5ecbe" \
    "58bef1365ff2f2b30657cae4ce961c6b1fcca30feb476e02987958df89a9cd02aa3d61bf497f287f63b884671348" \
    "07b7160ec14623134da182cfd161f03042280f33ece23a311a4e70fa466d43bed1b1074adb31dc01e3d6b6b21891670f7d731209dc62" \
    "3e588e16662d44deb2dbf02896a20eb35cb5b667f77a9ca7644e6aef74&unitid=7320"

rangelist=(201,250)

#图片保存到word文档所在全路径
docfilepath="E:\\doctransfer\\中国城市发展史.docx"
#图片初始下载保存路径
downloadfilepath="C:\\Users\\chenp\\Downloads"



driver.get(url)
time.sleep(3)
n=0
for i in range(rangelist[0],rangelist[1]+1):
    driver.find_element_by_id('pageInput').clear()
    driver.find_element_by_id('pageInput').send_keys(str(i))
    driver.find_element_by_name('input').click()
    time.sleep(1)
    n=n+1
    for m in range(0,10):
        js="return "+"document.getElementsByClassName('readerTip')["+str(n-1)+"].style.display"
        display=driver.execute_script(js)
        if display=="none":
          break
        time.sleep(1)


    # xpath='//*[@ id ="reader"]/div/div['+str(i-100)+']/input'
    # ele=driver.find_element_by_xpath(xpath)
    mouse.mouse_r_click(500, 500)
    # ActionChains(driver).context_click(ele).perform()
    time.sleep(1)
    mouse.key_input("v")
    time.sleep(1)
    dialog = win32gui.FindWindow('#32770', u'另存为')  # 查找对话框
    # ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, 'ComboBoxEx32', None)
    # ComboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, 'ComboBox', None)
    DUIViewWndClassName=win32gui.FindWindowEx(dialog, 0, 'DUIViewWndClassName', None)
    DirectUIHWND=win32gui.FindWindowEx(DUIViewWndClassName, 0, 'DirectUIHWND', None)
    FloatNotifySink=win32gui.FindWindowEx( DirectUIHWND, 0, 'FloatNotifySink', None)
    ComboBox=win32gui.FindWindowEx(FloatNotifySink, 0, 'ComboBox', None)
    Edit = win32gui.FindWindowEx(ComboBox, 0, 'Edit', None)  # 上面三句依次寻找对象，直到找到输入框Edit对象的句柄
    button = win32gui.FindWindowEx(dialog, 0, 'Button', None)  # 确定按钮Button

    win32gui.SendMessage(Edit, win32con.WM_SETTEXT, None, str(i))  # 往输入框输入

    win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, button)  # 按确认button
    time.sleep(1)
    # driver.find_element_by_id('nextPage').click()

#改文件名字

def renamefile(filepath):

    filelist = os.listdir(filepath)
    while len(filelist) != 0:
        curfile = filelist.pop(0)

        if curfile[0:6] == "ss2jpg":
         # 提取出文件编号
            i = 1
            while curfile[8:8 + i].isdigit():
                i += 1
            if i == 1:
                pagenum = '0'
            else:
                pagenum = curfile[8:8 + i - 1]

            pagenum = str(int(pagenum) +rangelist[0])

            # 重命名文件
            os.rename(os.path.join(filepath, curfile), os.path.join(filepath, pagenum + '.png'))
    pass

#文档路径、图片路径、起始图片编号、终止图片编号
def insertintodoc(docfilepath,downloadfilepath,startnum,endnum):

    #打开一个已存在的文件
    doc = Document(docfilepath)

    for i in range(startnum,endnum+1):
        pngpath=os.path.join(downloadfilepath,str(i)+'.png')
        doc.add_picture(pngpath, width=Inches(5.7), height=Inches(9.3))

    # 如果路径和打开文件路径一致，是为追加内容到文档中，否则重建一个文件，每次都会清除再插入内容
    doc.save(docfilepath)

    #删除下载的图片
    pass

renamefile(downloadfilepath)
insertintodoc(docfilepath,downloadfilepath,rangelist[0],rangelist[1])

# for i in range(1,51):
#
#     xpath = '//*[@id="jcopeLightBox"]/div[1]/div[2]/div[1]/img[1]'
#     ele = driver.find_element_by_xpath(xpath)
#     ActionChains(driver).context_click(ele).perform()
#     mouse.key_input("v")
#     time.sleep(2)
#     dialog = win32gui.FindWindow('#32770', u'另存为')  # 查找对话框
#     # ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, 'ComboBoxEx32', None)
#     # ComboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, 'ComboBox', None)
#     DUIViewWndClassName=win32gui.FindWindowEx(dialog, 0, 'DUIViewWndClassName', None)
#     DirectUIHWND=win32gui.FindWindowEx(DUIViewWndClassName, 0, 'DirectUIHWND', None)
#     FloatNotifySink=win32gui.FindWindowEx( DirectUIHWND, 0, 'FloatNotifySink', None)
#     ComboBox=win32gui.FindWindowEx(FloatNotifySink, 0, 'ComboBox', None)
#     Edit = win32gui.FindWindowEx(ComboBox, 0, 'Edit', None)  # 上面三句依次寻找对象，直到找到输入框Edit对象的句柄
#     button = win32gui.FindWindowEx(dialog, 0, 'Button', None)  # 确定按钮Button
#
#     win32gui.SendMessage(Edit, win32con.WM_SETTEXT, None, str(i))  # 往输入框输入上传文件的绝对地址
#     win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, button)  # 按确认button
#     time.sleep(2)
#     driver.find_element_by_id('lbNext').click()
#     time.sleep(1)
