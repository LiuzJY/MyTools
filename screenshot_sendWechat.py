# -*- coding:utf-8 -*-

import locale
import win32api
import win32clipboard as clipboard
import win32con
import win32gui
from datetime import datetime, timedelta
from time import sleep
from selenium import webdriver

from PIL import Image
from io import BytesIO
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


def login(dr):
    dr.maximize_window()
    dr.get("www.baidu.com")

    dr.find_element(By.NAME, 'user').send_keys("migu")
    dr.find_element(By.NAME, 'password').send_keys("migu")
    dr.find_element(By.NAME, 'password').send_keys(Keys.ENTER)


def screenshot(dr):
    img = dr.find_element(By.CSS_SELECTOR, '.react-grid-layout')
    dr.save_screenshot("screenshot.png")

    left = img.location['x']
    top = img.location['y']
    right = img.location['x'] + img.size['width']
    bottom = img.location['y'] + img.size['height']

    im = Image.open('screenshot.png')
    im = im.crop((left, top, right, bottom))  # 对浏览器截图进行裁剪
    im.save('monitor.png')


def msg():
    now = datetime.now()
    late = now - timedelta(hours=0.5)

    locale.setlocale(locale.LC_CTYPE, 'chinese')    # 时间及日期以中文格式显示
    time_msg = late.strftime('%m月%d日 %H:%M' + '-' + now.strftime('%H:%M'))
    msg_txt = "*****系统 "+time_msg+" 系统运行正常"
    return msg_txt


def txt_ctrl_v(txt_str):
    clipboard.OpenClipboard()
    clipboard.EmptyClipboard()
    clipboard.SetClipboardData(win32con.CF_UNICODETEXT, txt_str)
    clipboard.CloseClipboard()


def img_ctrl_v():
    img = Image.open("monitor.png")
    output = BytesIO()
    img.convert("RGB").save(output, "BMP")  # 以BMP格式保存流
    data = output.getvalue()[14:]  # bmp文件头14个字节丢弃
    print(data)
    output.close()
    clipboard.OpenClipboard()  # 打开剪贴板
    clipboard.EmptyClipboard()  # 先清空剪贴板
    clipboard.SetClipboardData(win32con.CF_DIB, data)  # 将图片放入剪贴板
    clipboard.CloseClipboard()


def send_msg(send_win):
    win32api.keybd_event(17, 0, 0, 0)  # 有效，按下CTRL
    sleep(1)  # 需要延时
    win32gui.SendMessage(send_win, win32con.WM_KEYDOWN, 86, 0)  # V
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)  # 放开CTRL
    sleep(1)  # 缓冲时间
    win32gui.SendMessage(send_win, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  # 回车发送


if __name__ == '__main__':
    driver = webdriver.Chrome()
    login(driver)
    sleep(3)
    screenshot(driver)
    driver.quit()

    msg = msg()
    # print(msg)
    win = win32gui.FindWindow('ChatWnd', '文件传输助手')
    print('%x' % win)
    win32gui.SetForegroundWindow(win)

    txt_ctrl_v(msg)
    send_msg(win)

    img_ctrl_v()
    send_msg(win)
    # while True:
    #     driver = webdriver.Chrome()
    #     login(driver)
    #     sleep(3)
    #     screenshot(driver)
    #     driver.quit()
    #
    #     win = win32gui.FindWindow('ChatWnd', '文件传输助手')
    #     print('%x' % win)
    #     win32gui.SetForegroundWindow(win)
    #
    #     txt_ctrl_v(msg())
    #     send_msg(win)
    #
    #     img_ctrl_v()
    #     send_msg(win)
    #
    #     sleep(10)
