import pyautogui
import time
import win32api
import win32con
import win32gui
import win32clipboard as w
from PIL import ImageGrab
import requests
import base64

def FindWindow(chatroom):
    win = win32gui.FindWindow('WeChatMainWndForPC',chatroom)
    print("找到窗口句柄：%x" % win)
    if win != 0:
        win32gui.ShowWindow(win, win32con.SW_SHOWMINIMIZED)
        win32gui.ShowWindow(win, win32con.SW_SHOWNORMAL)
        win32gui.ShowWindow(win, win32con.SW_SHOW)
        win32gui.SetWindowPos(win, win32con.HWND_TOP, 0, 0, 500, 700, win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(win)  # 获取控制
        time.sleep(1)
        tit = win32gui.GetWindowText(win)
        print('已启动【'+str(tit)+'】窗口')
    else:
        print('找不到【%s】窗口' % chatroom)
        exit()

# 设置和粘贴剪贴板
def ClipboardText(ClipboardText):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, ClipboardText)
    w.CloseClipboard()
    time.sleep(1)
    win32api.keybd_event(17,0,0,0)
    win32api.keybd_event(86,0,0,0)
    win32api.keybd_event(86,0,win32con.KEYEVENTF_KEYUP,0)
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)

# 模拟发送动作
def SendMsg():
    win32api.keybd_event(18, 0, 0, 0)
    win32api.keybd_event(83,0,0,0)
    win32api.keybd_event(18,0,win32con.KEYEVENTF_KEYUP,0)
    win32api.keybd_event(83,0,win32con.KEYEVENTF_KEYUP,0)

# 模拟发送微信文本消息
def SendWxMsg(wxid,sendtext):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索微信
    ClipboardText(wxid)
    time.sleep(1)
    # 进入聊天窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 粘贴文本内容
    ClipboardText(sendtext)
    # 发送
    SendMsg()
    print('已发送')
    # 关闭微信窗口
    time.sleep(1)
    pyautogui.moveTo(683, 16)
    pyautogui.click()

# 模拟发送文件消息（图片、文档、压缩包等）
def SendWxFileMsg(wxid,imgpath):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索微信
    ClipboardText(wxid)
    time.sleep(1)
    # 进入聊天窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 选择文件
    pyautogui.moveTo(373, 570)
    pyautogui.click()
    ClipboardText(imgpath)
    time.sleep(1)
    pyautogui.moveTo(784, 509)
    pyautogui.click()
    # 发送
    SendMsg()
    print('已发送')
    # 关闭微信窗口
    time.sleep(1)
    pyautogui.moveTo(683, 16)
    pyautogui.click()

# 转发群里最新的一条消息
def ZhuanfaMsg(wxid,groupname):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索群
    ClipboardText(groupname)
    time.sleep(1)
    # 进入群窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 开始转发
    pyautogui.moveTo(484, 439)
    time.sleep(1)
    pyautogui.rightClick()
    pyautogui.moveTo(543, 454)
    time.sleep(1)
    pyautogui.click()
    # 搜索用户
    ClipboardText(wxid)
    time.sleep(1)
    pyautogui.moveTo(828, 406)
    pyautogui.click()
    time.sleep(1)
    # 确定转发
    pyautogui.moveTo(1108, 755)
    pyautogui.click()

# 获取你的个人信息（昵称、微信号）
def GetYourInfo():
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 点击你的头像
    pyautogui.moveTo(28, 56)
    pyautogui.click()
    time.sleep(1)
    # 用户信息截图
    userinfo = (20, 60, 319, 284)
    userinfo_img = ImageGrab.grab(userinfo)
    userinfo_img.save('userinfo.png')
    # 识别用户信息截图
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
    f = open('userinfo.png', 'rb')
    img = base64.b64encode(f.read())
    params = {"image":img}
    access_token = '24.21008e8b243c5b2864a7cf0583d907e1.2592000.1661322017.282335-24796078'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print (response.json())
    time.sleep(1)
    # 关闭微信窗口
    time.sleep(1)
    pyautogui.moveTo(683, 16)
    pyautogui.click()

# 获取好友微信的个人信息（昵称、微信号）
def GetFriendInfo(wxid):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索微信
    ClipboardText(wxid)
    time.sleep(1)
    # 进入聊天窗口
    pyautogui.moveTo(160, 93)
    pyautogui.click()
    time.sleep(1)
    # 点击右上角···
    pyautogui.moveTo(678, 43)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(801, 51)
    pyautogui.click()
    # 用户信息截图
    userinfo = (802, 54, 1085, 331)
    userinfo_img = ImageGrab.grab(userinfo)
    userinfo_img.save('userinfo.png')
    # 识别用户信息截图
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
    f = open('userinfo.png', 'rb')
    img = base64.b64encode(f.read())
    params = {"image":img}
    access_token = '24.21008e8b243c5b2864a7cf0583d907e1.2592000.1661322017.282335-24796078'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print (response.json())
    time.sleep(1)
    # 关闭微信窗口
    time.sleep(1)
    pyautogui.moveTo(683, 16)
    pyautogui.click()
    pyautogui.click()

# 获取群人数
def GetCharRoomUserNum(groupname):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索群
    ClipboardText(groupname)
    time.sleep(1)
    # 进入群窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 群人数区域截图
    userinfo = (310, 0, 659, 47)
    userinfo_img = ImageGrab.grab(userinfo)
    userinfo_img.save('chatroom.png')
    # 识别群人数截图
    # 开发文档:https://cloud.baidu.com/doc/OCR/s/Ck3h7y2ia
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
    f = open('chatroom.png', 'rb')
    img = base64.b64encode(f.read())
    params = {"image":img}
    access_token = '24.21008e8b243c5b2864a7cf0583d907e1.2592000.1661322017.282335-24796078'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print (response.json())
    time.sleep(1)
    # 关闭微信窗口
    time.sleep(1)
    pyautogui.moveTo(683, 16)
    pyautogui.click()


# 发布群公告
def AddGorupNotice(groupname,NoticeText):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索群
    ClipboardText(groupname)
    time.sleep(1)
    # 进入群窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 定位到群名称
    pyautogui.moveTo(363, 33)
    pyautogui.click()
    # 定位到群公告
    time.sleep(1)
    pyautogui.moveTo(731, 509)
    pyautogui.click()
    # 粘贴群公告内容
    ClipboardText(NoticeText)
    # 确认发布群公告
    time.sleep(1)
    pyautogui.moveTo(288, 500)
    pyautogui.click()
    pyautogui.moveTo(312, 297)
    pyautogui.click()

# 邀请好友进群
def ReqFriendsToGroup(groupname,wxid):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索群
    ClipboardText(groupname)
    time.sleep(1)
    # 进入群窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 定位到群名称
    pyautogui.moveTo(363, 33)
    pyautogui.click()
    # 定位到邀请
    time.sleep(1)
    pyautogui.moveTo(852, 300)
    pyautogui.click()
    # 搜索好友
    ClipboardText(wxid)
    # 发出邀请
    time.sleep(1)
    pyautogui.moveTo(232, 197)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(493, 561)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(395, 427)
    pyautogui.click()

# 获取好友最新的聊天记录
def GetChatRecord(wxid):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 搜索微信号
    pyautogui.moveTo(166, 38)
    pyautogui.click()
    ClipboardText(wxid)
    time.sleep(1)
    pyautogui.moveTo(197, 123)
    pyautogui.click()
    # 聊天内容区域截图
    userinfo = (314, 68, 683, 549)
    userinfo_img = ImageGrab.grab(userinfo)
    userinfo_img.save('chatrecord.png')
    # 识别当前聊天窗口截图
    # 开发文档:https://cloud.baidu.com/doc/OCR/s/Ck3h7y2ia
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
    f = open('chatrecord.png', 'rb')
    img = base64.b64encode(f.read())
    params = {"image":img}
    access_token = '24.21008e8b243c5b2864a7cf0583d907e1.2592000.1661322017.282335-24796078'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print (response.json())

# 置顶群或好友
def SetTop(groupname_wxid):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索群或好友
    ClipboardText(groupname_wxid)
    time.sleep(1)
    # 进入窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 打开设置
    pyautogui.moveTo(684, 38)
    pyautogui.click()
    time.sleep(1)
    # 置顶
    pyautogui.moveTo(914, 227)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(682, 18)
    pyautogui.click()
    print('已将'+str(groupname_wxid)+'置顶')


# 添加微信
def AddWx(wxid):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 定位到添加微信位置
    pyautogui.moveTo(25, 151)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(278, 39)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(153, 39)
    pyautogui.click()
    time.sleep(1)
    # 搜索微信号
    ClipboardText(wxid)
    time.sleep(1)
    pyautogui.moveTo(183, 91)
    pyautogui.click()
    time.sleep(2)
    # 对搜索微信号结果进行截图
    userinfo = (306, 68, 565, 240)
    userinfo_img = ImageGrab.grab(userinfo)
    userinfo_img.save('addwx.png')
    # 对搜索微信号结果进行识别
    # 开发文档:https://cloud.baidu.com/doc/OCR/s/Ck3h7y2ia
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
    f = open('addwx.png', 'rb')
    img = base64.b64encode(f.read())
    params = {"image":img}
    access_token = '24.21008e8b243c5b2864a7cf0583d907e1.2592000.1661322017.282335-24796078'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response:
        print (response.json())
    # 添加
    time.sleep(1)
    pyautogui.moveTo(435, 203)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(300, 621)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(278, 37)
    pyautogui.click()
    print('已向'+str(wxid)+'发送添加好友请求')
    pyautogui.moveTo(682, 18)
    pyautogui.click()

# 发送文本消息（微信号或微信昵称或备注，需要发送的文本消息）
# SendWxMsg('牧耀佑七','测试')

# 发送文件消息（图片、文档、压缩包等）
# SendWxFileMsg('微信号',r"文件路径")

# 转发群里最新的一条消息（微信号或微信昵称或备注，群名称）
# ZhuanfaMsg('微信号','群名称')

# 获取你的个人信息（昵称、微信号）
# GetYourInfo('微信号')

# 获取好友微信的个人信息（昵称、微信号）
# GetFriendInfo('微信号')

# 获取微信群人数
# GetCharRoomUserNum('群名称')

# 发布群公告
# AddGorupNotice('群名称','Python发布群公告')

# 邀请好友进群
# ReqFriendsToGroup('群名称','cbzqx88')

# 获取好友最新的聊天记录
# GetChatRecord('微信号')

# 置顶群或好友
# SetTop('微信号')

# 添加微信
# AddWx('微信号')

foods = open('foods.txt','r',encoding='utf-8').read()
SendWxMsg('牧耀佑七',foods)
