# -*- coding: UTF-8 -*-
# @Time : 2022/2/6 7:12
# @Author : HJL
import os
import time
from datetime import datetime
import requests
import psutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import smtplib
from email.mime.text import MIMEText

accounts = [
    {'username': '学号1',
     'password': '密码1',
     'receiver': '邮箱1'  # 要有邮箱域名
     },
    {'username': '学号2',
     'password': '密码2',
     'receiver': '邮箱2'  # 要有邮箱域名
     },
    # 支持多个账户批量填报
]


"""  登录返校通  """
def login(username, password):
    # 创建一个参数对象，用来控制chrome以无界面模式打开
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(10)  # 等待防止网络不稳定引起的报错
    # 打开返校通登录界面
    driver.get(r"http://fanxiaotong.jiangnan.edu.cn/passport/login")
    # 跳转到e江南授权登录界面
    driver.find_element_by_xpath('/html/body/div[2]/div/div/div/div/div/div/div/div/div[1]/div[1]/a').click()
    try:
        # 输入e江南账号密码
        user_name = driver.find_element_by_id('username')
        user_name.clear()
        user_name.send_keys(username)
        pass_word = driver.find_element_by_id('pwd')
        pass_word.clear()
        pass_word.send_keys(password)
        driver.find_element_by_class_name('loginbutton').click()
    except Exception:
        # 获取不到元素，已登录
        pass
    # 跳转到日报填表界面
    driver.get(r'http://fanxiaotong.jiangnan.edu.cn/daily/fill')
    return driver


"""  提交日报  """
def submit_report(driver):
    try:
        # 点击提交按钮
        driver.find_element_by_xpath('//*[@id="form"]/div[22]/div/button').click()
    except Exception:
        print('提交失败！')
        exit(-1)
    return driver


"""  获取填报状态  """
def get_status(driver, account):
    html = driver.execute_script("return document.documentElement.outerHTML")
    success_msg = time.strftime('%Y-%m-%d', time.localtime()) + '日报'
    # 检查今日是否填报
    if success_msg in html:
        # 检查学号是否正确，防止填报错误
        driver.get('http://fanxiaotong.jiangnan.edu.cn/home')
        student_id = driver.find_element_by_xpath('/html/body/div[2]/div/div[1]/div/div/div[1]/div/p[2]').text[-10:]
        if student_id == account['username']:
            return success_msg[:-2] + '\n账号' + student_id + ' 返校通填报成功！'
    else:
        return '填报失败！失败日期：' + success_msg[:-2]


"""  邮件发送工具  """
class EmailSendTool:
    def __init__(self, qq='你的QQ', auth_code='你的QQ邮箱授权码'):
        self.__qq = qq
        self.__auth_code = auth_code
        self.__server = None
        self.__sender = None
        self.login()  # 登陆邮箱

    def login(self):
        self.__sender = self.__qq + '@qq.com'
        try:
            self.__server = smtplib.SMTP_SSL('smtp.qq.com', 465)  # 连接QQ邮箱的smtp服务器，和对应端口
            self.__server.login(self.__sender, self.__auth_code)  # 登入
        except:
            print('登录失败!')

    def send_msg(self, receiver, content):
        msg = MIMEText(content, 'plain', 'utf-8')
        msg['From'] = '返校通小助手'  # 发件人
        msg['Subject'] = '今日填报'  # 邮件主体，标题
        msg['To'] = receiver  # 接受人
        try:
            self.__server.sendmail(self.__sender, receiver, msg.as_string())  # 从谁发送给谁，内容是什么
            print(receiver + "发送成功!")
        except Exception:
            print(receiver + "发送失败!")

    def quit(self):
        self.__server.quit()


"""  检查网络状态  """
def isConnected():
    try:
        requests.get("http://www.baidu.com", timeout=2)
    except Exception:
        return False
    return True


if __name__ == "__main__":
    # 是否进行自动关机
    shutdown = False
    ran_seconds = (datetime.now() - datetime.fromtimestamp(psutil.boot_time())).total_seconds()
    if ran_seconds < 5 * 60:  # 运行时间不够5分钟，可知为通过BIOS启动
        shutdown = True
    # 检查网络连接是否正常
    while not isConnected():
        time.sleep(1)
    # 初始化邮件发送工具，一般来讲不会出现异常
    email = EmailSendTool()
    for account in accounts:
        username = account['username']
        password = account['password']
        receiver = account['receiver']
        # 防止发生异常
        try:
            driver = login(username, password)
            driver = submit_report(driver)
            status_msg = get_status(driver, account)
            email.send_msg(receiver, status_msg)
            driver.quit()
        except Exception as err:
            # 发送出错信息
            email.send_msg(receiver, '填报失败！失败时间：' + str(datetime.now()) + '\n异常信息：' + str(err))
        # 等待2秒
        time.sleep(2)
    email.quit()
    if shutdown:
        # 5秒后关机
        os.system('shutdown /s /t 5')
