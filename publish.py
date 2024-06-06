# update: 2024-5-10
# version: python 3.12,selenium 4.0.0 ,谷歌浏览器124.0.X.X
# 使用说明：①手动将发布信息填入产品信息表，版本跟oss_url需要更新，其他属性固定
#         ②升级提醒文件需要放到指定目录"C:\Users\Administrator\Desktop\升级提醒"，并且需要改名“升级提醒”

from selenium import webdriver
from selenium.webdriver.common.by import By
import xlrd
import time
import pywinauto
from pywinauto.keyboard import send_keys

# 定义一个方法，只有win平台需要使用
def update_down(down, name):
    # 下载器
    xia_zai = web.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[2]/div/div/form/div[8]/div/div/div[1]/input')
    xia_zai.send_keys(down)
    web.find_element(By.XPATH, '/html/body/div[3]/div[1]/div[1]/ul').click()
    time.sleep(1)
    # launch name
    launch = web.find_element(By.XPATH, '//*[@id="app"]/section/section/main/div[2]/div/div/form/div[10]/div/div/input')
    launch.clear()
    launch.send_keys(name)
    time.sleep(1)

# 打开数据表
table_data = xlrd.open_workbook("产品信息表.xls")
# 选择sheet
sheet = table_data.sheet_by_index(0)
# 行总计
raw = sheet.nrows

# 创建一个新的Chrome浏览器实例，并将其赋值给变量web
web = webdriver.Chrome()

# 打开网站
web.get("https://cbs-view.afirstsoft.cn/#/Login")
# 输入用户名
user = web.find_element(By.XPATH,'//*[@id="app"]/div/div/form/div[1]/div/div/input')
user.clear()
user.send_keys("zouhuawei")
# 输入密码
password = web.find_element(By.XPATH,'//*[@id="app"]/div/div/form/div[2]/div/div/input')
password.clear()
password.send_keys("QLuZhpaZoddHzAk")
#登录
web.find_element(By.XPATH,'//*[@id="app"]/div/div/form/div[3]/div/button').click()
time.sleep(3)
# 发布管理
web.find_element(By.XPATH,'//*[@id="app"]/section/section/aside/ul/div/li[2]').click()
time.sleep(2)
web.find_element(By.XPATH,'//*[@id="app"]/section/section/aside/ul/div/li[2]/ul/div/li[3]').click()
time.sleep(2)

for row in range(1,2):
    row_date = sheet.row_values(row)
    platform = row_date[0]
    PID = int(row_date[1])
    channel = row_date[2]
    down_loader = row_date[3]
    launch_name = row_date[4]
    version = row_date[5]
    oss_url = row_date[6]

    # 新增发布
    web.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[2]/div/div/div[1]/div[1]/button[1]/span').click()
    time.sleep(2)
    # 选择主产品
    web.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[2]/div/div/form/div[1]/div/div/div[1]/input').send_keys(PID)
    time.sleep(1)
    web.find_element(By.XPATH,'/html/body/div[2]/div[1]/div[1]/ul').click()
    time.sleep(4)
    # 判断平台，win需要下载器，mac不需要下载器
    if platform == "win":
        update_down(down_loader,launch_name)
    # 软件版本号
    ban_ben = web.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[2]/div/div/form/div[11]/div/div/input')
    ban_ben.clear()
    ban_ben.send_keys(version)
    # 安装包url
    an_url = web.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[2]/div/div/form/div[12]/div/div[1]/input')
    an_url.clear()
    an_url.send_keys(oss_url)
    # 升级提醒
    web.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[2]/div/div/form/div[17]/div/label/span/span').click()
    time.sleep(2)
    # 导入升级提醒文件
    web.find_element(By.XPATH,'//*[contains(@id,"el-collapse-head")]/div/div/div/div/button').click()
    # 使用pywinauto来选择文件
    app = pywinauto.Desktop()
    # # 选择文件上传的窗口
    dlg = app["打开"]
    # dlg.print_control_identifiers()
    #
    # # 选择文件地址输入框
    dlg["Toolbar3"].click()
    send_keys(r"C:\Users\Administrator\Desktop\升级提醒")
    send_keys("{VK_RETURN}")
    #
    # # 选中文件名输入框
    dlg["文件名(&N):Edit"].type_keys("通用提醒.xlsx")
    #  点击打开
    dlg["打开(&O)"].click()
    time.sleep(1)
    # 点击发布
    # web.find_element(By.XPATH,'//*[@id="app"]/section/section/main/div[2]/div/div/form/div[19]/div/button/span').click()
    print(str(PID) + channel + '发布成功')
    time.sleep(5)