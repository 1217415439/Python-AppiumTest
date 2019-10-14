import csv
import time
import unittest
from appium import webdriver
from appium.webdriver.common.touch_action import TouchAction
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
import openpyxl as xl


class MyTestCase(unittest.TestCase):
    def setUp(self):
        # 设备初始化配置
        desired_caps = {}
        desired_caps['platformName'] = 'Android'
        desired_caps['platformVersion'] = '9'
        desired_caps['deviceName'] = '0123456789ABCDEF'
        desired_caps['appPackage'] = 'com.android.launcher3'
        desired_caps['appActivity'] = 'com.android.launcher3.Launcher'
        desired_caps['appWaitActivity'] = 'com.android.launcher3.Launcher'
        # 打开设备连接
        self.driver = webdriver.Remote('http://localhost:4725/wd/hub', desired_caps)
        # 测试结果标志位
        self.result = 'True'
        # 计数器
        self.count = 0
        # 屏幕分辨率
        self.size = []
        # 状态值
        self.state = 'PASS'
        # 当前时间戳
        self.currentTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

    def tearDown(self):
        self.driver.quit()

    # 保存数据到excel表的方法
    @staticmethod
    def SaveDataToEXCEL(count, currenttime, state):
        # 打开excel表
        writexl = xl.load_workbook("OKGoogleResult.xlsx")
        # 定位到第一个sheet
        writeSheet = writexl.active
        # 行位置
        strCount = str(count+1)
        # 列位置
        cloum_1 = 'A' + strCount
        cloum_2 = 'B' + strCount
        # 写数据
        writeSheet[cloum_1] = currenttime
        writeSheet[cloum_2] = state
        # 保存数据
        writexl.save("OKGoogleResult.xlsx")

    @staticmethod
    def waitElement(driver, time, element_by, element, msg):
        """
        等待元素出现
        :param driver: driver
        :param time: 等待时间
        :param element_by: 元素类型
        :param element: 元素关键字
        :param msg: 输出信息
        :return:
        """
        WebDriverWait(driver, time). \
            until(expected_conditions.presence_of_element_located((element_by, element)), msg)

    # 获取屏幕分辨率
    def getSize(self):
        x = self.driver.get_window_size()['width']
        y = self.driver.get_window_size()['height']
        return x, y

    # 统计OkGoogle唤醒成功次数
    def testOkGoogle(self):
        while 1:
            try:
                # 等待唤醒成功弹出activity
                self.waitElement(self.driver, 900, By.ID, "com.google.android.googlequicksearchbox:id/immersive_view_container", "没有发现xxxx...")
                # 计数器加一
                self.count += 1
                # 记录当前时间
                self.currentTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                # 等待一秒
                time.sleep(1)
                # 保存数据到excel
                self.SaveDataToEXCEL(self.count, self.currentTime, self.state)
                # 获取屏幕分辨率
                size = self.getSize()
                # 通过分辨率计算点击屏幕位置，从而关闭该弹出activity
                TouchAction(self.driver).tap(x=size[0] * 0.1, y=size[1] * 0.5).perform()
            except TimeoutException:
                self.result = 'false'
                break
        # 判定测试结果
        self.assertEqual(self.result, True)


if __name__ == '__main__':
    unittest.main()
