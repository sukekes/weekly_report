# encoding: utf-8
"""
@python version: 3.7
@author: mufeng12138
@file: main.py
@time: 2020/5/1 11:15
"""

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time
import xlrd
import pandas as pd


class case:
    def __init__(self):
        # 设置调试模式：0则运行窗口，其他则只读取用例信息
        self.status_debug = 0
        # 设置是否无头
        self.status_head = 1
        # 设置等待时间
        self.wait_time = 10
        if self.status_debug == 0:
            self.init_driver()
        # 读取预置信息
        self.path_config = r"D:\Desktop\0520case\config.xlsx"
        config_list = self.get_config()

        self.url = config_list[0]
        self.name = config_list[1]
        self.pswd = config_list[2]
        self.project_name = config_list[3]
        self.version = config_list[4]
        self.case_path = config_list[5]
        self.module_name = config_list[6]
        # 状态文本设置
        self.status_pass = 1
        self.status_fail = 2
        self.status_block = 3
        # self.end_time = list[7]
        # self.uploader = list[8]
        # self.verifier = list[9]

    def get_config(self):
        # 读取excel中的data列作为信息源
        data = pd.read_excel(self.path_config)
        results = data["data"]
        return results

    def init_driver(self):
        if self.status_head == 0:
            # 配置chrome的参数
            options = Options()
            options.add_argument('--headless')
            # options.add_argument('--disable-gpu')
            self.driver = webdriver.Chrome(options=options)
        else:
            self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        # self.driver.set_window_size(960, 1080)
        # self.driver.set_window_position(0, 0)

    def login(self):
        if self.status_debug == 0:
            self.driver.get(self.url)
            self.driver.find_element_by_name("userBase.userName").clear()
            self.driver.find_element_by_name("userBase.userName").send_keys(self.name)
            self.driver.find_element_by_name("userBase.password").clear()
            self.driver.find_element_by_name("userBase.password").send_keys(self.pswd)
            self.driver.find_element_by_id("submit_button").click()
            self.wait()

    def wait(self):
        # 隐式等待
        self.driver.implicitly_wait(self.wait_time)

    def locate_case(self):
        if self.status_debug == 0:
            # 我的任务
            self.driver.find_element_by_id("MENU_MYPJ").click()
            self.wait()
            # 搜索项目
            self.driver.find_element_by_name("projectInformation.projectName").send_keys(self.project_name)
            self.driver.find_element_by_id("pjt_search").click()
            self.wait()
            # 点击项目
            # self.driver.find_element_by_xpath("//td[@title = '%s']" % self.project_name).click()
            self.driver.find_element_by_xpath("//*[@id='1']/td[2]/a/font").click()
            # self.driver.find_element_by_id("1").click()
            # self.driver.find_element_by_link_text(self.project_name).click()
            self.wait()
            # time.sleep(self.wait_time)

            # 我的用例
            self.driver.find_element_by_id("t1").click()

    def get_nums(self):
        data = pd.read_excel(self.case_path)
        case_pass = data["no"][data["results"] == self.status_pass]
        case_pass = case_pass.reset_index(drop=True)
        print(case_pass)
        # record_this_month=
        #   record[
        #       (record['WTGS_CODE']==set)&
        #       (record['FAULT_CODE'].isin(fault_list))
        #   ]
        return case_pass

    def get_block_case(self):
        data = pd.read_excel(self.case_path)
        case_block = data["no"][data["results"] == self.status_block]
        case_block = case_block.reset_index(drop=True)
        return case_block

    def get_fail_case(self):
        data = pd.read_excel(self.case_path)
        case_fail = data["no"][data["results"] == self.status_fail]
        case_fail = case_fail.reset_index(drop=True)
        return case_fail

    def run(self):
        # 准备数据
        # 筛选block
        block_case = self.get_block_case()
        # print(block_case)
        # for i in range(len(block_case)):
        #     print(block_case[i])

        # 筛选fail
        fail_case = self.get_fail_case()
        # print(fail_case)
        # 剩余all pass

        if self.status_debug == 0:
            # 选择待提交case
            selection = Select(self.driver.find_element_by_id("state_name")).select_by_value("2")
            # selection.select_by_value("4")
            # selection.select_by_value("2")
            # self.driver.find_element_by_id("testcase_btn")
            self.wait()

            print(len(block_case))

            for i in range(len(block_case)):
                print(block_case[i])
                self.driver.find_element_by_id("testcaseCode").send_keys("C", block_case[i])
                self.driver.find_element_by_id("testcase_btn")
                # try:
                for j in range(11):
                    # row = self.driver.find_element_by_id("%d" % j)
                    row = self.driver.find_element_by_id("//*[@id='%d']/td[5]" % j)
                    print(row.title)
                # except


            # 执行
            # self.driver.find_element_by_xpath("//*[@id='1']/td[17]/a[2]").click()

            # src = self.driver.find_element_by_id("code")
            # print(src.get_attribute("title"), "\n", src.text, "\n", src.tag_name, "\n", "test")

            time.sleep(self.wait_time)
            print("run done")

    def __del__(self):
        # print("析构")
        if self.status_debug == 0:
            self.driver.close()

    # def get_zip(self, project_name):
    #     self.driver.find_element_by_xpath('''//li[@id = "MENU_PJ"]''').click()
    #     for i in range(1, len(project_name)):
    #         print(project_name[i])
    #         title = "".join(project_name[i])
    #         self.driver.find_element_by_xpath("//td[@title = '%s']" % title).click()
    #         self.driver.implicitly_wait(5)
    #         self.driver.find_element_by_id("project_defect").click()
    #         self.driver.implicitly_wait(5)
    #         self.driver.find_element_by_id("all").click()
    #         self.driver.implicitly_wait(5)
    #         time.sleep(3)
    #         try:
    #             self.driver.find_element_by_id("defect_report_btn").click()
    #         except selenium.common.exceptions.ElementClickInterceptedException:
    #             self.driver.find_element_by_id("popup_panel").click()
    #             self.driver.find_element_by_id("all").click()
    #             self.driver.implicitly_wait(5)
    #             self.driver.find_element_by_id("defect_report_btn").click()
    #         # exit()
    #         self.get_back()
    #         self.get_back()
    #     self.driver.close()

    def get_back(self):
        self.driver.back()
        self.wait()


if __name__ == "__main__":
    # 生成对象
    spiderman = case()
    # 登录
    spiderman.login()
    # 我的任务 - 项目 - 我的用例 - 选择版本
    spiderman.locate_case()
    # 获取excel的编号
    # nums = spiderman.get_nums()

    # print(nums.reset_index(drop=True)[4])
    # print(nums[3])
    # print(type(nums))

    # print(nums[0])
    # length_of_nums = len(nums)
    # print(range(length_of_nums))
    # print(range(4))
    # for i in range(len(nums)):
    #     print(nums.reset_index(drop=True)[i])
    # 执行用例
    spiderman.run()
    # 收尾
    # spiderman.end()
