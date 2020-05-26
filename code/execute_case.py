# encoding: utf-8
"""
@python version: 3.7
@author: mufeng12138
@file: main.py
@time: 2020/5/1 11:15
"""

from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import xlrd

import os
import zipfile
import pandas as pd


class case:
    def __init__(self):
        self.status_debug = 0
        self.wait_time = 5
        if self.status_debug == 0:
            self.init_driver()
        self.path_config = r"D:\Desktop\0520case\config.xlsx"
        config_list = self.get_config()

        self.url = config_list[0]
        self.name = config_list[1]
        self.pswd = config_list[2]
        self.project_name = config_list[3]
        self.version = config_list[4]
        self.case_path = config_list[5]
        self.module_name = config_list[6]
        self.status_pass = 1
        self.status_fail = 2
        self.status_block = 3
        # self.end_time = list[7]
        # self.uploader = list[8]
        # self.verifier = list[9]

    def get_config(self):
        data = pd.read_excel(self.path_config)
        results = data["data"]
        return results

    def init_driver(self):
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
            self.driver.implicitly_wait(self.wait_time)

    def locate_case(self):
        if self.status_debug == 0:
            # 我的任务
            self.driver.find_element_by_id("MENU_MYPJ").click()
            self.driver.implicitly_wait(self.wait_time)
            # 搜索并进入项目
            self.driver.find_element_by_name("projectInformation.projectName").send_keys(self.project_name)
            self.driver.find_element_by_id("pjt_search").click()
            self.driver.implicitly_wait(self.wait_time)
            # self.driver.find_element_by_xpath("//td[@title = '%s']" % self.project_name).click()
            self.driver.find_element_by_xpath("//*[@id='1']/td[2]/a/font").click()
            # self.driver.find_element_by_id("1").click()
            # self.driver.find_element_by_link_text(self.project_name).click()
            self.driver.implicitly_wait(self.wait_time)
            time.sleep(self.wait_time)

            # 我的用例
            self.driver.find_element_by_id("t1").click()

    def get_nums(self):
        data = pd.read_excel(self.case_path)
        case_pass = data["no"][data["results"] == self.status_pass]
        # case_fail = data["no"][data["results"] == self.status_fail]
        # case_block = data["no"][data["results"] == self.status_block]
        # print(case_pass)
        # print(case_fail)
        # print(case_block)
        # record_this_month=
        #   record[
        #       (record['WTGS_CODE']==set)&
        #       (record['FAULT_CODE'].isin(fault_list))
        #   ]
        return case_pass

    def run(self, nums):
        if self.status_debug == 0:
            # print(nums)
            selection = Select(self.driver.find_element_by_id("state_name"))
            # selection.select_by_value("4")
            selection.select_by_value("4")
            self.driver.find_element_by_id("testcase_btn")
            self.driver.implicitly_wait(self.wait_time)
            time.sleep(self.wait_time)
            print("run")

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
        self.driver.implicitly_wait(self.wait_time)


if __name__ == "__main__":
    # 生成对象
    spiderman = case()
    # 登录
    spiderman.login()
    # 我的任务 - 项目 - 我的用例 - 选择版本
    spiderman.locate_case()
    # 获取excel的编号
    nums = spiderman.get_nums()

    # print(nums.reset_index(drop=True)[4])
    # print(nums[3])
    # print(type(nums))

    # print(nums[0])
    # length_of_nums = len(nums)
    # print(range(length_of_nums))
    # print(range(4))
    for i in range(len(nums)):
        print(nums.reset_index(drop=True)[i])
    # 执行用例
    spiderman.run(nums)
    # 收尾
    # spiderman.end()
