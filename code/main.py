# encoding: utf-8
"""
@python version: 3.7
@author: mufeng12138
@file: main.py
@time: 2020/5/1 11:15
"""

from selenium import webdriver
import time
import logging
import xlrd

import os
import shutil
import zipfile
from os.path import join, getsize
import pandas as pd
import openpyxl

class report():
    def __init__(self):
        # self.init_driver()
        self.path_config = r"D:\Desktop\0501report\test\config.xlsx"
        list = self.get_config()

        self.url = list[0]
        self.name = list[1]
        self.passwd = list[2]
        self.file_path = list[3]
        self.download_path = list[4]
        # print(self.download_path)
        self.dst_dir = list[5]
        self.start_time = list[6]
        self.end_time = list[7]
        self.uploader = list[8]
        self.verifier = list[9]

    def get_config(self):
        data = pd.read_excel(self.path_config)
        results = data["data"]
        return results

    def init_driver(self):
        self.driver = webdriver.Chrome()
        self.driver.set_window_size(960, 1080)
        self.driver.set_window_position(0, 0)

    def login(self):
        self.driver.get(self.url)
        self.driver.find_element_by_name("userBase.userName").clear()
        self.driver.find_element_by_name("userBase.userName").send_keys(self.name)
        self.driver.find_element_by_name("userBase.password").send_keys(self.passwd)
        self.driver.find_element_by_id("submit_button").click()
        self.driver.implicitly_wait(5)
        # time.sleep(5)
        # driver.close()
        # logging.info("fuc")

    def get_zip(self, project_name):
        self.driver.find_element_by_xpath('''//li[@id = "MENU_PJ"]''').click()
        for i in range(1, len(project_name)):
            # # 获取页数
            # num_of_pages = self.driver.find_element_by_id("sp_1_pjt_jqgrid_pager").text
            # for page in num_of_pages-1:
            #     print("page:", page)

            print(project_name[i])
            # print(type(i))
            title = "".join(project_name[i])
            # print(type(title))
            # self.driver.find_elements_by_link_text("".join(project_name[i])).click()
            # self.driver.find_elements_by_xpath('''//td[@title = "("".join(project_name[i]))"]''').click()
            self.driver.find_element_by_xpath("//td[@title = '%s']" % title).click()
            self.driver.implicitly_wait(5)
            # self.driver.find_element_by_class_name("fa fa-bug").click()
            # self.driver.find_element_by_class_name("fa fa-bug").click()
            self.driver.find_element_by_id("project_defect").click()
            self.driver.implicitly_wait(5)
            self.driver.find_element_by_id("all").click()
            # self.driver.find_element_by_id("all").click()
            self.driver.implicitly_wait(5)
            time.sleep(3)
            try:
                self.driver.find_element_by_id("defect_report_btn").click()
            except selenium.common.exceptions.ElementClickInterceptedException:
                self.driver.find_element_by_id("popup_panel").click()
                self.driver.find_element_by_id("all").click()
                self.driver.implicitly_wait(5)
                self.driver.find_element_by_id("defect_report_btn").click()
            # exit()
            self.getback()
            self.getback()
        self.driver.close()

    def getback(self):
        self.driver.back()
        self.driver.implicitly_wait(5)

    def read_excel(self):
        data = xlrd.open_workbook(self.file_path)  # 打开excel表格，参数为文件路径
        sheet_names = data.sheet_names()  # 获取所有sheet的名称
        table = data.sheet_by_name(sheet_names[0])  # 通过名称获取表格
        rows = table.nrows  # 获取总行数
        # print(rows)
        # print("range(rows):", range(rows))
        project_name = list(range(rows))
        # for i in project_name:
        #     print("project_name", project_name[i])
        # print("project_name", project_name)
        for i in range(1,rows):  # 获取每一行的数据
            # print(i)

            row_content = table.row_values(i)
            project_name[i] = row_content
            # print("row_content", row_content)
            # project_name.append(row_content)
            # print("project_name", project_name)
        # print(len(project_name))
        # for i in range(len(project_name)):
        #     print("".join(project_name[i]))
        return project_name
            # password = int(row_content[1])  # 读取的数字会变成float类型 如12345变为 12345.0

    # def decode(self):
    def unzip_file(self):
        zip_src = self.download_path
        dst_dir = self.dst_dir
        for filename in os.listdir(zip_src):
            filename = zip_src + "\\" + filename
            r = zipfile.is_zipfile(filename)
            # print(r)
            # print("filename", filename)
            if r:
                fz = zipfile.ZipFile(filename, 'r')
                for file in fz.namelist():
                    fz.extract(file, dst_dir)

            else:
                print('This is not zip')


    def get_data(self):
        dir = self.dst_dir
        for filename in os.listdir(dir):
            # print(filename)
            filename = dir + "\\" + filename
            # print(filename)
            # 跳过非xlsx文件
            r = zipfile.is_zipfile(filename)
            # print(r)
            if r:

                # print(filename)
                src = pd.read_excel(filename, header = 2, sheet_name = 0)
                # print(src)

                result = [self.increased_filter(src), self.closed_filter(src), self.still_open_filter(src)]
                # result = self.test(src)
                print("*"*50)
                print(filename.split("\\")[-1].split(".")[-2])
                # name = filename.split("\\")[-1].split(".")[-2]
                # print(name.split("-缺陷-")[0], "-", name.split("-缺陷-")[1])
                print(self.start_time, "-", self.end_time)
                print("新增\t关闭\t剩余\t")
                print(result)
                # os.remove(filename)
                # print(data.shape[1])
            else:
                continue

    def test(self, src):
        try:
            data = src.loc[src["提交人"] == self.uploader]
            up_time = "提交时间"
            # print(data[up_time])
            # 转变为时间戳

            data[up_time] = pd.to_datetime(data[up_time])
            # print(data[up_time])
            print(pd.to_datetime(self.start_time))
            print(pd.to_datetime(self.end_time))
            data = data[(data[up_time] >= pd.to_datetime(self.start_time)) & (data[up_time] <= pd.to_datetime(self.end_time))]
            # print("data_increased_issue", data.shape[0])
            return data.shape[0]
        except KeyError:
            print("未找到对应字段")
            return


    def closed_filter(self, src):
        try:
            # 关闭
            # s_date = datetime.datetime.strptime('20050606', '%Y%m%d').date()
            # e_date = datetime.datetime.strptime('20071016', '%Y%m%d').date()
            # df = df[(df['tra_date'] >= s_date) & (df['tra_date'] <= e_date)]
            # print("d f", df)

            data = src.loc[src["状态"] == "已关闭"] \
                .loc[src["验证人"] == self.verifier]
            # data = src.loc[src["状态"] == "已关闭"] \
            #     .loc[src["验证人"] == self.verifier].loc[]
            # .loc[src["提交人"] == self.uploader] \
            this_time = "验证时间"
            data[this_time] = pd.to_datetime(data[this_time])
            # print("thistime:", data[this_time])
            # print(type(data[this_time]))
            data = data[(data[this_time] >= pd.to_datetime(self.start_time)) & (data[this_time] <= pd.to_datetime(self.end_time))]
            # print("data:", data)
            # print("data_closed", data.shape[0])
            return data.shape[0]
        except KeyError:
            print("未找到对应字段")
            return

    def increased_filter(self, src):
        try:
            data = src.loc[src["提交人"] == self.uploader]
            # 转变为时间戳
            this_time = "提交时间"
            data[this_time] = pd.to_datetime(data[this_time])
            data = data[(data[this_time] >= pd.to_datetime(self.start_time)) & (data[this_time] <= pd.to_datetime(self.end_time))]
            # print(data)
            # print("data_increased_issue", data.shape[0])
            return data.shape[0]
        except KeyError:
            print("未找到对应字段")
            return

    def still_open_filter(self, src):
        try:
            # 未关闭
            data = src.loc[(src["状态"] == "待修复") | (src["状态"] == "待回归")]
            # .loc[src["提交人"] == self.uploader] \
            # print("data_still_open", data.shape[0])
            return data.shape[0]

        except KeyError:
            print("未找到对应字段")
            return

    def age_10_to_30(self, a):
        return 10 <= a < 40

    # def status_close(self, test):
    #     return test == "已关闭"

    def data_to_excel(self):
        f = openpyxl.Workbook()
        sheet1 = f.create_sheet()
    # def get_word(self,driver):
    # def mf_output(self,driver):

    def get_increased_issue(self):
        print("hello")

    def remove_files(self):
        zip_src = self.download_path
        for filename in os.listdir(zip_src):
            r = zipfile.is_zipfile(filename)
            if r:
                os.remove(filename)
            else:
                continue
        for filename in os.listdir(self.dst_dir):
            r = zipfile.is_zipfile(filename)
            if r:
                os.remove(filename)
            else:
                continue


if __name__ == "__main__":
    # 创建对象，生成driver
    report = report()
    # 登录页面
    # report.login()
    # 获取项目名称
    # project_name = report.read_excel()
    # 根据项目名称从平台中下载对应缺陷
    # report.get_zip(project_name)
     # 保险起见，等待下载完成
    # time.sleep(3)
    # 解压zip包
    report.unzip_file()
    # 数据获取
    report.get_data()
     # report.get_increased_issue()
    # report.remove_files()
    # 生成报告
     # report.data_to_excel()
