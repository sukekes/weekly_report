# encoding: utf-8
"""
@python version: 3.7
@author: mufeng12138
@file: main.py
@time: 2020/5/1 11:15
"""

from selenium import webdriver
import time
import xlrd

import os
import zipfile
import pandas as pd


class report:
    def __init__(self):
        # {
        #     status_data:{
        #         0:"文档",
        #         1:"每日，默认当天数据",
        #         2:"每周，默认今天周五"
        #     }
        # }
        self.status_data = 2
        # self.init_driver()
        # self.path_config = r"D:\Desktop\0501report\test\config.xlsx"
        self.path_config = r"./config.xlsx"
        list = self.get_config()

        self.url = list[0]
        self.name = list[1]
        self.passwd = list[2]
        # self.file_path = list[3]
        self.file_path = "./projects.xlsx"
        self.download_path = list[3]
        # self.dst_dir = list[5]
        self.dst_dir = "./dst_dir"

        # 获取当前时间
        self.this_date = time.strftime('%Y.%m.%d', time.localtime(time.time())).replace(".", "")
        if self.status_data == 0:
            self.start_time = list[4]
            self.end_time = list[5]
        elif self.status_data == 1:
            self.start_time = int(self.this_date)
            self.end_time = str(self.start_time + 1)
            self.start_time = str(self.start_time)
        elif self.status_data == 2:
            self.start_time = int(self.this_date) - 4
            self.end_time = str(self.start_time + 5)
            self.start_time = str(self.start_time)
        else:
            print("date error")
            exit()

        self.uploader = list[6]
        self.verifier = list[7]

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

    def get_zip(self, project_name):
        self.driver.find_element_by_xpath('''//li[@id = "MENU_PJ"]''').click()
        for i in range(1, len(project_name)):
            print(project_name[i])
            title = "".join(project_name[i])
            self.driver.find_element_by_xpath("//td[@title = '%s']" % title).click()
            self.driver.implicitly_wait(5)
            self.driver.find_element_by_id("project_defect").click()
            self.driver.implicitly_wait(5)
            self.driver.find_element_by_id("all").click()
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
            self.get_back()
            self.get_back()
        self.driver.close()

    def get_back(self):
        self.driver.back()
        self.driver.implicitly_wait(5)

    def read_excel(self):
        data = xlrd.open_workbook(self.file_path)  # 打开excel表格，参数为文件路径
        sheet_names = data.sheet_names()  # 获取所有sheet的名称
        table = data.sheet_by_name(sheet_names[0])  # 通过名称获取表格
        rows = table.nrows  # 获取总行数
        project_name = list(range(rows))
        for i in range(1, rows):  # 获取每一行的数据
            row_content = table.row_values(i)
            project_name[i] = row_content
        return project_name

    def unzip_file(self):
        zip_src = self.download_path
        dst_dir = self.dst_dir
        for filename in os.listdir(zip_src):
            filename = zip_src + "\\" + filename
            r = zipfile.is_zipfile(filename)
            if r:
                fz = zipfile.ZipFile(filename, 'r')
                for file in fz.namelist():
                    fz.extract(file, dst_dir)

            else:
                print('This is not zip')

    def get_data(self):
        _dir = self.dst_dir
        for filename in os.listdir(_dir):
            filename = _dir + "\\" + filename
            # 跳过非xlsx文件
            r = zipfile.is_zipfile(filename)
            if r:
                src = pd.read_excel(filename, header=2, sheet_name=0)
                result = [self.increased_filter(src), self.closed_filter(src), self.still_open_filter(src)]
                print("*" * 50)
                print(filename.split("\\")[-1].split(".")[-2])
                print(self.start_time, "-", self.end_time)
                print("新增\t关闭\t剩余\t")
                print(result)
            else:
                continue

    def test(self, src):
        try:
            data = src.loc[src["提交人"] == self.uploader]
            up_time = "提交时间"

            data[up_time] = pd.to_datetime(data[up_time])
            print(pd.to_datetime(self.start_time))
            print(pd.to_datetime(self.end_time))
            data = data[
                (data[up_time] >= pd.to_datetime(self.start_time)) & (data[up_time] <= pd.to_datetime(self.end_time))]
            return data.shape[0]
        except KeyError:
            print("未找到对应字段")
            return

    def closed_filter(self, src):
        try:
            data = src.loc[src["状态"] == "已关闭"] \
                .loc[src["验证人"] == self.verifier]
            this_time = "验证时间"
            data[this_time] = pd.to_datetime(data[this_time])
            data = data[(data[this_time] >= pd.to_datetime(self.start_time)) & (
                        data[this_time] <= pd.to_datetime(self.end_time))]
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
            data = data[(data[this_time] >= pd.to_datetime(self.start_time)) & (
                        data[this_time] <= pd.to_datetime(self.end_time))]
            return data.shape[0]
        except KeyError:
            print("未找到对应字段")
            return

    def still_open_filter(self, src):
        try:
            # 未关闭
            data = src.loc[(src["状态"] == "待修复") | (src["状态"] == "待回归")]
            return data.shape[0]

        except KeyError:
            print("未找到对应字段")
            return

    def remove_files(self):
        zip_src = self.download_path
        for filename in os.listdir(zip_src):
            r = zipfile.is_zipfile(zip_src + "\\" + filename)
            if r:
                # os.remove(zip_src + "\\" + filename)
                # print(filename, "deleted")
                continue
            else:
                continue
        for filename in os.listdir(self.dst_dir):
            r = zipfile.is_zipfile(self.dst_dir + "\\" + filename)
            # print(filename)
            if r:
                os.remove(self.dst_dir + "\\" + filename)
                print(filename, "deleted")
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
    # 无痕获取
    report.remove_files()
    # 生成报告
    # report.data_to_excel()
