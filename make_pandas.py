#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File  : make_pandas.py
# @Author: shijiu.Xu
# @Date  : 2020/10/22 
# @SoftWare  : PyCharm

import pandas as pd
import openpyxl


class MakePandas():
    """
        封装Pandas对Excel表的处理。
        自用功能
    """
    def __init__(self, file_path):
        self.path = file_path
        self.book = openpyxl.load_workbook(self.path)
        self.excel_writer = pd.ExcelWriter(self.path, engine='openpyxl')
        self.excel_writer.book = self.book
        # self.excel_writer.sheets = dict((ws.title, ws) for ws in self.book.worksheets)

    # 获取指定的sheet表数据
    def get_data(self, sheet): return pd.read_excel(self.path, sheet_name=sheet).values

    # 添加一个空的sheet
    def add_sheet(self, sheet_name, columns):
        data = pd.DataFrame(columns=columns)
        data.to_excel(self.excel_writer, sheet_name=sheet_name, index=False)
        self.excel_writer.save()

    # 在指定的sheet内追加一行数据。
    def append_data_to_sheet(self, sheet_name, row_list):
        sheet = pd.read_excel(self.excel_writer, sheet_name=sheet_name)
        sheet_data = pd.DataFrame(sheet)
        print(sheet_data)
        sheet_data.loc[sheet_data.shape[0]] = row_list  # 与原数据同格式
        sheet_data.to_excel(self.excel_writer, sheet_name=sheet_name, index=False, header=True)
        self.excel_writer.save()