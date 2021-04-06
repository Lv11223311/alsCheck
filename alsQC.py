# -*- coding: utf-8 -*-
# @Time     :2021/4/6 11:22
# @Author   ：Jinbo.lv


import pandas as pd

inputFile = input("输入文件")
outputFile = input("输入文件")

# 读取各sheet
Forms = pd.read_excel(inputFile, sheet_name="Forms")
Fields = pd.read_excel(inputFile, sheet_name="Fields")
DataDictionaries = pd.read_excel(inputFile, sheet_name="DataDictionaries")
DataDictionaryEntries = pd.read_excel(inputFile, sheet_name="DataDictionaryEntries")
Checks = pd.read_excel(inputFile, sheet_name="Checks")
CheckSteps = pd.read_excel(inputFile, sheet_name="CheckSteps")
CheckActions = pd.read_excel(inputFile, sheet_name="CheckActions")
Derivations = pd.read_excel(inputFile, sheet_name="Derivations")
Matrices = pd.read_excel(inputFile, sheet_name="Matrices")