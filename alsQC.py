# -*- coding: utf-8 -*-
# @Time     :2021/4/6 11:22
# @Author   ：Jinbo.lv
# 测试文件：E:\pythonProject\Jupyter notebook\CRF_Library_Phase1-3_Oncology_Draft.xlsx

import pandas as pd

inputFile = input ("输入文件")

# 读取各sheet
# df = pd.ExcelFile(inputFile)
# print(df.sheet_names) # 读取表名称
Forms = pd.read_excel (inputFile, sheet_name="Forms")
Fields = pd.read_excel (inputFile, sheet_name="Fields")
DataDictionaries = pd.read_excel (inputFile, sheet_name="DataDictionaries")
DataDictionaryEntries = pd.read_excel (inputFile, sheet_name="DataDictionaryEntries")
Checks = pd.read_excel (inputFile, sheet_name="Checks")
CheckSteps = pd.read_excel (inputFile, sheet_name="CheckSteps")
CheckActions = pd.read_excel (inputFile, sheet_name="CheckActions")
Derivations = pd.read_excel (inputFile, sheet_name="Derivations")
Matrices = pd.read_excel (inputFile, sheet_name="Matrices")

# Forms 校验
Forms = Forms.drop (Forms[Forms.IsSignatureRequired == "True"].index)
Forms['issue'] = "Signature required."

# DataDictionaryEntries预处理
pattern = "^(\-|\+)?\d+(\.\d+)?$"
DataDictionaryEntries = DataDictionaryEntries.drop(\
    DataDictionaryEntries[DataDictionaryEntries.DataDictionaryName == ""].index)\
    [['DataDictionaryName', 'CodedData', 'Ordinal', 'UserDataString', 'Specify']]

dic1 = DataDictionaries['DataDictionaryName']

dic2 = DataDictionaryEntries.copy()
dic2['OrdinalDic'] = dic2['Ordinal']
dic2['FormatLength'] = dic2['CodedData'].apply(str).str.len()
dic2['CodeType'] = None
dic2.loc[dic2['CodedData'].apply(str).str.match(pattern)==True,'CodeType'] = "Num"
dic2.loc[dic2['CodedData'].apply(str).str.match(pattern)==False,'CodeType'] = "Char"
dic2 = dic2[['DataDictionaryName', 'CodedData', 'UserDataString', 'Specify', 'OrdinalDic', 'FormatLength', 'CodeType']]

dic3 = dic2.copy()
dic3['CodedData'] = dic2['CodedData'].apply(str)
dic3 = dic3.groupby(['DataDictionaryName'])['CodedData'].apply('|'.join).reset_index()
dic3['Specify'] = None
dic3['FormatLength'] = None
dic3['FormatDic'] = None
for i in range(len(dic3)):
    # print("判断有无Specify")
    if True in list(dic2.loc[dic2['DataDictionaryName']==dic3['DataDictionaryName'][i]]['Specify']):
        dic3.loc[i, 'Specify'] = True
    else:
        dic3.loc[i, 'Specify'] = False
    # print("======================")
    # print("判断最大的FormatLength")
    dic3.loc[i, 'FormatLength'] = max(list(dic2.loc[dic2['DataDictionaryName']==dic3['DataDictionaryName'][i]]['FormatLength']))
    # print("======================")
    # print("创建FormatDic")
    if 'Char' in list(dic2.loc[dic2['DataDictionaryName']==dic3['DataDictionaryName'][i]]['CodeType']):
        dic3.loc[i, 'FormatDic'] = '$' + str(dic3.loc[i, 'FormatLength'])
    else:
        dic3.loc[i, 'FormatDic'] = str(dic3.loc[i, 'FormatLength'])


# Fields 预处理
Fields.DataFormatResult = None
Fields['DataFormatChar'] = ''
Fields.loc[Fields.DataFormat.str.match(pattern)==True,'DataFormatResult'] = "Num"
Fields.loc[Fields.DataFormat.str.match("^\$")==True,'DataFormatResult'] = "Char"
Fields.loc[Fields.DataFormat.str.match("^\$")==True,'DataFormatChar'] = "$"
Fields['DataFormat1'] = None
Fields['DataFormat1'] = Fields['DataFormat1'].mask(Fields['DataFormatResult'] == "Num",  Fields.DataFormat)
Fields['DataFormat1'] = Fields['DataFormat1'].mask(Fields['DataFormatResult'] == "Char", Fields.DataFormat.str[1:])
Fields.DefaultValueResult = None
Fields.loc[Fields.DefaultValue.str.match(pattern)==True,'DefaultValueResult'] = "Num"
Fields.loc[Fields.DefaultValue.str.match(pattern)==False,'DefaultValueResult'] = "Char"
Fields['DefaultValue1'] = None  # Num控件默认值长度
Fields['DefaultValue1'] = Fields['DefaultValue1'].mask(Fields['DefaultValueResult'] == 'Num', Fields['DefaultValue'].str.strip('-.').str.len())
Fields['DefaultValue1'] = Fields['DefaultValue1'].mask(Fields['DefaultValueResult'] == 'Char', Fields['DefaultValue'].str.len())
Fields = Fields[['FormOID','FieldOID','Ordinal','DraftFieldNumber', 'DraftFieldName', 'DraftFieldActive' ,'VariableOID', 'DataFormat', 'DataDictionaryName', 'UnitDictionaryName', 'CodingDictionary', 'ControlType', 'AcceptableFileExtensions' ,'IndentLevel',
         'PreText', 'FixedUnit', 'HeaderText', 'HelpText', 'SourceDocument', 'IsLog', 'DefaultValue', 'SASLabel', 'SASFormat', 'EproFormat', 'IsRequired', 'QueryFutureDate', 'IsVisible', 'IsTranslationRequired', 'AnalyteName', 'IsClinicalSignificance',
         'QueryNonConformance', 'OtherVisits', 'CanSetRecordDate', 'CanSetDataPageDate', 'CanSetInstanceDate', 'CanSetSubjectDate', 'DoesNotBreakSignature', 'LowerRange', 'UpperRange', 'NCLowerRange', 'NCUpperRange', 'ViewRestrictions',
         'EntryRestrictions', 'ReviewGroups', 'IsVisualVerify', 'FDownloadedFromObjectId', 'FSourceObjectId', 'VDownloadedFromObjectId', 'VSourceObjectId', 'FSourceUrlId', 'VSourceUrlId', 'DataFormatResult', 'DataFormatChar', 'DataFormat1', 'DefaultValue1']]

Fields = Fields.drop(Fields[Fields.AnalyteName==""].index)
Fields['isLabForm'] = 1

Fields = Fields.drop(Fields[Fields.VariableOID==""].index)
Fields['isDupSasLabel'] = 1

Forms.rename(columns={'OID': 'FormOID'})
Check_Fields = pd.merge(Fields, dic3,on='DataDictionaryName',how='outer')
Check_Fields = pd.merge(Check_Fields, Derivations, on=['VariableOID', 'FormOID', 'FieldOID'],how='outer')
