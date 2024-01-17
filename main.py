import openpyxl
import pandas as pd
from translate import Translator
#######################old version##########################################
df = pd.read_excel('output.xlsx', header=0, dtype={'交易账户号': str})
df = df[['客户 ID', '申请日期', '身份证件号码', '全名-英文拼音', '交易账户号', '审核时间', '称谓' ,'出生日期', '中文地址']]
translator = Translator(to_lang="en", from_lang="zh")
# Loop through the rows and print each row
for index, row in df.iterrows():
    workbook = openpyxl.load_workbook('tem.xlsx')
    worksheet = workbook.active
    worksheet[f'B{4}'] = row['全名-英文拼音']
    worksheet[f'B{5}'] = row['出生日期']
    worksheet[f'B{6}'] = row['称谓']
    worksheet[f'B{7}'] = row['交易账户号']
    worksheet[f'B{8}'] = row['审核时间']
    worksheet[f'B{9}'] = 'ID Card'
    worksheet[f'B{10}'] = row['身份证件号码']
    translation = translator.translate(row['中文地址'])
    worksheet[f'B{12}'] = translation
    workbook.save('{0}.xlsx'.format(row['交易账户号']))
