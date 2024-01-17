import openpyxl
import pandas as pd
from translate import Translator
####################第二步##############################
df = pd.read_excel('output.xlsx', header=0, dtype={'交易账户号': str})
df = df[['客户 ID', '申请日期', '身份证件号码', '全名-英文拼音', '交易账户号', '审核时间' ,'出生日期','中文地址']]
translator = Translator(to_lang="en", from_lang="zh")
# Loop through the rows and print each row
for index, row in df.iterrows():
    workbook = openpyxl.load_workbook('tem2.xlsx')
    worksheet = workbook.active
    worksheet[f'C{4}'] = row['全名-英文拼音']
    worksheet[f'C{5}'] = row['客户 ID']
    worksheet[f'C{6}'] = row['申请日期']
    worksheet[f'C{7}'] = row['出生日期']
    worksheet[f'A{25}'] = 'Date: {0}'.format(row['审核时间'])
    worksheet[f'B{18}'] = 'Date: {0}'.format(row['审核时间'])
    translation = translator.translate(row['中文地址'])
    worksheet[f'C{12}'] = translation
    workbook.save('{0}.xlsx'.format(row['交易账户号']))
