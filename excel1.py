import openpyxl
import pandas as pd
####################第一步##############################
#input = source.xlsx
#output = output.xlsx
df = pd.read_excel('source.xlsx', header=0, dtype={'交易账户号': str})
df = df[['客户 ID', '申请日期', '身份证件号码', '全名-英文拼音', '交易账户号', '审核时间', '称谓' ,'出生日期', '中文地址']]
for index, row in df.iterrows():
    if row['称谓'] == 'Mr':
        df.at[index, '称谓'] = 'male'
    elif row['称谓'] == 'Ms':
        df.at[index, '称谓'] = 'female'
    if pd.isna(row['审核时间']):
        df.at[index, '审核时间'] = ''
    else:
        df.at[index, '审核时间'] = str(row['审核时间'].strftime('%Y-%m-%d'))
    if pd.isna(row['申请日期']):
        df.at[index, '申请日期'] = ''
    else:
        df.at[index, '申请日期'] = str(row['申请日期'].strftime('%Y-%m-%d'))
        #print(df.at[index, '审核时间'])
df.dropna(subset=['审核时间'], inplace=True)

#print(df)
#headernames = ["acc", "finish_day"]
#df2 = pd.read_excel('working record.xlsx', header=None,names=headernames,dtype={'acc': str})
#print(df2)

#df3 = df[df['交易账户号'].str.contains('|'.join(df2['acc'].astype(str)), na=False)]
#print(df3)
df.to_excel('output.xlsx', index=False)
'''
# Loop through the rows and print each row
for index, row in df.iterrows():
    workbook = openpyxl.load_workbook('tem.xlsx')
    worksheet = workbook.active
    worksheet[f'B{4}'] = row[0]
    worksheet[f'B{5}'] = row[1]
    worksheet[f'B{6}'] = row[2]
    worksheet[f'B{7}'] = row[3]
    worksheet[f'B{8}'] = row[4]
    worksheet[f'B{9}'] = row[5]
    worksheet[f'B{10}'] = row[6]
    workbook.save('{0}.xlsx'.format(row[3]))

'''