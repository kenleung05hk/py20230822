from win32com import client
import pandas as pd
####################第三步##############################

df = pd.read_excel('output.xlsx', header=0)
df = df[['交易账户号']]
print(df)

# Loop through the rows and print each row
for index, row in df.iterrows():
    excel = client.Dispatch("Excel.Application")
    excel.Visible = False #Visible
    sheets = excel.Workbooks.Open('C:\\Users\\KohleKen\\PycharmProjects\\pythonProject\\{0}'.format(row[0]))
    work_sheets = sheets.Worksheets[0]
    work_sheets.ExportAsFixedFormat(0, 'C:\\Users\\KohleKen\\PycharmProjects\\pythonProject\\pdf update\\{0}'.format(row[0]))
    excel.Quit()

    #workbook.save('{0}.xlsx'.format(row[3]))



