import fitz
import openpyxl
import pandas as pd

df = pd.read_excel('wrong list.xlsx',header=0)
print(df)

# Loop through the rows and print each row
for index, row in df.iterrows():
    doc = fitz.open(row[0])
    rect = fitz.Rect(270, 660, 370, 710)
    for Page in doc:
        Page.insert_image(rect, filename="signature.png")
    doc.save('C:\\Users\\KohleKen\\PycharmProjects\\pythonProject\\\pdf update\\{0}'.format(row[0]))
    #workbook.save('{0}.xlsx'.format(row[3]))






'''
doc = fitz.open("testpdf.pdf")
rect = fitz.Rect(270,660,370,710)
for Page in doc:
    Page.insert_image(rect, filename = "signature.png")
doc.save("result.pdf")
'''