import os
import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('C:\\Users\\KohleKen\\PycharmProjects\\pythonProject\\pdf update\\output.xlsx')

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    # Get the old and new file names
    old_name = os.path.join('C:\\Users\\KohleKen\\PycharmProjects\\pythonProject\\pdf update', row['Old Name'])
    new_name = os.path.join('C:\\Users\\KohleKen\\PycharmProjects\\pythonProject\\pdf update', row['New Name'])

    # Rename the file
    os.rename(old_name, new_name)