import os
import pandas as pd

# Set the directory path where the xlsx files are located
directory_path = 'C:\\Users\\KohleKen\\OneDrive - Kohle Services Limited\\桌面\\出金報表'

# Loop through all the files in the directory
for filename in os.listdir(directory_path):
    if filename.endswith('.xlsx'):
        # Read the xlsx file
        df = pd.read_excel(os.path.join(directory_path, filename), engine='openpyxl')
        # Print the file name
        print(f'{df.iloc[4, 1]}筆  {filename} {df.iloc[3, 1]}')
        #sum_of_df = df[4].sum()
        #print(f"The sum of column '4' is {sum_of_df}.")

#查表function
target=input("target:")
for filename in os.listdir(directory_path):
    if filename.endswith('.xlsx'):
        # Read the xlsx file
        print(f' {filename} ')
        df = pd.read_excel(os.path.join(directory_path, filename), skiprows=7, dtype={4: str})
        # find rows that contain the value 5
        rows_with_value = df.loc[df.eq(target).any(axis=1)]
        # if no rows contain the value, print "NA"
        if rows_with_value.empty:
            print("NA")
        else:
            # print the resulting rows
            print(f"Rows containing the value:\n{rows_with_value}")
input()