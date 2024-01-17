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

##############################改名部分################################
# Find all files that start with 'KH2024'
kh2024_files = []
for filename in os.listdir(directory_path):
    if filename.endswith('.xlsx') and filename.startswith('KH2024'):
        kh2024_files.append(filename)

# Loop through all the KH2024 files
for kh_file in kh2024_files:
    # Read the KH2024 file
    kh_df = pd.read_excel(os.path.join(directory_path, kh_file))
    kh_value = kh_df.iloc[3, 1]

    # Loop through all the 2024 files
    for filename in os.listdir(directory_path):
        if filename.endswith('.xlsx') and filename.startswith('2024'):
            other_df = pd.read_excel(os.path.join(directory_path, filename))
            other_value = other_df.iloc[3, 1]

            # If the value in cell D4 of the KH2024 file matches the value in cell D4 of the 2024 file
            if kh_value == other_value:
                # Remove the file extension before appending the new suffix
                new_filename = os.path.splitext(kh_file)[0] + '_已完成.xlsx'
                # Rename the KH2024 file to include the name of the 2024 file and the suffix '_已完成'
                os.rename(os.path.join(directory_path, kh_file), os.path.join(directory_path, kh_file[:-5] + ' ' + filename[:-5] + ' 已完成.xlsx'))

