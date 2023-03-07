import pandas as pd
import os

# set the directory path
folder_path = r'd:\FineTec\00_working\pmis\\'

# create an empty list to store the dataframes
dfs = []

# loop through each file and append the "일일출역현황" dataframe to the list
for file in os.listdir(folder_path):
    if file.endswith('.xlsx'):
        filepath = os.path.join(folder_path, file)
        df = pd.read_excel(filepath, sheet_name='일일출역현황', encoding='cp949')
        dfs.append(df)

# combine the dataframes into a single dataframe
combined_df = pd.concat(dfs, ignore_index=True)

# write the combined data to a new .xlsx file
combined_df.to_excel(os.path.join(folder_path, 'combined.xlsx'), index=False)
