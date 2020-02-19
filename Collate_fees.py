#!/usr/bin/env python

'''
Collate_fees.ipynb
A. J. McCulloch, February 2020
'''

####################################################################################################
# Import modules
####################################################################################################

import pandas as pd # Required for dataframes
import numpy as np # Required from numerical operations

file = '2016_Fees.xlsx' # File to import
xlsx = pd.ExcelFile(file) # Open the Excel file

# Concatinate all sheets from the excel spreadsheet into a dataframe
for count, s in enumerate(sheets):
    if count == 0: # Initialise the dataframe with the first sheet
        df = pd.read_excel(file, s, header=1, index_col = 2, usecols = 'A:J', skipfooter = 18)
        df = df.drop(np.NaN)
        df['Course'] = s
    else: # Join all other data frames to the inital frame
        df_temp = pd.read_excel(file, s, header=1, index_col = 2, usecols = 'A:J', skipfooter = 18)
        df_temp = df_temp.drop(np.NaN)
        df_temp['Course'] = s

        df = pd.concat([df, df_temp]) # Concatinate the frames

# Get the cost of living data which is differently formatted
df_CoL = pd.read_excel(file, xlsx.sheet_names[-1], header=3, index_col = 0, usecols = 'H:J', skipfooter = 0)

df = pd.merge(df, df_CoL, left_index=True, right_index=True) # Merge the cost of living
df.index.name = 'INST' # Rename the index as it is lost during the merge

df = df.set_index(['Course', df.index]) # Set a multi-index index
df = df.sort_values(by=['Course', 'INST']) # Sort the data by Course then institution

df.to_csv('Cleaned-fee.csv') # Export the cleaned file to csv
