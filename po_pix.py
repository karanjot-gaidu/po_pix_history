import pandas as pd
import numpy as np
import datetime

# read excel files and make a dataframe
df_po = pd.read_excel('PODL.xlsx')
df_pix = pd.read_excel('Master_Gain_Loss_Data.xlsx')

df_po.dropna(axis = 0, how = 'all', inplace = True)

# make the first row as columns and update the dataframe values with rest
df_pix.columns = df_pix.iloc[0]
df_pix = df_pix[1:]

# filter the pix dataframe based on SKU
df_filtered_pix = df_pix[df_pix.SKU.isin(df_po.SKU)]

# remove columns which are NaN
df_filtered_pix = df_filtered_pix.iloc[:,df_filtered_pix.columns.notna()]

# keep first 2 and last 15 columns (SKU, Description, last 14 days, Total)
df_filtered_pix.drop(df_filtered_pix.iloc[:,2:-15], axis=1, inplace=True)
df_filtered_pix

# add variance column onto filtered data by merging the dataframes and assign to new dataframe
final_df = pd.merge(df_filtered_pix,df_po[['PO Number','SKU','Variance']],on='SKU',how='left')

# remove the time from column headings
final_df.rename(columns={col: col.strftime('%Y-%m-%d') if isinstance(col, datetime.datetime) else col for col in final_df.columns}, inplace=True)

# move PO Number to the front
final_df = final_df[['PO Number'] + [c for c in final_df if c not in['PO Number']]]

# Update the total with sum of last 10 days pix
final_df['TOTAL'] = final_df.iloc[:, 3:-2].sum(axis=1)

# remove all NaN values
final_df.fillna("",inplace=True)

final_df.sort_values(by=['PO Number', 'SKU'], ascending=[True, True], inplace=True)
print(final_df)



final_df.to_excel('result.xlsx', index=False)
