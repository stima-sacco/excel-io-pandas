import pandas as pd
import warnings

branch_column_ids = [
    'Global_Dimension_1_Code', 
    'branch_Code', 
    'branch', 
    'Branch', 
    'Branch_code', 
    'Branch_Code'
]
with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")
    df = pd.read_excel(r'C:\_Temp\DataCleanup\Members_with_Duplicate_P_I_N_NUMBERS.xls', dtype=object, index_col=0)#, engine='openpyxl')
    #print(df['Branch'])
    data_frame = df.loc[df['Branch'] == 'CBD']

    writer = pd.ExcelWriter('demo.xls', engine="openpyxl")
    data_frame.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.save()