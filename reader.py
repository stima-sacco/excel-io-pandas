import os
import pandas as pd
import warnings

class data_cleanup(object):
    def __init__(self, base_url):
        self.base_url = base_url

        self.branches = [
            'CBD',
            'ELD',
            'EMB',
            'KAWI',
            'KSM',
            'MSA',
            'NBI',
            'NKR',
            'NONBRANCH',
            'OLK'
        ]

        self.branch_column_ids = [
            'Global_Dimension_1_Code', 
            'branch_Code', 
            'branch', 
            'Branch', 
            'Branch_code', 
            'Branch_Code'
        ]

    def create_branch_folders(self):
        for branch in self.branches:
            fld_path = ''.join([self.base_url, '\\', branch])

            if os.path.exists(fld_path) == False:
                os.mkdir(fld_path)

    def create_brach_data_cleanup_output_file(self, file_name):
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            df = pd.read_excel(r'C:\_Temp\DataCleanup\Members_with_Duplicate_P_I_N_NUMBERS.xls', dtype=object, index_col=0)#, engine='openpyxl')

            data_frame = df.loc[df['Branch'] == 'CBD']

            writer = pd.ExcelWriter('demo.xls', engine="openpyxl")
            data_frame.to_excel(writer, sheet_name="Sheet1", index=False)
            writer.save()