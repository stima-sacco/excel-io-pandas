import os
import pandas as pd
import warnings

class data_cleanup(object):
    def __init__(self, output_folders_source_url, output_folders_destination_url):
        self.output_folders_source_url = output_folders_source_url
        self.output_folders_destination_url = output_folders_destination_url

        f = open('missing-columns.txt', 'w')
        f.close()
        
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
            'OLK',
            'HEAD NB'
        ]

        self.branch_column_ids = [
            'Global_Dimension_1_Code', 
            'branch_Code', 
            'branch', 
            'Branch', 
            'Branch_code', 
            'Branch_Code'
        ]

        #self.create_branch_folders_if()

        print('START...')
        for fl in os.listdir(self.output_folders_source_url):
            source_file_name = ''.join([self.output_folders_source_url, '\\', fl])
            print(fl + '...')
            for branch in self.branches:
                #print('    ' + branch + '...')
                branch_folder_path = ''.join([self.output_folders_destination_url, '\\', branch])
                if os.path.exists(branch_folder_path) == False:
                    os.mkdir(branch_folder_path)

                self.create_brach_data_cleanup_output_file(branch_folder_path, source_file_name)
        print('FINISHED!')
    def create_branch_folders_if(self):
        for branch in self.branches:
            fld_path = ''.join([self.output_folders_destination_url, '\\', branch])

            if os.path.exists(fld_path) == False:
                os.mkdir(fld_path)

    def get_valid_column_key(self, data_frame):
        key_found = False
        valid_id = ''

        for id in self.branch_column_ids:
            if id == 'Branch':
                id = 'Branch'
            
            if id in data_frame.columns:
                valid_id = id
                key_found = True
                break

            else:
                key_found = False
                valid_id = ''

        return_value = {
            'key_found' : key_found,
            'valid_id': valid_id
        }

        return return_value

    def create_brach_data_cleanup_output_file(self, branch_folder_path, source_file_name):
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")

            while True:
                #source_file_name = ''.join([self.output_folders_source_url, '\\', file_name])
                df = pd.read_excel(source_file_name, dtype=object, index_col=0)#, engine='openpyxl')

                return_value = self.get_valid_column_key(df)

                if return_value['key_found'] == False:
                    #print('Missing Column Headers : ' + source_file_name)
                    f = open('missing-columns.txt', 'a')
                    f.writelines(source_file_name +"\n")
                    f.close()
                    break

                branch_name = os.path.basename(branch_folder_path)
                # if branch_name == 'NONBRANCH':
                #     branch_name = 'NaN'

                valid_id = return_value['valid_id']

                if branch_name == 'NONBRANCH':
                    data_frame = df[df[valid_id].isna()]
                else:
                    data_frame = df.loc[df[valid_id] == branch_name]

                if data_frame.empty:
                    break

                output_file_name = ''.join([branch_folder_path, '\\', os.path.basename(source_file_name)])
                writer = pd.ExcelWriter(output_file_name, engine="openpyxl")
                data_frame.to_excel(writer, sheet_name="Sheet1", index=False)
                writer.save()
                break


if __name__ == '__main__':
    output_folders_source_url = r'C:\_Temp\DataCleanup'
    #output_folders_source_url = r'C:\_python\data_cleanup\Source'
    output_folders_destination_url = r'C:\_python\data_cleanup\Destination'

    dc = data_cleanup(
        output_folders_source_url,
        output_folders_destination_url
    )