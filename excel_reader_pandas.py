import os
import pandas as pd
import warnings
import wx
import pythoncom

class Form(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None, title='Data Cleanup', size=(720,400))
        self.Pan = wx.Panel(self, -1)

        self.lcMissingColumns = []
        self.output_folders_source_url = ''
        self.output_folders_destination_url = ''

        self.source_path_selected = False        
        self.destination_path_selected = False
        # f = open('missing-columns.txt', 'w')
        # f.close()
        
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
            'HEAD NB',
            'CHANEXPE'
        ]

        self.branch_column_ids = [
            'Global_Dimension_1_Code', 
            'branch_Code', 
            'branch', 
            'Branch', 
            'Branch_code', 
            'Branch_Code',
            'Transaction_Branch',
            'BRANCH'
        ]
        # f = open('missing-columns.txt', 'w')
        # f.close()

        self.lblSourcePath = wx.StaticText(self.Pan, label='Source Path', pos=(10, 20), size=(70, 20))
        #self.setCustomFont(self.lblSourcePath)
        self.txtSourcePath = wx.TextCtrl(self.Pan, pos=(110, 20), size=(430, 20))
        self.btnSelectSourcePath = wx.Button(self.Pan, id=1, label='...', pos=(545, 20), size=(20,20))
        self.btnSelectSourcePath.Bind(wx.EVT_BUTTON, self.evt_selectPath)

        self.lblDestinationPath = wx.StaticText(self.Pan, label='Destination Path', pos=(10, 45), size=(90, 20))
        self.txtDestinationPath = wx.TextCtrl(self.Pan, pos=(110, 45), size=(430, 20))
        self.btnSelectDestinationPath = wx.Button(self.Pan, id=2, label='...', pos=(545, 45), size=(20,20))
        self.btnSelectDestinationPath.Bind(wx.EVT_BUTTON, self.evt_selectPath)
        
        self.lblProgress = wx.StaticText(self.Pan, label='Progress', pos=(10, 80), size=(90, 20))
        self.tcProgress = wx.TextCtrl(self.Pan, pos=(110, 80), size=(530, 70), style=wx.TE_MULTILINE|wx.TE_READONLY)
        
        self.lblNoColumnFiles = wx.StaticText(self.Pan, label='No Column Files', pos=(10, 160), size=(90, 20))
        self.tcNoColumnFiles = wx.TextCtrl(self.Pan, pos=(110, 160), size=(530, 70), style=wx.TE_MULTILINE|wx.TE_READONLY)

        self.btnExtract = wx.Button(self.Pan, id=3, label='Extract', pos=(10, 330), size=(70,20))
        self.btnExtract.Enable(False)
        self.btnExtract.Bind(wx.EVT_BUTTON, self.evt_selectPath)

        self.btnClose = wx.Button(self.Pan, id=4, label='Close', pos=(620, 330), size=(70,20))
        self.btnClose.Bind(wx.EVT_BUTTON, self.evt_selectPath)

        self.Pan.Bind(wx.EVT_SIZE, self.Evt_Resize)

    def Evt_Resize(self, evt):
        nHeight = self.GetSize()[1]
        nWidth = self.GetSize()[0]
        
        width = int((self.GetSize().Width / 2) + 200)
        self.txtSourcePath.SetSize((width, 20))
        self.txtDestinationPath.SetSize((width, 20))
        # self.tcProgress.SetSize((width, 70))
        # self.tcNoColumnFiles.SetSize((width, 70))

        button_left = self.txtSourcePath.Position[0] + self.txtSourcePath.Size[0] + 5
        self.btnSelectSourcePath.SetPosition((button_left, 20))
        self.btnSelectDestinationPath.SetPosition((button_left, 45))

        self.btnExtract.SetPosition((10, (nHeight -65)))
        self.btnClose.SetPosition(((nWidth - 100), (nHeight -65)))

        self.lblNoColumnFiles.SetPosition((10, int(nHeight/2) + 20))
        self.tcNoColumnFiles.SetPosition((110, int(nHeight/2) + 20))
        self.tcNoColumnFiles.SetSize((width, int(nHeight/2 - 90)))

        self.tcProgress.SetSize((width, int(nHeight/2 - 70)))
    def evt_selectPath(self, evt):
        button_id = evt.GetId()

        if button_id == 1:
            dlg = wx.DirDialog(self, message="Select Source Folder")
            if dlg.ShowModal() == wx.ID_OK:
                self.source_path_selected = True
                if self.source_path_selected and self.destination_path_selected:
                    self.btnExtract.Enable(True)
                dirname = dlg.GetPath()
                self.output_folders_source_url = dirname
                self.txtSourcePath.Value = dirname
                
            dlg.Destroy()

        elif button_id == 2:
            dlg = wx.DirDialog(self, message="Select Destination Folder")
            if dlg.ShowModal() == wx.ID_OK:
                self.destination_path_selected = True
                if self.source_path_selected and self.destination_path_selected:
                    self.btnExtract.Enable(True)
                dirname = dlg.GetPath()
                self.txtDestinationPath.Value = dirname
                self.output_folders_destination_url = dirname
                
            dlg.Destroy()

        elif button_id == 3:
            self.startExtraction()
    
        elif button_id == 4:
            self.Close()
            
    def setCustomFont(self, control):
        font = wx.Font(15, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        font_clr = wx.Colour(60, 60, 60, alpha=wx.ALPHA_OPAQUE)
        
        #control.SetTextForeground(font_clr)
        control.SetFont(font)

    def startExtraction(self):
        self.tcProgress.Value = self.tcProgress.Value + '\n' + 'START...'
        for fl in os.listdir(self.output_folders_source_url):
            source_file_name = ''.join([self.output_folders_source_url, '\\', fl])
            self.tcProgress.Value = self.tcProgress.Value + '\n' + fl
            for branch in self.branches:
                #print('    ' + branch + '...')
                branch_folder_path = ''.join([self.output_folders_destination_url, '\\', branch])
                if os.path.exists(branch_folder_path) == False:
                    os.mkdir(branch_folder_path)

                self.create_brach_data_cleanup_output_file(branch_folder_path, source_file_name)
                pythoncom.PumpWaitingMessages()

        lc = set(self.lcMissingColumns)

        for fl in lc:
            self.tcNoColumnFiles.Value = self.tcNoColumnFiles.Value + '\n' + fl

        self.tcProgress.Value = self.tcProgress.Value + '\n' + 'FINISHED!'
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
                #df = pd.read_excel(source_file_name, dtype=object, index_col=0)#, engine='openpyxl')
                df = pd.read_excel(source_file_name, dtype=object)#, index_col=0)#, engine='openpyxl')

                return_value = self.get_valid_column_key(df)

                if return_value['key_found'] == False:
                    #print('Missing Column Headers : ' + source_file_name)
                    self.lcMissingColumns.append(source_file_name)
                    # f = open('missing-columns.txt', 'a')
                    # f.writelines(source_file_name +"\n")
                    # f.close()
                    break

                branch_name = os.path.basename(branch_folder_path)
                # if branch_name == 'NONBRANCH':s
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
                data_frame.to_excel(writer, sheet_name="Sheet1")#, index=False)
                writer.save()
                break

if __name__ == '__main__':
    app = wx.App()
    f = Form().Show()
    app.MainLoop()
# if __name__ == '__main__':
#     output_folders_source_url = r'C:\_Temp\DataCleanup'
#     #output_folders_source_url = r'C:\_python\data_cleanup\Source'
#     output_folders_destination_url = r'C:\_python\data_cleanup\Destination'

#     dc = data_cleanup(
#         output_folders_source_url,
#         output_folders_destination_url
#     )