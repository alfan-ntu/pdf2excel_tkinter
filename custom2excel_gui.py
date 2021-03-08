# File: fileseldlg.py
#    http://infohost.nmt.edu/tcc/help/pubs/tkinter//dialogs.html#tkFileDialog
#    http://tkinter.unpythonic.net/wiki/tkFileDialog
#
# Note:
#    there are a variety of options for the FileDialog
#    see above references for more information

from tkinter import *
from tkinter import ttk
import tkinter.filedialog as fdlg
from tkinter import messagebox
import utilities

from gui_construct import MsgPanel, OperationPanel, FileSelectionPanel


# Create a new FileSelDlgDemo class based on the class ttk.Frame
class FileSelDlgDemo(ttk.Frame):
    def __init__(self, isapp=True, name='fileseldlgdemo'):
        ttk.Frame.__init__(self, name=name)
        self.pack(expand=Y, fill=BOTH)
        self.master.title('彙總稅單轉檔工具')
        self.isapp = isapp
        self._create_widgets()

    # actual constructs the GUI by calling class instructors defined in gui_construct.py
    def _create_widgets(self):
        if self.isapp:
            # MsgPanel(self,
            #          ["Enter a file name in the entry box or click on the 'Browse' ",
            #           "buttons to select a file name using the file selection dialog."])
            MsgPanel(self, ["雅博會計師事務所\n",
                            "海關稅單彙總清單轉檔工具"])
            fs_panel = FileSelectionPanel(self)
            OperationPanel(self, fs_panel)



if __name__ == '__main__':
    FileSelDlgDemo().mainloop()
