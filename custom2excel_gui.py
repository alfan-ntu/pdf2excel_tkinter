"""
custom2excel_gui.py
============

This is the GUI entry point of the extraction and conversion tool which extracts 稅單號碼，
報單號碼，納稅義務人統編 and 稅金金額 from 彙總稅單稅單清單.

.. note::
    This documentation follows the reStructuredText (reST) format, which is
    compatible with Sphinx autodoc.

:author: Maoyi Fan
:email: maoyi.fan@yapro.com.tw
:date: 2025-03-09
:version: 1.1
:license: MIT
:history:
    - 1.1 (2025-03-09) - To handle exceptional case where only one record on the
                         last page
Example:
      To run this script with or without optimization flag -O
      - no debug information nor intermediate text files
      > python custom2excel_gui.py

      - with debug information, debug_output.txt and intermediate text files, including
      - 彙總清單.text : complete text information in the 彙總清單.pdf
      - 彙總清單.txt : 稅單 related text information formed in different columns to generate
                      its correspondent Excel file
      > python -O custom2excel_gui.py

      To run the executable of this GUI tool, ensure subdirectory .\\images is in the same
      directory as that custom2excel_gui.exe is in
      > .\\custom2excel_gui.exe

"""
#
# References:
#    http://infohost.nmt.edu/tcc/help/pubs/tkinter//dialogs.html#tkFileDialog
#    http://tkinter.unpythonic.net/wiki/tkFileDialog
#
# Note:
#    there are a variety of options for the FileDialog
#    see above references for more information

from tkinter import *
from tkinter import ttk

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
