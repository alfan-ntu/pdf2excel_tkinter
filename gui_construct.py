# File: gui_construct.py
# References:
#    http://hg.python.org/cpython/file/4e32c450f438/Lib/tkinter/simpledialog.py
#    http://docs.python.org/py3k/library/inspect.html#module-inspect
#
# Icons sourced from:
#    http://findicons.com/icon/69404/deletered?width=16#
#    http://findicons.com/icon/93110/old_edit_find?width=16#
#
# This file is imported by the Tkinter Demos

from tkinter import *
from tkinter import ttk
from tkinter.simpledialog import Dialog
from tkinter import messagebox
from PIL import Image, ImageTk
import tkinter.filedialog as fdlg
import utilities
import constant
import customs_pdf_2_excel
import subprocess
import os
import pdb


class MsgPanel(ttk.Frame):
    def __init__(self, master, msgtxt):
        ttk.Frame.__init__(self, master)
        self.pack(side=TOP, fill=X)
        msg = Label(self, wraplength='4i', justify=CENTER, font="標楷體 14")
        msg['text'] = ''.join(msgtxt)
        msg.pack(fill=X, padx=5, pady=5)


class FileSelectionPanel(ttk.Frame):
    def __init__(self, master):
        self.fsel_panel = Frame(master)
        self.fsel_panel.pack(side=TOP, fill=BOTH, expand=Y)
        self.upper_frame = ttk.Frame(self.fsel_panel)
        self.bottom_frame = ttk.Frame(self.fsel_panel)
        # draw the file selection label/entry and button in the upper frame
        self.lbl_input = ttk.Label(self.upper_frame, width=constant.LABEL_WIDTH,
                                   text="選取彙總稅單清單", font="標楷體 12")
        self.ent_input = ttk.Entry(self.upper_frame, width=constant.ENTRY_WIDTH)
        self.btn = ttk.Button(self.upper_frame, text="選擇",
                              command=lambda i="open", e=self.ent_input: self._file_dialog(i, e))
        self.lbl_input.pack(side=LEFT)
        self.ent_input.pack(side=LEFT, expand=Y, fill=X)
        self.btn.pack(side=LEFT, padx=constant.H_SPACE, fill=X)
        self.upper_frame.pack(fill=X, padx='1c', pady=3)
        # draw the label and entry in the bottom_frame to accept output file name
        self.lbl_output = ttk.Label(self.bottom_frame, width=constant.LABEL_WIDTH,
                                    text="彙總清單Excel名稱", font="標楷體 12")
        self.ent_output = ttk.Entry(self.bottom_frame, width=constant.ENTRY_WIDTH)
        self.lbl_output.pack(side=LEFT)
        self.ent_output.pack(side=LEFT)
        self.bottom_frame.pack(fill=X, padx='1c', pady=3)
        # setup class attribute fn
        self.fn = None

    # File select button call-back function
    def _file_dialog(self, type, ent):
        # pdb.set_trace()
        print("File selector button is clicked")
        opts = {'initialfile': self.ent_input.get(),
                'filetypes': (('Adobe Acrobat Document', '.pdf'),
                              ('All files', '.*'))}
        opts['title'] = '選取轉檔彙總清單...'
        self.fn = fdlg.askopenfilename(**opts)
        fn = self.fn
        # debug purpose
        print("Selected pdf file for conversion is ", fn)
        # update the entry
        if fn:
            self.ent_input.delete(0, END)
            self.ent_input.insert(END, fn)
            # debug purpose only
            base_fn = utilities.extract_file(fn)
            fn_ext = utilities.extract_file_extension(base_fn)
            fn_file = utilities.extract_file_name(base_fn)
            output_filename = fn_file + ".xlsx"
            # debug purpose only
            print("Output file name is ", output_filename)
            print("Output directory is ", utilities.extract_directory(fn))
            self.ent_output.delete(0, END)
            self.ent_output.insert(END, output_filename)


class OperationPanel(ttk.Frame):
    def __init__(self, master, fs_panel):
        ttk.Frame.__init__(self, master)
        self.pack(side=BOTTOM, fill=X)  # resize with parent
        # separator widget
        sep = ttk.Separator(orient=HORIZONTAL)
        # Dismiss button
        im = Image.open('images//exit.png')  # image file
        imh = ImageTk.PhotoImage(im)  # handle to file
        dismissBtn = ttk.Button(text='離開', image=imh, command=self.winfo_toplevel().destroy)
        dismissBtn.image = imh  # prevent image from being garbage collected
        dismissBtn['compound'] = LEFT  # display image to left of label text

        # 'See Code' button
        im = Image.open('images//view.png')
        imh = ImageTk.PhotoImage(im)
        # codeBtn = ttk.Button(text='檢視', image=imh, command=lambda: view_output(self.master))
        codeBtn = ttk.Button(text='檢視', image=imh, command=lambda fp=fs_panel: view_output(fp))
        codeBtn.image = imh
        codeBtn['compound'] = LEFT

        # "Convert" button
        im = Image.open('images//convert.png')
        imh = ImageTk.PhotoImage(im)
        convertBtn = ttk.Button(text='轉換', image=imh, default=ACTIVE,
                                command=lambda fp=fs_panel: convert_customs_pdf(fp))
        convertBtn.image = imh
        convertBtn['compound'] = LEFT
        convertBtn.focus()

        # About button
        im = Image.open('images//info.png')
        imh = ImageTk.PhotoImage(im)
        abtBtn = ttk.Button(text='關於', image=imh, command=lambda: show_about_info())
        abtBtn.image = imh
        abtBtn['compound'] = LEFT

        # position and register widgets as children of this frame
        sep.grid(in_=self, row=0, columnspan=4, sticky=EW, pady=5)
        abtBtn.grid(in_=self, row=1, column=0, sticky=W, padx=5, pady=5)
        convertBtn.grid(in_=self, row=1, column=1, sticky=E, padx=5, pady=5)
        codeBtn.grid(in_=self, row=1, column=2, sticky=E, padx=5, pady=5)
        dismissBtn.grid(in_=self, row=1, column=3, sticky=E, padx=5, pady=5)

        # set resize constraints
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # bind <Return> to demo window, activates 'See Code' button;
        # <'Escape'> activates 'Dismiss' button
        self.winfo_toplevel().bind('<Return>', lambda x: codeBtn.invoke())
        self.winfo_toplevel().bind('<Escape>', lambda x: dismissBtn.invoke())


def show_about_info():
    messagebox.showinfo(title="海關稅單彙總清單轉檔工具",
                        message="1. 選取PDF格式海關稅單彙總清單\n2. 點選轉換按鍵進行轉檔\n3. 點選檢視按鍵進行檢視")


def convert_customs_pdf(fp):
    working_dir = utilities.extract_directory(fp.fn)
    input_fn = utilities.extract_file(fp.fn)
    output_fn = fp.ent_output.get()
    customs_pdf_2_excel.pdf_2_excel(working_dir, input_fn, output_fn)
    msg = "PDF轉檔完成-"+output_fn+"\n使用檢視功能檢視轉出Excel"
    messagebox.showinfo(title="海關稅單彙總清單轉檔工具", message=msg)


def view_output(fp):
    working_dir = utilities.extract_directory(fp.fn)
    output_fn = fp.ent_output.get()
    target_excel = working_dir + "/" + output_fn
    if os.path.exists('C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'):
        subprocess.call(['C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE', target_excel])
    elif os.path.exists('C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE'):
        subprocess.call(['C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE', target_excel])
    else:
        messagebox.showinfo(title="海關稅單彙總清單轉檔工具", message="無法找到Excel的安裝")


