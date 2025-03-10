"""
customs_pdf_2_excel.py
============

This module provides core function to extract contents from a .pdf file, form
a table and store it into an Excel file.

.. note::
    This documentation follows the reStructuredText (reST) format, which is
    compatible with Sphinx autodoc.

:author: Maoyi Fan
:email: maoyi.fan@yapro.com.tw
:date: 2025-03-09
:version: 1.0
:license: MIT
:history:
    - 1.0 (2025-03-09) - To handle exceptional case where only one record on the
                         last page
"""

import os
import pdfminer
import pdfminer.high_level
import pdfminer.layout
import pdfminer.layout
import class_set
import utilities
import constant
import excel_output
import pdb


def pdf_2_excel(working_dir, input_fn, output_fn):
    # declare lists of table columns
    tax_bill_list = []                  # 稅單號碼
    declaration_form_list = []          # 報單號碼
    tax_ID_list = []                    # 納稅義務人統一標號
    tax_amount_list = []                # 金額
    first_page = True                   # end of Tax_ID of the 1st page differs from the rest page
    tax_bill_or_not = True
    tax_bill_info_OK = False            # 稅單號碼, 報單號碼 extracted
    tax_bill_entry = False              # 稅單資料輸入
    decl_form_entry = False             # 報單資料輸入
    tax_ID_entry = False                # 統一編號輸入
    tax_amount_entry = False            # 報單金額輸入
    # sanity check of the input parameters
    if input_fn == "":
        return False

    es = class_set.entry_setting(tax_bill_entry, decl_form_entry, tax_ID_entry, \
                                 tax_amount_entry, tax_bill_info_OK)
    # ifObj : file object of the input pdf file
    file_to_open = working_dir + "/" + input_fn
    print("Source file to open : ", file_to_open)
    ifObj = open(file_to_open, "rb")
    stripped_file_name = utilities.extract_file_name(input_fn)
    stripped_path_name = working_dir + "/" + stripped_file_name
    # prepare tfObj, file object of the intermediate text file
    textoutput_file = stripped_path_name + ".text"
    print("彙總稅單文字內容(text):", textoutput_file)
    if os.path.isfile(textoutput_file):
        os.remove(textoutput_file)
    tfObj = open(textoutput_file, "wb")

    # pdb.set_trace()
    # prepare ofObj, file object of the output tab-separated txt file
    if output_fn == "":
        output_file = stripped_path_name + ".txt"
    else:
        stripped_file_name = utilities.extract_file_name(output_fn)
        output_file = stripped_path_name + ".txt"
        print("彙總稅單清單轉出名稱(tab separated)", output_file)
    ofObj = open(output_file, "w", encoding="utf-8")

    # prepare to parse the input pdf file
    all_texts = None
    detect_vertical = None
    word_margin = None
    char_margin = None
    line_margin = None
    boxes_flow = None
    strip_control = False
    output_type = 'text'
    layoutmode = 'normal'

    laparams = pdfminer.layout.LAParams()
    for param in ("all_texts", "detect_vertical", "word_margin", "char_margin", "line_margin", "boxes_flow"):
            paramv = locals().get(param, None)
            if paramv is not None:
                setattr(laparams, param, paramv)
    #
    # extract text from the input .pdf file
    pdfminer.high_level.extract_text_to_fp(ifObj, tfObj, **locals())
    tfObj.close()

    if __debug__ is False:
        dbgFileObj = open("debug_output.txt", "w", encoding="utf-8")

    tfObj = open(textoutput_file, "r", encoding="utf-8")
    print_flag = False
    # reading input file line-by-line
    tfStr = tfObj.readline()
    l = 1
    while tfStr:
        # for x in tfStr:
        #    print(x.encode("utf-8").decode("utf-8", "ignore"))
        #
        # determining the state of entries
        #
        #        if tfStr.strip("\n") == constant.FILE_HEADER:

        if tfStr.strip('\n') == constant.FILE_HEADER:
            # 彙總稅單稅單清單
            if __debug__ is False:
                print("File Header: ", tfStr.strip('\n'))
        #        elif tfStr.strip('\n') == constant.FILE_TAILER:
        # 總筆數"
        #            print("File Tailer: ", tfStr.strip('\n'))
        elif tfStr.strip('\n') == constant.BEGINNING_DECLARATION_ID:
            # 報單號碼;
            # Description: pdfminer reads text in the order of storage, not display
            #              the results of text parsing group 稅單號碼 and 報單號碼 under
            #              報單號碼
            #
            # === Normal Case ===
            # Parse results:
            # 報單號碼
            #
            # ABI31130803796  <---- 稅單號碼
            #
            # ABF213H16A9766  <---- 報單號碼
            #
            # ABI31130790050  <---- 稅單號碼
            #
            # ABF213H16A4990  <---- 報單號碼
            # ....
            # 納稅義務人統編
            #
            # 29060646        <---- 納稅義務人統編
            #
            # === Exceptional Case ===
            # Description: If the last page contains only one record, 稅單號碼,報單號碼 is grouped into
            #              納稅義務人統編
            # Parse results:
            # 報單號碼
            #
            # 納稅義務人統編
            #
            # ABI31146190113  <---- 稅單號碼
            #
            # ABF214H16A5556  <---- 報單號碼
            #
            # 29060646        <---- 納稅義務人統編
            #
            #
            print_flag = True
            tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry = \
                es.clear_current_setting()
            tax_bill_entry = True
            tax_bill_info_OK = False
            es.set_current_entry(tax_bill_entry, decl_form_entry, tax_ID_entry, \
                                 tax_amount_entry, tax_bill_info_OK)
            if __debug__ is False:
                print("Beginning declaration ID: ", tfStr.strip('\n'))
        elif tfStr.strip('\n') == constant.BEGINNING_TAX_ID_COLUMN:  # also END_DECLARATION_ID
            # 納稅義務人統編
            print_flag = True
            tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry = \
                es.clear_current_setting()
            tax_ID_entry = True
            es.set_current_entry(tax_bill_entry, decl_form_entry, tax_ID_entry, \
                                 tax_amount_entry, tax_bill_info_OK)
            if __debug__ is False:
                print("Tax ID: ", tfStr.strip('\n'))
        elif tfStr.strip('\n') == constant.BEGINNING_AMOUNT_COLUMN:
            # 金額
            print_flag = True
            tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry = \
                es.clear_current_setting()
            tax_amount_entry = True
            es.set_current_entry(tax_bill_entry, decl_form_entry, tax_ID_entry, \
                                 tax_amount_entry, tax_bill_info_OK)
            if __debug__ is False:
                print("Amount: ", tfStr.strip('\n'))
        elif tfStr.strip('\n') == constant.END_AMOUNT_COLUMN_P1:
            print_flag = True
            tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry = \
                es.clear_current_setting()
            if __debug__ is False:
                dbgFileObj.write("<<<<製表日期>>>>\n")
        elif tfStr[:2] == constant.END_TAX_ID_P2:
            print_flag = True
            tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry = \
                es.clear_current_setting()
            if __debug__ is False:
                dbgFileObj.write("<<<<頁碼>>>>\n")
        elif tfStr.strip('\n') == constant.RECORD_COUNT:
            print_flag = True
            tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry = \
                es.clear_current_setting()
            if __debug__ is False:
                dbgFileObj.write("<<<<總筆數>>>>\n")
        #
        # processing per state machine
        #
        valid_entry, tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry, tax_bill_info_OK = \
            es.get_current_setting()
        # if print_flag is True:
        #    print_flag = False
        #    print(es.get_current_setting())
        #
        # state machine processing column entries
        if tax_bill_entry is True:
            if print_flag is True: # print_flag is True means this line is the column tag itself, i.e. 報單號碼, 納稅義務人統編....
                                   # rather than the data contents
                print_flag = False
                #  print("Column tag: 處理稅單、報單資料")
                if __debug__ is False:
                    dbgFileObj.write(tfStr)
            else:
                if tfStr.strip('\n') != "":
                    if tax_bill_or_not is True: # 稅單號碼
                        tax_bill_list.append(tfStr.strip('\n'))
                        tax_bill_or_not = False
                    else:                       # 報單號碼
                        declaration_form_list.append(tfStr.strip('\n'))
                        tax_bill_or_not = True
                        es.tax_bill_info_OK = True  # completed at least one record
                    if __debug__ is False:
                        dbgFileObj.write(tfStr)
        elif tax_ID_entry is True:
            if print_flag is True:
                print_flag = False
                #  print("Column tag 處理統一編號")
                if __debug__ is False:
                    dbgFileObj.write(tfStr)
            else:
                if tfStr.strip('\n') != "":
                    # Handling Exceptional Case Here!!
                    if tax_bill_info_OK is not True:
                        print(f"Hit exceptional case@{tfStr.strip('\n')}", end="")
                        if tax_bill_or_not is True:
                            print(f"---> 稅單號碼")
                            tax_bill_list.append(tfStr.strip('\n'))
                            tax_bill_or_not = False
                        else:
                            print(f"---> 報單號碼")
                            declaration_form_list.append(tfStr.strip('\n'))
                            tax_bill_or_not = True
                            es.tax_bill_info_OK = True
                    else:   # Normal Cases
                        tax_ID_list.append(tfStr.strip('\n'))
                    if __debug__ is False:
                        dbgFileObj.write(tfStr)
        elif tax_amount_entry is True:
            if print_flag is True:
                print_flag = False
                #  print("Column tag 金額")
                if __debug__ is False:
                    dbgFileObj.write(tfStr)
            else:
                if tfStr.strip('\n') != "":
                    tax_amount_list.append(tfStr.strip('\n'))
                    if __debug__ is False:
                        dbgFileObj.write(tfStr)

        tfStr = tfObj.readline()
        l = l+1

    if __debug__ is False:
        dbgFileObj.write("tax_bill_list length:" + str(len(tax_bill_list)) + "\n")
        dbgFileObj.write("declaration_form_list:" + str(len(declaration_form_list)) + "\n")
        dbgFileObj.write("tax_ID_list:" + str(len(tax_ID_list)) + "\n")
        dbgFileObj.write("tax_amount_llist:" + str(len(tax_amount_list)) + "\n")

#
# compose output file by combining the four lists collected in
# the above state machine
#
    if len(tax_bill_list) == len(declaration_form_list) == len(tax_ID_list) == \
        len(tax_amount_list):
        for i in range(0, len(tax_bill_list)):
            combined_string = declaration_form_list[i] + "\t" + tax_bill_list[i] + \
                            "\t" + tax_ID_list[i] + "\t" + tax_amount_list[i] + "\n"
            ofObj.write(combined_string)
        excel_output.generate_excel_output(stripped_path_name, tax_bill_list, declaration_form_list, tax_ID_list, tax_amount_list)
    #
    # Close file handlers and remove
    #
    if ifObj.closed is False:
        ifObj.close()

    if ofObj.closed is False:
        ofObj.close()

    if __debug__ is True:
        if os.path.isfile(output_file):
            os.remove(output_file)

    if tfObj.closed is False:
        tfObj.close()

    if __debug__ is True:
        if os.path.isfile(textoutput_file):
            os.remove(textoutput_file)

    if __debug__ is False:
        dbgFileObj.close()