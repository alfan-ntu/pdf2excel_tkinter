import xlsxwriter


def generate_excel_output(stripped_path_name, tax_bill_list, declaration_form_list, tax_ID_list, tax_amount_list):
    # prepare spread sheet list, which stores row lists of
    # elements, tax_bill_list, declaration_form_list, tax_ID_list
    # and tax_amount_list
    spread_sheet = []
    row = []
    column_width = [0,0,0,0]
    for i in range(0, len(tax_bill_list)):
        row.append(declaration_form_list[i])
        if len(declaration_form_list[i]) > column_width[0]:
            column_width[0] = len(declaration_form_list[i])
        row.append(tax_bill_list[i])
        if (len(tax_bill_list[i])) > column_width[1]:
            column_width[1] = len(tax_bill_list[i])
        row.append(tax_ID_list[i])
        if (len(tax_ID_list[i])) > column_width[2]:
            column_width[2] = len(tax_ID_list[i])
        row.append(tax_amount_list[i])
        if (len(tax_amount_list[i])) > column_width[3]:
            column_width[3] = len(tax_amount_list[i])
        spread_sheet.append(row)
        row = []

    generate_excel(stripped_path_name, spread_sheet, column_width)


def generate_excel(stripped_path_name, spread_sheet, column_width):
    workbook = xlsxwriter.Workbook(stripped_path_name + ".xlsx")
    worksheet = workbook.add_worksheet()
    tax_bill_column = ""
    customs_bill_column = ""
    tax_ID_column = ""
    tax_amount_column = ""
    row = 0
    col = 0
    header_row_format = workbook.add_format({'align': 'center'})
    #
    # Notes: setting the row format in a separate statement automatically hide the row!!
    # worksheet.set_row(0, 0, header_row_format, {'hidden': False})
    #
    worksheet.write(row, col, "稅單號碼", header_row_format)
    worksheet.write(row, col+1, "報單號碼", header_row_format)
    worksheet.write(row, col+2, "統一編號", header_row_format)
    worksheet.write(row, col+3, "金額", header_row_format)
    row += 1
    for tax_bill_column, customs_bill_column, tax_ID_column, tax_amount_column in spread_sheet:
        worksheet.write(row, col, tax_bill_column)
        worksheet.write(row, col+1, customs_bill_column)
        worksheet.write(row, col+2, tax_ID_column)
        worksheet.write(row, col+3, tax_amount_column)
        row += 1
    # adjust column width based on the content length of each column
    for i in range(4):
        worksheet.set_column(i, i, column_width[i]+5)

    workbook.close()
    print("產出彙總稅單清單轉檔(.xslx):" + stripped_path_name + ".xlsx")
