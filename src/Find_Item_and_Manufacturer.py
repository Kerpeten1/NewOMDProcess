import xlrd


def get_rows_and_columns_of_omd_excel2(item, item_description):
    print("ITEM: ", item)
    pcs_item = check_if_pcs(item)
    if pcs_item == "not PCS":
        return
    print("returned item: ", pcs_item)
    formatted_item = formatting_item_number(pcs_item)
    item = formatted_item
    item_description_manufacturer = get_rows_and_columns_of_omd_excel(item, item_description)
    return item_description_manufacturer


def check_if_pcs(item):
    # item is not a list of items, it is one item number (type: string)
    alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","Ü","Ö","Ä"]
    if "-" in item:
        item = "not PCS"
        return item
    for j in range(0, len(alphabet)-1):
        if alphabet[j] in item:
            item = "not PCS"
            return item
    return item


def formatting_item_number(item):
    if "." in item:
        item.replace(".", "")
    formatted_item = item[:6]
    print("Item " + item + " formatted!")
    return formatted_item


# https://stackoverflow.com/questions/14944623/python-xrld-read-rows-and-columns
# übergabe: item + description
# return item + description + manufacturer
# path
def get_rows_and_columns_of_omd_excel(this_item, this_description):
    item = this_item
    description = this_description
    workbook = xlrd.open_workbook(r"C:\Users\M261651\Desktop\Dokumente\Files für Pycharm Projekte\OMD Prozess\ExcelFile\Original manufacturer for Disclosures_OMD3.xlsx")
    worksheet = workbook.sheet_by_name('OMT-Articles')
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    curr_row = 0
    list_all_rows_all = []
    list_all_rows = []

    while curr_row < num_rows:
        curr_row += 1
        curr_cell = -1
        list_current_row = []

        while curr_cell < num_cells:
            curr_cell += 1
            cell_value = worksheet.cell_value(curr_row, curr_cell)
            list_current_row.append(str(cell_value))  # #.replace("\n", " "))
        if list_current_row[0][:6] == str(item):    # wenn zeile index 0 ist item
            list_all_rows.append(list_current_row)
            list_all_rows_all.append([list_all_rows[-1][0]])
    print(list_all_rows_all)
    item_and_manufacturer = filter_item_and_manufacturer(list_all_rows, description)
    print(item_and_manufacturer)
    return item_and_manufacturer


def filter_item_and_manufacturer(list_all_rows, description):
    item_and_manufacturer = []
    for i in range(0,len(list_all_rows)):
        item_and_manufacturer.append((list_all_rows[i][0][:6], description, list_all_rows[i][3]))
    return item_and_manufacturer


