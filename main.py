import pandas as pd
import openpyxl

PBT = "personal bulk task.xlsx"

def lookup(file_name, value, col_find, col_give, sheet_id=0, minrow=0, maxrow=-1, colnumber=False):
    """attempts to find the passed value in the column col_find within the spreadsheet, and when it does, returns
       a list of the corresponding values in the same row within col_give. sheet_id can be either the name of a
       page in the spreadsheet, or else a 0-indexed number of that page. when colnumber is true, it uses 0-indexed
       IDs for the columns instead of the column names (like VLOOKUP). minrow is exclusive and maxrow is exclusive,
       based on a 0-indexed list of the rows starting with the first data row"""

    df = pd.read_excel(file_name, sheet_id)  # data frame of the relevant sheet

    if colnumber:
        col_find = df.columns[col_find]
        col_give = df.columns[col_give]
    cf = df[col_find]
    if maxrow < 0:
        maxrow = cf.size
    found_rows = []
    for i in range(minrow, maxrow):
        if (str(cf[i]) == str(
                value)):  # if the string version of a column value in the lookup column is the same as the
            found_rows += [i]  # string version of the passed value, take note of the row
            # print(str(cf[i]), str(value), str(cf[i] == str(value)))
    cg = df[col_give]
    ret_values = [cg[v] for v in found_rows]
    return ret_values


def column_list(file_name, sheet_id=0):
    return pd.read_excel(file_name, sheet_id).columns


def all_column_lists(file_name):
    num_tabs = len(openpyxl.load_workbook(file_name).sheetnames)
    ret_list = []
    for i in range(num_tabs):
        ret_list += [column_list(file_name, i)]
    return ret_list


def all_col_test():
    a = all_column_lists("personal bulk task.xlsx")
    for b in a:
        print(b)


def export_csv(src_file, sheet_id, prefix=""):
    pd.read_excel(src_file, sheet_id).to_csv(prefix + str(sheet_id) + ".csv", index=False)


def export_all_csv(src_file, prefix=""):
    wb = openpyxl.load_workbook(src_file)
    for x in wb.sheetnames:
        export_csv(src_file, x, prefix)


#get row by id
#search entire document for some string ("filter")


def lookup_test():
    print(lookup("personal bulk task.xlsx", 331, "Report 1 ID", "Second Report Name", "Commonality"))
    print([lookup("personal bulk task.xlsx", x, 1, 4, 1, colnumber = True) for x in range(331,341)])
    print(lookup("personal bulk task.xlsx", 643376, "Report ID", "Report Path", 3))
    print(lookup("personal bulk task.xlsx", 41, "Analyzer Task ID", "Folder Path", "Task Details", maxrow = 4))



if __name__ == '__main__':
    export_all_csv(PBT, "biswabir")
    #export_all_csv("personal bulk task.xlsx", prefix="testhello")
    #lookup_test()
    #all_col_test()

    #for sheet in (openpyxl.load_workbook("personal bulk task.xlsx").sheetnames):
    #    print(sheet + ":")
    #    cols = column_list("personal bulk task.xlsx", sheet)
    #    if("Report ID" in cols and "Report Name" in cols):
    #        print(" ", lookup("personal bulk task.xlsx", 632214, "Report ID", "Report Name", sheet))
    #    else:
    #        print("  no relevant columns")

