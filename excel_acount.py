from pypinyin import pinyin, lazy_pinyin, Style
import xlrd
import xlsxwriter

book = xlrd.open_workbook('开发区奎山街道小程序账号.xlsx')
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=2, colx=3)))
expenses = []
dup = {}
offset = 1000000000000000000
for rx in range(sh.nrows):
    # print(sh.row(rx))
    init = sh.row_values(rx)
    a89 = sh.row_values(rx, 8, 10)
    szm = ''
    for element in a89:
        t = lazy_pinyin(element, style=Style.FIRST_LETTER)
        for x in t:
            szm += x
    if szm in dup.keys():
        dup[szm] += 1
        init[11] = szm + str(dup[szm])
    else:
        init[11] = szm
        dup[szm] = 0
    # init[0] = int(init[0])
    init[0] = rx+offset


    init[1] = '076945'  # 租户ID
    init[4] = 3
    init[12] = 'aecd82d7f8f28062c94e9682781155dc1f1f818f'  # 密码888888
    init[13] = '1264841812109684737'  # 角色ID 采集员
    init[14] = '1216548691450490881'  # create_user 管理员
    init[15] = '1216544938320191490'  # create_dept 管理员
    init[17] = '1216548691450490881'  # update_user 管理员
    init[19] = 1  # status
    init[20] = 0  # is_deleted

    init[2] = '1266206068415537154'
    init[3] = '0,1216544938320191490,1216557952289173505,1216553945722220546,1266205960651284481,1266206068415537154'
    init[16] = '2020/5/29  16:50:00'  # create_time
    init[18] = '2020/5/29  16:50:00'  # update_time

    expenses.append(init)

# print(expenses)
workbook = xlsxwriter.Workbook('奎山街道.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0
num_format = workbook.add_format({'num_format': '0'})

for item in expenses:
    col = 0
    row += 1
    for cc in item:
        if col == 0:
            worksheet.write(row, col, cc, num_format)
        else:
            worksheet.write(row, col, cc)
        col += 1

workbook.close()
