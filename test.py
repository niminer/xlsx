from pypinyin import pinyin, lazy_pinyin, Style
import xlrd
import xlsxwriter

book = xlrd.open_workbook("莒县农村社区情况表.xlsx")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=2, colx=3)))
expenses = []
y = []
dup = {}
for rx in range(sh.nrows):
    # print(sh.row(rx))
    init = sh.row_values(rx)
    jiedao = init[0]
    cun = init[4]
    jihe=[jiedao,cun]
    expenses.append(jihe)


print(expenses)

t = []
for x in expenses:
    if x[0]!="":
        zhi=x[0]
    t1 = x[1].split('，')
    for xx in t1:
        t.append([zhi, xx])
# t=y[2].split('，')


print(t)
#
workbook = xlsxwriter.Workbook('t.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

for item in t:
    col = 0
    row += 1
    for cc in item:
        worksheet.write(row, col, cc)
        col+=1

workbook.close()

# a = lazy_pinyin('中心',style=Style.FIRST_LETTER)
# b=''
# for x in a:
#     b+=x
# print(b.title())
