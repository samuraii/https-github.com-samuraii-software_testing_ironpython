import string
import time
import random
import os.path
import getopt
import sys
import clr
from model.group import Group

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["num of groups", "file name"])
except getopt.GetoptError as err:
    getopt.usage()
    sys.exit(2)

n = 5
f_name = 'data/groups.xlsx'

for opt, val in opts:
    if opt == '-n':
        n = int(val)
    elif opt == '-f':
        f = val


def random_string(prefix, max_len):
    symbols = string.ascii_letters + string.digits + " " * 10
    return prefix + ''.join([random.choice(symbols) for i in range(random.randrange(max_len))])


testdata = [Group(name='')] + [
    Group(name=random_string('name', 10))
    for i in range(n)
]

file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', f_name)

excel = Excel.ApplicationClass()
excel.visible = True

workbook = excel.Workbooks.Add()
sheet = workbook.ActiveSheet

for i in range(len(testdata)):
    sheet.Range[("A%s" % (i + 1))].Value2 = testdata[i].name

workbook.SaveAs(file)

excel.Quit()

time.sleep(10)
