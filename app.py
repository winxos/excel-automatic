from openpyxl import load_workbook
import os
def get_filter_files(dir, ext=None):
    allfiles = []
    needExtFilter = (ext != None)
    for root, dirs, files in os.walk(dir):
        for filepath in files:
            filepath = os.path.join(root, filepath)
            extension = os.path.splitext(filepath)[1][1:]
            if needExtFilter and extension in ext:
                allfiles.append(filepath)
            elif not needExtFilter:
                allfiles.append(filepath)
    return allfiles
fs = get_filter_files("./考勤/",["xlsx"])
for f in fs:
    w = load_workbook(f)
    s1=w.worksheets[0]
    for r in s1.rows:
        if r[6].value == "缺勤" or r[6].value == "迟到":
            print(r[0].value,r[6].value)
fs = get_filter_files("./作业/",["xlsx"])
for f in fs:
    if "一课一文" in f:
        continue
    w = load_workbook(f)
    s1=w.worksheets[0]
    for r in s1.rows:
        if r[9].value == "0":
            print(r[2].value,r[9].value)
w = load_workbook("./作业/一课一文.xlsx")
s1=w.worksheets[0]
for r in s1.rows:
    print(r[2].value,r[9].value)