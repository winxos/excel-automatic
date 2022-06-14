from openpyxl import load_workbook
import os
stus = {}
def getfiles(suffix):                   #suffix为str，示例:   '.txt'      '.json'       '.h5'
    res = []
    for file in os.listdir("."):
            name,suf = os.path.splitext(file)
            if suf == suffix:
                res.append(file)
    return(res)
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
        if r[6].value == "缺勤":
            print(r[0].value,r[6].value)
            if r[0].value not in stus:
                stus[r[0].value] = [0,0,0,0]
            stus[r[0].value][0] +=1
        elif r[6].value == "迟到":
            if r[0].value not in stus:
                stus[r[0].value] = [0,0,0,0]
            stus[r[0].value][1] +=1
fs = get_filter_files("./作业/",["xlsx"])
for f in fs:
    if "一课一文" in f:
        continue
    w = load_workbook(f)
    s1=w.worksheets[0]
    for r in s1.rows:
        if r[9].value == "0":
            print(r[2].value,r[9].value)
            if r[2].value not in stus:
                stus[r[2].value] = [0,0,0,0]
            stus[r[2].value][2] +=1
w = load_workbook("./作业/一课一文.xlsx")
s1=w.worksheets[0]
for r in s1.rows:
    if r[3].value == "1971":
        print(r[2].value,r[9].value)
        if r[2].value not in stus:
            stus[r[2].value] = [0,0,0,0]
        stus[r[2].value][3] = int(r[9].value)
print(stus)
fs=getfiles(".xlsx")
if len(fs) != 1:
    print("无点名册")
    exit(0)
w = load_workbook(fs[0])
s1=w.worksheets[0]
for r in s1.rows:
    if str(r[0].value).isdigit():
        if r[3].value in stus:
            if stus[r[3].value][0]>0:
                s1.merge_cells("F%d:L%d"%(r[0].row,r[0].row))
                s1["F%d"%r[0].row].value = "缺勤%d次"%stus[r[3].value][0]
            if stus[r[3].value][1]>0:
                s1.merge_cells("M%d:R%d"%(r[0].row,r[0].row))
                s1["M%d"%r[0].row].value = "迟到%d次"%stus[r[3].value][1]
            r[20].value = 20 - stus[r[3].value][0]*5 - stus[r[3].value][1] *2
            r[27].value = 20 - stus[r[3].value][2]*5
            r[32].value = stus[r[3].value][3]
w.save("new_%s"%fs[0])