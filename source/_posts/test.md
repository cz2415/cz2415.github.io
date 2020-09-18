---
title: test
date: 2020-09-18 11:47:43
tags:
---

excel文件对比
from openpyxl import *


def compareExcel(ename1, ename2):

    print("Begin")
    print("Comparing", ename1, ename2)
    fileSame = True

    wb1 = load_workbook(filename=ename1, read_only=True)
    wb2 = load_workbook(filename=ename2, read_only=True)
    sn1 = wb1.sheetnames
    sn2 = wb2.sheetnames

    if (sn1 != sn2):
        print("两个excel 的 sheet 名不同")
        print(ename1, " sheet names:", sn1)
        print(ename2, " sheet names:", sn2)
    else:
        sn = sn1
        for wsn in sn:
            ws1 = wb1[wsn]
            ws2 = wb2[wsn]
            c = ws1.max_column
            r = ws1.max_row
            if ((ws2.max_column != c) or (ws2.max_row != r)):
                print("SHEET ", wsn,
                      ": 行数或列数不同!")
                fileSame = False
            else:
                flag = True
                for i in range(1, r+1):
                    for j in range(1, c+1):
                        c1 = ws1.cell(i, j)
                        c2 = ws2.cell(i, j)
                        if (c1):
                            if (c2):
                                if (c1.value != c2.value):
                                    if ((wsn == "Internal Info") and ((i == 4) and (j == 2)) or ((i == 5) and (j == 3))):
                                        continue
                                    print(c1.coordinate)
                                    print("v1:", c1.value)
                                    print("v2:", c2.value)
                                    print(
                                        "-------------------------------------------------------------------------------------------------------------")
                                    flag = False
                            else:
                                print("DIFFERDENT_TO_NONE at SHEET-",
                                      wsn, ": At (", i, ",", j, ")")
                                print("diff FROM", c1.value)
                                flag = False
                        else:
                            if (c2):
                                print("DIFFERDENT_TO_NONE at SHEET-",
                                      wsn, ": At (", i, ",", j, ")")
                                print("diff FROM", c2.value)
                                flag = False
                fileSame = fileSame and flag
        if fileSame:
            print("SAME_FILE:", ename1, ename2)

    print("Completed")


if __name__ == "__main__":
    compareExcel(r'C:\Users\USER\Desktop\04.xlsx',
                 r'C:\Users\USER\Desktop\05.xlsx')