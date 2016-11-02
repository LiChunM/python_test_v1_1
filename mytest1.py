import os
import xlrd
import subprocess


if __name__ == '__main__':
    namepath = r"C:\profiles\test.exe"
    p = subprocess.Popen(namepath, stdin = subprocess.PIPE,stdout = subprocess.PIPE, stderr = subprocess.PIPE, shell = False)
    p.stdin.write("344.95")
    p.stdin.write("yizhi")
    p.stdin.write("F:\laopo\yizhi           .TXT")
    p.stdin.write("766.025")
    p.stdin.write("1-1")
    p.stdin.write("F:\laopo\1\1               .TXT")
    fname = r"F:\laopo\1.xls"
    bk = xlrd.open_workbook(fname)
    table=bk.sheets()[0]
    cell=table.cell(1,1).value
    p.stdin.write(str(cell))
    p.stdin.write("C:\profiles\1-1.txt")
