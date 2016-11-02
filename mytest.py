import os
import xlrd
import subprocess


namepath = r"C:\profiles\test.exe"
p = subprocess.Popen(namepath, stdin = subprocess.PIPE,stdout = subprocess.PIPE, stderr = subprocess.PIPE, shell = False)
p.stdin.write("344.95\n") 
p.stdin.write("yizhi\n")
p.stdin.write("F:\\laopo\\yizhi.txt\n")
p.stdin.write("766.025\n")
p.stdin.write("1-1\n")
p.stdin.write("F:\\laopo\\1\\1.txt\n")
fname = r"F:\laopo\1.xls"
bk = xlrd.open_workbook(fname)
table=bk.sheets()[0]
cell=table.cell(1,1).value
p.stdin.write(str(cell)+"\n")
p.stdin.write("C:\\profiles\\1-3.TXT\n")
print p.stdout.read() 
