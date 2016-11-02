import os
import xlrd
import subprocess

def rename():
    path=r"F:\laopo\1"
    filelist=os.listdir(path)
    for files in filelist:
        Olddir=os.path.join(path,files);
        if os.path.isdir(Olddir):
            continue;
        filename=os.path.splitext(files)[0];
        newfilename = filename.strip()
        filetype=os.path.splitext(files)[1];
        Newdir=os.path.join(path,newfilename+filetype);
        os.rename(Olddir,Newdir);

rename();

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


