import xlrd
import xlwt
import subprocess
import time
import os
import re

def ReName(path):
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

def GetFileName(path):
    ListFileName = []
    filelist=os.listdir(path)
    for files in filelist:
        Olddir=os.path.join(path,files);
        if os.path.isdir(Olddir):
            continue;
        filename=os.path.splitext(files)[0];
        ListFileName.append(filename)
    return ListFileName

def GetOneListBuf(filename,Listlie1,Listlie2):
    SplitOnebuf = []
    fd = open(filename)
    for linebuf in fd:
        SplitOnebuf = linebuf.split()
        Listlie1.append(SplitOnebuf[0])
        Listlie2.append(SplitOnebuf[1])
    fd.close()



def WriteAOpenxls(wxlssheet,Name,Listlie1,Listlie2,num):
    index = 0
    wxlssheet.write(index,num*2,Name)
    for lie1 in Listlie1:
        index = index+1
        wxlssheet.write(index,num*2,lie1)
    index = 0
    for lie2 in Listlie2:
        index=index+1
        wxlssheet.write(index,1+num*2,lie2)
    
    

if __name__ =='__main__':
    num = 0
    ListFileName = []
    mypath = "C:\\laopo\\yang\\"
    ReName(mypath)
    ListFileName = GetFileName(mypath)
    wbk = xlwt.Workbook()
    wxlssheet = wbk.add_sheet("yang")
    for ename in ListFileName:
        Listlie1 = []
        Listlie2 = []
        GetOneListBuf(mypath+ename+".txt",Listlie1,Listlie2)
        WriteAOpenxls(wxlssheet,ename,Listlie1,Listlie2,num)
        num = num+1
        del Listlie1[:]
        del Listlie2[:]
    wbk.save(mypath+"yang.xls")





