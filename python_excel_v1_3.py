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
    ListFileYanName = []
    ListFileYinName = []
    mypathYan = "C:\\laopo\\yang\\"
    mypathYin = "C:\\laopo\\yin\\"
    ReName(mypathYan)
    ReName(mypathYin)
    ListFileYanName = GetFileName(mypathYan)
    ListFileYinName = GetFileName(mypathYin)
    for eyin in ListFileYinName:
        num = 0
        ListYinlie1 = []
        ListYinlie2 = []
        wbk = xlwt.Workbook()
        wxlssheet = wbk.add_sheet(eyin)
        GetOneListBuf(mypathYin+eyin+".txt",ListYinlie1,ListYinlie2)
        #print ListFileYanName
        for eyang in ListFileYanName:
            ListYanlie1 = []
            ListYanlie2 = []
            resalut =[]
            #print eyang
            GetOneListBuf(mypathYan+eyang+".txt",ListYanlie1,ListYanlie2)
            renum = 0
            for enum in ListYanlie2:
                resalut.append(str(float(ListYanlie2[renum])+float(ListYinlie2[renum])))
                renum =  renum+1
            WriteAOpenxls(wxlssheet,eyin+'-'+eyang,ListYanlie1,resalut,num)
            num = num+1
            del ListYanlie1[:]
            del ListYanlie2[:]
            del resalut[:]
        wbk.save(mypathYan+eyin+'-'+"yang.xls")
        del ListYinlie1[:]
        del ListYinlie2[:]




