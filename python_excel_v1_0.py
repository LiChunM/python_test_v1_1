import os
import xlrd
import subprocess
import time

def GetFilesList(path):
    dirList = []
    dirInitName = []
    filelist=os.listdir(path)
    for files in filelist:
        Olddir=os.path.join(path,files);
        if(os.path.isdir(Olddir)):
            dirInitName.append(int(files))
    dirInitName.sort()
    for i in dirInitName:
       Newdir = os.path.join(path,str(i));
       dirList.append(Newdir)
    return dirList
    
def rename(path):
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

def ExecDoCalc(exepath,yizhipath,outname1,outname2,txtpath,cell,outpath):
    p = subprocess.Popen(exepath, stdin = subprocess.PIPE,stdout = subprocess.PIPE, stderr = subprocess.PIPE, shell = False)
    p.stdin.write("344.95\n") 
    p.stdin.write("yizhi\n")
    p.stdin.write(yizhipath+"\n")
    #print yizhipath
    p.stdin.write("766.025\n")
    p.stdin.write(outname1+"-"+outname2+"\n")
    #print outname1+"-"+outname2
    p.stdin.write(txtpath+"\n")
    #print txtpath
    p.stdin.write(str(cell)+"\n")
    #print cell
    p.stdin.write(outpath+outname1+"-"+outname2+".txt"+"\n")
    #print outpath+outname1+"-"+outname2+".txt"
    
def GetFilesListName(path):
    dirListName = []
    dirInitName = []
    dirNewName  = []
    filelist=os.listdir(path)
    for files in filelist:
        Olddir=os.path.join(path,files);
        if(os.path.isdir(Olddir)):
                dirListName.append(files)
    for filenames in dirListName:
        dirInitName.append(int(filenames))
    dirInitName.sort()
    for i in dirInitName:
        dirNewName.append(str(i))
    return dirNewName

def GetFileName(path):
    ListFileName = []
    ListInitName = []
    ListNewName = []
    filelist=os.listdir(path)
    for files in filelist:
        Olddir=os.path.join(path,files);
        if os.path.isdir(Olddir):
            continue;
        filename=os.path.splitext(files)[0];
        if "yin" in filename:
            continue;
        ListFileName.append(filename)
    for filenames in ListFileName:
        ListInitName.append(int(filenames))
    ListInitName.sort()
    for i in ListInitName:
        ListNewName.append(str(i))
    return ListNewName





if __name__ =='__main__':
    num = 0
    echang = 1
    elie = 1
    fname = r"C:\laopo\1.xls"
    bk = xlrd.open_workbook(fname)
    table=bk.sheets()[0]
    MainDirList1 = []
    MainDirList1 = GetFilesList("C:\\laopo")
    for elist1 in MainDirList1:
        rename(elist1)
    MainDirList2 = []
    MainDirList2 = GetFilesListName("C:\\laopo")
    for elist2 in MainDirList2:
        SuDirList1 = []
        SuDirList1 = GetFileName(MainDirList1[num])
        for elist3 in SuDirList1:
            txtpath = MainDirList1[num]+"\\"+ elist3 + ".TXT"
            cell=table.cell(elie,echang).value
            #print elie
            ExecDoCalc(r"C:\laopo\test.exe",r"C:\laopo\yizhi.txt",elist2,elist3,txtpath,cell,MainDirList1[num]+"\\")
            time.sleep(1) 
            #break
            elie+=1
        num+=1
        echang+=1
        elie=1
        #print echang
        del SuDirList1[:]
        #break
   
    
    
