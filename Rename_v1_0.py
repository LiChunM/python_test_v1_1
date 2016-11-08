# -*- coding: cp936 -*-
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
    
def Rename(path):
    filelist=os.listdir(path)
    for files in filelist:
        Olddir=os.path.join(path,files);
        if os.path.isdir(Olddir):
            continue;
        filename=os.path.splitext(files)[0];
        if "-" in filename:
            newfilename = filename.split('-',2)[2];
            filetype=os.path.splitext(files)[1];
            Newdir=os.path.join(path,newfilename+filetype);
            os.rename(Olddir,Newdir);
        
if __name__ =='__main__':
    Filelist = GetFilesList("C:\\pro")
    for files in Filelist:
        Rename(files)
    
   
   
    
    
