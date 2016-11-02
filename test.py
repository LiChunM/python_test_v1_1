Python 2.7 (r27:82525, Jul  4 2010, 09:01:59) [MSC v.1500 32 bit (Intel)] on win32
Type "copyright", "credits" or "license()" for more information.
>>> import subprocess
>>> p = subprocess.Popen("app2.exe", stdin = subprocess.PIPE, 
stdout = subprocess.PIPE, stderr = subprocess.PIPE, shell = False)   
