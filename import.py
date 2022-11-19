import PyPDF2
import os
import sys

print("Python Program to print list the files in a directory.")
 
Direc = input(r"Enter the path of the folder: ")
print(f"Files in the directory: {Direc}")
 
files = os.listdir(Direc)
files = [f for f in files if os.path.isfile(Direc+'/'+f)] #Filtering only the files.
file = open('output1.txt', 'a')
sys.stdout = file
print(*files, sep="\n")
file.close()

