from tkinter import *
from tkinter import filedialog,Tk,constants,messagebox
import os


fromCompdir=[]
toDirComp=[]
dir = []
root = Tk()
root.title('PyComp by Rajat')
# Options for buttons
button_opt = {'fill': constants.BOTH, 'padx': 5, 'pady': 5}
# Define asking directory button


def SelectDirectory():
    def openDirectory():
      count = 0
      dirname = filedialog.askdirectory(parent=root, initialdir='/home/', title='select from_compare directory')
      print(dirname)
      if dirname!='':
         dir.append(dirname)
      else:
        count=1
        messagebox.showinfo("Error", "Please select a directory !!")

      if (len(dir)==1 and count!=1):

        Button(root, text='select to_compare directory', fg='black', command=openDirectory).pack(**button_opt)
      if (len(dir) != 2):
        root.mainloop()
      else:
        messagebox.showinfo("processing","please sit back...we will inform you when process completes !!")
        root.destroy()

    if(len(dir)==0):

        Button(root, text='select from compare directory', fg='black', command=openDirectory).pack(**button_opt)
        root.mainloop()

    root.mainloop()
    return dir



def listFromDirFile(fromComp_dirString):
    for fromComp_dirStrn in os.listdir(fromComp_dirString):
        if fromComp_dirStrn.endswith(".docx"):
            fromCompdir.append(os.path.join(fromComp_dirString, fromComp_dirStrn))
    return fromCompdir

def listToDirFile(toComp_dirString):
    for toComp_dirStrn in os.listdir(toComp_dirString):
        if toComp_dirStrn.endswith(".docx"):
            toDirComp.append(os.path.join(toComp_dirString, toComp_dirStrn))
    return toDirComp


def findCommonFiles(path1, path2):
    files1 = set(os.path.relpath(os.path.join(root, file), path1) for root, _, files in os.walk(path1) for file in files)
    files2 = set(os.path.relpath(os.path.join(root, file), path2) for root, _, files in os.walk(path2) for file in files)
    return files1 & files2


def safeStr(obj):
    try: return str(obj)
    except UnicodeEncodeError:
        return obj.encode('ascii', 'ignore').decode('ascii')
    except:
        return ""