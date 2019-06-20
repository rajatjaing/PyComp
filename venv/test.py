# try:
#     from xml.etree.cElementTree import XML
# except ImportError:
#     from xml.etree.ElementTree import XML
# import zipfile
#
#
# """
# Module that extract text from MS XML Word document (.docx).
# (Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
# """
#
# WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
# PARA = WORD_NAMESPACE + 'p'
# TEXT = WORD_NAMESPACE + 'r w:rsidRPr'
#
#
# def get_docx_text(path):
#     """
#     Take the path of a docx file as argument, return the text in unicode.
#     """
#     document = zipfile.ZipFile(path)
#     xml_content = document.read('word/document.xml')
#     document.close()
#     tree = XML(xml_content)
#
#     paragraphs = []
#     for paragraph in tree.getiterator(PARA):
#         texts = [node.text
#                  for node in paragraph.getiterator(TEXT)
#                  if node.element]
#         if texts:
#             paragraphs.append(''.join(texts))
#
#     return '\n\n'.join(paragraphs)
#
# a=get_docx_text("C:\\Users\\qwerty\\Desktop\\templates\\test.docx")
# print(a)


# import os
# for file in os.listdir("C:\\Users\\qwerty\\Desktop\\templates"):
#     if file.endswith(".xml"):
#         os.path.join("C:\\Users\\qwerty\\Desktop\\templates", file)

# group2 = [1,2,3]
# group1 = [1,2]
#
# for k in group1:
#     for v in group2:
#         if k == v:
#             print(k)


import os
#
# def findCommonDeep(path1, path2):
#     files1 = set(os.path.relpath(os.path.join(root, file), path1) for root, _, files in os.walk(path1) for file in files)
#     files2 = set(os.path.relpath(os.path.join(root, file), path2) for root, _, files in os.walk(path2) for file in files)
#     return files1 & files2
#
# print(findCommonDeep("C:\\Users\\qwerty\\Desktop\\templates\\RightFolder","C:\\Users\\qwerty\\Desktop\\templates\\leftFolder"))

import zipfile
# from docx import Document
#
# document = Document("C:\\Users\\qwerty\\Desktop\\templates\\toCompare" + "\\" + "Rebranding_AVAYA_temp_v3.1.docx")
# sections = document.sections
# for section in sections:
#   print(section.orientation)



# import os
# from tkinter import filedialog,Tk
#
# from pip._vendor.distlib.compat import raw_input
#
# toplevel = Tk()
# toplevel.withdraw()
# filename = filedialog.askopenfilename()
# if os.path.isfile(filename):
#     for line in open(filename,'r'):
#         print (line),
# else: print ('No file chosen')
# raw_input('Ready, push Enter')


import tkinter
from tkinter import *
from tkinter import filedialog,Tk,constants,messagebox

dir = []
root = Tk()
root.title('PyComp by Rajat')
# Options for buttons
button_opt = {'fill': constants.BOTH, 'padx': 5, 'pady': 5}
# Define asking directory button



def openDirectory():
  count = 0
  dirname = filedialog.askdirectory(parent=root, initialdir='/home/', title='select from compare directory')
  print(dirname)
  if dirname!='':
     dir.append(dirname)
  else:
    count=1
    messagebox.showinfo("Error", "Please select a directory !!")

  if (len(dir)==1 and count!=1):

    Button(root, text='select to compare directory', fg='black', command=openDirectory).pack(**button_opt)
  if (len(dir) != 2):
    root.mainloop()
  else:
    messagebox.showinfo("processing","Check the to compare dir for Results")
    root.destroy()

if(len(dir)==0):

    Button(root, text='select from compare directory', fg='black', command=openDirectory).pack(**button_opt)
    root.mainloop()


# if(len(dir)==1):
#     Button(root, text='select to compare directory', fg='black', command=openDirectory).pack(**button_opt)


root.mainloop()


