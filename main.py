from bs4 import BeautifulSoup
from xml.dom.minidom import parseString
#specific to extracting information from word documents
import zipfile
#to pretty print our xml:
import xml.dom.minidom
from Def import *
from tkinter import messagebox

# ----------------------------------------------deffining path for both directory--------------------------------------


fromCompdir=[]
toDirComp=[]
dire=[]
fromComp_dirString_old=''
toComp_dirString_old=''

dire=SelectDirectory()
fromComp_dirStrn=''
toComp_dirStrn=''
length = len(dire)
for i in range(length):

    if(i==0):
        fromComp_dirString_old=str(dire[i])
    if(i==1):
        toComp_dirString_old=str(dire[i])


print("this is from :"+fromComp_dirString_old)
print("this is to : "+toComp_dirString_old)

fromComp_dirString=fromComp_dirString_old.replace("/","\\")
toComp_dirString=toComp_dirString_old.replace("/","\\")

print("this is  new from :"+fromComp_dirString)
print("this is  new to : "+toComp_dirString)

# ----------------------------------------------getting files from directory--------------------------------------

fromCompdir=listFromDirFile(fromComp_dirString)
if len(fromCompdir)==0:
    messagebox.showinfo("Error !!", "try again...Empty Directory  " + fromComp_dirString)

    print(fromCompdir)

else:
    # ----------------------------------------------getting files from directory--------------------------------------

    toDirComp=listToDirFile(toComp_dirString)
    if len(toDirComp) == 0:
        messagebox.showinfo("Error !!", "try again..Empty Directory  " + toComp_dirString)
        print(toDirComp)
    else:
        # ------------------------------------------finding Common file-----------------------------------------------------

        commonFilesSet=findCommonFiles(fromComp_dirString,toComp_dirString)

        #------------------------------------------------ this is main program------------------------------------------------------------------------


        for commFile in commonFilesSet:
            print("this is common file : "+commFile)

            document_from = zipfile.ZipFile(fromComp_dirString+"\\"+commFile)
            document_to = zipfile.ZipFile(toComp_dirString + "\\" + commFile)
            Result_comp_path=toComp_dirString+'\\Comp_Result\\'
            Image_comp_path=toComp_dirString+'\\'+'Comp_Images\\'+commFile+'\\'
            if not os.path.exists(Result_comp_path):
                os.makedirs(Result_comp_path)

            writeFile = open(Result_comp_path + commFile + "_result" + ".txt", "a+")
            imagePackName=[]

            dataDict_from = {}
            dataDict_to = {}

            imageDict_from = {}
            imageDict_to = {}

            Port_count = 0
            Land_count = 0
            portrait = "Portrait"
            # -------------------------------------------------from which file to be compared---------------------------------------------

            for docPackageName in document_from.namelist():


                if str(docPackageName).find('word/media/') != -1:
                    imagePackName.append(docPackageName)
                    # print(imagePackName)

                else:
                    uglyXml = xml.dom.minidom.parseString(document_from.read(docPackageName)).toprettyxml(indent='  ')
                    # print(docPackageName+"\n")
                    text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
                    prettyXml = text_re.sub('>\g<1></', uglyXml)

                    tagString=''
                    soup = BeautifulSoup(prettyXml,'xml')
                    Texttags = soup.find_all('w:r')
                    OrientationTag = soup.find_all('w:sectPr')

                    wrTagSumm=[]
                    for wrTags in Texttags:
                        wrTagSumm.append(str(wrTags))


                    orientSumm_from=[]
                    for sectTags in OrientationTag:
                        orientSumm_from.append(str(sectTags))


                    for tagsVal in wrTagSumm:

                     # print("\n"+tagsVal)
                     TagSoup = BeautifulSoup(tagsVal, "xml")
                     # print("\n"+TagSoup.prettify())
                     soupString=TagSoup.prettify()
                     soupData = BeautifulSoup(soupString,'xml')
                     textTagValue= soupData.find('t')
                     fontTagValue = soupData.find('rFonts')
                     colorTagValue = soupData.find('color')
                     tagValue = soupData.find('r')
                     # and str(fontValue) != 'None' and str(colorValue) != 'None'
                     if str(textTagValue) != 'None' and str(tagValue) != 'None' :
                         # Summary="Text is: "+str(textValue)+"\n"+"Font is: "+str(fontValue)+"\n"+"color is :"+str(colorValue)+"\n"
                         # print(tagValue)

                         for a in TagSoup.find_all('r'):
                             colorValueF = a.find("color")
                             fontValueF = a.find("rFonts")
                             sizeValueF = a.find("sz")
                             textValueF = a.find("t")

                             if colorValueF:
                                 color=colorValueF["val"]
                                 # print(color)
                             else :
                                 color="NONE"
                                 # print(color)

                             if textValueF :
                                 text=textValueF.text
                                 # print(text)
                             else :
                                 text="NONE"
                                 # print(text)

                             if fontValueF :
                                  font=fontValueF["ascii"]
                                  # print(font)
                             else:
                                 font="NONE"
                                 # print(font)

                             if sizeValueF:
                                 sizeString=sizeValueF["val"]
                                 size=int(int(sizeString)/2)
                                 # print(size)
                             else :
                                 size="NONE"
                                 # print(size)
                             # Summary = "Text is: " + str(text) + "\n" + "Font is: " + str(font) + "\n" + "color is :" + str(color) + "\n"+"size is :" +str(size)+"\n"+"\n"
                             # # writeFile.write("\r\n"+Summary)
                             # print(Summary)

                             if text!=' ':
                                 dataDict_from[text] = []
                                 dataDict_from[text].append(font)
                                 dataDict_from[text].append(color)
                                 dataDict_from[text].append(size)

                                 # writeFile.write("\r\n" +str(dataDict_from))

                                 print(dataDict_from)
                                 # dataDict_from.clear()


                    for orientVal in orientSumm_from:
                        orntSoup = BeautifulSoup(orientVal, "xml")
                        OrntsoupString = orntSoup.prettify()
                        OrntsoupData = BeautifulSoup(OrntsoupString, 'xml')
                        OrntTagValue = OrntsoupData.find('pgSz')
                        HeaderRefTagValue = OrntsoupData.find('headerReference')
                        footerRefTagValue = OrntsoupData.find('footerReference')
                        if str(OrntTagValue) != 'None':
                            for b in orntSoup.find_all('sectPr'):
                                OrntValueF = b.find("pgSz")
                                HeaderReftValueF = b.find("headerReference")
                                FooterReftValueF = b.find("footerReference")
                                if OrntValueF:
                                    print(OrntValueF)
                                    orientation = OrntValueF.get("orient")
                                    if str(orientation)!='None':
                                        print("orientation is : "+str(orientation))
                                        if HeaderReftValueF:
                                            print(HeaderReftValueF)
                                            headerID = HeaderReftValueF.get("id")
                                            if str(headerID) != 'None':
                                                print("header ID is : " + str(headerID))
                                        if FooterReftValueF:
                                            print(FooterReftValueF)
                                            footerID = FooterReftValueF.get("id")
                                            if str(footerID) != 'None':
                                                print("footer ID is : " + str(footerID))
                                    else:
                                        portrait = "Portrait"
                                        print("orientation is : "+portrait)

            # -------------------------------------------------for file to be compared---------------------------------------------

            for docPackageName in document_to.namelist():


                if str(docPackageName).find('word/media/') != -1:
                    imagePackName.append(docPackageName)
                    # print(imagePackName)

                else:
                    uglyXml = xml.dom.minidom.parseString(document_to.read(docPackageName)).toprettyxml(indent='  ')
                    # print(docPackageName+"\n")
                    text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
                    prettyXml = text_re.sub('>\g<1></', uglyXml)

                    tagString=''
                    soup = BeautifulSoup(prettyXml,'xml')
                    Texttags = soup.find_all('w:r')
                    OrientationTag = soup.find_all('w:sectPr')

                    wrTagSumm=[]
                    for wrTags in Texttags:
                        wrTagSumm.append(str(wrTags))


                    orientSumm_to=[]
                    for sectTags in OrientationTag:
                        orientSumm_to.append(str(sectTags))


                    for tagsVal in wrTagSumm:

                     # print("\n"+tagsVal)
                     TagSoup = BeautifulSoup(tagsVal, "xml")
                     # print("\n"+TagSoup.prettify())
                     soupString=TagSoup.prettify()
                     soupData = BeautifulSoup(soupString,'xml')
                     textTagValue= soupData.find('t')
                     fontTagValue = soupData.find('rFonts')
                     colorTagValue = soupData.find('color')
                     tagValue = soupData.find('r')
                     # and str(fontValue) != 'None' and str(colorValue) != 'None'
                     if str(textTagValue) != 'None' and str(tagValue) != 'None' :
                         # Summary="Text is: "+str(textValue)+"\n"+"Font is: "+str(fontValue)+"\n"+"color is :"+str(colorValue)+"\n"
                         # print(tagValue)

                         for a in TagSoup.find_all('r'):
                             colorValueF = a.find("color")
                             fontValueF = a.find("rFonts")
                             sizeValueF = a.find("sz")
                             textValueF = a.find("t")

                             if colorValueF:
                                 color=colorValueF["val"]
                                 # print(color)
                             else :
                                 color="NONE"
                                 # print(color)

                             if textValueF :
                                 text=textValueF.text
                                 # print(text)
                             else :
                                 text="NONE"
                                 # print(text)

                             if fontValueF :
                                  font=fontValueF["ascii"]
                                  # print(font)
                             else:
                                 font="NONE"
                                 # print(font)

                             if sizeValueF:
                                 sizeString=sizeValueF["val"]
                                 size=int(int(sizeString)/2)
                                 # print(size)
                             else :
                                 size="NONE"
                                 # print(size)
                             # Summary = "Text is: " + str(text) + "\n" + "Font is: " + str(font) + "\n" + "color is :" + str(color) + "\n"+"size is :" +str(size)+"\n"+"\n"
                             # # writeFile.write("\r\n"+Summary)
                             # print(Summary)

                             if text!=' ':
                                 dataDict_to[text] = []
                                 dataDict_to[text].append(font)
                                 dataDict_to[text].append(color)
                                 dataDict_to[text].append(size)

                                 # writeFile.write("\r\n" +str(dataDict_to))

                                 print(dataDict_to)
                                 # dataDict_to.clear()



                    for orientVal in orientSumm_to:
                        orntSoup = BeautifulSoup(orientVal, "xml")
                        OrntsoupString = orntSoup.prettify()
                        OrntsoupData = BeautifulSoup(OrntsoupString, 'xml')
                        OrntTagValue = OrntsoupData.find('sectPr')
                        # HeaderRefTagValue = OrntsoupData.find('headerReference')
                        # footerRefTagValue = OrntsoupData.find('footerReference')
                        if str(OrntTagValue) != 'None':
                            for b in orntSoup.find_all('sectPr'):
                                OrntValueF = b.find("pgSz")
                                # HeaderReftValueF = b.find("headerReference")
                                # FooterReftValueF = b.find("footerReference")
                                if OrntValueF:
                                    print(OrntValueF)
                                    orientation = OrntValueF.get("orient")
                                    if str(orientation)!='None':
                                        Land_count=Land_count+1
                                        print("orientation is : "+str(orientation))

                                    else:
                                        Port_count = Port_count + 1
                                        print("orientation is : " + portrait)

            writeFile.write("Developed By : Rajat "+"\r\r\n\n")
            writeFile.write("\r\n" + str(Port_count) + " Page : " + portrait)
            writeFile.write("\r\n" + str(Land_count) + " Page : " + "Landscape")

                                        # if HeaderReftValueF:
                                        #     print(HeaderReftValueF)
                                        #     headerID = HeaderReftValueF.get("id")
                                        #     if str(headerID) != 'None':
                                        #         print("header ID is : " + str(headerID))
                                        # if FooterReftValueF:
                                        #     print(FooterReftValueF)
                                        #     footerID = FooterReftValueF.get("id")
                                        #     if str(footerID) != 'None':
                                        #         print("footer ID is : " + str(footerID))




            # --------------------------------------comparing two result dictonary-------------------------------------------------



            fontBool=False
            sizeBool=False
            ColorBool=False
            keyBool=True

            for keyFrom, valuesFrom in dataDict_from.items():
                for keyTo,valuesTo in dataDict_to.items():
                    if str(keyFrom)==str(keyTo):
                        keyBool=True
                    else:
                        keyBool=False
                    if valuesFrom[0]==valuesTo[0]:
                        fontBool=True
                    else:
                        fontBool=False
                    if valuesFrom[1]==valuesTo[1]:
                        ColorBool=True
                    else:
                        ColorBool=False
                    if valuesFrom[2] == valuesTo[2]:
                        sizeBool=True
                        print('Similar files no changes')
                        # writeFile.write("\r\n" + " -------------: similar files :-------")
                    else:
                        sizeBool=False

            if(fontBool==True and ColorBool==True and sizeBool==True and keyBool==True):
                writeFile.write("\r\n" + " -------------: similar files :-------")
                fontBool=False
                ColorBool=False
                sizeBool=False

            for keyFrom, valuesFrom in dataDict_from.items():
                       for keyTo,valuesTo in dataDict_to.items():
                        if str(keyFrom)==str(keyTo):
                            if valuesFrom[0]!=valuesTo[0]:
                                writeFile.write("\r\n" + keyTo+" font is different ie. : "+valuesTo[0])
                            if valuesFrom[1]!=valuesTo[1]:
                                writeFile.write("\r\n" + keyTo+" color is different ie. : "+valuesTo[1])
                            if valuesFrom[2]!=valuesTo[2]:
                                writeFile.write("\r\n" + keyTo+" size is different ie. : "+str(valuesTo[2]))

            dataDict_to.clear()
            dataDict_from.clear()

        # --------------------------------------------------Comparing Images-----------------------------------------------

            if not os.path.exists(Image_comp_path):
                os.makedirs(Image_comp_path)
            count=0
            for imagepkg in imagePackName:
                image1 = document_from.open(imagepkg).read()
                count=count+1
                f = open(Image_comp_path+commFile+'_BT_Images_'+str(count)+".png",'wb')
                f.write(image1)

        messagebox.showinfo("Success !!","Files compared Successfully,Please Check "+toComp_dirString_old)



