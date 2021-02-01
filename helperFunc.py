#HELPER FUNCTIONS FOR AUTOMATION
import xlrd
from docx import *
from docx.shared import Pt,Inches
import difflib
import json
import glob




def readxls(sh):
    bigTable=[]
    for rx in range(2,sh.nrows):
        if sh.cell_value(rowx=rx, colx=0):
            bigTable.append({
                "head" : sh.cell_value(rowx=rx, colx=0),
                "ly" : int(sh.cell_value(rowx=rx, colx=2)),      #ly:leptomereies ypovolhs ki einai enas ari8mos (1,6)
                "smallTable" : []   
            })
        bigTable[int(sh.cell_value(rowx=rx, colx=2))-1]["smallTable"].append({"text" : sh.cell_value(rowx=rx, colx=1)}) #adding a dict to the list smalltable for every row in xls file the other keys to this this dict will be added when doc is read.
    return bigTable


def printBigTable(bigTable):
    for category in bigTable:
        print("\n"+"*************" + category["head"]+ str(category["ly"]) + "*************"+"\n")
        for line in category["smallTable"]:
            print("line : ",line)


def docParagraphsAdjust(docArray):
    for index,p in enumerate(docArray):
        if ((p.text) == "Σχεδίαση"):
            return docArray[index:]

def printParagraphs(paragraphs):
    for i in paragraphs:
        print(i.text)

def similar(seq1, seq2):
    return difflib.SequenceMatcher(a=seq1.lower().strip(), b=seq2.lower().strip()).ratio()
    

def postBigTable(bigTable,sarg,dse): #sa=search argument in my case the text from paragraphs list        dse: "ΔΙΕΥΘΕΠ" so be careful to have the same key value to the dict
    for category in bigTable:
        for row in category["smallTable"]:
            if similar(row["text"],sarg)>0.93 :
                row[dse] = True

#this is the function that checks in the doc which lines have diefthep synolika and ekp photos and which do not
def getFromDoc(adjustedParagraphs,bigTable):
    count=0
    indexing=[]
    for category in bigTable:
        for row in category["smallTable"]: 
            boolean = True                      #maybe hould delte that check it some point
            for i,p in enumerate(adjustedParagraphs): 
                ratio = similar(row["text"],p.text)          
                if ratio > 0.93:
                    indexing.append(i)
                    count+=1
                    boolean= False
    indexing.sort()     #we have to sort because the excel  has some lines in different order than the word
    for i in range(len(indexing)-1):
        for r in range(indexing[i]+1,indexing[i+1]):
            if  "ΔΙΕΥΘΕΠ" in adjustedParagraphs[r].text :
                postBigTable(bigTable,adjustedParagraphs[indexing[i]].text,"Diefthep")
            if  "ΣΥΝΟΛΙΚΑ" in adjustedParagraphs[r].text :
                postBigTable(bigTable,adjustedParagraphs[indexing[i]].text,"All")
            if  "ΕΚΠΑΙΔΕΥΟΜΕΝΟΙ" in adjustedParagraphs[r].text :
                postBigTable(bigTable,adjustedParagraphs[indexing[i]].text,"Trainees")


#########         Json saving and opening files   (used only for developing reasons)       #########
def saveJson(bigTable):
    with open(r'c:\Users\user\Documents\Python Scripts\automation_excel_word\bigtable.json', 'w', encoding='utf8') as fout:
        json.dump(bigTable, fout,ensure_ascii=False)

def openJson():
    with open(r'c:\Users\user\Documents\Python Scripts\automation_excel_word\bigtable.json', 'r', encoding='utf8') as fout:
        bigTable = json.load(fout)
    return bigTable


#########         image processing       #########
def imagePath(folder,category,line):
    path = r"C:\Users\user\Documents\Python Scripts\automation_excel_word\ΣΚΟΥΤΕΡΗΣ\Figures" + "\\" + folder
    length = len(path)
    path=path  +"\\*.png" 
    listOfphotos = glob.glob(path) 
    string = str(category) + "." + str(line)
    for path in listOfphotos:
        if string in path[length:]:
            return path

