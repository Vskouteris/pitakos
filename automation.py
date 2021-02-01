from helperFunc import *



book = xlrd.open_workbook(r"C:\Users\user\Documents\Python Scripts\automation_excel_word\categories_erotimatologio_taas_cax_pittakos_mytilinaios_21_1.xls")
sh = book.sheet_by_index(0)
doc_erwtimatologio = Document(r"C:\Users\user\Documents\Python Scripts\automation_excel_word\erwtimatologio.docx")   
paragraphs = doc_erwtimatologio.paragraphs
adjustedParagraphs = docParagraphsAdjust(paragraphs)


# #READING FROM FROM THE XLS FILE  AND CREATING A LIST OF DICTINARIES
# bigTable = readxls(sh)
# #READING FROM FROM THE DOC FILE  AND UPDATE THE DICTIONARIES INSIDE SMALLTABLE LIST OF EACH CATEGORY-HEAD
# getFromDoc(adjustedParagraphs,bigTable)
# saveJson(bigTable)

bigTable = openJson()
# #PRINTING THE DATA I HAVE ALREADY READ
# printBigTable(bigTable)
imagePath("Diefthep",1,4)
           

#OUTPUTS
outputDoc = Document()
for category in bigTable:
    p =outputDoc.add_paragraph()
    run = p.add_run(category["head"])
    run.bold = True
    run.font.size = Pt(15)
    run.font.name = "Arial"
    smallTable = category["smallTable"]
    for count,line in enumerate(smallTable):
        p =outputDoc.add_paragraph()
        run = p.add_run(line["text"])
        run.bold = True
        run.font.name = "Arial"
        if "Diefthep" in line:
            p =outputDoc.add_paragraph()
            run = p.add_run("ΔΙΕΥΘΕΠ")
            run.font.name = "Arial"
            p.style = 'List Bullet'
            outputDoc.add_picture(imagePath("Diefthep",category["ly"],count+1),width=Inches(4.5))
        if "Trainees" in line:
            p =outputDoc.add_paragraph()
            run = p.add_run("ΕΚΠΑΙΔΕΥΟΜΕΝΟΙ")
            run.font.name = "Arial"
            p.style = 'List Bullet'
            outputDoc.add_picture(imagePath("Trainees",category["ly"],count+1),width=Inches(4.5))
        if "All" in line:
            p =outputDoc.add_paragraph()
            run = p.add_run("ΣΥΝΟΛΙΚΑ")
            run.font.name = "Arial"
            p.style = 'List Bullet'
            outputDoc.add_picture(imagePath("All",category["ly"],count+1),width=Inches(4.5))

        
outputDoc.add_page_break()
outputDoc.save(r'c:\Users\user\Documents\Python Scripts\automation_excel_word\demo.docx')
    


