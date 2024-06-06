import os #import os function to remove files
from PyPDF4 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os

def create_page_pdf(num, tmp):
    c = canvas.Canvas(tmp, pagesize=letter)
    for i in range(1, num + 1):
        c.setFont('Courier', 9)
        c.drawCentredString(396, 20, f"Page {i} of {num}")
        c.showPage()
    c.save()

def add_page_numbers(pdf_path, newpath):
    tmp = "__tmp.pdf"
    writer = PdfWriter()
    
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        n = len(reader.pages)

        create_page_pdf(n, tmp)

        with open(tmp, "rb") as ftmp:
            number_pdf = PdfReader(ftmp)

            for p in range(n):
                page = reader.pages[p]
                number_layer = number_pdf.pages[p]
                page.mergePage(number_layer)
                writer.add_page(page)

            if len(writer.pages) > 0:
                with open(newpath, "wb") as f:
                    writer.write(f)
    
    os.remove(tmp)


def pdftoc(input_pdf, output_pdf,max_title_lines = 4):
    from pathlib import Path # Brings in path from pathlib important for reading and deleting local files
    from PyPDF2 import PdfReader, PdfWriter # Brings in PyPDF2 Gives us the ability to read, write and annotate pdfs
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle #Following lines imports tools from reportlab used to create pdfs with python
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Flowable
    from reportlab.platypus import PageBreak
    from reportlab.platypus.tableofcontents import TableOfContents
    from reportlab.platypus import BaseDocTemplate, Frame,FrameBreak, PageTemplate, Paragraph
    from functools import partial
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.rl_config import defaultPageSize
    from reportlab.lib.units import inch
    from PyPDF2.generic import AnnotationBuilder
    import pandas
    import xlrd

    writer = PdfWriter() #Shortens Pdfwriter function to just writer for ease of calling
    reader = PdfReader(input_pdf) # Reads imput pdf file
    excelc1 = pandas.read_excel('st_abbrev.xls', usecols =[1]) # Read the Full text column of st_abbrevs
    excelc2 = pandas.read_excel('st_abbrev.xls',usecols=[2]) #Reads the Abbrev column of st_abbrevs
    new_list2 =[] # Creates list to put values from abbrevs in 
    for name, values in excelc2.iterrows(): # iterates thru the abbrev column
        new_list2.append(values[0]) # appends each abbrev
    new_list1 =[] # creates an empty list for full text to be read from
    for name, values in excelc1.iterrows(): # itterates thru column of full text
        new_list1.append(values[0]) # appends all full text values to list
    bookmark_abbrevs = [] # Create list for indexs of abbrews to use for the bookmarks
    filename_abbrevs = [] # Creates lis for indexs of abbrevs to use for the filename
    text_df = pandas.read_excel('st_abbrev.xls') # creates data frame for frame to read fulltext and abbrevs
    text_df.drop(index=0) # Drops columns that don't exist
    for Scope,values in text_df.iterrows(): # Iterates thru the rows of the .xls to grab full text and abbrevs
        if values[0] !="Filename": # conditional statement to get all values that that are not for only the file name
            bookmark_abbrevs.append(Scope) # append the index of those values
        if values[0] !="Bookmark": # conditional statment to get all values that are not for only Bookmark
            filename_abbrevs.append(Scope) # appends index of those values
    bm_fulltext = [] # creates empty list to store full text string to iterate thru later
    bm_abbrev = [] # Creates empty list to store strings of abbrevs  
    fn_fulltext = [] # creates empty list to store full text string
    fn_abbrev = [] # creates empty list to store abbrevs for file names

    for obj in bookmark_abbrevs: # Loop to iterate thru all indexs of abbrevs to use in bookmarks
        bm_fulltext.append(new_list1[obj]) # append full text to list
        bm_abbrev.append(new_list2[obj]) # appends abbrev to list
    for item in filename_abbrevs: # loop to iterate thru all indexs of abbrevs to use in filenaming
        fn_fulltext.append(new_list1[item]) # append full text to list
        fn_abbrev.append(new_list2[item]) # appends abbrev to list

    if max_title_lines == 4:
        bookmarklist = [] #creates empty list for bookmark name and page index to be stored(Note: page index starts at 0 so page 1 is index 0 and so on)
        fnlist = [] # creates empty list for the TOC listings and page numebrs to be stored
        i = 0 # Sets begining index to page 1 
        while i <= len(reader.pages)-1: #While loop grabs TFL listing from page i and edits the text to be in a readable format
            page = reader.pages[i] # Reads page index i
    
            parts = [] # Creates empty list for text to be stored
            def visitor_body(text, cm, tm, fontDict, fontSize): # Visitor_body uses y cords to read text from pdf that fall in those cords
                y = tm[5] # tm[5] is used to set visisor to y cords if we wanted x we would use tm[4]
                if y > 463 and y < 515: # sets visitor to area which TFL title is found in most docs
                    parts.append(text) # appends read text to list
            page.extract_text(visitor_text=visitor_body) # Extracts the text from the page
            text_body = "".join(parts) # Joins text that was previously a list into a str
            text_body= text_body.split("\n") # splits text back in to a list while removing paragraph breaks
            j = 0
            k = 0
            while j == 0:
                if text_body[k].find("Table") == -1 and text_body[k].find("Figure") == -1 and text_body[k].find("Listing") == -1:
                    k+=1
                else:
                    j = k
            y = 0
    
            while y == 0:
                x = 0
                while x <= len(text_body)-1:
                    if text_body[x].find("Analysis Set") == -1:
                        if x == len(text_body)-1:
                            y = j + 1
                            x += 1
                        else:
                            x += 1
                    elif text_body[x].find("Analysis Set") != -1:
                        y = x
                        x = len(text_body) + 1
            text_body = text_body[j:y+1]
            body_text = text_body # Grabs the 2nd thru 5th items in the list the 1st is always "" and this removes text dulpication done sometimes by the reader
            body_text[0]=body_text[0].split() # splits the TFL classifaction from its index
            body_text[0][0]=body_text[0][0][0] #Grabs the first letter from TFL
            body_text[0]=" ".join(body_text[0])
            c = 0
            for text in body_text:
                if text.find(": ") == -1:
                    c+=1
                    pass
                else:
                    text = text.split(": ")
                    text = text[1]
                    text= "".join(text)
                    body_text[c] = text
                    c+=1
            
            
    
            body_text = '-'.join(body_text)
            k =0 
            bm_body_text = body_text
            while k <= len(bm_fulltext)-1:
                if body_text.find(bm_fulltext[k]) != -1:
                    bm_body_text = bm_body_text.replace(bm_fulltext[k],bm_abbrev[k])
                    k+=1
                else:
                    k+=1
            k =0 
            fn_body_text = body_text
            while k <= len(fn_fulltext)-1:
                if body_text.find(fn_fulltext[k]) != -1:
                    fn_body_text = fn_body_text.replace(fn_fulltext[k],fn_abbrev[k])
                    k+=1
                else:
                    k+=1
            
            bm_temp_list = [] # creates a temporay list to hold the bookmark title and its index in a list which can later be used to make a list of lists
    
            bm_temp_list.append(bm_body_text) # appends the title for index 0 of the list
            bm_temp_list.append(i) #appends the index of the title
            bookmarklist.append(bm_temp_list) # appends the list (title,index) to a list of lists
    
            fn_temp_list = []
    
            fn_temp_list.append(fn_body_text)
            fn_temp_list.append(i)
            fnlist.append(fn_temp_list)
            i+=1
            
        def remove_lists_with_same_first_as_previous(lst): # function used to remove title duplicates in other words only keeps the index where the title first appears
            updated_list = [] # creates empty list to put unique titles and page numbers in
            for sublist in lst: # iterates thru the bookmark list
                if not updated_list or sublist[0] != updated_list[-1][0]: # checks for uniquness 
                    updated_list.append(sublist) # adds unique element to list
            return updated_list # returns list of unique elements
    
        bookmarklist = remove_lists_with_same_first_as_previous(bookmarklist) # Makes the list only first appareance of each unique title
        fnlist = remove_lists_with_same_first_as_previous(fnlist)
        
        reader = PdfReader(input_pdf) # reads the first page of the PDF
        page = reader.pages[0] # same as above
        right_heading = [] # Creates empty list to put header information
        
        def visitor_body(text, cm, tm, fontDict, fontSize): #Function that only reads the top right margin of the pdf
            y = tm[5] # sets y cord
            x = tm[4] #sets x cord
            if y > 500 and y <720 : #creates subsection of pdf between y vals
                if x >300 and x<800: #creates subsection of pdf between x vals
                    right_heading.append(text)
        page.extract_text(visitor_text=visitor_body)

        if len(right_heading)==0:
            right_heading.append("")
            right_heading.append("")
            right_heading.append("")


        reader = PdfReader(input_pdf)  #Function that only reads the top left margin of the pdf
        page = reader.pages[0]
        left_heading = []
        
        def visitor_body(text, cm, tm, fontDict, fontSize):
            y = tm[5]
            x = tm[4]
            if y > 500 and y <720 :
                if x >0 and x<300:
                    left_heading.append(text)
        page.extract_text(visitor_text=visitor_body)
    

        Title = "TABLE OF CONTENTS" # Sets title for PDF to be TABLE OF CONTENTS
        UCB = 'UCB' # Sets Top right Header to be UCB
        lheader2 = left_heading[2]
        lheader3 = left_heading[4]
        
        if len(right_heading) ==2:
            rheader1 = ""
            rheader2 = right_heading[0]
            rheader3 = right_heading[1]
            
        elif len(right_heading) ==3:
            rheader1 = right_heading[0]
            rheader2 = right_heading[1]
            rheader3 = right_heading[2]
            
        else:
            print("Possible issue with reading right side of Header!, Code may need to be tweaked!")
            
        def myFirstPage(canvas, doc):
            canvas.saveState()
            canvas.setFont('Times-Roman',14)
            canvas.drawString(doc.width/2, 504, Title)
            canvas.setFont('Courier',9)
            canvas.drawString(63.36, 546.6, UCB)
            canvas.setFont('Courier',9)
            canvas.drawString(63.36, 539.4, lheader2)
            canvas.setFont('Courier',9)
            canvas.drawString(63.36, 532.2,lheader3 )
            canvas.restoreState()
            
        def on_remaining_pages(canvas,doc):
            canvas.saveState()
            canvas.setFont('Courier',9)
            canvas.drawString(22.35+2.54, 600, UCB)
            canvas.restoreState()
            
        class DelayedRef(Flowable):
            _ZEROSIZE = True

            def __init__(self, toc, *args):
                self.args = args
                self.toc = toc

            def wrap(self, w, h):
                return 0, 0

            def draw(self, *args, **kwd):
                self.toc.addEntry(*self.args)


        def header(canvas, doc, content):
            canvas.saveState()
            w, h = content.wrap(-1*doc.width, doc.topMargin)
            content.drawOn(canvas, doc.leftMargin, doc.height + doc.topMargin - h)
            canvas.restoreState()
            
        def temp_simple_toc():
            doc = SimpleDocTemplate("tempsimple_toc.pdf", 
                                    pagesize = [792.0,612.0],
                                    leftMargin=63.36,
                                    rightMargin=56.88,
                                    topMargin= 108,
                                    bottomMargin = 72)
            story = []
            styles = getSampleStyleSheet()
            toc = TableOfContents()
            toc.levelStyles = [
                ParagraphStyle(fontName='Times-Roman',
                               fontSize=14,
                               name='Heading1',
                               leftIndent=-20,
                               firstLineIndent=0,
                                spaceBefore=5,
                               leading=16),
                ParagraphStyle(fontName='Courier',
                               fontSize=9,
                               name='Heading2',
                               leftIndent=-20,
                               firstLineIndent=0,
                               spaceBefore=2.54,
                               leading=14.4,
                               textColor = (0,0,1)),
            ]
            s = Spacer(width=0, height=21.6)
            story.insert(1,s)
            story.append(toc)
            i = 0 
    

            while i <= len(remove_lists_with_same_first_as_previous(fnlist))-1:
                story.append(DelayedRef(toc, 1, fnlist[i][0], fnlist[i][1]))
                i += 1

            doc.multiBuild(story, onFirstPage=myFirstPage,onLaterPages=on_remaining_pages)

        temp_simple_toc()
        tempreader = PdfReader("tempsimple_toc.pdf") 
        

        Title = "TABLE OF CONTENTS"
        lheader1 = left_heading[0]
        lheader2 = left_heading[2]
        lheader3 = left_heading[4]
        
        if len(right_heading) ==2:
            rheader1 = ""
            rheader2 = right_heading[0]
            rheader3 = right_heading[1]
            
        elif len(right_heading) ==3:
            rheader1 = right_heading[0]
            rheader2 = right_heading[1]
            rheader3 = right_heading[2]
            
        else:
            print("Possible issue with reading right side of Header!, Code may need to be tweaked!")
            
        def myFirstPage(canvas, doc):
            canvas.saveState()
            canvas.setFont('Times-Roman',14)
            canvas.drawCentredString(396, 504, Title)
            canvas.setFont('Courier',9)
            canvas.drawString(63.36, 550.2, lheader1)
            canvas.drawString(63.36, 539.4, lheader2)
            canvas.drawString(63.36, 528.6,lheader3 )
            canvas.drawRightString(734.4, 550.2, rheader1)
            canvas.drawRightString(734.4, 539.4, rheader2)
            canvas.drawRightString(734.4,528.6,rheader3 )
            canvas.restoreState()
            
        def on_remaining_pages(canvas,doc):
            canvas.saveState()
            canvas.setFont('Courier',9)
            canvas.drawString(63.36, 550.2, lheader1)
            canvas.drawString(63.36, 539.4, lheader2)
            canvas.drawString(63.36, 528.6,lheader3 )
            canvas.drawRightString(734.4, 550.2, rheader1)
            canvas.drawRightString(734.4, 539.4, rheader2)
            canvas.drawRightString(734.4, 528.6,rheader3 )
            canvas.restoreState()
            
        class DelayedRef(Flowable):
            _ZEROSIZE = True

            def __init__(self, toc, *args):
                self.args = args
                self.toc = toc

            def wrap(self, w, h):
                return 0, 0

            def draw(self, *args, **kwd):
                self.toc.addEntry(*self.args)


        def header(canvas, doc, content):
            canvas.saveState()
            w, h = content.wrap(-1*doc.width, doc.topMargin)
            content.drawOn(canvas, doc.leftMargin, doc.height + doc.topMargin - h)
            canvas.restoreState()
            
        def simple_toc():
            doc = SimpleDocTemplate("simple_toc.pdf", 
                                    pagesize = [792.0,612.0],
                                    leftMargin=63.36,
                                    rightMargin=56.88,
                                    topMargin= 108,
                                    bottomMargin = 72)
            story = []
            styles = getSampleStyleSheet()
            toc = TableOfContents()
            toc.levelStyles = [
                ParagraphStyle(fontName='Times-Roman',
                               fontSize=14,
                               name='Heading1',
                               leftIndent=-20,
                               firstLineIndent=0,
                               spaceBefore=5,
                               leading=16),
                ParagraphStyle(fontName='Courier',
                               fontSize=9,
                               name='Heading2',
                               leftIndent=-20,
                               firstLineIndent=0,
                               spaceBefore=2.54,
                               leading=14.4,
                               textColor = (0,0,1)),
            ]
            s = Spacer(width=0, height=21.6)
            story.insert(1,s)
            story.append(toc)
            i = 0 
    

            while i <= len(remove_lists_with_same_first_as_previous(fnlist))-1:
                story.append(DelayedRef(toc, 1, fnlist[i][0], fnlist[i][1]+1+len(tempreader.pages)))
                i+=1

            doc.multiBuild(story, onFirstPage=myFirstPage,onLaterPages=on_remaining_pages)

        simple_toc()
            
            

        pdflist = ['simple_toc.pdf',input_pdf]

        for pdf in pdflist:
            writer.append(pdf)
    
        with open("merged-pdf.pdf", "wb") as v:
            writer.write(v)
    
        add_page_numbers("merged-pdf.pdf","merged-pdf1.pdf")
        

        tempreader = PdfReader("tempsimple_toc.pdf") 

        writer = PdfWriter()
        reader = PdfReader("merged-pdf1.pdf")
        i=0
        target_page_index = 0
        
        while i <= len(reader.pages)-1:
            page = reader.pages[i]
            writer.add_page(reader.pages[i])
            i+=1
        reader = PdfReader("merged-pdf.pdf")
        j=0
        
        while j <= len(tempreader.pages)-1:
            new_title_list = []
            page = reader.pages[j]
            parts = []
            
            def visitor_body(text, cm, tm, fontDict, fontSize):
                y = tm[5]
                if y > -2000 and y <400 :
                    parts.append(text)
            page.extract_text(visitor_text=visitor_body)
            text_body = "".join(parts)
            text_body= text_body.split("\n")
            
            for text in text_body:
                if text =="":
                    text_body.pop()
                else:
                    if text[0:2] =="F " or text[0:2] =="T " or text[0:2] =="L ":
                        new_title_list.append(text)
                
            if j == 0: # Creates links on the first page of TOC
                k=1
                for text in text_body: # C
                    if text[0] == " " or text[0] == "0" or text[0] == "1" or text[0] == "2" or text[0] == "3" or text[0] == "4" or text[0] == "5" or text[0] == "6" or text[0] == "7" or text[0] == "8" or text[0] == "9":
            
                        k += 1
                    elif text[0:2] =="F " or text[0:2] =="T " or text[0:2] =="L ":
                        x = 480-(14.5*k)
                        x1= 480-(14.5*(k-1))
                        annotation = AnnotationBuilder.link(rect=(45, x1, 730, x), target_page_index=bookmarklist[target_page_index][1]++len(tempreader.pages))
                        writer.add_annotation(page_number=j, annotation=annotation)
                        target_page_index +=1
                        k += 1 
                    else:
                        x = 480-(14.5*k)
                        x1= 480-(14.5*(k-1))
                        annotation = AnnotationBuilder.link(rect=(45, x1, 730, x), target_page_index=bookmarklist[target_page_index-1][1]+len(tempreader.pages))
                        writer.add_annotation(page_number=j, annotation=annotation)
                        k += 1

    
            elif j != 0: # Creates links for remaining pages of the TOC
                k=1
                for text in text_body:
                    if text[0] == " " or text[0] == "0" or text[0] == "1" or text[0] == "2" or text[0] == "3" or text[0] == "4" or text[0] == "5" or text[0] == "6" or text[0] == "7" or text[0] == "8" or text[0] == "9": # Creates spacing between linked TOC entries
            
                        k += 1
                    elif text[0:2] =="F " or text[0:2] =="T " or text[0:2] =="L ": # Creates linked box for first line or entire TOC entry 
                        x = 510-(14.5*k)
                        x1= 510-(14.5*(k-1))
                        annotation = AnnotationBuilder.link(rect=(45, x1, 730, x), target_page_index=bookmarklist[target_page_index][1]+len(tempreader.pages))
                        writer.add_annotation(page_number=j, annotation=annotation)
                        target_page_index +=1
                        k += 1 
                    else: # Creates linked box over text that has been wrapped
                        x = 510-(14.5*k)
                        x1= 510-(14.5*(k-1))
                        annotation = AnnotationBuilder.link(rect=(45, x1, 730, x), target_page_index=bookmarklist[target_page_index-1][1]+len(tempreader.pages))
                        writer.add_annotation(page_number=j, annotation=annotation)
                        k += 1
            j+=1
            x=0
        while x <= len(bookmarklist)-1:# Creates bookmarks in pdf by creating a list of text and index and writes them to the Final PDF
            temptext = bookmarklist[x][0] # pulls title for the current x value
            tempnum = bookmarklist[x][1] + len(tempreader.pages) # pulls page number for current x
            parent = writer.add_outline_item(temptext,tempnum) # writes bookmark to pdf with linked pagez
            x+=1

        with open(output_pdf, "wb") as qc: # Saves final PDF as output_pdf input
            writer.write(qc)

    os.remove("merged-pdf.pdf") # Removes merged-pdf
    os.remove("merged-pdf1.pdf") # removes merged-pdf1
    os.remove("tempsimple_toc.pdf") # removes tempsimple_toc
            

  
        
        
        
def pdf_bundle(Directory,outputpdf, eu = False): # Creates a bundled PDF for a specific directory of PDFS
    import os
    from PyPDF2 import PdfReader,PdfWriter
    writer = PdfWriter()
    file_list = []
    cwd = os.getcwd()
    directory = os.fsencode(Directory)
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        if filename.endswith(".pdf"):
            file_list.append(filename)
        
        else:
            continue
    file_list.sort()
    
    bookmark_title_list = []
    page_num_list = []
    os.chdir(Directory)
    i=0
    bpage = 0
    page_num_list.append(bpage)
    for name in file_list:
        reader = PdfReader(name)
        test =reader.outline
        bookmark_title_list.append(test[0].title)
        
        
        
        writer.append(name, import_outline=False)
        parent = writer.add_outline_item(bookmark_title_list[i],bpage)
        i+=1
        bpage += len(reader.pages)
        page_num_list.append(bpage)
        
    os.chdir(cwd)  
    with open(outputpdf, "wb") as v:
        writer.write(v)
    if eu == True:
        pdf_spliter(outputpdf,page_num_list,bookmark_title_list,"Test",file_list)
                    
        
        
        
        
def pdf_spliter(input_pdf, page_list, bm_list,Directory, file_list):
    from PyPDF2 import PdfReader,PdfWriter
    import os
    k=0
    cwd = os.getcwd()
    while k == 0:
        writer1 = PdfWriter()
        writer2 = PdfWriter()
        
        if os.path.getsize(input_pdf) >= 209715200:
            half_split = int(0.5*len(file_list))
            files_1 = []
            bm_1 = []
            pn_1 = []
            files_2 = []
            bm_2 = []
            pn_2 =[]
            for i in range(0,half_split):
                files_1.append(file_list[i])
                bm_1.append(bm_list[i])
                pn_1.append(page_list[i])

            for i in range(half_split,len(file_list)):
                files_2.append(file_list[i])
                bm_2.append(bm_list[i])
                pn_2.append(page_list[i]-page_list[half_split])
            
            input_pdf =input_pdf.split(".")
            input_pdf = input_pdf[0]
            f = 0
            os.chdir(Directory)
            for files in files_1:
                writer1.append(files,import_outline=False )
                parent = writer1.add_outline_item(bm_1[f],pn_1[f])
                f +=1
            os.chdir(cwd)
            num = 1
            with open(f"{input_pdf}_{num}.pdf", "wb") as c:
                writer1.write(c)
            num += 1
            os.chdir(Directory)
            f = 0
            for files in files_2:
                writer2.append(files,import_outline=False )
                parent = writer2.add_outline_item(bm_2[f],pn_2[f])
                f+=1
                    
            os.chdir(cwd)
            with open(f"{input_pdf}_{num}.pdf","wb") as d:
                writer2.write(d)
            if os.path.getsize(f"{input_pdf}_1.pdf") >= 209715200:
                new_bm1 =[]
                new_bm1 = bm_1
                new_pn1 = []
                new_pn1 = pn_1
                new_tl1 = []
                new_tl1 = files_1
                pdf_spliter(f"{input_pdf}_1.pdf",new_pn1,new_bm1,"Test",new_tl1)
            if os.path.getsize(f"{input_pdf}_2.pdf") >= 209715200:
                new_bm2 =[]
                new_bm2 = bm_2
                new_pn2 = []
                new_pn2 = pn_2
                new_tl2 = []
                new_tl2 = files_2
                pdf_spliter(f"{input_pdf}_2.pdf",new_pn2,new_bm2,"Test",new_tl2)
                
            k+=1
                
    