import PyPDF2
pdffileobj=open('in.pdf','rb')
pdfreader=PyPDF2.PdfReader(pdffileobj)
x=len(pdfreader.pages)
pageobj=pdfreader.pages[x-1]
text=pageobj.extract_text()
file1=open(r"out.txt","a")
file1.writelines(text)