import io
import re
import pandas as pd

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from PyPDF2 import PdfFileReader, PdfFileWriter

report_path = r'example_punch.pdf'
sorted_report_path = r'emp_sorted_report.pdf'

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = io.StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages,
                                  password=password,
                                  caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text

#extract the text from the pdf
report_txt = convert_pdf_to_txt(report_path)

#remove the cids
cid_pattern = re.compile(r'\(cid:\d\d\)\n')
cid_removed = re.sub(cid_pattern, '', report_txt)

#clean up the emp number
emp_id_pattern = re.compile(r'Emp\s#\s')
emp_id_fix = r'EmpID '
emp_fixed = re.sub(emp_id_pattern, emp_id_fix, cid_removed)

#replace all newlines w/ space
cleaned = emp_fixed.replace('\n', ' ')

#get all the employee ids
ids = re.findall(r'EmpID\s[\d]+', cleaned)

#put emp ids in a df
df = pd.DataFrame({'Emp_ID': ids})

#extract the numbers
df['Emp_ID'] = df['Emp_ID'].str.extract(r'(\d+)').astype(int)

#get the number of pages of employee timesheets
emp_pages = len(df)

#get the total number of pages
with open(report_path, 'rb') as infile:
    reader = PdfFileReader(infile)
    total_pages = reader.getNumPages()

#calculate the report page length
report_pages = total_pages - emp_pages

#add the report page(s)
rep_length = report_pages
while rep_length:
    place = 0 - rep_length
    df = df.append({'Emp_ID': place}, ignore_index=True)
    rep_length -= 1

#sort by emp id
df.sort_values(by=['Emp_ID'], inplace=True)

#put the index into a list
page_order = df.index.tolist()

#reorder pages into a new pdf
writer = PdfFileWriter()
with open(report_path, 'rb') as infile:
    
    reader = PdfFileReader(infile)
    for entry in page_order:
        writer.addPage(reader.getPage(entry))

    with open(sorted_report_path, 'wb') as outfile:
        writer.write(outfile)