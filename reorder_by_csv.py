import io
import re
import pandas as pd

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from PyPDF2 import PdfFileReader, PdfFileWriter

report_path = r'example_punch.pdf'
sorted_report_path = r'sorted_report.pdf'
mapping_path = r'mapping.csv'

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
emp_id_df = pd.DataFrame({'Emp_ID': ids})

#extract the numbers
emp_id_df['Emp_ID'] = emp_id_df['Emp_ID'].str.extract(r'(\d+)').astype(int)

#add the report page
emp_id_df = emp_id_df.append({'Emp_ID': 0}, ignore_index=True)

#read mapping csv to a df
mapping_df = pd.read_csv(mapping_path)

#rename the columns
mapping_df.columns = ['Emp_ID', 'Sort_Key']

#left outer join the two df 
combo_df = pd.merge(emp_id_df, mapping_df, on='Emp_ID', how='left')

#change all NaN to large number string
combo_df = combo_df.fillna('9999-999999999')

#remove the dash in Sort_Key
combo_df['Sort_Key'] = combo_df['Sort_Key'].str.replace('-', '', regex=False)

#convert Sort_Key to numbers
combo_df['Sort_Key'] = pd.to_numeric(combo_df['Sort_Key'], errors='coerce')

#change it so the report will come first in the df
combo_df.loc[combo_df.Emp_ID == 0, 'Sort_Key'] = 0

#sort by Sort_Key then Emp_ID
combo_df.sort_values(by=['Sort_Key', 'Emp_ID'], inplace=True)

#put the page order into a list
page_order = combo_df.index.tolist()

#reorder the pages to a new pdf
writer = PdfFileWriter()
with open(report_path, 'rb') as infile:
    
    reader = PdfFileReader(infile)
    for entry in page_order:
        writer.addPage(reader.getPage(entry))

    with open(sorted_report_path, 'wb') as outfile:
        writer.write(outfile)