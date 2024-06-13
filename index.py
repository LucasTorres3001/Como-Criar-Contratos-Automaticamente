from docx import Document
import pandas as pd
from datetime import datetime

document_name = 'Contrato.docx'
database = 'informações.xlsx'

document = Document(docx=document_name)
df = pd.read_excel(database)

for line in df.index:
    username = df.loc[line, 'Nome']
    item1 = df.loc[line, 'Item1']
    item2 = df.loc[line, 'Item2']
    item3 = df.loc[line, 'Item3']
    references = {
        'XXXX': username,
        'YYYY': item1,
        'ZZZZ': item2,
        'WWWW': item3,
        'DD': str(datetime.now().day),
        'MM': str(datetime.now().month),
        'AAAA': str(datetime.now().year)
    }
    for paragraph in document.paragraphs:
        for code in references:
            value = references[code]
            paragraph.text = paragraph.text.replace(code, value)
            
    document.save(path_or_stream=f'Contract - {username}.docx')