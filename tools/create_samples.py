from pathlib import Path
from openpyxl import Workbook
import fitz

def create_samples(folder):
    folder=Path(folder); folder.mkdir(parents=True,exist_ok=True)
    revisions={
        'previous':[('9001-FT-101','4-20 mA','0-100'),('9001-PT-201','4-20 mA','0-10')],
        'revised':[('9001-FT-101','4-20 mA','0-200'),('9001-TT-301','4-20 mA','0-300')],
    }
    for name,rows in revisions.items():
        book=Workbook(); sheet=book.active; sheet.title='Instrument Index'
        sheet.append(['Tag Number','Signal','Range'])
        for row in rows: sheet.append(row)
        book.save(folder/(name+'.xlsx')); book.close()
        document=fitz.open(); page=document.new_page()
        page.insert_text((40,40),'SYNTHETIC EXAMPLE - NOT A PROJECT DOCUMENT',fontsize=12)
        for i,(tag,signal,scale) in enumerate(rows):
            page.insert_text((40,90+i*30),f'{tag}  Signal: {signal}  Range: {scale}',fontsize=11)
        document.set_metadata({'title':'Synthetic Revision Example','author':'Synthetic example generator'})
        document.save(folder/(name+'.pdf')); document.close()
    return folder

if __name__=='__main__': print('Created synthetic revision pairs in',create_samples(Path(__file__).resolve().parents[1]/'samples'))
