import importlib.util
from pathlib import Path
import sys
import unittest
from uuid import uuid4
import fitz

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from create_samples import create_samples
spec=importlib.util.spec_from_file_location('analyzer',ROOT/'app.py')
app=importlib.util.module_from_spec(spec); spec.loader.exec_module(app)

class ChangeAnalysisTest(unittest.TestCase):
    def setUp(self):
        self.folder=create_samples(ROOT/'outputs'/('test-'+uuid4().hex))

    def test_excel_revision_comparison(self):
        before=(self.folder/'previous.xlsx').read_bytes()
        old,old_notes=app.load(self.folder/'previous.xlsx'); new,new_notes=app.load(self.folder/'revised.xlsx')
        self.assertEqual(old_notes+new_notes,[])
        results=app.analyze(old,new)
        self.assertEqual({r[0] for r in results if r[1]=='ADDED'},{'9001-TT-301'})
        self.assertEqual({r[0] for r in results if r[1]=='REMOVED'},{'9001-PT-201'})
        changes=[r for r in results if r[1]=='CHANGED']
        self.assertEqual(len(changes),4)
        self.assertEqual({r[5] for r in changes},{'Data Sheet','DCS Scaling','Alarm / Interlock','FAT Procedure'})
        self.assertTrue(all(r[3:5]==('0-100','0-200') for r in changes))
        self.assertEqual(before,(self.folder/'previous.xlsx').read_bytes())
        self.assertEqual(app.analyze(old,old),[])

    def test_pdf_context_and_no_text_warning(self):
        old,notes=app.load(self.folder/'previous.pdf'); new,_=app.load(self.folder/'revised.pdf')
        results=app.analyze(old,new)
        self.assertTrue(notes)
        self.assertIn(('9001-FT-101','CHANGED'),{r[:2] for r in results})
        blank=self.folder/'blank.pdf'
        with fitz.open() as pdf:
            pdf.new_page(); pdf.save(blank)
        records,warnings=app.load(blank)
        self.assertEqual(records,{})
        self.assertTrue(any('no selectable text' in w for w in warnings))

    def test_unmapped_field_remains_review_prompt(self):
        self.assertEqual(app.classify('Unmapped Attribute'),['Engineering Review (rule not mapped)'])

if __name__=='__main__': unittest.main()
