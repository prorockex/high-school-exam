import importlib.util, json, tempfile, unittest, sys
sys.dont_write_bytecode = True
from pathlib import Path

spec = importlib.util.spec_from_file_location('archive', Path(__file__).resolve().parents[1] / 'tools/import_pdf_archive.py')
archive = importlib.util.module_from_spec(spec); spec.loader.exec_module(archive)

class ArchiveTests(unittest.TestCase):
    def test_local_import_and_repeatability(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp); files = root/'input'; app = root/'app'; files.mkdir(); app.mkdir()
            (app/'index.html').write_text('<script id="pdf-catalog" type="application/json">{"papers":[]}</script>')
            pdf = b'%PDF-1.4\nTEST ONLY\n%%EOF'
            (files/'中文</script>.pdf'.replace('/', '_')).write_bytes(pdf)
            (files/'same.pdf').write_bytes(pdf)
            (files/'invalid.pdf').write_bytes(b'<html>not PDF</html>')
            (files/'partial.pdf').write_bytes(b'%PDF-1.4\npartial')
            result = archive.import_archive(files, app, 'https://example.test/test-fixture', '測試學校')
            self.assertEqual(result, {'imported':1,'duplicates':1,'failed':2,'total':1})
            catalog = json.loads((app/'pdf-catalog.json').read_text()); paper = catalog['papers'][0]
            self.assertEqual((app/paper['path']).read_bytes(),pdf); self.assertEqual(paper['school'],'測試學校')
            self.assertEqual(archive.import_archive(files,app)['imported'],0)
            self.assertEqual(len(list((app/'pdfs').glob('*.pdf'))),1)
            self.assertEqual(list((app/'pdfs').glob('*.part')),[])
            (app/paper['path']).write_bytes(b'corrupt')
            self.assertEqual(archive.import_archive(files, app)['imported'],1)
            self.assertEqual((app/paper['path']).read_bytes(),pdf)
    def test_source_validation(self):
        with tempfile.TemporaryDirectory() as temp:
            with self.assertRaises(ValueError): archive.import_archive(Path(temp),Path(temp)/'app','javascript:alert(1)')

if __name__=='__main__': unittest.main()
