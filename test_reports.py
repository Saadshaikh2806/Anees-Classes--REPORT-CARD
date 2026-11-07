import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import pandas as pd
from PyPDF2 import PdfReader
from reportlab.pdfgen import canvas
from report_utils import prepare_subjective
from merge import merge_pdfs

ROOT = Path(__file__).resolve().parent


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ReportTests(unittest.TestCase):
    def test_maxima_and_rows(self):
        data = pd.DataFrame({'Name': ['First', 'Second'], 'Subject1': ['Maths'] * 2,
                             'Subjective Date1': ['01-Jan'] * 2,
                             'Subjective Marks1': ['', '35/70']}, index=[5, 9])
        with patch('builtins.input', side_effect=AssertionError('Unnecessary prompt')):
            maxima = prepare_subjective(data)
        self.assertEqual(list(data.index), [5, 9])
        self.assertEqual(data.at[5, 'Subjective Marks1'], 'Absent')
        self.assertEqual(data.at[9, 'Subjective Marks1'], '35.0/70.0')
        self.assertEqual(maxima[('Maths', '01-Jan')], 70)

    def test_default_and_invalid_maximum(self):
        data = pd.DataFrame({'Name': ['Test'], 'Subject1': ['Maths'], 'Subjective Marks1': ['50']})
        with patch('builtins.input', side_effect=['0', 'nan', '']):
            prepare_subjective(data)
        self.assertEqual(data.at[0, 'Subjective Marks1'], '50.0/100.0')

    def test_generators(self):
        row = {'Name': 'Test / Student', 'Roll No': '001', 'Class': '12', 'Batch': 'NEET',
               'District': 'Pune', 'SELF': '', 'FATHER': '', 'MOTHER': '',
               'Subject1': 'Maths', 'Subjective Date1': '01-Jan', 'Subjective Marks1': '50/100',
               'Subject2': 'Maths', 'Subjective Date2': '02-Jan', 'Subjective Marks2': 'Absent',
               'Exam1': '01 JAN', 'NEET1': '360', 'Maths1': '150', 'GAT1': '200', 'English1': '100',
               'Exam2': '02 JAN', 'NEET2': 'Absent', 'Maths2': 'Absent', 'GAT2': 'Absent', 'English2': 'Absent'}
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            csv = folder / 'input.csv'
            pd.DataFrame([row]).to_csv(csv, index=False)
            for name in ['main_MID-TERM', 'main-NDA', 'B4_main']:
                with self.subTest(generator=name):
                    generator = load(name).ReportCardGenerator()
                    generator.cwd = str(folder / name)
                    output = Path(generator.cwd) / 'Report_Cards'
                    output.mkdir(parents=True)
                    with patch('builtins.input', return_value='100'):
                        generator.generate_report_cards(str(csv))
                    pdfs = list(output.glob('*.pdf'))
                    self.assertEqual(len(pdfs), 1)
                    pdf = PdfReader(str(pdfs[0]))
                    self.assertEqual(len(pdf.pages), 1 if name == 'main_MID-TERM' else 2)
                    text = pdf.pages[0].extract_text()
                    self.assertIn('25.00%', text)
                    self.assertIn('Absent', text)

    def test_merge_repeat_and_order(self):
        with tempfile.TemporaryDirectory() as directory:
            for number in [10, 2, 1]:
                pdf = canvas.Canvas(str(Path(directory) / f'{number}_test_report_card.pdf'))
                pdf.drawString(50, 50, str(number))
                pdf.save()
            for _ in range(2):
                merge_pdfs(directory, 'Merged_Report_Cards.pdf')
                pages = PdfReader(str(Path(directory) / 'Merged_Report_Cards.pdf')).pages
                self.assertEqual([p.extract_text().strip() for p in pages], ['1', '2', '10'])


if __name__ == '__main__':
    unittest.main()
