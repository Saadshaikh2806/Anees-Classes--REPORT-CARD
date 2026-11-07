import os
import re
import tempfile
from pathlib import Path
from PyPDF2 import PdfMerger


def merge_pdfs(folder_path, output_filename):
    folder = Path(folder_path).resolve()
    output = folder / output_filename
    def order(path):
        return [int(part) if part.isdigit() else part.lower()
                for part in re.split(r'(\d+)', path.name)]
    files = sorted((p for p in folder.iterdir() if p.is_file()
                    and p.name.lower().endswith('_report_card.pdf')
                    and p.resolve() != output.resolve()), key=order)
    if not files:
        raise ValueError(f'No individual report cards found in {folder}')
    merger = PdfMerger()
    temporary = None
    try:
        for pdf in files:
            merger.append(str(pdf))
        with tempfile.NamedTemporaryFile(dir=folder, suffix='.tmp', delete=False) as target:
            temporary = target.name
            merger.write(target)
        merger.close()
        os.replace(temporary, output)
        print(f'Merged {len(files)} report cards into {output}')
    finally:
        merger.close()
        if temporary and os.path.exists(temporary):
            os.remove(temporary)


if __name__ == '__main__':
    merge_pdfs(Path(__file__).resolve().parent / 'Report_Cards', 'Merged_Report_Cards.pdf')
