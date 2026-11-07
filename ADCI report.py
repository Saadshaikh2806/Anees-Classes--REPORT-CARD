"""ADCI report: choose a generator and create individual student PDFs."""
from pathlib import Path
import runpy


def main():
    print('ADCI report')
    print('1. Mid-term / subjective report')
    print('2. NDA report')
    print('3. NEET / JEE report')
    choices = {'1': 'main_MID-TERM.py', '2': 'main-NDA.py', '3': 'B4_main.py'}
    choice = input('Choose report type (default 1): ').strip() or '1'
    if choice not in choices:
        raise SystemExit('Invalid choice. Enter 1, 2, or 3.')
    runpy.run_path(str(Path(__file__).resolve().parent / choices[choice]), run_name='__main__')


if __name__ == '__main__':
    main()
