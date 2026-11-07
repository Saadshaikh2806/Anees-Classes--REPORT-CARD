"""Shared CSV validation and mark handling for report generators."""
import math
import re
import pandas as pd


def is_absent(value):
    return pd.isna(value) or str(value).strip().lower() in {'', 'nan', 'absent', '-', '—', 'ab', 'a'}


def read_students(path):
    data = pd.read_csv(path, dtype=str).fillna('')
    data.columns = data.columns.str.strip()
    data = data.rename(columns={'Disrict': 'District'})
    missing = [key for key in ['Name', 'Roll No', 'Class', 'Batch'] if key not in data]
    if missing:
        raise ValueError('Missing required CSV columns: ' + ', '.join(missing))
    if data.empty:
        raise ValueError('The CSV contains no students.')
    for key in ['Name', 'Roll No', 'Class', 'Batch']:
        data[key] = data[key].str.strip()
    if data['Name'].eq('').any():
        raise ValueError('Every student must have a name.')
    if 'District' not in data:
        data['District'] = 'N/A'
    for short, long in [('SELF', 'Student Contact No.'), ('FATHER', 'Father/Guardian Contact No.'), ('MOTHER', 'Mother/Guardian Contact No.')]:
        if short not in data:
            data[short] = data[long] if long in data else 'N/A'
        if long not in data:
            data[long] = data[short]
    for column in data:
        if re.fullmatch(r'(?:Maths|GAT|English|NEET|JEE)(?: Marks)?\d+', column):
            data[column] = data[column].map(lambda value: 'Absent' if is_absent(value) else value)
            for value in data[column]:
                if value != 'Absent' and not math.isfinite(float(value)):
                    raise ValueError(f'Invalid objective marks in {column}: {value}')
    return data


def prepare_subjective(data):
    maxima = {}
    columns = [c for c in data if re.fullmatch(r'Subjective Marks\d+', c)]
    def exam(row, column):
        suffix = column.removeprefix('Subjective Marks')
        return (str(row.get('Subject' + suffix, column)).strip(), str(row.get('Subjective Date' + suffix, '')).strip())
    for column in columns:
        for _, row in data.iterrows():
            value = row[column]
            if is_absent(value) or '/' not in str(value):
                continue
            mark, maximum = map(float, str(value).split('/'))
            if not math.isfinite(mark) or not math.isfinite(maximum) or maximum <= 0 or mark < 0 or mark > maximum:
                raise ValueError(f'Invalid marks for {row["Name"]}: {value}')
            key = exam(row, column)
            if key in maxima and maxima[key] != maximum:
                raise ValueError(f'Conflicting maximum marks for {key}: {maxima[key]} and {maximum}')
            maxima[key] = maximum
    for column in columns:
        for index, row in data.iterrows():
            key = exam(row, column)
            if key not in maxima:
                while True:
                    answer = input(f'Enter maximum marks for {key[0]} {key[1]} (default 100): ').strip()
                    try:
                        maximum = float(answer or '100')
                        if math.isfinite(maximum) and maximum > 0:
                            maxima[key] = maximum
                            break
                    except ValueError:
                        pass
                    print('Enter a positive finite number.')
            value = row[column]
            if is_absent(value):
                data.at[index, column] = 'Absent'
            else:
                mark = float(str(value).split('/')[0])
                if not math.isfinite(mark) or not 0 <= mark <= maxima[key]:
                    raise ValueError(f'Invalid marks for {row["Name"]}: {value}')
                data.at[index, column] = f'{mark}/{maxima[key]}'
    return maxima


def subject_maximum(maxima, student, number):
    return maxima[(str(student.get(f'Subject{number}', f'Subjective Marks{number}')).strip(), str(student.get(f'Subjective Date{number}', '')).strip())]


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(value)).replace(' ', '_').strip('. ') or 'student'
