# ADCI report

Run from PowerShell:

```powershell
cd "C:\Users\ADCI\Desktop\SAAD\REPORT CARD"
python -m pip install -r requirements.txt
python "ADCI report.py"
```

Enter the path to your CSV when prompted. Use `main-NDA.py` for NDA reports or
`B4_main.py` for NEET/JEE reports. Each script saves PDFs in the `Report_Cards`
folder beside the script, regardless of the shell's working directory. Existing
files with the same generated name are replaced. CSV sources are never changed.

CSV input requires `Name`, `Roll No`, `Class`, and `Batch`. Optional district and
contacts use `District`, `SELF`, `FATHER`, `MOTHER`; the longer contact names used
by older generators are also accepted. Roll numbers retain leading zeroes.

Subjective exams use sequential columns `Subjective Date1`, `Subject1`,
`Subjective Marks1`, then the same names ending in 2, 3, etc. Prefer marks such as
`55/70`. Maximum marks are inferred from other rows for the same subject and date.
If no maximum is available, the generator asks; Enter means 100. Invalid or
conflicting maximum marks raise an error rather than producing misleading totals.

Absences (empty marks, `Absent`, `AB`, `A`, `-`, or `—`) count as zero marks and
retain their maximum marks in the percentage denominator. Reports still display
`Absent`. For example, 50/100 followed by an absence out of 100 gives 25%.
An exam column present in the CSV is treated as a scheduled exam for each student;
use separate CSV files when students have different exam schedules.

NDA objective columns use `Exam1`, `Maths1`, `GAT1`, `English1`, with optional
`NEET1` and `JEE1`; NEET/JEE reports use `Exam1`, `NEET1`, `JEE1`. Repeat with
sequential numbers for later exams. Objective marks are numeric or absent. Fixed
maximum marks are Maths 300, GAT 400, English 200, NEET 720, and JEE 360.
Keep `logo.png`, `fonts`, and (for NDA/NEET/JEE reports) `back.pdf` beside the scripts.

Combine individual reports:

```powershell
python merge.py
```

The merger includes only filenames ending in `_report_card.pdf`, sorts serial
numbers numerically, and excludes its output. Re-running it does not duplicate
previously merged pages. Generate different report batches in separate output
folders or move completed PDFs before generating another batch.

Run regression checks:

```powershell
python -m unittest discover -s . -p test_reports.py -v
```

Checks use temporary synthetic data and do not replace student reports.
