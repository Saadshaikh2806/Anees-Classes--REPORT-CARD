import pandas as pd
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import re
from PyPDF2 import PdfReader, PdfWriter

class ReportCardGenerator:
    def __init__(self):
        self.cwd = os.getcwd()
        self.Make_Directories(self.cwd + "/Report_Cards")
        self.setup_fonts()
        self.institute_details = {
            "name": "ANEES DEFENCE CAREER INSTITUTE",
            "address_line1": "S.no.125, Sai Niketan Colony,Near Madhuban Society,",
            "address_line2": "Alandi Rd, Vishrant Wadi, Pune, Maharashtra 411015",
            "phone": "+91 9890826344, +91 9021305751",
            "email": "kutty.anees@gmail.com",
            "website": "https://www.aneesclasses.com"
        }
        self.logo_path = "C:/Users/ADCI/Desktop/REPORT CARD/logo.png"
        self.table_color = colors.Color(0.1, 0.3, 0.1, alpha=0.1)  # Pale dark green
        self.second_page_pdf = "C:/Users/ADCI/Desktop/REPORT CARD/back.pdf"  # Update this path
        self.weightage = {
            'Objective': 0.3,
            'Subjective': 0.3,
            'Co-curricular': 0.2,
            'Extra-curricular': 0.2
        }

    def Make_Directories(self, path):
        os.makedirs(path, exist_ok=True)

    def setup_fonts(self):
        fonts_dir = os.path.join(os.path.dirname(__file__), 'fonts')
        pdfmetrics.registerFont(TTFont('Roboto', os.path.join(fonts_dir, 'Roboto-Regular.ttf')))
        pdfmetrics.registerFont(TTFont('Roboto-Bold', os.path.join(fonts_dir, 'Roboto-Bold.ttf')))
        pdfmetrics.registerFont(TTFont('Roboto-Italic', os.path.join(fonts_dir, 'Roboto-Italic.ttf')))
        pdfmetrics.registerFont(TTFont('Roboto-Black', os.path.join(fonts_dir, 'Roboto-Black.ttf')))
        pdfmetrics.registerFont(TTFont('Roboto-BlackItalic', os.path.join(fonts_dir, 'Roboto-BlackItalic.ttf')))

    def load_data(self, file_path):
        df = pd.read_csv(file_path)
        if 'District' not in df.columns:
            df['District'] = 'N/A'

        df['Name'] = df['Name'].astype(str)
        df['Roll No'] = df['Roll No'].astype(str)
        df['Class'] = df['Class'].astype(str)
        df['Batch'] = df['Batch'].astype(str)
        df['District'] = df['District'].astype(str)

        def extract_date_and_code(exam_string):
            if pd.isna(exam_string):
                return pd.Series({'Date': '', 'Exam Code': ''})

            months = {
                'JAN': 'Jan', 'FEB': 'Feb', 'MAR': 'Mar', 'APR': 'Apr',
                'MAY': 'May', 'JUN': 'Jun', 'JUL': 'Jul', 'AUG': 'Aug',
                'SEP': 'Sep', 'OCT': 'Oct', 'NOV': 'Nov', 'DEC': 'Dec'
            }

            date_pattern = r'\b(\d{1,2})\s*([A-Z]{3})\b|\b([A-Z]{3})\s*(\d{1,2})\b'
            match = re.search(date_pattern, exam_string.upper())

            if match:
                day, month, month_alt, day_alt = match.groups()
                if day and month:
                    date = f"{day.zfill(2)}-{months[month]}"
                elif month_alt and day_alt:
                    date = f"{day_alt.zfill(2)}-{months[month_alt]}"
                else:
                    return pd.Series({'Date': '', 'Exam Code': exam_string})

                exam_code = re.sub(date_pattern, '', exam_string).strip()
                return pd.Series({'Date': date, 'Exam Code': exam_code})
            else:
                return pd.Series({'Date': '', 'Exam Code': exam_string})

        for i in range(1, 100):  # Assume a large upper limit
            exam_col = f'Exam{i}'
            neet_col = f'NEET{i}'
            jee_col = f'JEE{i}'
            
            if exam_col not in df.columns:
                break
            
            df[[f'Date{i}', f'Exam Code{i}']] = df[exam_col].apply(extract_date_and_code)
            
            if neet_col in df.columns:
                df[f'NEET Marks{i}'] = df[neet_col]
            if jee_col in df.columns:
                df[f'JEE Marks{i}'] = df[jee_col]
            
            # Drop the original columns to avoid duplication
            columns_to_drop = [col for col in [exam_col, neet_col, jee_col] if col in df.columns]
            df = df.drop(columns=columns_to_drop)

        return df

    def generate_report_cards(self, file_path):
        self.data = self.load_data(file_path)
        self.adjust_weightage()
        for _, student in self.data.iterrows():
            self.generate_report_card(student)
        print(f"Report cards generated and saved in {self.cwd}/Report_Cards")

    def adjust_weightage(self):
        columns = self.data.columns
        if not any('Subjective' in col for col in columns):
            self.weightage['Objective'] += self.weightage['Subjective']
            self.weightage['Subjective'] = 0
        if 'Co-curricular Attendance' not in columns and 'Extra-curricular Attendance' not in columns:
            self.weightage['Objective'] += self.weightage['Co-curricular'] + self.weightage['Extra-curricular']
            self.weightage['Co-curricular'] = 0
            self.weightage['Extra-curricular'] = 0
        elif 'Co-curricular Attendance' not in columns:
            self.weightage['Extra-curricular'] += self.weightage['Co-curricular']
            self.weightage['Co-curricular'] = 0
        elif 'Extra-curricular Attendance' not in columns:
            self.weightage['Co-curricular'] += self.weightage['Extra-curricular']
            self.weightage['Extra-curricular'] = 0

    def generate_report_card(self, student):
        batch = student['Batch'].upper()
        if batch not in ['JEE', 'NEET']:
            print(f"Skipping student {student['Name']} with unknown batch: {batch}")
            return

        file_name = f"{student['Roll No']}_{student['Name'].replace(' ', '_')}_{batch}_report_card.pdf"
        file_path = os.path.join(self.cwd, "Report_Cards", file_name)

        # Generate the first page
        c = canvas.Canvas(file_path, pagesize=A4)
        self.draw_report_card(c, student, batch)
        c.save()

        # Attach the second page
        self.attach_second_page(file_path)

    def attach_second_page(self, report_card_path):
        # Read the generated report card
        report_card = PdfReader(report_card_path)
        
        # Read the second page PDF
        second_page = PdfReader(self.second_page_pdf)

        # Create a new PDF writer
        output = PdfWriter()

        # Add the report card page
        output.add_page(report_card.pages[0])

        # Add the second page
        output.add_page(second_page.pages[0])

        # Save the combined PDF
        with open(report_card_path, 'wb') as output_file:
            output.write(output_file)

    def draw_report_card(self, c, student, batch):
        width, height = A4

        c.setFillColorRGB(0, 0, 0)  # Black
        c.setFont("Roboto-Bold", 18)
        c.drawString(2 * cm, height - 2 * cm, self.institute_details["name"])
        c.setFont("Roboto-Bold", 10)
        c.drawString(2 * cm, height - 2.75 * cm, self.institute_details["address_line1"])
        c.drawString(2 * cm, height - 3.25 * cm, self.institute_details["address_line2"])
        c.drawString(2 * cm, height - 3.75 * cm, f"Phone: {self.institute_details['phone']} | Email: {self.institute_details['email']}")
        c.drawString(2 * cm, height - 4.25 * cm, f"Website: {self.institute_details['website']}")

        # Logo
        logo = ImageReader(self.logo_path)
        c.drawImage(logo, width - 7.5 * cm, height - 5 * cm, width=6 * cm, height=4 * cm, mask='auto')

        # Report Card Title
        c.setFont("Roboto-Bold", 16)
        c.drawString(2 * cm, height - 5.5 * cm, f"{batch} STUDENT'S PERFORMANCE REPORT CARD")

        # Student Info
        self.draw_info_box(c, student, 2 * cm, height - 9 * cm, width - 4 * cm, 3 * cm)

        # Exam Results
        self.draw_exam_results(c, student, batch, 2 * cm, height - 15 * cm, width - 6 * cm, 6 * cm)

        # Subject Marks and Activities
        self.draw_subject_marks_and_activities(c, student, 2 * cm, height - 21 * cm, width - 4 * cm, 5 * cm, heading_y=height - 16 * cm)

        # Total Marks
        self.draw_total_marks(c, student, batch, 2 * cm, height - 26 * cm, width - 4 * cm, 6 * cm)

        # Contact Info
        self.draw_contact_info(c, student, 11.5 * cm, 20.8 * cm, width - 4 * cm, 4 * cm)

        # Footer
        c.setFont("Roboto-Italic", 8)
        c.drawString(2 * cm, 1 * cm, "This is an official document of ANEES DEFENCE CAREER INSTITUTE")

    def draw_info_box(self, c, student, x, y, width, height):
        c.setStrokeColorRGB(0, 0, 0)  # Black border
        c.rect(x, y, width, height)
        c.setFont("Roboto-Black", 12)
        c.drawString(x + 0.5 * cm, y + height - 0.75 * cm, f"Name: {student['Name'].upper()}")
        c.drawString(x + 0.5 * cm, y + height - 1.5 * cm, f"Roll No: {student['Roll No']}")
        c.drawString(x + width / 5, y + height - 1.5 * cm, f"Class: {student['Class'].upper()}")
        c.drawString(x + 0.5 * cm, y + height - 2.25 * cm, f"Batch: {student['Batch']}")
        if 'District' in student:
            c.drawString(x + width / 5, y + height - 2.25 * cm, f"District: {student['District'].upper()}")
        else:
            c.drawString(x + width / 5, y + height - 2.25 * cm, "District: N/A")

    def draw_exam_results(self, c, student, batch, x, y, width, height):
        c.setFont("Roboto-Black", 14)
        c.drawString(x, y + height - 1.5 * cm, f"{batch} OBJECTIVE EXAM RESULTS (Top 4): ")
    
        data = [["Date", f"{batch} Marks (Out of {720 if batch == 'NEET' else 360})"]]
    
        exam_data = {}
        for i in range(1, 100):
            date_key = f'Date{i}'
            marks_key = f'{batch} Marks{i}'
    
            if date_key not in student:
                break
    
            date = student[date_key]
            marks = self.format_marks(student.get(marks_key, ""))
    
            if date not in exam_data:
                exam_data[date] = {"Marks": "Absent"}
    
            if marks and marks != "Absent":
                exam_data[date]["Marks"] = marks
    
        # Select top 4 dates
        top_4_exam_data = self.select_top_4_dates(exam_data)
    
        for date, results in top_4_exam_data.items():
            data.append([
                date,
                results["Marks"]
            ])
    
        # Adjust these values to shift the table left and modify its width
        left_margin = 0.0 * cm  # Further reduced to shift more to the left
        table_width = width + 2 * cm  # Increased to use even more of the available width
    
        col_widths = [
            table_width * 0.3,  # Date
            table_width * 0.7,  # Marks
        ]
    
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.table_color),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Roboto-Black'),
            ('FONTNAME', (0, 1), (-1, -1), 'Roboto-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('ROUNDEDCORNERS', [10, 10, 10, 10]),
        ]))
        table.wrapOn(c, table_width, height)
        table.drawOn(c, x + left_margin, y)

    def draw_subject_marks_and_activities(self, c, student, x, y, width, height, heading_y):
        c.setFont("Roboto-Black", 14)
        c.drawString(x, heading_y, "SUBJECTIVE, CO-CURRICULAR, AND EXTRA-CURRICULAR RESULTS:")

        # Subjective Exam Results
        data = [["Date", "Subject", "Marks"]]
        exam_data = {}
        i = 1
        while True:
            date, subject, marks = f"Subjective Date{i}", f"Subject{i}", f"Subjective Marks{i}"
            if subject not in student:
                break
            if pd.notna(student[subject]):
                date_value = student[date] if pd.notna(student[date]) else f"Exam {i}"
                marks_value = student[marks] if pd.notna(student[marks]) else "-"
                if marks_value != "-" and marks_value != "Absent":
                    try:
                        mark, total = map(int, str(marks_value).split('/'))
                        score = (mark / total) * 100
                    except ValueError:
                        score = 0
                else:
                    score = 0
                
                if date_value not in exam_data:
                    exam_data[date_value] = []
                exam_data[date_value].append((student[subject], marks_value, score))
            i += 1

        # Sort exams by score and select top 4
        sorted_exams = sorted(exam_data.items(), key=lambda x: max(item[2] for item in x[1]) if x[1] else 0, reverse=True)
        top_4_exams = sorted_exams[:4]

        # Fill remaining slots if needed
        while len(top_4_exams) < 4:
            placeholder_date = f"Exam {len(top_4_exams) + 1}"
            top_4_exams.append((placeholder_date, [("N/A", "-", 0)]))

        # Populate the data for the table
        for date, subjects in top_4_exams:
            for subject, marks, _ in subjects:
                data.append([date, subject, marks])

        col_widths = [width * 0.15, width * 0.20, width * 0.15]
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.table_color),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Roboto-Black'),
            ('FONTNAME', (0, 1), (-1, -1), 'Roboto-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('ROUNDEDCORNERS', [10, 10, 10, 10]),
        ]))
        table_width = sum(col_widths)
        table.wrapOn(c, table_width, height)
        table.drawOn(c, x, y)

        # Calculate Curricular attendance
        curricular_attendance = self.calculate_curricular_attendance(student)

        # Co-curricular and Extra-curricular Activities
        co_curricular_attendance = self.get_attendance(student, [
            'Co-curricular Attendance', 'Co-curricular Attendance ',
            'Co-curricular'
        ])
        extra_curricular_attendance = self.get_attendance(student, [
            'Extra-curricular Attendance', 'Extra-curricular Attendance ',
            'Extra-curricular'
        ])

        # Calculate overall attendance
        overall_attendance = self.calculate_overall_attendance(curricular_attendance, co_curricular_attendance, extra_curricular_attendance)

        activities_data = [
            ["Activities", "Attendance"],
            ["Curricular", curricular_attendance],
            ["Co-curricular", co_curricular_attendance],
            ["Extra-curricular", extra_curricular_attendance],
            ["Overall", overall_attendance]
        ]

        activities_table = Table(activities_data, colWidths=[width * 0.2, width * 0.2])

        activities_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.table_color),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Roboto-Black'),
            ('FONTNAME', (0, 1), (-1, -1), 'Roboto-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('ROUNDEDCORNERS', [10, 10, 10, 10]),
            ('BACKGROUND', (0, 4), (1, 4), self.get_color(overall_attendance)),
        ]))

        activities_x = x + table_width + 1.7 * cm  # Add some spacing between tables
        activities_table.wrapOn(c, width * 0.35, height)
        activities_table.drawOn(c, activities_x, y)

        # Add a vertical line between Subjective and Activities tables
        line_x = x + table_width + 0.75 * cm  # Position the line between tables
        line_start_y = y + height - 0.5 * cm  # Start from the top of the section
        line_end_y = y - 0.3 * cm  # End slightly below the tables
        c.setStrokeColor(colors.black)  # Use a light grey color for subtlety
        c.setLineWidth(0.5)  # Set a thin line width
        c.line(line_x, line_start_y, line_x, line_end_y)

        # Performance level note with pale green background and curved edges
        note_width = width * 0.4 
        note_height = 2.5 * cm
        note_x = x + width - note_width + 0 * cm
        note_y = y - 3 * cm

        c.setFillColor(colors.Color(0.1, 0.3, 0.1, alpha=0.1))  # Pale green color
        c.roundRect(note_x, note_y, note_width, note_height, 5, stroke=0, fill=1)

        c.setFont("Roboto-BlackItalic", 8)
        c.setFillColor(colors.black)
        text_object = c.beginText(note_x + 0.4 * cm, note_y + note_height - 0.2 * cm)
        note_text = """
       Performance Levels:
       - Green: Best Performance - Exemplary.
       - Orange: Good - Can Improve.
       - Red: Needs Focus - Requires Improvement.
        """
        for line in note_text.split('\n'):
            text_object.textLine(line.strip())
        c.drawText(text_object)

    def calculate_curricular_attendance(self, student):
        attended = 0
        total = 0
        
        # Count objective exams
        for i in range(1, 100):
            for subject in ['NEET', 'JEE']:
                mark_key = f'{subject} Marks{i}'
                if mark_key in student:
                    total += 1
                    if pd.notna(student[mark_key]) and student[mark_key] not in ["Absent", "-"]:
                        attended += 1

        # Count subjective exams
        for i in range(1, 100):
            mark_key = f'Subjective Marks{i}'
            if mark_key in student:
                total += 1
                if pd.notna(student[mark_key]) and student[mark_key] not in ["Absent", "-"]:
                    attended += 1

        return f"{attended}/{total}" if total > 0 else "-"

    def get_attendance(self, student, possible_keys):
        for key in possible_keys:
            if key in student:
                value = student[key]
                if pd.notna(value) and value != "-":
                    # If the value is already in the format "x/y", return it as is
                    if isinstance(value, str) and '/' in value:
                        return value
                    # If it's a single number, assume it's the number of classes attended out of 10
                    elif isinstance(value, (int, float)) or (isinstance(value, str) and value.replace('.', '').isdigit()):
                        attended = int(float(value))
                        return f"{attended}/10"
        return '-'

    def calculate_overall_attendance(self, curricular, co_curricular, extra_curricular):
        def parse_attendance(attendance):
            if isinstance(attendance, str) and '/' in attendance:
                attended, total = map(int, attendance.split('/'))
                return attended, total
            return 0, 0

        curr_attended, curr_total = parse_attendance(curricular)
        co_curr_attended, co_curr_total = parse_attendance(co_curricular)
        extra_curr_attended, extra_curr_total = parse_attendance(extra_curricular)

        total_attended = curr_attended + co_curr_attended + extra_curr_attended
        total_classes = curr_total + co_curr_total + extra_curr_total

        if total_classes == 0:
            return '-'
        
        return f"{total_attended}/{total_classes}"

    def draw_total_marks(self, c, student, batch, x, y, width, height):
        c.setFont("Roboto-Black", 14)
        c.drawString(x, y + height - 2 * cm, "OVERALL EVALUATION: ")

        exam_marks = []
        exam_count = 0

        i = 1
        while True:
            exam_key = f"{batch} Marks{i}"

            if exam_key not in student:
                break

            if pd.notna(student.get(exam_key)) and student[exam_key] != "Absent":
                exam_marks.append(float(student[exam_key]))
                exam_count += 1

            i += 1

        exam_total = exam_count * (720 if batch == 'NEET' else 360)
        exam_sum = sum(exam_marks)

        def get_color_for_subject(score, total_score, subject, exam_count):
            Traffic_red = colors.Color(0.85, 0.1, 0.1)   # Lighter red
            Traffic_amber = colors.Color(0.85, 0.55, 0)  # Lighter amber
            Traffic_green = colors.Color(0.1, 0.75, 0.2) # Lighter green

            if exam_count == 0:
                return Traffic_red

            if subject == 'NEET':
                green_threshold = total_score * (451 / 720)
                orange_threshold = total_score * (401 / 720)
            elif subject == 'JEE':
                green_threshold = total_score * (201 / 360)
                orange_threshold = total_score * (181 / 360)
            elif subject == 'Subjective':
                green_threshold = 76  # 76% and above
                orange_threshold = 51  # Between 50% and 75%
            else:
                return colors.white  # Default color if subject is not recognized

            if score >= green_threshold:
                return Traffic_green
            elif score >= orange_threshold:
                return Traffic_amber
            else:
                return Traffic_red

        # Subjective Evaluation
        subjective_marks = []
        subjective_total = []
        subjective_count = 0
        i = 1
        while True:
            marks_key = f"Subjective Marks{i}"
            if marks_key not in student:
                break
            marks = student[marks_key]
            if pd.notna(marks) and marks != "Absent":
                mark, total = map(int, str(marks).split('/'))
                subjective_marks.append(mark)
                subjective_total.append(total)
                subjective_count += 1
            i += 1
        subjective_sum = sum(subjective_marks)
        subjective_total_sum = sum(subjective_total)
        subjective_percent = (subjective_sum / subjective_total_sum) * 100 if subjective_total_sum else 0

        objective_sum = exam_sum
        objective_total = exam_total
        objective_percent = (objective_sum / objective_total) * 100 if objective_total else 0

        subjective_percent = (subjective_sum / subjective_total_sum) * 100 if subjective_total_sum else 0

        co_curricular_percent = self.calculate_activity_percentage(student, 'Co-curricular')
        extra_curricular_percent = self.calculate_activity_percentage(student, 'Extra-curricular')

        # Set base weightage
        curriculum_weight = 0.60
        activity_weight = 0.40

        # Determine if objective and subjective components are present
        objective_present = objective_total > 0
        subjective_present = subjective_total_sum > 0

        # Distribute curriculum weightage
        if objective_present and subjective_present:
            objective_weight = subjective_weight = curriculum_weight / 2  # 30% each
        elif objective_present:
            objective_weight = curriculum_weight
            subjective_weight = 0
        elif subjective_present:
            objective_weight = 0
            subjective_weight = curriculum_weight
        else:
            objective_weight = subjective_weight = 0

        # Determine if co-curricular and extra-curricular activities are valid (not hyphenated)
        co_curricular_valid = co_curricular_percent != 0  # 0 indicates a hyphen was present
        extra_curricular_valid = extra_curricular_percent != 0  # 0 indicates a hyphen was present
        
        # Distribute activity weightage
        if co_curricular_valid and extra_curricular_valid:
            co_curricular_weight = extra_curricular_weight = activity_weight / 2
        elif co_curricular_valid:
            co_curricular_weight = activity_weight
            extra_curricular_weight = 0
        elif extra_curricular_valid:
            co_curricular_weight = 0
            extra_curricular_weight = activity_weight
        else:
            # If both activities are invalid, redistribute the weight to curriculum
            if objective_present and subjective_present:
                objective_weight += activity_weight / 2
                subjective_weight += activity_weight / 2
            elif objective_present:
                objective_weight += activity_weight
            elif subjective_present:
                subjective_weight += activity_weight
            co_curricular_weight = extra_curricular_weight = 0

        # Final adjusted weightage
        final_weightage = {
            'Objective': objective_weight,
            'Subjective': subjective_weight,
            'Co-curricular': co_curricular_weight,
            'Extra-curricular': extra_curricular_weight
        }

        # Calculate overall percentage
        overall_percent = (
            objective_percent * final_weightage['Objective'] +
            subjective_percent * final_weightage['Subjective'] +
            co_curricular_percent * final_weightage['Co-curricular'] +
            extra_curricular_percent * final_weightage['Extra-curricular']
        )
        normal_style = ParagraphStyle('Normal', fontName='Roboto-Bold', fontSize=10, leading=12, alignment=TA_LEFT)

        data = []
        if exam_count > 0:
            data.append([Paragraph(f"<font name='Roboto-Black'>{batch} Total:</font> {exam_sum:.2f} out of <font name='Roboto-bold'>{exam_total}</font> (<font name='Roboto-bold'>{exam_count}</font> exams)", normal_style)])
        if subjective_count > 0:
            data.append([Paragraph(f"<font name='Roboto-Black'>Subjective Total:</font> {subjective_sum:.2f} out of <font name='Roboto-bold'>{subjective_total_sum}</font> (<font name='Roboto-bold'>{subjective_count}</font> exams)", normal_style)])
        data.append([Paragraph(f"<font name='Roboto-Black'>Overall Percentage:</font> {overall_percent:.2f}% (out of 100%)", normal_style)])

        table = Table(data, colWidths=[width - 7 * cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), self.table_color),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONT', (0, 0), (-1, -1), 'Roboto-Bold', 10),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROUNDEDCORNERS', [10, 10, 10, 10]),
        ]))

        for i, row in enumerate(data):
            if f'{batch} Total:' in row[0].text:
                table.setStyle(TableStyle([('BACKGROUND', (0, i), (0, i), get_color_for_subject(exam_sum, exam_total, batch, exam_count))]))
            elif 'Subjective Total:' in row[0].text:
                table.setStyle(TableStyle([('BACKGROUND', (0, i), (0, i), get_color_for_subject(subjective_percent, 100, 'Subjective', subjective_count))]))
            elif 'Overall Percentage:' in row[0].text:
                table.setStyle(TableStyle([
                    ('LINEBEFORE', (0, i), (0, i), 2, colors.black),
                    ('LINEAFTER', (0, i), (0, i), 2, colors.black),
                    ('LINEBELOW', (0, i), (0, i), 2, colors.black),
                ]))

        table.wrapOn(c, width, height)
        table.drawOn(c, x, y + height - 5 * cm)
        
        # Add note about overall percentage calculation
        note_style = ParagraphStyle('Note', fontName='Roboto-Italic', fontSize=8, leading=10, alignment=TA_LEFT)
        note_text = f"""
        <font name='Roboto-BlackItalic'>Note on Overall Percentage Calculation:</font>
        The overall percentage is weighted as follows:
        • {final_weightage['Objective'] * 100:.1f}% Objective ({batch})
        • {final_weightage['Subjective'] * 100:.1f}% Subjective
        • {final_weightage['Co-curricular'] * 100:.1f}% Co-curricular activities
        • {final_weightage['Extra-curricular'] * 100:.1f}% Extra-curricular activities
        """
        note = Paragraph(note_text, note_style)
        note_width = width - 7 * cm
        note_height = 2 * cm
        note.wrapOn(c, note_width, note_height)
        note.drawOn(c, x, y + height - 8.1 * cm)

        def center_text_over_line(text, line_start, line_end, y_position):
            text_width = c.stringWidth(text, "Roboto-Bold", 10)
            line_center = (line_start + line_end) / 2
            text_start = line_center - (text_width / 2)
            return text_start

        c.setFont("Roboto-Bold", 10)
        line_start = x + width - 6.5 * cm
        line_end = x + width - 0.5 * cm

        text_x = center_text_over_line("Parent's Signature", line_start, line_end, y + height - 9.2 * cm)
        c.drawString(text_x, y + height - 8.0 * cm, "Parent's Signature")
        c.line(line_start, y + height - 7.5 * cm, line_end, y + height - 7.5 * cm)

        text_x = center_text_over_line("Cluster's Signature", line_start, line_end, y + height - 7.2 * cm)
        c.drawString(text_x, y + height - 6.0 * cm, "Cluster's Signature")
        c.line(line_start, y + height - 5.5 * cm, line_end, y + height - 5.5 * cm)

    def draw_contact_info(self, c, student, x, y, width, height):
        c.setFont("Roboto-Black", 13)
        c.drawString(x, y + height - 1.8 * cm, "Contact Information: ")
        c.setFont("Roboto-Bold", 11)

        def format_phone(number):
            if pd.isna(number):
                return "N/A"
            try:
                return str(int(float(number)))
            except ValueError:
                return str(number)  # Return as is if it's already a string

        student_contact = format_phone(student['Student Contact No.'])
        father_contact = format_phone(student['Father/Guardian Contact No.'])
        mother_contact = format_phone(student['Mother/Guardian Contact No.'])

        c.drawString(x, y + height - 2.5 * cm, f"Student Contact: {student_contact}")
        c.drawString(x, y + height - 3 * cm, f"Father/Guardian Contact: {father_contact}")
        c.drawString(x, y + height - 3.5 * cm, f"Mother/Guardian Contact: {mother_contact}")

    def get_color(self, attendance):
        Traffic_red = colors.Color(0.85, 0.1, 0.1)   # Lighter red
        Traffic_amber = colors.Color(0.85, 0.55, 0)  # Lighter amber
        Traffic_green = colors.Color(0.1, 0.75, 0.2) # Lighter green
        
        if isinstance(attendance, str) and '/' in attendance:
            try:
                attended, total = map(int, attendance.split('/'))
                percentage = (attended / total) * 100
            except (ValueError, ZeroDivisionError):
                return colors.white  # Default color if there's an error parsing the string
        elif isinstance(attendance, (int, float)):
            percentage = attendance
        else:
            return colors.white  # Default color for unexpected input types
        
        if percentage < 51:
            return Traffic_red
        elif percentage < 76:
            return Traffic_amber
        else:
            return Traffic_green

    def format_marks(self, marks):
        if pd.isna(marks) or marks == "Absent":
            return "Absent"
        try:
            return f"{float(marks):.2f}"
        except ValueError:
            return marks

    def calculate_activity_percentage(self, student, activity_type):
        attendance = self.get_attendance(student, [f'{activity_type} Attendance', f'{activity_type} Attendance ', activity_type])
        if attendance == '-':
            return 0
        if isinstance(attendance, str) and '/' in attendance:
            attended, total = map(int, attendance.split('/'))
            return (attended / total) * 100 if total > 0 else 0
        return 0  # Return 0 for any other unexpected data type

    def select_top_4_dates(self, exam_data):
        date_scores = []
        absent_dates = []
        for date, results in exam_data.items():
            score = results['Marks']
            if score and score != "Absent":
                try:
                    if isinstance(score, str) and '/' in score:
                        numerator, denominator = score.split('/')
                        score = (float(numerator) / float(denominator)) * 100
                    else:
                        score = float(score)
                    date_scores.append((date, score))
                except ValueError:
                    absent_dates.append(date)
            else:
                absent_dates.append(date)

        # Sort dates by score in descending order
        sorted_dates = sorted(date_scores, key=lambda x: x[1], reverse=True)

        # Select top 4 dates
        top_4_dates = [date for date, _ in sorted_dates[:4]]

        # If we have fewer than 4 dates, add absent dates
        while len(top_4_dates) < 4 and absent_dates:
            top_4_dates.append(absent_dates.pop(0))

        # If we still have fewer than 4 dates, add placeholder dates
        while len(top_4_dates) < 4:
            placeholder_date = f"Exam {len(top_4_dates) + 1}"
            top_4_dates.append(placeholder_date)
            exam_data[placeholder_date] = {"Marks": "Absent"}

        return {date: exam_data[date] for date in top_4_dates}

def main():
    generator = ReportCardGenerator()

    print("Welcome to the ANEES DEFENCE CAREER INSTITUTE Report Card Generator")
    print("Please provide the path to your CSV file:")

    file_path = input().strip()

    # Remove the '&' if it's at the start of the path
    if file_path.startswith('&'):
        file_path = file_path[1:].strip()

    # Remove surrounding quotes if present
    if (file_path.startswith('"') and file_path.endswith('"')) or (file_path.startswith("'") and file_path.endswith("'")):
        file_path = file_path[1:-1]

    # Replace backslashes with forward slashes
    file_path = file_path.replace('\\', '/')

    if not os.path.exists(file_path):
        print(f"Error: The file {file_path} does not exist.")
        return

    if not file_path.lower().endswith('.csv'):
        print("Error: Please provide a CSV file.")
        return

    print(f"Generating report cards from {file_path}...")
    generator.generate_report_cards(file_path)

if __name__ == "__main__":
    main()