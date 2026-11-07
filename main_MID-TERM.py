from report_utils import read_students, prepare_subjective, subject_maximum, is_absent, safe_filename
import pandas as pd
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfReader, PdfWriter
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
import io
import matplotlib
matplotlib.use('Agg')

class ReportCardGenerator:
    def __init__(self):
        self.cwd = os.path.dirname(os.path.abspath(__file__))
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
        self.logo_path = os.path.join(self.cwd, "logo.png")
        self.table_color = colors.Color(0.2, 0.4, 0.8, alpha=0.1)  # Light blue with low opacity
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
        df = read_students(file_path)

        # Convert student info columns to string
        df['Name'] = df['Name'].astype(str)
        df['Roll No'] = df['Roll No'].astype(str)
        df['Class'] = df['Class'].astype(str)
        df['Batch'] = df['Batch'].astype(str)
        df['District'] = df['District'].astype(str)

        # Convert all mark columns to string to avoid dtype warnings
        for col in df.columns:
            if 'Marks' in col:
                df[col] = df[col].astype(str)
                # Replace 'nan' strings with empty strings
                df[col] = df[col].replace('nan', '')

        return df

    def generate_report_cards(self, file_path):
        self.data = self.load_data(file_path)

        # Collect maximum values for marks that don't have them
        self.collect_max_values()

        # Add serial numbers to students and generate individual report cards
        for i, (_, student) in enumerate(self.data.iterrows(), 1):
            # Create a copy of the student data with the serial number added
            student_data = student.copy()
            student_data['Serial No'] = i
            self.generate_report_card(student_data)

        print(f"Report cards generated and saved in {self.cwd}/Report_Cards")

    def collect_max_values(self):
        self.subject_maxima = prepare_subjective(self.data)

    def generate_report_card(self, student):
        # Create a unique filename using both Serial No and Roll No to ensure uniqueness
        file_name = f"{student['Serial No']}_{safe_filename(student['Roll No'])}_{safe_filename(student['Name'])}_report_card.pdf"
        file_path = os.path.join(self.cwd, "Report_Cards", file_name)

        # Check if file already exists and remove it to avoid duplicates
        if os.path.exists(file_path):
            os.remove(file_path)

        # Generate the report card
        c = canvas.Canvas(file_path, pagesize=A4)
        self.draw_report_card(c, student)
        c.save()

    def draw_report_card(self, c, student):
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
        c.drawString(2 * cm, height - 5.5 * cm, "STUDENT'S PRELIUM EXAM REPORT CARD")

        # Student Info
        self.draw_info_box(c, student, 2 * cm, height - 9 * cm, width - 4 * cm, 3 * cm)

        # Exam Results (now showing all subjective exams and total marks)
        self.draw_exam_results(c, student, 2 * cm, height - 15 * cm, width - 6 * cm, 6 * cm)

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
        c.drawString(x + 0.5 * cm, y + height - 1.5 * cm, f"Serial No: {student['Serial No']}")
        c.drawString(x + width / 4, y + height - 1.5 * cm, f"Class: {student['Class'].upper()}")
        c.drawString(x + 0.5 * cm, y + height - 2.25 * cm, f"Batch: {student['Batch']}")
        c.drawString(x + width / 4, y + height - 2.25 * cm, f"District: {student['District'].upper()}")

    def draw_exam_results(self, c, student, x, y, width, height):
        c.setFont("Roboto-Black", 14)

        data = [["DATE", "SUBJECT", "MARKS"]]

        subjective_marks = []
        subjective_total = []

        for i in range(1, 100):  # Assuming a maximum of 99 subjective exams
            date_key = f'Subjective Date{i}'
            subject_key = f'Subject{i}'
            marks_key = f'Subjective Marks{i}'

            if date_key not in student or subject_key not in student or marks_key not in student:
                break

            date = student[date_key]
            subject = student[subject_key]
            marks = self.format_marks(student[marks_key])

            # Skip "Absent" marks for calculations but still display them
            if marks == "Absent" or str(marks).lower() == "absent":
                # Keep the mark as "Absent" for display
                marks = "Absent"
                subjective_marks.append(0.0)
                subjective_total.append(subject_maximum(self.subject_maxima, student, i))
            elif '/' in str(marks):
                try:
                    parts = str(marks).split('/')
                    if len(parts) == 2:
                        mark_str, total_str = parts
                        # Check for empty or non-numeric values
                        if mark_str.strip() and total_str.strip():
                            try:
                                mark = float(mark_str)
                                total = float(total_str)
                                if not pd.isna(mark) and not pd.isna(total) and total > 0:
                                    subjective_marks.append(mark)
                                    subjective_total.append(total)
                                    # Format for display
                                    marks = f"{mark:.2f}/{total:.0f}"
                                else:
                                    print(f"Warning: Invalid mark or total in '{marks}'. Skipping.")
                            except (ValueError, TypeError):
                                print(f"Warning: Could not parse mark value '{marks}'. Skipping.")
                        else:
                            print(f"Warning: Empty mark or total in '{marks}'. Skipping.")
                    else:
                        print(f"Warning: Invalid format for mark '{marks}'. Skipping.")
                except Exception as e:
                    print(f"Warning: Error processing mark '{marks}': {e}. Skipping.")
            else:
                # At this point, marks without max values should have been processed
                # in collect_max_values method, but just in case, handle it here too
                try:
                    # Check if it's a valid number
                    if str(marks).strip() and not str(marks).lower() == 'nan':
                        mark = float(marks)
                        if not pd.isna(mark):
                            print(f"Mark '{marks}' for {subject} doesn't include a maximum value.")
                            total_input = input(f"Please enter the maximum value for {subject} mark {mark} (default 100): ")
                            total = 100.0  # Default value
                            if total_input.strip():
                                try:
                                    total = float(total_input)
                                except ValueError:
                                    print(f"Invalid input '{total_input}', using default value 100.")

                            if total > 0:
                                subjective_marks.append(mark)
                                subjective_total.append(total)
                                # Update the marks format for display
                                marks = f"{mark:.2f}/{total:.0f}"
                            else:
                                print(f"Warning: Total value must be greater than 0. Skipping.")
                        else:
                            marks = "Absent"
                    else:
                        marks = "Absent"
                except ValueError:
                    print(f"Warning: Could not parse mark value '{marks}'. Marking as Absent.")
                    marks = "Absent"

            data.append([date, subject, marks])

        table_width = width + 2 * cm
        col_widths = [table_width * 0.25, table_width * 0.50, table_width * 0.25]

        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.table_color),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Roboto-Bold'),  # Changed to Roboto-Bold for all cells
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('FONTSIZE', (0, 0), (-1, 0), 11),  # Increased font size for the header (row 0)
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROUNDEDCORNERS', [5, 5, 5, 5]),
        ]))
        table.wrapOn(c, table_width, height)
        table.drawOn(c, x, y)

        # Calculate total marks and percentage
        subjective_sum = sum(subjective_marks)
        subjective_total_sum = sum(subjective_total)
        subjective_percent = (subjective_sum / subjective_total_sum) * 100 if subjective_total_sum else 0

        # Draw a box for total marks and percentage
        box_width = table_width
        box_height = 1.5 * cm
        box_x = x
        box_y = y - box_height - 0.2 * cm

        # Draw subtle gradient-like background
        c.saveState()
        p = c.beginPath()
        p.roundRect(box_x, box_y, box_width, box_height, 5)
        c.clipPath(p, stroke=1, fill=0)

        # Create a gradient-like effect
        num_steps = 20
        for i in range(num_steps):
            y_pos = box_y + (i / num_steps) * box_height
            color = colors.Color(
                self.table_color.red,
                self.table_color.green,
                self.table_color.blue,
                alpha=(i / num_steps) * 0.2
            )
            c.setFillColor(color)
            c.rect(box_x, y_pos, box_width, box_height / num_steps, fill=1, stroke=0)

        c.restoreState()

        c.setFillColor(colors.black)
        c.setFont("Roboto-Bold", 10)

        # Estimate text height (adjust this value if needed)
        estimated_text_height = 10 * 1.2  # Font size * 1.2 for some padding

        # Calculate vertical center and add an upward offset
        vertical_offset = 0.1 * cm  # Adjust this value to shift the text up or down
        vertical_center = box_y + (box_height - estimated_text_height) / 2 + vertical_offset

        # Draw total marks
        total_marks_text = f"Total Marks: {subjective_sum} out of {subjective_total_sum}"
        c.drawString(box_x + 0.3 * cm, vertical_center, total_marks_text)

        # Draw percentage
        percentage_text = f"Overall Percentage: {subjective_percent:.2f}%"
        c.drawString(box_x + box_width / 2 + 0.3 * cm, vertical_center, percentage_text)

        # Add subtle divider
        c.setStrokeColorRGB(0.7, 0.7, 0.7)  # Light gray
        c.setLineWidth(0.5)
        c.line(box_x + box_width / 2, box_y + 0.2 * cm, box_x + box_width / 2, box_y + box_height - 0.2 * cm)

        # Prepare data for the pie chart
        subjects = []
        marks = []
        percentages = []

        for i in range(1, 100):  # Assuming a maximum of 99 subjective exams
            subject_key = f'Subject{i}'
            marks_key = f'Subjective Marks{i}'

            if subject_key not in student or marks_key not in student:
                break

            subject = student[subject_key]
            mark_str = str(student[marks_key])

            # Skip empty, missing, or "Absent" marks
            if pd.isna(student[marks_key]) or mark_str.strip() == "" or mark_str == "—" or mark_str.lower() == "absent":
                continue

            # Process the mark
            try:
                if '/' in mark_str:
                    mark_value, total = map(float, mark_str.split('/'))
                    # Only include valid marks (non-zero total)
                    if total > 0:
                        subjects.append(subject)
                        marks.append(mark_value)
                        percentages.append((mark_value / total) * 100)
                else:
                    # This should not happen at this point, but just in case
                    print(f"Warning: Mark '{mark_str}' for {subject} in pie chart doesn't include a maximum value.")
                    continue
            except (ValueError, TypeError) as e:
                print(f"Warning: Could not parse mark value '{mark_str}' for pie chart. Error: {e}. Skipping.")
                continue

        # Create the pie chart only if we have valid data
        valid_data = False

        if subjects and percentages and len(subjects) == len(percentages):
            # Filter out any NaN values or zeros in percentages
            valid_indices = []
            for i, p in enumerate(percentages):
                if not pd.isna(p) and p > 0:
                    valid_indices.append(i)

            if valid_indices:
                valid_subjects = [subjects[i] for i in valid_indices]
                valid_percentages = [percentages[i] for i in valid_indices]

                # Double-check that we have valid data
                if valid_subjects and valid_percentages and len(valid_subjects) == len(valid_percentages):
                    valid_data = True

                    plt.figure(figsize=(6, 7))  # Decreased overall size
                    sns.set_style("whitegrid")
                    sns.set_palette("deep")

                    try:
                        plt.pie(valid_percentages, labels=valid_subjects, autopct='%1.1f%%', startangle=90,
                                wedgeprops=dict(width=0.5, edgecolor='white'))
                        plt.title("Subject-wise Performance", fontsize=16, fontweight='bold', pad=20)

                        # Maintain font size for labels and percentages
                        plt.rcParams['font.size'] = 10
                        plt.rcParams['font.weight'] = 'bold'

                        # Add a legend with maintained font size
                        plt.legend(valid_subjects, title="Subjects", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1),
                                   fontsize=10, title_fontsize=12)

                        plt.tight_layout()
                    except Exception as e:
                        print(f"Error creating pie chart: {e}")
                        valid_data = False

        if not valid_data:
            # Create an empty figure with a message if no valid data
            plt.figure(figsize=(6, 7))
            plt.text(0.5, 0.5, "No subject data available for pie chart",
                     horizontalalignment='center', verticalalignment='center',
                     fontsize=14, fontweight='bold')
            plt.axis('off')  # Hide axes

        # Save the plot to a bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=300, bbox_inches='tight')
        buf.seek(0)
        plt.close()

        # Draw the graph on the PDF
        graph_width = width * 0.9  # Slightly decreased width
        graph_height = 10 * cm  # Decreased height
        graph_x = x + (width - graph_width) / 2  # Center the graph horizontally
        graph_y = box_y - graph_height - 0.5 * cm

        c.drawImage(ImageReader(buf), graph_x, graph_y, width=graph_width, height=graph_height)

        # Add a note explaining the chart with maintained font size
        note_y = graph_y - 0.5 * cm
        c.setFont("Roboto-Italic", 9)
        c.drawString(graph_x, note_y, "Note: The pie chart represents the proportion of marks obtained in each subject relative to their respective totals.")

    def format_marks(self, marks):
        # Handle missing data as "Absent"
        if pd.isna(marks) or marks == "—" or str(marks).strip() == "" or str(marks).lower() == 'nan':
            return "Absent"

        # Handle "Absent" string directly
        if str(marks).lower() == "absent":
            return "Absent"

        try:
            # If marks already contains a slash (formatted as value/total)
            if '/' in str(marks):
                mark_value, total = map(float, str(marks).split('/'))
                return f"{mark_value:.2f}/{total:.0f}"
            # Otherwise, just format the number
            return f"{float(marks):.2f}"
        except ValueError:
            return marks

    def draw_contact_info(self, c, student, x, y, width, height):
        # Note: width parameter is kept for consistency with other methods but not used directly
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

        student_contact = format_phone(student['SELF'])
        father_contact = format_phone(student['FATHER'])
        mother_contact = format_phone(student['MOTHER'])

        c.drawString(x, y + height - 2.5 * cm, f"Student's Contact: {student_contact}")
        c.drawString(x, y + height - 3 * cm, f"Father's Contact: {father_contact}")
        c.drawString(x, y + height - 3.5 * cm, f"Mother's Contact: {mother_contact}")

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

    def calculate_activity_percentage(self, student, activity_type):
        attendance = self.get_attendance(student, [f'{activity_type} Attendance', f'{activity_type} Attendance ', activity_type])
        if attendance == '-':
            return 0
        if isinstance(attendance, str) and '/' in attendance:
            attended, total = map(int, attendance.split('/'))
            return (attended / total) * 100 if total > 0 else 0
        return 0  # Return 0 for any other unexpected data type

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
