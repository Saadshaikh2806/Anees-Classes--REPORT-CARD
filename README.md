# ANEES DEFENCE CAREER INSTITUTE Report Card Generator

## Overview

This project is a Python-based Report Card Generator designed specifically for the ANEES DEFENCE CAREER INSTITUTE. It automates the process of creating comprehensive student report cards from CSV data, including academic performance, attendance, and extra-curricular activities.

## Features

- Generates individual PDF report cards for each student
- Includes both objective and subjective exam results
- Calculates and displays attendance for curricular, co-curricular, and extra-curricular activities
- Provides an overall evaluation based on weighted percentages of different activities
- Uses color-coding to visually represent performance levels
- Supports multiple exams and subjects
- Includes contact information for students and guardians
- Attaches a custom second page to each report card

## Requirements

- Python 3.7+
- pandas
- reportlab
- PyPDF2

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/your-username/report-card-generator.git
   cd report-card-generator
   ```

2. Install the required packages:
   ```
   pip install pandas reportlab PyPDF2
   ```

3. Ensure you have the necessary font files in a `fonts` directory:
   - Roboto-Regular.ttf
   - Roboto-Bold.ttf
   - Roboto-Italic.ttf
   - Roboto-Black.ttf
   - Roboto-BlackItalic.ttf

## Usage

1. Prepare your CSV file with student data. The CSV should include columns for:
   - Name
   - Roll No
   - Class
   - Batch
   - District
   - Exam details (Date, Subject, Marks)
   - Attendance information
   - Contact numbers

2. Run the script:
   ```
   python main.py
   ```

3. When prompted, enter the path to your CSV file.

4. The script will generate individual PDF report cards for each student in the `Report_Cards` directory.

## Customization

- You can modify the `institute_details` in the `__init__` method of the `ReportCardGenerator` class to update the institute's information.
- Adjust the `logo_path` and `second_page_pdf` variables to use your own logo and second page template.
- Modify the color scheme by changing the `table_color` and other color definitions in the script.

## Contributing

Contributions to improve the Report Card Generator are welcome. Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For any queries or support, please contact the ANEES DEFENCE CAREER INSTITUTE IT department.
