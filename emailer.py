"""
Converts a CSV file to emails to send to professors.
Will install pypiwin32 and pypandoc if not already installed.
"""

header_length = 11  # Number of items in the CSV header
table_cols = ('CRN', 'Subject and Course Number', 'Student ID', 'Student Last Name', 'Student First Name')

import csv
import os
import pathlib

try:
    import win32com.client as win32
except ModuleNotFoundError:
    os.system('python -m pip install pypiwin32')

try:
    import pypandoc
except ModuleNotFoundError:
    os.system('python -m pip install pypandoc')

# Test values
pwd = pathlib.Path(__file__).parent
subject = 'Enter email subject line'
email_file = pwd / 'test_template.txt'
css_file = pwd / 'style.css'
csv_file = pwd / 'test_csv.csv'
attachment_file = pwd / 'test_attachment.pdf'

def main(subject, email_file, css_file, csv_file, attachment_file, test):
    # Read CSV
    faculty2students = read_csv(csv_file)

    # Make emails
    style = read_style(css_file)
    student_count = 0
    print_header()
    for email_count, (faculty_email, students) in enumerate(faculty2students.items()):
        body = format_email_body(email_file, style, students, table_cols)
        if not test:
            send_email(faculty_email, subject, body, attachment_file)
        student_count += len(students)
        print(f'{len(students):^13}    {faculty_email}')
    print_footer(email_count, student_count)

def read_csv(csv_file):
    """Pass student info CSV to transformation function and return result."""
    with open(csv_file) as file:
        csv_obj = csv.reader(file)
        col_names = next(csv_obj)  # Get
        faculty2students = make_faculty2students(col_names, csv_obj)
        return faculty2students

def make_faculty2students(keys, rows_of_vals):
    """Transform student info csv into dictionary where keys are faculty emails."""
    faculty2students = {}
    for row_of_vals in rows_of_vals:
        student = dict(zip(keys, row_of_vals))
        if (faculty_email := student['Faculty Email']) not in faculty2students:
            faculty2students[faculty_email] = []
        faculty2students[faculty_email].append(student)
    # Sort the student list by CRN
    for students in faculty2students.values():
        students.sort(key = lambda student: student['CRN'])
    return faculty2students

def read_style(css_file):
    with open(css_file) as file:
        return ''.join(file.readlines())

def format_email_body(email_file, style, students, table_cols):
    """Apply format dictionary to email template."""
    with open(email_file) as file:
        template = ''.join(file.readlines())
    # Create a string format dictionary from the first student.
    format_dict = dict(students[0])
    # Create the table
    format_dict['table'] = make_table(students, table_cols)
    email = template.format(**format_dict)
    email_html = pypandoc.convert_text(email, 'html', format='md')
    return style + email_html

def make_table(students, table_cols):
    """Create HTML table."""
    header = make_row(table_cols)
    hline = make_row(['-']*len(table_cols))
    row_cells = [[student[key] for key in table_cols] for student in students]
    rows = [make_row(cell) for cell in row_cells]
    table = '\n'.join([header] + [hline] + rows)
    return table

def make_row(cells):
    return '|'.join(cell for cell in cells)

def send_email(to, subject, body, attachment):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Attachments.Add(attachment)
    mail.Send()

def print_header():
    print(header := f'student count    recipient')
    print('-'*(20 + len(header)))

def print_footer(send_count, student_count):
    print()
    print(f'{send_count+1:>3} emails sent')
    print(f'{student_count:>3} students')

#import inspector
#inspector.apply_decorator(globals())

if __name__ == '__main__':
    main(subject, email_file, css_file, csv_file, attachment_file, test=True)
