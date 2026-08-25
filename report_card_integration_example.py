"""
Report Card Generation - Integration Example
This file shows how to integrate the dynamic report_card_dynamic.html template
with your existing Flask application.

Usage:
1. Copy the generate_report_card function to your app.py
2. Implement get_student_data() based on your Excel structure
3. Update the route to match your URL structure
"""

from flask import Flask, render_template, request, send_file
from weasyprint import HTML
from datetime import datetime
import os
import pandas as pd


def generate_report_card_for_student(student_data, output_dir='generated_pdfs'):
    """
    Generate a PDF report card for a single student.
    
    Args:
        student_data (dict): Dictionary containing all student information
        output_dir (str): Directory to save the generated PDF
        
    Returns:
        str: Path to the generated PDF file
    """
    
    # Prepare data for template
    template_data = {
        'student_name': student_data['name'],
        'class_name': student_data['class'],
        'section': student_data.get('section', 'N/A'),
        'roll_number': student_data.get('roll_number', 'N/A'),
        'student_id': student_data.get('student_id', 'N/A'),
        'academic_year': student_data.get('academic_year', '2025-2026'),
        'school_name': student_data.get('school_name', 'YOUR SCHOOL NAME'),
        'school_address': student_data.get('school_address', 'School Address, City, State - PIN'),
        'dob': student_data.get('dob', 'N/A'),
        
        # Subjects with marks - this should come from your database
        'subjects': student_data.get('subjects', []),
        
        # Calculated values
        'total_marks': student_data.get('total_marks', 0),
        'max_marks': student_data.get('max_marks', 500),
        'percentage': student_data.get('percentage', 0.0),
        'grade': student_data.get('grade', 'N/A'),
        'grade_points': student_data.get('grade_points', 'N/A'),
        'result': student_data.get('result', 'PASS'),
        
        # Optional sections
        'attendance': student_data.get('attendance'),
        'remarks': student_data.get('remarks'),
        'co_scholastic': student_data.get('co_scholastic'),
        
        'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    # Render HTML template with Jinja2
    # Make sure report_card_dynamic.html is in your templates folder
    html_content = render_template('report_card_dynamic.html', **template_data)
    
    # Generate PDF using WeasyPrint
    pdf = HTML(string=html_content).write_pdf()
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Create filename
    safe_name = student_data['name'].replace(' ', '_').replace('/', '_')
    pdf_filename = f"{output_dir}/report_card_{safe_name}_{student_data.get('roll_number', 'N/A')}.pdf"
    
    # Save PDF
    with open(pdf_filename, 'wb') as f:
        f.write(pdf)
    
    return pdf_filename


def get_student_data_from_excel(excel_file, student_identifier, identifier_type='roll_number'):
    """
    Fetch student data from Excel file.
    
    Args:
        excel_file (str): Path to marks_database.xlsx
        student_identifier: Roll number or student ID to search for
        identifier_type (str): 'roll_number' or 'student_id'
        
    Returns:
        dict: Student data dictionary
    """
    # Read Excel file
    df = pd.read_excel(excel_file)
    
    # Find student row
    if identifier_type == 'roll_number':
        student_row = df[df['roll_number'] == student_identifier]
    else:
        student_row = df[df['student_id'] == student_identifier]
    
    if student_row.empty:
        raise ValueError(f"Student with {identifier_type}={student_identifier} not found")
    
    # Convert to dict (take first row if multiple matches)
    data = student_row.iloc[0].to_dict()
    
    # Transform into template-ready format
    student_data = {
        'name': data.get('student_name', data.get('name', 'Unknown')),
        'class': data.get('class', 'N/A'),
        'section': data.get('section', 'N/A'),
        'roll_number': data.get('roll_number', 'N/A'),
        'student_id': data.get('student_id', 'N/A'),
        'dob': data.get('dob', 'N/A'),
        
        # Build subjects list from columns
        'subjects': [],
        'total_marks': 0,
        'max_marks': 0,
    }
    
    # Example: If your Excel has columns like 'Maths_Marks', 'Science_Marks', etc.
    # Adjust column names based on your actual Excel structure
    subject_columns = {
        'Mathematics': 'maths_marks',
        'Science': 'science_marks',
        'English': 'english_marks',
        'Hindi': 'hindi_marks',
        'Social Studies': 'social_studies_marks',
        'Computer Science': 'cs_marks',
    }
    
    total = 0
    max_total = 0
    
    for subject_name, column_name in subject_columns.items():
        if column_name in data:
            marks = data[column_name]
            if pd.notna(marks):
                marks = int(marks)
                total += marks
                max_total += 100  # Assuming 100 marks per subject
                
                student_data['subjects'].append({
                    'name': subject_name,
                    'theory': marks,
                    'practical': 'N/A',
                    'total': marks
                })
    
    student_data['total_marks'] = total
    student_data['max_marks'] = max_total
    student_data['percentage'] = round((total / max_total * 100), 2) if max_total > 0 else 0
    
    # Calculate grade based on percentage
    student_data['grade'] = calculate_grade(student_data['percentage'])
    student_data['grade_points'] = calculate_grade_points(student_data['percentage'])
    
    return student_data


def calculate_grade(percentage):
    """Calculate grade based on CBSE pattern"""
    if percentage >= 90:
        return 'A1'
    elif percentage >= 80:
        return 'A2'
    elif percentage >= 70:
        return 'B1'
    elif percentage >= 60:
        return 'B2'
    elif percentage >= 50:
        return 'C1'
    elif percentage >= 40:
        return 'C2'
    elif percentage >= 33:
        return 'D'
    else:
        return 'E'


def calculate_grade_points(percentage):
    """Calculate grade points based on CBSE pattern"""
    if percentage >= 95:
        return 10.0
    elif percentage >= 90:
        return 9.0
    elif percentage >= 80:
        return 8.0
    elif percentage >= 70:
        return 7.0
    elif percentage >= 60:
        return 6.0
    elif percentage >= 50:
        return 5.0
    elif percentage >= 40:
        return 4.0
    elif percentage >= 33:
        return 3.0
    else:
        return 0.0


# Example Flask route to add to your app.py
"""
@app.route('/generate_report/<class_name>/<roll_number>')
def generate_report(class_name, roll_number):
    try:
        # Get student data from Excel
        student_data = get_student_data_from_excel(
            'marks_database.xlsx',
            roll_number,
            'roll_number'
        )
        
        # Generate PDF
        pdf_path = generate_report_card_for_student(student_data)
        
        # Send file to user
        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=f"report_card_{student_data['name']}.pdf"
        )
        
    except Exception as e:
        return f"Error generating report card: {str(e)}", 500


@app.route('/generate_all_reports/<class_name>')
def generate_all_reports(class_name):
    """Generate report cards for all students in a class"""
    try:
        # Read all students from Excel
        df = pd.read_excel('marks_database.xlsx')
        class_students = df[df['class'] == class_name]
        
        generated_files = []
        
        for index, row in class_students.iterrows():
            student_data = get_student_data_from_excel(
                'marks_database.xlsx',
                row['roll_number'],
                'roll_number'
            )
            
            pdf_path = generate_report_card_for_student(student_data)
            generated_files.append(pdf_path)
        
        return f"Generated {len(generated_files)} report cards:<br>" + "<br>".join(generated_files)
        
    except Exception as e:
        return f"Error: {str(e)}", 500
"""


if __name__ == '__main__':
    # Test the function
    print("Report Card Generation Module")
    print("=" * 50)
    print("This file provides helper functions for report card generation.")
    print("Import these functions into your app.py file.")
    print("")
    print("Key functions:")
    print("1. generate_report_card_for_student(student_data)")
    print("2. get_student_data_from_excel(excel_file, student_id)")
    print("3. calculate_grade(percentage)")
    print("4. calculate_grade_points(percentage)")
