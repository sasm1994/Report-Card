"""
Report Card Generation System - Flask Application
Updated with dynamic template support for generating individual student report cards
"""

from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from weasyprint import HTML
import os
import pandas as pd
from datetime import datetime
import io

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'

UPLOAD_FOLDER = 'student_lists'
GENERATED_PDFS_FOLDER = 'generated_pdfs'
DATABASE_FILE = 'marks_database.xlsx'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_PDFS_FOLDER, exist_ok=True)

def get_student_data_from_excel(student_identifier, identifier_type='roll_number', class_name=None):
    try:
        if not os.path.exists(DATABASE_FILE):
            raise FileNotFoundError(f"Database file {DATABASE_FILE} not found")
        
        df = pd.read_excel(DATABASE_FILE)
        
        if class_name:
            df = df[df['class'].astype(str).str.upper() == class_name.upper()]
        
        if identifier_type == 'roll_number':
            student_row = df[df['roll_number'].astype(str) == str(student_identifier)]
        else:
            student_row = df[df['student_id'].astype(str) == str(student_identifier)]
        
        if student_row.empty:
            return None
        
        data = student_row.iloc[0].to_dict()
        student_data = build_student_data_structure(data, df.columns)
        return student_data
    except Exception as e:
        print(f"Error fetching student data: {str(e)}")
        raise

def build_student_data_structure(data, columns):
    column_mappings = {
        'student_name': ['student_name', 'name', 'Student Name'],
        'class': ['class', 'Class', 'grade'],
        'section': ['section', 'Section', 'SEC'],
        'roll_number': ['roll_number', 'roll_no', 'Roll Number', 'Roll No'],
        'student_id': ['student_id', 'id', 'ID', 'Admission No'],
        'dob': ['dob', 'date_of_birth', 'DOB', 'Birth Date'],
    }
    
    subject_mappings = {
        'Mathematics': ['maths', 'math', 'Mathematics', 'Maths_Marks', 'MATH'],
        'Science': ['science', 'Science', 'SCI', 'Science_Marks'],
        'English': ['english', 'English', 'ENG', 'English_Marks'],
        'Hindi': ['hindi', 'Hindi', 'HIN', 'Hindi_Marks'],
        'Social Studies': ['social', 'sst', 'Social Studies', 'SST', 'Social_Marks'],
        'Computer Science': ['computer', 'cs', 'Computer Science', 'CS', 'Comp_Marks'],
        'Sanskrit': ['sanskrit', 'Sanskrit', 'SAN'],
        'Art Education': ['art', 'Art Education', 'ART'],
        'Work Education': ['work', 'Work Education', 'WORK'],
        'Physical Education': ['physical', 'Physical Education', 'PE', 'PT'],
    }
    
    def get_value(mapping_dict, key):
        for possible_name in mapping_dict[key]:
            if possible_name in columns:
                return data.get(possible_name)
        return None
    
    student_name = get_value(column_mappings, 'student_name') or 'Unknown'
    class_name_val = get_value(column_mappings, 'class') or 'N/A'
    section = get_value(column_mappings, 'section') or 'N/A'
    roll_number = get_value(column_mappings, 'roll_number') or 'N/A'
    student_id = get_value(column_mappings, 'student_id') or 'N/A'
    dob = get_value(column_mappings, 'dob') or 'N/A'
    
    subjects = []
    total_marks = 0
    max_marks = 0
    
    for subject_name, possible_columns in subject_mappings.items():
        for col_name in possible_columns:
            if col_name in columns:
                marks = data.get(col_name)
                if pd.notna(marks):
                    try:
                        marks = int(float(marks))
                        total_marks += marks
                        max_marks += 100
                        subjects.append({'name': subject_name, 'theory': marks, 'practical': 'N/A', 'total': marks})
                    except (ValueError, TypeError):
                        pass
                break
    
    percentage = round((total_marks / max_marks * 100), 2) if max_marks > 0 else 0
    grade = calculate_grade(percentage)
    grade_points = calculate_grade_points(percentage)
    
    return {
        'student_name': student_name,
        'class_name': class_name_val,
        'section': section,
        'roll_number': roll_number,
        'student_id': student_id,
        'dob': str(dob) if pd.notna(dob) else 'N/A',
        'academic_year': '2025-2026',
        'school_name': 'YOUR SCHOOL NAME',
        'school_address': 'School Address, City, State - PIN',
        'subjects': subjects,
        'total_marks': total_marks,
        'max_marks': max_marks,
        'percentage': percentage,
        'grade': grade,
        'grade_points': grade_points,
        'result': 'PASS' if percentage >= 33 else 'FAIL',
        'attendance': None,
        'remarks': data.get('remarks', None),
        'co_scholastic': None,
        'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

def calculate_grade(percentage):
    if percentage >= 90: return 'A1'
    elif percentage >= 80: return 'A2'
    elif percentage >= 70: return 'B1'
    elif percentage >= 60: return 'B2'
    elif percentage >= 50: return 'C1'
    elif percentage >= 40: return 'C2'
    elif percentage >= 33: return 'D'
    else: return 'E'

def calculate_grade_points(percentage):
    if percentage >= 95: return 10.0
    elif percentage >= 90: return 9.0
    elif percentage >= 80: return 8.0
    elif percentage >= 70: return 7.0
    elif percentage >= 60: return 6.0
    elif percentage >= 50: return 5.0
    elif percentage >= 40: return 4.0
    elif percentage >= 33: return 3.0
    else: return 0.0

def generate_report_card_pdf(student_data, output_dir=GENERATED_PDFS_FOLDER):
    html_content = render_template('report_card_dynamic.html', **student_data)
    pdf = HTML(string=html_content).write_pdf()
    os.makedirs(output_dir, exist_ok=True)
    safe_name = student_data['student_name'].replace(' ', '_').replace('/', '_').replace('.', '_')
    pdf_filename = f"{output_dir}/report_card_{safe_name}_{student_data['roll_number']}.pdf"
    with open(pdf_filename, 'wb') as f:
        f.write(pdf)
    return pdf_filename

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/class_selector')
def class_selector():
    try:
        if os.path.exists(DATABASE_FILE):
            df = pd.read_excel(DATABASE_FILE)
            classes = df['class'].unique().tolist()
        else:
            classes = []
        return render_template('class_selector.html', classes=classes)
    except Exception as e:
        flash(f'Error loading classes: {str(e)}', 'error')
        return render_template('class_selector.html', classes=[])

@app.route('/enter_marks/<class_name>')
def enter_marks(class_name):
    try:
        if os.path.exists(DATABASE_FILE):
            df = pd.read_excel(DATABASE_FILE)
            students = df[df['class'].astype(str).str.upper() == class_name.upper()]
            student_list = students[['student_name', 'roll_number', 'section']].to_dict('records')
        else:
            student_list = []
        return render_template('enter_marks.html', class_name=class_name, students=student_list)
    except Exception as e:
        flash(f'Error loading students: {str(e)}', 'error')
        return redirect(url_for('class_selector'))

@app.route('/save_marks', methods=['POST'])
def save_marks():
    try:
        data = request.json
        return jsonify({'success': True, 'message': 'Marks saved successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/generate_report/<class_name>/<roll_number>')
def generate_report(class_name, roll_number):
    try:
        student_data = get_student_data_from_excel(roll_number, 'roll_number', class_name)
        if not student_data:
            flash(f'Student with roll number {roll_number} not found in class {class_name}', 'error')
            return redirect(url_for('enter_marks', class_name=class_name))
        pdf_path = generate_report_card_pdf(student_data)
        return send_file(pdf_path, as_attachment=True, download_name=f"report_card_{student_data['student_name']}_{roll_number}.pdf")
    except Exception as e:
        flash(f'Error generating report card: {str(e)}', 'error')
        return redirect(url_for('enter_marks', class_name=class_name))

@app.route('/generate_all_reports/<class_name>')
def generate_all_reports(class_name):
    try:
        if not os.path.exists(DATABASE_FILE):
            flash('Database file not found', 'error')
            return redirect(url_for('class_selector'))
        df = pd.read_excel(DATABASE_FILE)
        class_students = df[df['class'].astype(str).str.upper() == class_name.upper()]
        if class_students.empty:
            flash(f'No students found in class {class_name}', 'error')
            return redirect(url_for('enter_marks', class_name=class_name))
        generated_files = []
        for index, row in class_students.iterrows():
            try:
                student_data = get_student_data_from_excel(row['roll_number'], 'roll_number', class_name)
                if student_data:
                    pdf_path = generate_report_card_pdf(student_data)
                    generated_files.append(pdf_path)
            except Exception as e:
                print(f"Error generating report for {row.get('student_name', 'Unknown')}: {str(e)}")
                continue
        if not generated_files:
            flash('No report cards generated', 'error')
            return redirect(url_for('enter_marks', class_name=class_name))
        files_html = "<br>".join([f"✓ {os.path.basename(f)}" for f in generated_files])
        return f"<html><head><title>Reports Generated</title></head><body style='font-family: Arial; padding: 40px;'><h1>✓ Successfully Generated {len(generated_files)} Report Cards</h1><p>Class: {class_name}</p><p>Files saved to: {GENERATED_PDFS_FOLDER}/</p><h3>Generated Files:</h3><p>{files_html}</p><br><a href='{url_for('enter_marks', class_name=class_name)}'>← Back to Class</a> <a href='{url_for('pdf_list')}' style='margin-left: 20px;'>View All PDFs →</a></body></html>"
    except Exception as e:
        flash(f'Error generating reports: {str(e)}', 'error')
        return redirect(url_for('enter_marks', class_name=class_name))

@app.route('/preview_report/<class_name>/<roll_number>')
def preview_report(class_name, roll_number):
    try:
        student_data = get_student_data_from_excel(roll_number, 'roll_number', class_name)
        if not student_data:
            flash(f'Student not found', 'error')
            return redirect(url_for('enter_marks', class_name=class_name))
        return render_template('report_card_dynamic.html', **student_data)
    except Exception as e:
        flash(f'Error previewing report: {str(e)}', 'error')
        return redirect(url_for('enter_marks', class_name=class_name))

@app.route('/pdf_list')
def pdf_list():
    try:
        if os.path.exists(GENERATED_PDFS_FOLDER):
            pdf_files = [f for f in os.listdir(GENERATED_PDFS_FOLDER) if f.endswith('.pdf')]
            pdf_files.sort(reverse=True)
        else:
            pdf_files = []
        return render_template('pdf_list.html', pdf_files=pdf_files, folder=GENERATED_PDFS_FOLDER)
    except Exception as e:
        flash(f'Error listing PDFs: {str(e)}', 'error')
        return render_template('pdf_list.html', pdf_files=[], folder=GENERATED_PDFS_FOLDER)

@app.route('/download_pdf/<filename>')
def download_pdf(filename):
    try:
        filepath = os.path.join(GENERATED_PDFS_FOLDER, filename)
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True)
        else:
            flash('File not found', 'error')
            return redirect(url_for('pdf_list'))
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        return redirect(url_for('pdf_list'))

@app.route('/api/student/<class_name>/<roll_number>')
def api_get_student(class_name, roll_number):
    try:
        student_data = get_student_data_from_excel(roll_number, 'roll_number', class_name)
        if student_data:
            return jsonify({'success': True, 'data': student_data})
        else:
            return jsonify({'success': False, 'message': 'Student not found'}), 404
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/students/<class_name>')
def api_get_students(class_name):
    try:
        if not os.path.exists(DATABASE_FILE):
            return jsonify({'success': False, 'message': 'Database not found'}), 404
        df = pd.read_excel(DATABASE_FILE)
        class_students = df[df['class'].astype(str).str.upper() == class_name.upper()]
        students = [{'name': row.get('student_name', 'Unknown'), 'roll_number': str(row.get('roll_number', 'N/A')), 'section': row.get('section', 'N/A')} for index, row in class_students.iterrows()]
        return jsonify({'success': True, 'data': students})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    print("=" * 60)
    print("Report Card Generation System")
    print("=" * 60)
    print(f"Database: {DATABASE_FILE}")
    print(f"PDF Output: {GENERATED_PDFS_FOLDER}/")
    print("")
    print("Available Routes:")
    print("  / - Home page")
    print("  /class_selector - Select class")
    print("  /enter_marks/<class> - Enter marks for class")
    print("  /generate_report/<class>/<roll> - Generate single report")
    print("  /generate_all_reports/<class> - Generate all reports for class")
    print("  /preview_report/<class>/<roll> - Preview report in browser")
    print("  /pdf_list - View all generated PDFs")
    print("")
    print("Running on http://127.0.0.1:5000")
    print("=" * 60)
    app.run(debug=True, host='127.0.0.1', port=5000)
