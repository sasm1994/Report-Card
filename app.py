from flask import Flask, render_template, request, jsonify, send_file
from master_database import MarksDatabase, get_subjects, get_subject_group
import os
import pdfkit

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "static", "logo.png")

db = MarksDatabase()

# ════════════════════════════════════════════════════════════
# TEMPLATE ROUTER — picks correct report card template
# ════════════════════════════════════════════════════════════

def get_template_for_class(class_name):
    group = get_subject_group(class_name)
    return {
        "NUR": "NUR Template.html",
        "KG": "KG Template.html",
        "CLASS_1_2": "1-2 Template.html",
        "CLASS_3_9": "3-9 Template.html",
    }.get(group, "3-9-Template.html")


# ════════════════════════════════════════════════════════════
# HOME — class selection
# ════════════════════════════════════════════════════════════

@app.route('/')
def home():
    classes = db.get_all_classes()
    return render_template('home.html', classes=classes)


# ════════════════════════════════════════════════════════════
# MARKS ENTRY PAGE
# ════════════════════════════════════════════════════════════

@app.route('/marks-entry/<class_name>')
def marks_entry(class_name):
    students = db.get_students(class_name)
    subjects = db.get_subject_list(class_name)
    group    = get_subject_group(class_name)

    # Sessions, terms, exam types for dropdowns
    sessions     = ["2025-26", "2026-27"]
    terms        = ["1", "2"]
    examinations = ["Periodic Test", "Unit Test",
                    "Half Yearly", "Semester I",
                    "Semester II", "Annual Exam"]

    # Section extracted from sheet name e.g. Class_3_B → B
    try:
        section = class_name.split("_")[-1]
    except Exception:
        section = "A"

    # Class label e.g. Class_3_B → Class 3
    try:
        parts       = class_name.split("_")
        class_label = " ".join(parts[:-1]).replace("_", " ")
    except Exception:
        class_label = class_name

    return render_template(
        'marks_entry.html',
        class_name   = class_name,       # e.g. "Class_3_B"
        class_label  = class_label,      # e.g. "Class 3"
        section      = section,          # e.g. "B"
        students     = students,         # list of student dicts
        subjects     = subjects,         # list of column keys for this class
        group        = group,            # "NUR" / "KG" / "CLASS_1_2" / "CLASS_3_9"
        sessions     = sessions,
        terms        = terms,
        examinations = examinations,
    )


# ════════════════════════════════════════════════════════════
# API — GET student marks
# ════════════════════════════════════════════════════════════

@app.route('/api/marks/<class_name>/<int:roll_no>', methods=['GET'])
def api_get_marks(class_name, roll_no):
    students = db.get_students(class_name)
    for student in students:
        if student['RollNo'] == roll_no:
            return jsonify(student)
    return jsonify({"error": "Student not found"}), 404


# ════════════════════════════════════════════════════════════
# API — SAVE student marks
# ════════════════════════════════════════════════════════════

@app.route('/api/marks/<class_name>/<int:roll_no>', methods=['POST'])
def api_save_marks(class_name, roll_no):
    data = request.get_json()
    if not data:
        return jsonify({"success": False,
                        "message": "No data received"}), 400

    success, msg = db.save_marks(class_name, roll_no, data)
    return jsonify({"success": success, "message": msg})


# ════════════════════════════════════════════════════════════
# API — GET all students in a class (used by marks_entry.html JS)
# ════════════════════════════════════════════════════════════

@app.route('/api/class/<class_name>/students')
def api_students(class_name):
    students = db.get_students(class_name)
    return jsonify(students)


# ════════════════════════════════════════════════════════════
# GENERATE PDFs for entire class
# ════════════════════════════════════════════════════════════

@app.route('/class/<class_name>/generate_pdfs')
def generate_pdfs(class_name):
    students = db.get_students(class_name)
    template = get_template_for_class(class_name)
    output_dir = os.path.join(BASE_DIR, "generated_pdfs", class_name)
    os.makedirs(output_dir, exist_ok=True)

    generated = []
    errors = []

    try:
        config = pdfkit.configuration(
            wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
        )
    except Exception as e:
        return render_template(
            "pdf_list.html",
            class_name=class_name,
            files=[],
            errors=[f"wkhtmltopdf not found: {str(e)}"]
        )

    options = {
        "page-size": "A4",
        "orientation": "Portrait",
        "encoding": "UTF-8",
        "margin-top": "5mm",
        "margin-bottom": "5mm",
        "margin-left": "5mm",
        "margin-right": "5mm",
        "enable-local-file-access": None
    }

    for student in students:
        try:
            html = render_template(
                template,
                student=student,
                class_name=class_name,
                logo_path=LOGO_PATH.replace("\\", "/")
            )

            safe_name = str(student["StudentName"]).replace(" ", "_").replace("/", "_")
            filename = f"{student['RollNo']}_{safe_name}.pdf"
            filepath = os.path.join(output_dir, filename)

            pdfkit.from_string(html, filepath, options=options, configuration=config)
            generated.append(filename)

        except Exception as e:
            errors.append(f"{student.get('StudentName', 'Unknown Student')}: {str(e)}")

    return render_template(
        "pdf_list.html",
        class_name=class_name,
        files=generated,
        errors=errors
    )

# ════════════════════════════════════════════════════════════
# LIST generated PDFs
# ════════════════════════════════════════════════════════════

@app.route('/class/<class_name>/pdfs')
def list_pdfs(class_name):
    pdf_dir = os.path.join(BASE_DIR, "generated_pdfs", class_name)
    files   = []

    if os.path.exists(pdf_dir):
        files = sorted([
            f for f in os.listdir(pdf_dir)
            if f.endswith(".pdf")
        ])

    return render_template(
        'pdf_list.html',
        class_name = class_name,
        files      = files,
        errors     = [],
    )


# ════════════════════════════════════════════════════════════
# DOWNLOAD a PDF
# ════════════════════════════════════════════════════════════

@app.route('/class/<class_name>/pdfs/<filename>')
def download_file(class_name, filename):
    pdf_dir  = os.path.join(BASE_DIR, "generated_pdfs", class_name)
    filepath = os.path.join(pdf_dir, filename)

    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        return jsonify({"error": "File not found"}), 404


# ════════════════════════════════════════════════════════════
# EXPORT class marks to CSV
# ════════════════════════════════════════════════════════════

@app.route('/class/<class_name>/export')
def export_marks(class_name):
    output_file = os.path.join(BASE_DIR, f"{class_name}_marks.csv")
    success, msg = db.export_to_csv(class_name, output_file)

    if success:
        return send_file(output_file, as_attachment=True)
    else:
        return jsonify({"error": msg}), 400


# ════════════════════════════════════════════════════════════
# RUN
# ════════════════════════════════════════════════════════════

if __name__ == '__main__':
    app.run(debug=True, port=5000)