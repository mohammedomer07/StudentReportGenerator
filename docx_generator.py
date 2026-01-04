from docx import Document
import os
import sys


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


TEMPLATE_PATH = resource_path("assets/Progress Report.docx")


def calculate_grade(marks):
    try:
        marks = int(marks)
    except:
        return ""
    if marks >= 22:
        return "A+"
    elif marks >= 18:
        return "A"
    elif marks >= 15:
        return "B"
    elif marks >= 10:
        return "C"
    else:
        return "D"


def replace_placeholders_in_paragraph(paragraph, replacements):
    # 🚫 Skip paragraphs that contain images
    for run in paragraph.runs:
        if run.element.xpath(".//w:drawing"):
            return

    # Ensure at least one run exists
    if not paragraph.runs:
        paragraph.add_run("")

    full_text = "".join(run.text for run in paragraph.runs)

    for key, value in replacements.items():
        if key in full_text:
            full_text = full_text.replace(key, str(value))

    # Clear existing runs safely
    for run in paragraph.runs:
        run.text = ""

    paragraph.runs[0].text = full_text


def replace_everywhere(doc, replacements):
    for paragraph in doc.paragraphs:
        replace_placeholders_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_placeholders_in_paragraph(paragraph, replacements)


def generate_student_docx(student, report_title, output_folder="output"):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    doc = Document(TEMPLATE_PATH)

    # ---------------- PLACEHOLDER DATA ----------------
    replacements = {
        "{{student_name}}": student.get("student_name", ""),
        "{{father_name}}": student.get("father_name", ""),
        "{{mother_name}}": student.get("mother_name", ""),
        "{{gender}}": student.get("gender", ""),
        "{{age}}": student.get("age", ""),
        "{{standard}}": student.get("standard", ""),
        "{{roll_no}}": student.get("roll_no", ""),
        "{{address}}": student.get("address", ""),

        "{{report_title}}": report_title,
        "{{year}}": student.get("year", ""),
        "{{contact_number}}": student.get("contact_number", ""),

        "{{total_marks}}": student.get("total_marks", ""),
        "{{percentage}}": f"{student.get('percentage', '')}%", 

        "{{total_classes}}": student.get("total_classes", ""),
        "{{classes_attended}}": student.get("classes_attended", ""),
        "{{attendance_percentage}}": f"{student.get('attendance_percentage', '')}%",
    }

    # ✅ SAFE PLACEHOLDER REPLACEMENT
    replace_everywhere(doc, replacements)

    # ---------------- SUBJECT TABLE ----------------
    for table in doc.tables:
        for row in table.rows:
            subject = row.cells[0].text.strip().lower()

            if subject == "english":
                row.cells[1].text = "25"
                row.cells[2].text = str(student.get("english", ""))
                row.cells[3].text = calculate_grade(student.get("english", 0))

            elif subject in ["maths", "mathematics", "math"]:
                row.cells[1].text = "25"
                row.cells[2].text = str(student.get("math", ""))
                row.cells[3].text = calculate_grade(student.get("math", 0))

            elif subject == "science":
                row.cells[1].text = "25"
                row.cells[2].text = str(student.get("science", ""))
                row.cells[3].text = calculate_grade(student.get("science", 0))

    filename = f"{student.get('student_name', 'student').replace(' ', '_')}.docx"
    path = os.path.join(output_folder, filename)
    doc.save(path)

    return path
