from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import sys


# ===============================
# PYINSTALLER-SAFE RESOURCE PATH
# ===============================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ===============================
# FONT CONFIG (GLOBAL, REQUIRED)
# ===============================
FONT_NAME = "Poppins"
FONT_PATH = resource_path("assets/fonts/Poppins-Regular.ttf")

if FONT_NAME not in pdfmetrics.getRegisteredFontNames():
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))


# ===============================
# PDF GENERATION
# ===============================
def generate_student_report(student, report_title, output_folder="output"):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_name = f"{student['student_name'].replace(' ', '_')}.pdf"
    file_path = os.path.join(output_folder, file_name)

    c = canvas.Canvas(file_path, pagesize=A4)
    width, height = A4

    # HEADER
    c.setFont(FONT_NAME, 16)
    c.drawCentredString(width / 2, height - 50, student["institution"])

    c.setFont(FONT_NAME, 13)
    c.drawCentredString(width / 2, height - 80, report_title)

    # STUDENT DETAILS
    c.setFont(FONT_NAME, 10)
    y = height - 130

    fields = [
        ("Student Name", student["student_name"]),
        ("Father's Name", student["father_name"]),
        ("Mother's Name", student["mother_name"]),
        ("Age", student["age"]),
        ("Gender", student["gender"]),
        ("Parents Contact", student["parents_contact"]),
        ("Region", student["region"]),
        ("Class / Standard", student["standard"]),
    ]

    for label, value in fields:
        c.drawString(50, y, f"{label} : {value}")
        y -= 18

    # BASELINE (OPTIONAL)
    if student.get("baseline_english") not in [None, 0]:
        y -= 20
        c.setFont(FONT_NAME, 11)
        c.drawString(50, y, "Baseline Assessment Score (25 Marks)")
        y -= 15
        c.line(50, y, 500, y)

        y -= 20
        c.setFont(FONT_NAME, 10)
        for subj, key in [
            ("English", "baseline_english"),
            ("Math", "baseline_math"),
            ("Science", "baseline_science"),
        ]:
            c.drawString(50, y, subj)
            c.drawString(300, y, str(student.get(key, "")))
            y -= 18

        c.drawString(50, y, f"Total : {student.get('baseline_total', '')}")
        c.drawString(300, y, f"{student.get('baseline_percentage', '')}%")

    # MONTHLY
    y -= 30
    c.setFont(FONT_NAME, 11)
    c.drawString(50, y, "Monthly Assessment (25 Marks)")
    y -= 15
    c.line(50, y, 500, y)

    y -= 20
    c.setFont(FONT_NAME, 10)
    for subj, key in [
        ("English", "english"),
        ("Math", "math"),
        ("Science", "science"),
    ]:
        c.drawString(50, y, subj)
        c.drawString(300, y, str(student[key]))
        y -= 18

    y -= 10
    c.drawString(50, y, f"Total Marks : {student['total']}")
    c.drawString(50, y - 20, f"Percentage : {student['percentage']}%")

    c.showPage()
    c.save()

    return file_path
