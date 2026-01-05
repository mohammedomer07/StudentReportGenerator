from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import sys
import math


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
# FONT CONFIG (SAFE)
# ===============================
FONT_NAME = "Poppins"
FALLBACK_FONT = "Helvetica"
FONT_PATH = resource_path("assets/fonts/Poppins-Regular.ttf")

FONT_IN_USE = FALLBACK_FONT

try:
    if os.path.exists(FONT_PATH):
        if FONT_NAME not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
        FONT_IN_USE = FONT_NAME
except Exception:
    FONT_IN_USE = FALLBACK_FONT


# ===============================
# HELPERS
# ===============================
def clean_value(val):
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    return str(val).strip()


def safe_mark(val):
    if val is None:
        return "ABSENT"

    if isinstance(val, str):
        v = val.strip().upper()
        if v in ["AB", "ABSENT"]:
            return "ABSENT"
        if v == "":
            return ""
        try:
            return int(float(v))
        except Exception:
            return "ABSENT"

    if isinstance(val, float):
        if math.isnan(val):
            return "ABSENT"
        return int(val)

    if isinstance(val, int):
        return val

    return "ABSENT"


def calc_total_and_percentage(english, math, science):
    marks = []
    for m in [english, math, science]:
        if isinstance(m, int):
            marks.append(m)

    total = sum(marks)
    percentage = round((total / 75) * 100, 2) if marks else 0
    return total, percentage


def calc_attendance_percentage(attended, total):
    try:
        if total in [0, None, ""]:
            return 0
        return round((int(attended) / int(total)) * 100, 2)
    except Exception:
        return 0


# ===============================
# PDF GENERATION
# ===============================
def generate_student_report(student, report_title, output_folder="output/PDF"):
    os.makedirs(output_folder, exist_ok=True)

    file_name = f"{student['student_name'].replace(' ', '_')}.pdf"
    file_path = os.path.join(output_folder, file_name)

    c = canvas.Canvas(file_path, pagesize=A4)
    width, height = A4

    # ===== HEADER =====
    c.setFont(FONT_IN_USE, 14)
    c.drawCentredString(width / 2, height - 60, report_title)

    # ===== STUDENT DETAILS =====
    c.setFont(FONT_IN_USE, 10)
    y = height - 110

    details = [
        ("Student Name", clean_value(student.get("student_name"))),
        ("Father's Name", clean_value(student.get("father_name"))),
        ("Mother's Name", clean_value(student.get("mother_name"))),
        ("Gender", clean_value(student.get("gender"))),
        ("Age", clean_value(student.get("age"))),
        ("Roll No", clean_value(student.get("roll_no"))),
        ("Class", clean_value(student.get("standard"))),
        ("Year", clean_value(student.get("year"))),
        ("Contact No", clean_value(student.get("contact_number"))),
        ("Address", clean_value(student.get("address"))),
    ]

    for label, value in details:
        c.drawString(50, y, f"{label} : {value}")
        y -= 16

    # ===== MARKS =====
    y -= 15
    c.setFont(FONT_IN_USE, 11)
    c.drawString(50, y, "Subject-wise Marks (Out of 25)")
    y -= 10
    c.line(50, y, 500, y)

    y -= 18
    c.setFont(FONT_IN_USE, 10)

    eng = safe_mark(student.get("english"))
    math_m = safe_mark(student.get("math"))
    sci = safe_mark(student.get("science"))

    c.drawString(50, y, f"English : {eng}")
    y -= 16
    c.drawString(50, y, f"Math : {math_m}")
    y -= 16
    c.drawString(50, y, f"Science : {sci}")

    total, percentage = calc_total_and_percentage(eng, math_m, sci)

    y -= 20
    c.drawString(50, y, f"Total Marks : {total}")
    y -= 16
    c.drawString(50, y, f"Percentage : {percentage}%")

    # ===== ATTENDANCE =====
    y -= 25
    c.setFont(FONT_IN_USE, 11)
    c.drawString(50, y, "Attendance")
    y -= 10
    c.line(50, y, 500, y)

    y -= 18
    c.setFont(FONT_IN_USE, 10)

    total_classes = student.get("total_classes", 0)
    attended = student.get("classes_attended", 0)
    attendance_percentage = calc_attendance_percentage(attended, total_classes)

    c.drawString(50, y, f"Total Classes : {total_classes}")
    y -= 16
    c.drawString(50, y, f"Classes Attended : {attended}")
    y -= 16
    c.drawString(50, y, f"Attendance Percentage : {attendance_percentage}%")

    c.showPage()
    c.save()

    return file_path
