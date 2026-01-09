import pandas as pd
import math


def clean_number(value):
    """
    - Converts 1.0 → 1
    - Keeps 0 as 0
    - Converts NaN → ""
    """
    if pd.isna(value):
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
    return value


def clean_mark(value):
    """
    - 'Ab' / 'AB' → 'ABSENT'
    - NaN → ''
    - 0 stays 0
    - Numbers stay numbers
    """
    if pd.isna(value):
        return ""
    if isinstance(value, str) and value.strip().lower() == "ab":
        return "ABSENT"
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
    return value


def read_students_excel(excel_path):
    df = pd.read_excel(excel_path)

    students = []

    for _, row in df.iterrows():
        student = {
            "student_name": str(row.get("Student's Name", "")).strip(),
            "father_name": str(row.get("Father's Name", "")).strip(),
            "mother_name": str(row.get("Mother's Name", "")).strip(),
            "gender": str(row.get("Gender", "")).strip(),
            "age": clean_number(row.get("Age", "")),
            "standard": str(row.get("Standard", "")).strip(),
            "roll_no": clean_number(row.get("Roll No", "")),
            "address": str(row.get("Address", "")).strip(),
            "year": clean_number(row.get("Year", "")),
         #   "remark": str(row.get("Remark", "")).strip(),
            "contact_number": clean_number(row.get("Contact Number", "")),

            "english": clean_mark(row.get("English", "")),
            "math": clean_mark(row.get("Math", "")),
            "science": clean_mark(row.get("Science", "")),

            "total_classes": clean_number(row.get("Total Classes", 0)),
            "classes_attended": clean_number(row.get("Classes Attended", 0)),
        }

        students.append(student)

    return students
