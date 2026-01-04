import pandas as pd


def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def read_students_excel(file_path):
    df = pd.read_excel(file_path)
    students = []

    for _, row in df.iterrows():
        english = int(row.get("English", 0) or 0)
        math = int(row.get("Math", 0) or 0)
        science = int(row.get("Science", 0) or 0)

        total_marks = english + math + science
        percentage = round((total_marks / 75) * 100, 2)

        total_classes = int(row.get("Total Classes", 0) or 0)
        classes_attended = int(row.get("Classes Attended", 0) or 0)

        attendance_percentage = ""
        if total_classes > 0:
            attendance_percentage = round(
                (classes_attended / total_classes) * 100, 2
            )

        student = {
            "student_name": clean(row.get("Student's Name")),
            "father_name": clean(row.get("Father's Name")),
            "mother_name": clean(row.get("Mother's Name")),
            "gender": clean(row.get("Gender")),
            "age": clean(row.get("Age")),
            "standard": clean(row.get("Standard")),
            "roll_no": clean(row.get("Roll No")),
            "address": clean(row.get("Address")),
            "year": clean(row.get("Year")),
            "contact_number": clean(row.get("Contact Number")),

            "english": english,
            "math": math,
            "science": science,

            "total_marks": total_marks,
            "percentage": percentage,

            "total_classes": total_classes,
            "classes_attended": classes_attended,
            "attendance_percentage": attendance_percentage,
        }

        students.append(student)

    return students
