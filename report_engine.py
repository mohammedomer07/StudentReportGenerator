import os
from excel_reader import read_students_excel
from docx_generator import generate_student_docx
from pdf_generator import generate_student_report


def generate_reports_from_excel(excel_path, report_title, export_format):
    students = read_students_excel(excel_path)

    if not students:
        return 0

    base_output = "output"
    docx_output = os.path.join(base_output, "DOCX")
    pdf_output = os.path.join(base_output, "PDF")

    # Create folders safely
    os.makedirs(docx_output, exist_ok=True)
    os.makedirs(pdf_output, exist_ok=True)

    count = 0

    for student in students:
        if export_format == "DOCX":
            generate_student_docx(
                student,
                report_title,
                output_folder=docx_output
            )

        elif export_format == "PDF":
            generate_student_report(
                student,
                report_title,
                output_folder=pdf_output
            )

        count += 1

    return count
