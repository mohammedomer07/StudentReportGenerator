import os
from excel_reader import read_students_excel
from docx_generator import generate_student_docx
from pdf_generator import generate_student_report
from docx2pdf import convert


def generate_reports_from_excel(excel_path, report_title, export_format):
    students = read_students_excel(excel_path)
    if not students:
        return []

    base = "output"
    docx_dir = os.path.join(base, "DOCX")
    pdf_dir = os.path.join(base, "PDF")
    docx2pdf_dir = os.path.join(base, "DOCX2PDF")

    os.makedirs(docx_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(docx2pdf_dir, exist_ok=True)

    generated = []

    for student in students:

        # 1️⃣ DOCX (Template)
        if export_format == "DOCX":
            path = generate_student_docx(student, report_title, docx_dir)
            generated.append(path)

        # 2️⃣ PDF (No Template – reportlab)
        elif export_format == "PDF":
            path = generate_student_report(student, report_title, pdf_dir)
            generated.append(path)

        # 3️⃣ DOCX ➜ PDF (Template + Word required)
        elif export_format == "DOCX2PDF":
            docx_path = generate_student_docx(student, report_title, docx2pdf_dir)
            pdf_path = os.path.join(
                pdf_dir,
                os.path.basename(docx_path).replace(".docx", ".pdf")
            )
            convert(docx_path, pdf_path)
            generated.append(pdf_path)

    return generated
