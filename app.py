import tkinter as tk
from tkinter import filedialog, messagebox
from report_engine import generate_reports_from_excel
import os
import webbrowser


class ReportCardApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Omer Report Generator")
        self.root.geometry("500x420")
        self.root.resizable(False, False)

        self.excel_path = ""

        # ===== TITLE =====
        title = tk.Label(
            root,
            text="Student Report Card Generator",
            font=("Arial", 16, "bold")
        )
        title.pack(pady=15)

        # ===== EXCEL FILE LABEL =====
        self.file_label = tk.Label(
            root,
            text="No Excel file selected",
            fg="gray"
        )
        self.file_label.pack(pady=8)

        # ===== SELECT FILE BUTTON =====
        select_btn = tk.Button(
            root,
            text="Select Excel File",
            width=25,
            command=self.select_file
        )
        select_btn.pack(pady=5)

        # ===== REPORT TITLE INPUT =====
        title_label = tk.Label(
            root,
            text="Report Title (e.g. JULY RESULT REPORT)"
        )
        title_label.pack(pady=(10, 0))

        self.report_title_entry = tk.Entry(root, width=40)
        self.report_title_entry.pack(pady=5)

        # ===== EXPORT OPTION =====
        self.export_format = tk.StringVar(value="DOCX")

        export_frame = tk.Frame(root)
        export_frame.pack(pady=10)

        tk.Radiobutton(
            export_frame,
            text="Export as DOCX",
            variable=self.export_format,
            value="DOCX"
        ).pack(side="left", padx=15)

        tk.Radiobutton(
            export_frame,
            text="Export as PDF",
            variable=self.export_format,
            value="PDF"
        ).pack(side="left", padx=15)

        # ===== GENERATE BUTTON =====
        generate_btn = tk.Button(
            root,
            text="Generate Report Cards",
            width=25,
            bg="#4CAF50",
            fg="white",
            command=self.generate_reports
        )
        generate_btn.pack(pady=20)

        # ===== STATUS LABEL =====
        self.status_label = tk.Label(
            root,
            text="Waiting for action...",
            fg="blue"
        )
        self.status_label.pack(pady=5)

        # ===== WATERMARK =====
        credit = tk.Label(
            root,
            text="By Mohammed Omer for NEIEA",
            fg="blue",
            cursor="hand2",
            font=("Arial", 9, "underline")
        )
        credit.pack(side="bottom", anchor="e", padx=10, pady=5)

        credit.bind(
            "<Button-1>",
            lambda e: webbrowser.open("https://instagram.com/_mromer_/")
        )

    # ===== FILE SELECT =====
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if file_path:
            self.excel_path = file_path
            self.file_label.config(text=os.path.basename(file_path), fg="black")
            self.status_label.config(text="Excel file selected.")

    # ===== SUCCESS DIALOG WITH VIEW NOW =====
    def show_success_dialog(self, message):
        dialog = tk.Toplevel(self.root)
        dialog.title("Success")
        dialog.geometry("380x160")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        label = tk.Label(
            dialog,
            text=message,
            wraplength=340,
            justify="center"
        )
        label.pack(pady=25)

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)

        view_btn = tk.Button(
            btn_frame,
            text="View Now",
            width=12,
            command=lambda: self.open_output_folder(dialog)
        )
        view_btn.pack(side="left", padx=10)

        ok_btn = tk.Button(
            btn_frame,
            text="OK",
            width=12,
            command=dialog.destroy
        )
        ok_btn.pack(side="left", padx=10)

    def open_output_folder(self, dialog=None):
        output_path = os.path.abspath("output")
        os.startfile(output_path)
        if dialog:
            dialog.destroy()

    # ===== GENERATE REPORTS =====
    def generate_reports(self):
        if not self.excel_path:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return

        report_title = self.report_title_entry.get().strip()
        if not report_title:
            report_title = "STUDENT REPORT CARD"

        export_format = self.export_format.get()

        try:
            self.status_label.config(text="Generating report cards...")
            self.root.update_idletasks()

            count = generate_reports_from_excel(
                self.excel_path,
                report_title,
                export_format
            )

            if count == 0:
                raise Exception("No valid student records found in the Excel file.")

            self.status_label.config(
                text=f"{count} report card(s) generated successfully!"
            )

            self.show_success_dialog(
                f"{count} report card(s) generated as {export_format}.\nCheck the output folder."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="Error occurred.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportCardApp(root)
    root.mainloop()
