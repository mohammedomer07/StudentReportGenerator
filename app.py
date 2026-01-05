import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from report_engine import generate_reports_from_excel
import os, threading, time, webbrowser


class ReportCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NEIEA Report Generator")
        self.root.geometry("520x460")
        self.root.resizable(False, False)

        self.excel_path = ""
        self.export_format = tk.StringVar(value="DOCX")

        tk.Label(root, text="Student Report Card Generator",
                 font=("Arial", 16, "bold")).pack(pady=15)

        self.file_label = tk.Label(root, text="No Excel file selected", fg="gray")
        self.file_label.pack(pady=5)

        tk.Button(root, text="Select Excel File", width=25,
                  command=self.select_file).pack(pady=5)

        tk.Label(root, text="Report Title (e.g. JULY RESULT REPORT)").pack()
        self.report_title_entry = tk.Entry(root, width=42)
        self.report_title_entry.pack(pady=5)

        frame = tk.Frame(root)
        frame.pack(pady=10)

        tk.Radiobutton(frame, text="Export as DOCX (Template)",
                       variable=self.export_format, value="DOCX").pack(anchor="w")

        tk.Radiobutton(frame, text="Export as PDF (No Template / Design)",
                       variable=self.export_format, value="PDF").pack(anchor="w")

        tk.Radiobutton(frame, text="Export as PDF (From DOCX – Requires Microsoft Word)",
                       variable=self.export_format, value="DOCX2PDF").pack(anchor="w")

        tk.Button(root, text="Generate Report Cards",
                  bg="#4CAF50", fg="white", width=25,
                  command=self.generate).pack(pady=20)

        self.status = tk.Label(root, text="Waiting for action...", fg="blue")
        self.status.pack()

        credit = tk.Label(root, text="By Mohammed Omer for NEIEA",
                          fg="blue", cursor="hand2",
                          font=("Arial", 9, "underline"))
        credit.pack(side="bottom", anchor="e", padx=10, pady=5)
        credit.bind("<Button-1>",
                    lambda e: webbrowser.open("https://instagram.com/_mromer_/"))

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.excel_path = path
            self.file_label.config(text=os.path.basename(path), fg="black")

    def show_progress(self, total):
        self.popup = tk.Toplevel(self.root)
        self.popup.title("Processing")
        self.popup.geometry("420x180")
        self.popup.transient(self.root)
        self.popup.grab_set()

        self.bar = ttk.Progressbar(self.popup, maximum=total, length=350)
        self.bar.pack(pady=20)

        self.msg = tk.Label(self.popup, text="Contacting NASA 🚀")
        self.msg.pack()

        self.messages = [
            "Contacting NASA 🚀",
            "Waking Einstein 🧠",
            "Borrowing Tesla power ⚡",
            "Convincing Microsoft Word 😤",
            "Almost done 😌"
        ]

    def update_progress(self, i):
        self.bar["value"] = i
        self.msg.config(text=self.messages[i % len(self.messages)])

    def close_progress(self):
        self.popup.destroy()

    def generate(self):
        if not self.excel_path:
            messagebox.showerror("Error", "Select Excel file first.")
            return

        title = self.report_title_entry.get().strip() or "STUDENT REPORT CARD"
        mode = self.export_format.get()

        if mode == "DOCX2PDF":
            if not messagebox.askyesno(
                "Microsoft Word Required",
                "This option requires Microsoft Word.\n\nProceed?"
            ):
                return

        def task():
            try:
                files = generate_reports_from_excel(
                    self.excel_path, title, mode
                )

                self.root.after(0, self.show_progress, len(files))

                for i in range(len(files)):
                    time.sleep(0.4)
                    self.root.after(0, self.update_progress, i + 1)

                self.root.after(0, self.close_progress)

                def done():
                    if messagebox.askyesno(
                        "Success",
                        f"{len(files)} report(s) generated.\n\nOpen output folder?"
                    ):
                        os.startfile(os.path.abspath("output"))

                self.root.after(0, done)

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

        threading.Thread(target=task, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    ReportCardApp(root)
    root.mainloop()
