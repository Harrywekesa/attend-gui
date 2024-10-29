import pandas as pd
import sqlite3
from datetime import datetime
from tkinter import Tk, Label, Button, Entry, filedialog, StringVar, messagebox, OptionMenu
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import xlsxwriter
import os

# Database connection setup
conn = sqlite3.connect('attendance.db')
cursor = conn.cursor()

# Database tables
cursor.execute('''
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY,
        name TEXT,
        admission_number TEXT UNIQUE
    )
''')
cursor.execute('''
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY,
        student_id INTEGER,
        unit_name TEXT,
        date TEXT,
        is_present BOOLEAN,
        FOREIGN KEY(student_id) REFERENCES students(id)
    )
''')
conn.commit()

# 1. Upload students from Excel
def upload_students():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            if not set(['Name', 'Admission Number']).issubset(df.columns):
                messagebox.showerror("Error", "Excel file must contain 'Name' and 'Admission Number' columns.")
                return

            for _, row in df.iterrows():
                try:
                    cursor.execute("INSERT INTO students (name, admission_number) VALUES (?, ?)",
                                   (row['Name'], row['Admission Number']))
                except sqlite3.IntegrityError:
                    print(f"Skipping duplicate entry for admission number {row['Admission Number']}.")
            conn.commit()
            messagebox.showinfo("Success", "Students uploaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error processing file: {e}")

# 2. Take attendance for a specific unit
def take_attendance():
    unit_name = unit_name_entry.get()
    if not unit_name:
        messagebox.showwarning("Warning", "Please enter a unit name.")
        return

    today = datetime.today().strftime('%Y-%m-%d')
    cursor.execute("SELECT * FROM students")
    students = cursor.fetchall()

    for student in students:
        status = messagebox.askquestion("Attendance", f"Mark attendance for {student[1]} (Admission Number: {student[2]})", icon='question')
        is_present = (status == 'yes')
        cursor.execute("INSERT INTO attendance (student_id, unit_name, date, is_present) VALUES (?, ?, ?, ?)",
                       (student[0], unit_name, today, is_present))
    conn.commit()
    messagebox.showinfo("Success", "Attendance recorded successfully.")

# 3. Generate report
def generate_report():
    unit_name = unit_name_entry.get()
    output_format = report_format_var.get()

    if not unit_name:
        messagebox.showwarning("Warning", "Please enter a unit name.")
        return

    cursor.execute("""
        SELECT s.name, s.admission_number, a.date, a.is_present
        FROM students s
        LEFT JOIN attendance a ON s.id = a.student_id
        WHERE a.unit_name = ?
        ORDER BY s.name, a.date
    """, (unit_name,))
    records = cursor.fetchall()

    dates = sorted(set(record[2] for record in records if record[2]))

    if output_format == "Excel":
        generate_excel_report(records, unit_name, dates)
    elif output_format == "PDF":
        generate_pdf_report(records, unit_name, dates)
    else:
        messagebox.showerror("Error", "Unsupported format. Choose either 'Excel' or 'PDF'.")

def generate_excel_report(records, unit_name, dates):
    file_name = f"{unit_name}_Attendance_Report.xlsx"
    with xlsxwriter.Workbook(file_name) as workbook:
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'Name')
        worksheet.write('B1', 'Admission Number')
        for col_num, date in enumerate(dates, start=2):
            worksheet.write(0, col_num, date)

        row = 1
        current_name = ""
        for record in records:
            if record[0] != current_name:
                current_name = record[0]
                row += 1
                worksheet.write(row, 0, record[0])
                worksheet.write(row, 1, record[1])

            col = dates.index(record[2]) + 2 if record[2] in dates else None
            if col is not None:
                worksheet.write(row, col, 'X' if record[3] else 'O')

    messagebox.showinfo("Success", f"Excel report generated as {file_name}")

def generate_pdf_report(records, unit_name, dates):
    file_name = f"{unit_name}_Attendance_Report.pdf"
    c = canvas.Canvas(file_name, pagesize=letter)
    width, height = letter
    c.drawString(30, height - 40, f"Attendance Report for {unit_name}")
    c.drawString(30, height - 60, "Name")
    c.drawString(150, height - 60, "Admission Number")
    x_pos = 300
    for date in dates:
        c.drawString(x_pos, height - 60, date)
        x_pos += 50

    y = height - 80
    current_name = ""
    for record in records:
        if record[0] != current_name:
            current_name = record[0]
            y -= 20
            c.drawString(30, y, record[0])
            c.drawString(150, y, record[1])

        x_pos = 300 + dates.index(record[2]) * 50 if record[2] in dates else None
        if x_pos:
            c.drawString(x_pos, y, 'X' if record[3] else 'O')

        if y < 50:
            c.showPage()
            y = height - 80

    c.save()
    messagebox.showinfo("Success", f"PDF report generated as {file_name}")

# GUI setup
root = Tk()
root.title("Attendance Management System")
root.geometry("400x300")

# Upload students button
Label(root, text="Upload Student Data").pack(pady=10)
Button(root, text="Upload from Excel", command=upload_students).pack()

# Unit name entry
Label(root, text="Enter Unit Name").pack(pady=10)
unit_name_entry = Entry(root)
unit_name_entry.pack()

# Take attendance button
Button(root, text="Take Attendance", command=take_attendance).pack(pady=10)

# Report format selection
Label(root, text="Select Report Format").pack(pady=10)
report_format_var = StringVar(root)
report_format_var.set("Excel")
OptionMenu(root, report_format_var, "Excel", "PDF").pack()

# Generate report button
Button(root, text="Generate Report", command=generate_report).pack(pady=10)

root.mainloop()
