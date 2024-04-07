import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import openpyxl

# Function to calculate GPA
def calculate_gpa():
    total_grade_points = 0
    total_credits = 0
    
    for subject in subjects:
        credit = subjects[subject]['credit'].get()
        grade = subjects[subject]['grade'].get()
        
        if grade in grade_points:
            total_grade_points += grade_points[grade] * int(credit)
            total_credits += int(credit)
    
    if total_credits == 0:
        gpa = 0
    else:
        gpa = total_grade_points / total_credits
    
    result_label.config(text="GPA: {:.2f}".format(gpa))
    save_button.config(state=tk.NORMAL)

# Function to add a subject
def add_subject():
    subject_name = subject_entry.get()
    credit_value = credit_combobox.get()
    grade_value = grade_combobox.get()
    
    if subject_name and credit_value and grade_value:
        subjects[subject_name] = {'credit': tk.StringVar(), 'grade': tk.StringVar()}
        subjects[subject_name]['credit'].set(credit_value)
        subjects[subject_name]['grade'].set(grade_value)
        update_subjects_list()

# Function to update the subjects list
def update_subjects_list():
    table.delete(*table.get_children())  # Clear table before updating
    for subject in subjects:
        subject_name = subject
        credit_value = subjects[subject]['credit'].get()
        grade_value = subjects[subject]['grade'].get()
        table.insert("", "end", values=(subject_name, credit_value, grade_value))

# Function to save result to a file
def save_result():
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filename:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Subject Name", "Credit Value", "Grade"])
        for subject in subjects:
            credit_value = subjects[subject]['credit'].get()
            grade_value = subjects[subject]['grade'].get()
            ws.append([subject, credit_value, grade_value])
        wb.save(filename)

# Function to reset the form
def reset_form():
    subject_entry.delete(0, tk.END)
    credit_combobox.set('')
    grade_combobox.set('')

# Initialize tkinter
root = tk.Tk()
root.title("GPA Calculator")

# Dictionary to store subjects and their credit/grade
subjects = {}

# Dictionary for grade points
grade_points = {
    "A+": 4.00, "A": 4.00, "A-": 3.70,
    "B+": 3.30, "B": 3.00, "B-": 2.70,
    "C+": 2.30, "C": 2.00, "C-": 1.70,  
    "D+": 1.30, "D": 1.00, "E": 0.00,
    "I": 0.00
}

# Frame for adding subjects
add_subject_frame = tk.Frame(root)
add_subject_frame.pack(pady=10)

# Subject entry
subject_label = tk.Label(add_subject_frame, text="Subject Name:")
subject_label.grid(row=0, column=0)
subject_entry = tk.Entry(add_subject_frame)
subject_entry.grid(row=0, column=1)

# Credit combobox
credit_label = tk.Label(add_subject_frame, text="Credit Value:")
credit_label.grid(row=0, column=2)
credit_combobox = ttk.Combobox(add_subject_frame, values=["2", "3"])
credit_combobox.grid(row=0, column=3)
credit_combobox.current(0)

# Grade combobox
grade_label = tk.Label(add_subject_frame, text="Grade:")
grade_label.grid(row=0, column=4)
grade_combobox = ttk.Combobox(add_subject_frame, values=list(grade_points.keys()))
grade_combobox.grid(row=0, column=5)
grade_combobox.current(0)

# Add subject button
add_button = tk.Button(add_subject_frame, text="Add Subject", command=add_subject, bg="blue")
add_button.grid(row=0, column=6, padx=10)

# Green button to reset form
reset_button = tk.Button(add_subject_frame, text="Reset Form", command=reset_form, bg="blue")
reset_button.grid(row=0, column=7, padx=10)

# Frame for displaying subjects
display_subjects_frame = tk.Frame(root)
display_subjects_frame.pack()

# Subjects table
table = ttk.Treeview(display_subjects_frame, columns=("Subject Name", "Credit Value", "Grade"), show="headings")
table.heading("Subject Name", text="Subject Name", anchor="center")
table.heading("Credit Value", text="Credit Value", anchor="center")
table.heading("Grade", text="Grade", anchor="center")
table.column("Subject Name", anchor="center")  
table.column("Credit Value", anchor="center")  
table.column("Grade", anchor="center")  
table.pack(side="left", fill="both", expand=True)

# Scrollbar for the subjects table
scrollbar = ttk.Scrollbar(display_subjects_frame, orient="vertical", command=table.yview)
scrollbar.pack(side="right", fill="y")
table.configure(yscrollcommand=scrollbar.set)

# Frame for calculate button and save button
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# Calculate button
calculate_button = tk.Button(button_frame, text="Calculate GPA", command=calculate_gpa, bg="blue")
calculate_button.grid(row=0, column=0, padx=5)

# Save button
save_button = tk.Button(button_frame, text="Save Result", command=save_result, state=tk.DISABLED, bg="blue")
save_button.grid(row=0, column=1, padx=5)

# Result label
result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
