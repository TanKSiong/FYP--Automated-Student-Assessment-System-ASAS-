# FYP--Automated-Student-Assessment-System-ASAS

## 📌 Overview

The Automated Student Assessment System (ASAS) is a Python-based application developed to automate the student grading process. 

This system eliminates manual calculations by processing assessment data from Excel files and generating accurate results, including total marks, percentages, grades, and pass/fail status.

It is designed to improve efficiency, reduce human error, and provide consistent academic assessment results.

---

## 🚀 Features

- Upload and read student assessment data from Excel files
- Automatically calculate total marks
- Generate percentage scores for each student
- Determine pass or fail status
- Assign grades based on calculated results
- Generate summary reports
- Export processed results into a new Excel file
- Simple graphical user interface (GUI) for user interaction

---

## 🛠 Tech Stack

- Language: Python
- Libraries:
  - Pandas (data processing)
  - OpenPyXL (Excel file handling)
- Interface: Python GUI (e.g., Tkinter)

---

## 📂 System Workflow

1. User uploads an Excel file containing student marks  
2. System reads and validates the data  
3. System performs:
   - Total mark calculation  
   - Percentage calculation  
   - Grade assignment  
   - Pass/Fail determination  
4. System generates a summary report  
5. Processed results are exported as a new Excel file  

---

## ▶️ How to Run

1. Install required libraries:

```bash
pip install pandas openpyxl
```

2. Run the application:

```bash
python ASAS.py
```
3. Upload the required Excel file when prompted

---

📊 **Input Requirements**
Excel file must follow a structured format
Must contain:
Student identifiers
Assessment components (e.g., assignment, exam)
Valid numerical marks

---

**System Capabilities**
- Handles large datasets efficiently
- Reduces manual workload for lecturers
- Ensures accurate and consistent grading
- Provides quick result generation

---

⚠️**Limitations**
- No database integration
- Single-user operation
- Requires properly formatted Excel input
