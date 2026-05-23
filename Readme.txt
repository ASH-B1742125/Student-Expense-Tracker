# 📊 Expense Tracker with Excel Charts (Python + OpenPyXL)

## 📌 Overview
This project reads expense data from an Excel file (`student_expense.xlsx`), calculates total spending and category-wise breakdown, and generates a new Excel report (`expense_report.xlsx`) with visual charts.

It uses **OpenPyXL** to automate Excel processing and create charts directly inside the spreadsheet.

---

## ⚙️ Features
- Reads expense data from Excel
- Calculates:
  - Total expenses
  - Category-wise spending
- Generates a structured Excel report
- Adds visual charts:
  - 🥧 Pie Chart → Expense distribution (percentage-wise)
  - 📊 Bar Chart → Category comparison

---

## 📁 Input File Format (`student_expense.xlsx`)

| Date | Item | Category | Amount |
|------|------|----------|--------|
| 01-01-2025 | Lunch | Food | 200 |
| 02-01-2025 | Bus | Travel | 50 |

---

## 📤 Output File (`expense_report.xlsx`)

The generated report includes:
- Total spent summary
- Category-wise breakdown table
- Pie chart (expense distribution)
- Bar chart (category comparison)

---

## 📦 Installation

Install required dependency:

```bash
pip install openpyxl