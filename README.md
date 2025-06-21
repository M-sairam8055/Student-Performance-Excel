# 🎓 Student Performance Report Card & Dashboard System

This is a **dynamic Excel-based Student Report Card System** created by **M SAI RAM**, designed to automate student performance tracking and reporting.

It includes:
- Report card generation using dropdown-based selection
- Subject-wise marks analysis
- Grade calculation, PDF export with one click (via VBA)
- Interactive dashboard with charts and performance insights

---

## 📂 Project Structure

| File Name | Description |
|-----------|-------------|
| `Student_Performance_Project.xlsm` | Main Excel workbook with macros, report card, and dashboard |
| `ReportCard_Sample.pdf` | Exported sample PDF report card |
| `README.md` | This documentation file |

---

## ✅ Features

### 📋 Report Card System
- Select a **Student ID** from dropdown
- Auto-fetch:
  - Student Name
  - Class
  - Subject-wise marks
  - Total marks
  - Average marks
  - Grade based on average
- Professional header & footer with your branding
- Export to **PDF** using a single-click VBA button

### 📊 Dashboard
- **Bar Chart**: Top 5 performers based on total marks
- **Pie Chart**: Grade-wise distribution
- **Line Chart**: Subject-wise performance (student-level)
- Class-wise color coding
- Freeze panes for easy viewing

---

## 🧠 How It Works

### ➤ Formulas Used
- `VLOOKUP` / `INDEX-MATCH` for dynamic data fetch
- `IF`, `COUNTIF`, `AVERAGE`, `SUM` for calculations
- `Conditional Formatting` for class color coding

### ➤ VBA for PDF Export
```vba
Sub ExportReportCardToPDF()
    Dim ws As Worksheet
    Dim FilePath As String
    Dim StudentID As String

    Set ws = ThisWorkbook.Sheets("Report_Card")
    StudentID = ws.Range("B2").Value
    FilePath = Application.DefaultFilePath & "\ReportCard_" & StudentID & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=FilePath, Quality:=xlQualityStandard
    MsgBox "PDF exported successfully for " & StudentID
End Sub
