# Attendance-Project 
**Developed by Ahmed Abduljalil** ([GitHub: 0maido](https://github.com/0maido))  
**Last Updated**: February 2025  

A suite of tools for managing student attendance data with Excel integration.

---

## 📑 Table of Contents
- [Projects Overview](#-projects-overview)
- [Features](#-features)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Usage](#-usage)
- [Contributing](#-contributing)
- [License](#-license)

---

## 📂 Projects Overview

### 1. Weekly Report App (Attendance Aggregator)
**Date Created**: 2025-02-28  
A GUI application that:
- Processes multiple attendance Excel files
- Aggregates weekly absences and leaves
- Generates comprehensive reports with visual formatting

### 2. Attendance Processor
**Date Created**: 2025-02-25  
A processing tool that:
- Compares main student lists with attendance files
- Marks present/absent students with color coding
- Generates detailed attendance statistics

---

## 🌟 Features

### Weekly Report App
- Multi-file loading for different weekdays
- Automatic data validation and cleaning
- Intelligent merging of student records
- Custom Excel export with:
  - Color-coded headers (red for absences, green for leaves)
  - Multi-day summary statistics
  - Professional formatting with adjustable column widths
- Real-time data preview table
- Error handling with detailed tracebacks

### Attendance Processor
- Dual-file comparison system
- Visual feedback with color coding:
  - Green: Present students
  - Red: Absent students
- Range-specific processing (custom row/column selection)
- Statistical reporting:
  - Total students processed
  - Present/Absent counts
  - Data mismatch detection
- Temporary file handling with auto-cleanup

---

## 📋 Prerequisites

- Python 3.8+
- Microsoft Excel or compatible spreadsheet software
- Excel files structured with:
  - Header rows (first 2 rows ignored)
  - Student IDs in consistent columns
  - Attendance marked with 'A' (Absent) and 'L' (Leave)

---

## ⚙️ Installation

1. **Clone repository** (if available):
   ```bash
   git clone https://github.com/0maido/Attendance-Project.git
   cd Attendance-Project
   ```

2. **Install required packages**:
   ```bash
   pip install -r requirements.txt
   ```

---

## 🚀 Usage

### Weekly Report App
1. **Launch the application**:
   ```bash
   python weekly_report.py
   ```

2. **Load attendance files**:
   - Click weekday buttons (Monday-Thursday)
   - Select multiple Excel files for each day

3. **View results**:
   - Table shows students with absences/leaves
   - Totals calculated automatically

4. **Export report**:
   - Click "Export Report"
   - Choose save location (XLSX format)


### Attendance Processor
1. **Run the processor**:
   ```bash
   python attendance_processor.py
   ```

2. **Select files**:
   - Main File: Complete student list
   - Attendance File: Daily records

3. **Set processing ranges**:
   ```ini
   Attendance File Rows: 3-38 (typical)
   Main File Rows: 3-38 (typical)
   Columns: C (Attendance), F (Main)
   ```

4. **Process and export**:
   - Click "Process Attendance"
   - Review statistics
   - Export final report


---

## 🛠 Technical Specifications

| Component              | Weekly Report       | Attendance Processor |
|------------------------|---------------------|----------------------|
| Core Library           | pandas              | openpyxl             |
| GUI Framework          | Tkinter             | Tkinter              |
| Excel Engine           | openpyxl/xlsxwriter | openpyxl             |
| Data Validation        | Automatic           | Range-based          |
| Output Format          | formatted XLSX      | XLSX with stats      |

---

## 🤝 Contributing

1. Fork the repository
2. Add tests for new features
3. Submit pull request


---

## 📬 Contact
For support or feature requests:
- 📧 Email: [Your Email Here]
- 💬 GitHub Issues: [Create New Issue](https://github.com/0maido/Attendance-Project/issues)
```

**To Complete**:
1. Add your contact email
2. Replace placeholder image URLs with actual screenshots
3. Add system-specific notes if needed
4. Include sample Excel files in `/examples` folder

