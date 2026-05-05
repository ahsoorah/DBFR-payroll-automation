# Automated Payroll Report Formatter

Project Overview
---
This project is a Python-based automation utility developed to streamline the transformation of raw workforce management data into standardized, audit-ready financial reports.

Originally developed during a Fire Technology Internship with the City of Delray Beach Fire-Rescue Department, this tool was designed to bridge the gap between automated labor exports and the precise formatting requirements of municipal finance departments.

Performance Impact
---
Efficiency: Automates a complex manual formatting process, saving several hours of administrative labor every bi-weekly pay period cycle.

Accuracy: Eliminates human error in data mapping, type-casting, and row filtering, ensuring 100% consistency across reporting periods.

Technical Features
---
Hybrid Ingestion Engine: Supports both .xlsx and .csv source files. If a CSV is detected, the script dynamically builds a memory-resident workbook and applies automated type-casting for numerical analysis.

Intelligent Data Cleaning: Implements logic-based row filtering to identify and remove non-billable entries and administrative placeholders without disrupting the index integrity of the sheet.

Dynamic Header Generation: Scans the dataset to extract the earliest and latest timestamps, automatically generating pay period date ranges for the final report header.

XML Patching: Utilizes the zipfile library to perform a raw XML injection into the .xlsx structure. This allows the modification of attributes (such as x14ac:dyDescent) that are not exposed by high-level APIs like OpenPyXL, ensuring the output meets exact legacy layout standards.

Modular GUI: Features a Tkinter-based interface for streamlined user interaction, metadata collection, and error handling.

Privacy & Sanitization Notice
---
This is a sanitized proof-of-concept version of the original production script. To comply with municipal privacy standards and organizational security protocols, the following modifications have been made:
---
All department-specific logic and internal status codes have been generalized.

Hardcoded organizational references + branding have been removed.

Formatting standards have been adjusted to a generic professional template.

No actual personnel data is included or accessible through this repository.

Tech Stack
---
Language: Python 3.x

Libraries: * openpyxl: For advanced Excel manipulation and styling.

tkinter: For the graphical user interface.

zipfile & os: For low-level file system and XML operations.

How to Use
---
Install dependencies:
pip install openpyxl

Run the script:
---
python payroll_formatter.py

Process: Follow the GUI prompts to select your source file, enter the reporting unit, and save the formatted output.

Author
Suriyah Saravanan
Bachelor's in Management Information Systems (MIS), Cybersecurity
Florida Atlantic University
