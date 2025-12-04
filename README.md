
# Outcomes Report Automator

The **Outcomes Report Automator** is a Python-based application designed to automate the creation of academic Outcomes Reports.
It reads structured Excel files, validates the data, and generates polished PDF reports with consistent formatting.
This tool removes repetitive manual work and ensures accuracy for staff producing student outcome documentation.

---

## **📌 System Architecture Overview**

### **1. High-Level Workflow Diagram**

<img src="https://github.com/user-attachments/assets/ad099c54-0a50-40b0-bf86-3291c0d8f1b7" width="800">

This diagram gives an overview of the system’s end-to-end workflow.
It shows how input files move through validation, processing, and PDF generation.
Each stage highlights how the automation reduces manual tasks.
Useful for understanding the overall operation of the project.

---

### **2. Component Breakdown Diagram**

<img src="https://github.com/user-attachments/assets/bed4c418-d80f-4cba-9dba-aef38f2322e9" width="800">

This diagram explains the core internal components of the system.
It shows how input handling, validation modules, and report generation are separated.
Helps clarify responsibilities of each part of the codebase.
Useful when maintaining or extending the application.

---

### **3. Detailed Process Flowchart**

<img src="https://github.com/user-attachments/assets/1cd8c239-d253-4c26-a974-6219b26c37bd" width="800">

A step-by-step flowchart describing the processing sequence.
Follows the user journey from uploading the Excel file to downloading the final PDF.
Shows how functions call each other across the workflow.
Helpful for debugging or onboarding new developers.

---

### **4. Class Structure Diagram**

<img src="https://github.com/user-attachments/assets/f480e0fd-68c7-4a15-9c8f-c9043177797c" width="700">

A structural class diagram showing the main classes and relationships.
Illustrates how different components interact at object level.
Demonstrates a modular and maintainable architecture.
Useful when modifying or adding new features.

---

## **📊 Input Template & Validation**

### **5. Required Excel Template Format**

<img src="https://github.com/user-attachments/assets/6965a9bb-2a35-48c3-ba15-7a29ee12817d" width="700">

Shows the exact column layout expected by the system.
Ensures users prepare Excel files in the correct structure.
Prevents import errors caused by missing or renamed headers.
Essential for anyone supplying the input data.

---

### **6. Example of Validation / Error Highlighting**

<img src="https://github.com/user-attachments/assets/1a51678c-1f01-4b4f-bd44-719171ebe92b" width="600">

Displays how invalid entries are flagged during validation.
Incorrect or missing values are highlighted for quick correction.
This prevents faulty data from reaching the report stage.
Improves accuracy and reduces manual checks.

---

## **📂 Application Interface & Features**

### **7. Main Application Dashboard (Home Screen)**

<img src="https://github.com/user-attachments/assets/43b5d504-3dc3-4cc7-a087-56c66db4db8b" width="800">

Shows the main dashboard where users interact with the tool.
Options include uploading data, generating reports, and accessing help.
Designed for simplicity so staff can use it without technical knowledge.
Acts as the starting point for the whole workflow.

---

### **8. File Upload & Validation Interface**

<img src="https://github.com/user-attachments/assets/2330a2c4-733d-4a96-81d5-9eece81f68e1" width="800">

This interface confirms the uploaded file and displays validation results.
Shows success messages or any issues that need correcting.
Ensures clean data before generating PDFs.
Prevents incomplete or inaccurate output.

---

## **📄 Generated Report Output**

### **9. Generated PDF Report – Header Section**

<img src="https://github.com/user-attachments/assets/024fd9a4-7993-47e2-a830-1c898ad7aa20" width="800">

Shows the top of a generated PDF, including student identifiers.
Provides a consistent and professional layout.
Ensures reports follow the expected academic format.
Useful for confirming correct header rendering.

---

### **10. Generated PDF – Student Details Section**

<img src="https://github.com/user-attachments/assets/4c25ad10-6cec-4911-a35e-4701545b2cf3" width="800">

Shows the middle portion of the report with detailed student fields.
Includes demographic and programme information.
Designed to present data clearly and professionally.
Matches the format required by academic staff.

---

### **11. Generated PDF – Outcome Grades Section**

<img src="https://github.com/user-attachments/assets/83f7fbcf-005c-4387-b284-f4064252fda1" width="800">

Displays the outcomes or grades awarded to the student.
Clear tables help staff interpret performance quickly.
Ensures consistency across large batches of reports.
Supports multi-page outputs when needed.

---

### **12. Generated PDF – Additional Notes / Modules Section**

<img src="https://github.com/user-attachments/assets/da1837b6-6fff-4cf5-8345-622e766e632d" width="800">

Shows extended data such as module-specific outcomes.
Structured layout helps provide a comprehensive overview.
Provides space for additional academic notes.
Ensures all required fields appear in the final PDF.

---

### **13. Generated PDF – Final Page / Summary**

<img src="https://github.com/user-attachments/assets/8fe309f1-409e-46b7-b8b0-97e773572f96" width="800">

Displays the final portion of the report.
Includes closing details or additional summary fields.
Ensures the document ends neatly without layout issues.
Completes the full report structure.

---

## **🛠 Supporting Files**

### **14. Configuration File Example**

<img src="https://github.com/user-attachments/assets/36489e4b-51ba-44c2-bae2-2122f0d5444d" width="700">

Shows the configuration options used by the application.
Includes file paths, templates, and naming settings.
Allows the system to be customised for different environments.
Useful when deploying to new departments.

---

### **15. Output Folder Structure Example**

<img src="https://github.com/user-attachments/assets/9b077678-a6bb-445b-8cdf-8dec04cb8f11" width="700">

Shows how generated PDFs and logs are stored.
Provides a clean and organised folder layout.
Makes it easier for staff to locate individual reports.
Supports large batches of automated report generation.


