# 💊 Pharmacy Data Processing Application

This is a **standalone Flask-based desktop application** that processes pharmacy data by combining **BestRx insurance logs, vendor shipments, and conversion/master files** into a single comprehensive Excel report.  
It provides detailed insights into **purchased vs. billed quantities, package size differences, and financial comparisons**, while also generating helper sheets for compliance and auditing.

---

## 📋 Features
- **Multi-file Uploads**  
  Upload BestRx insurance logs, Kinray/vendor shipments, and conversion/master files.
  
- **Automated Data Processing**  
  Cleans, merges, and aggregates pharmacy data with accurate package size and billing calculations.

- **Comprehensive Excel Report**  
  Generates a detailed Excel file with:  
  - Processed Data (billed vs. purchased)  
  - **Needs to be Ordered** (negative package differences)  
  - **Do Not Order** (positive package differences)  
  - **Never Ordered – Check** (items purchased = 0)  

- **Smart Formatting**  
  Excel output includes merged headers, autosums, colored highlights for discrepancies, and page setup for printing.

- **Standalone GUI**  
  Runs locally using **Flask + pywebview**, packaged as a desktop-style app with no external server required.

---

## 🛠️ Tech Stack
- **Backend**: Flask (Python 3.x)  
- **Data Processing**: Pandas, OpenPyXL  
- **Frontend / GUI**: HTML templates + pywebview  
- **Others**: Tkinter (file handling), FlaskWebGUI  
