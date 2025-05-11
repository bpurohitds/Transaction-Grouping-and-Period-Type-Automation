# Transaction-Grouping-and-Period-Type-Automation
Automates the financial reconciliation and tagging process for deferred revenue entries. Consolidates data, classifies transactions, and determines the period type (Annual, TwoYear, ThreeYear) for each customer, improving accuracy, efficiency, and reducing manual effort for finance teams
Automated Financial Reconciliation & Tagging Process

\[Transaction Grouping and Period Type Classification\]

📌 **Objective**

This project automates a unique and highly specific reconciliation and tagging workflow used by the finance (BPnA) team. It replaces a time-consuming SOP-driven process with a clean, efficient, Python-based pipeline. Key tasks include:

- Consolidating Excel data from different departments
- Classifying transactions (Credit Note, Adjustment Note, Other)
- Mapping financial amounts and tagging logic
- Determining each customer's contract Period Type (Annual, TwoYear, ThreeYear, etc.)

**⚙️ Key Features**

• **End-to-End Automation**: Replaces 11+ hours of manual work with <20 minutes of Python execution

**• Tagging Logic**: Automates classification of transaction types using rule-based logic

• **Pivot Table Generation**: Summarizes relevant metrics and tagging breakdowns

**• Amount Mapping**: Maps CN/ADJN amounts to relevant entries, excluding GST

• **Period Type Assignment**: Classifies customers into contract types using service & transaction attributes

**🧠 Tools & Technologies**

- Python (Pandas, OpenPyXL, Pyxlsb)
- Jupyter Notebook
- Excel (.xlsx, .xlsb)
- Windows File System
- AI tools

**🗂 Repository Structure**

**Transaction-Grouping-Period-Type-Automation/**  
├── Part1_Reconciliation_Tagging.py # CN/ADJN tagging & mapping logic  
├── Part2_Period_Type_Classification.py # Customer period type classification  
├── sample_data/ # Optional dummy files (if shared)  
├── README.md # Project summary and structure  
└── LICENSE # No license (internal project only)

**💥 Impact**

- ⏱ **Saved Time**: From **11 hours → under 20 minutes**
- 🔍 **Improved Accuracy**: Eliminated human error from tagging/classification
- 🧾 **Audit-Ready Data**: Final output is structured for reporting & audit use
- 📈 **Boosted Productivity**
- : Finance team can focus on insights, not cleanup

**⚠️ Limitations**

• Cannot share actual data or exact column names (confidential)

• Script is tailored for internal data structure—may require adjustments for external/general use

⚠️ **Note**: This is an internal automation project not intended for external reuse. Shared here for portfolio demonstration only—code may not work without adaptation to your data structure.
