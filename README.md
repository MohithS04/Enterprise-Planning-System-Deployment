# Enterprise Planning System Deployment

Building a complete, production-ready Enterprise Performance Management (EPM) infrastructure entirely via Python automation. This repository contains the architecture, data validation logic, and generation scripts necessary to deploy standardized corporate FP&A templates and reporting dashboards.

## 🎯 Project Overview

This deployment package replaces manual, localized Excel reporting with a governed, centralized model. It constructs 10 interlocking deliverables designed for a multi-entity corporate environment:

*   **Standardized Data Capture:** Headcount, CapEx, OpEx, and Revenue forecasting logic.
*   **System Integrity:** Automated GL reconciliation control sheets and strict numeric data validation rules.
*   **Ad-Hoc Reporting:** Executive flash metrics, variance heat-mapping, and bridge charts.
*   **User Documentation:** Markdown-ready navigational guides and role-based training curriculums.

## 📁 Repository Structure

```
├── 01_Planning_Templates/
│   ├── 1.1_Annual_Budget_Template.xlsx
│   ├── 1.2_Rolling_Forecast_Template.xlsx
│   ├── 1.3_Departmental_Submission_Template.xlsx
│   └── 1.4_Consolidation_Template.xlsx
├── 02_Validation_and_Reconciliation/
│   ├── 2.1_GL_Reconciliation_Control_Sheet.xlsx
│   └── 2.2_Validation_and_Close_Checklist.xlsx
├── 03_Training_and_Support/
│   ├── 3.1_User_Support_Guide.md
│   ├── 3.2_Quick_Reference_Card.md
│   └── 3.3_Training_Tracks_and_Quizzes.md
├── 04_Reporting_Suite/
│   ├── 4.1_Monthly_BvA_Report.xlsx
│   ├── 4.2_Retail_Operational_Impact.xlsx
│   ├── 4.3_Monthly_Budget_Change_Log.xlsx
│   └── 4.4_Executive_Flash_Report.xlsx
├── scripts/
│   ├── generate_budget_template.py
│   ├── generate_change_log.py
│   ├── generate_consolidation.py
│   ├── generate_dept_submission.py
│   ├── generate_flash_report.py
│   ├── generate_gl_recon.py
│   ├── generate_monthly_bva.py
│   ├── generate_retail_impact.py
│   ├── generate_rolling_forecast.py
│   ├── generate_validation_checklist.py
│   └── requirements.txt
└── deploy_system.sh
```

## 🚀 Installation & Deployment

The entire suite of FP&A templates and reporting dashboards is generated via Python using the `xlsxwriter` library to ensure perfect cell locking, data validation dropdowns, and conditional formatting.

### Prerequisites
*   Python 3.8+
*   `pip` package manager

### Setup

1.  **Clone the Repository**
    ```bash
    git clone https://github.com/your-username/enterprise-planning-system.git
    cd enterprise-planning-system
    ```

2.  **Install Dependencies**
    ```bash
    pip install -r scripts/requirements.txt
    ```

3.  **Execute the Deployment Script**
    Run the master script to compile all 10 deliverables into their designated folders.
    ```bash
    chmod +x deploy_system.sh
    ./deploy_system.sh
    ```

## 📊 Modules & Capabilities

### Module 1: Custom Planning Templates
Automated generation of locked-down Excel interfaces for business users to input assumptions.
*   **Driver-Based Forecasting:** Automatic conversion of headcount entering variables into fully-loaded payroll calculations.
*   **Dynamic Data Validation:** Restrict inputs using lookup lists (GL Account + Cost Center mappings) and numerical checks (blocking negative expense entries).
*   **Consolidation Engine:** Intercompany elimination logging and dynamic FX translation matrices.

### Module 2: Data Validation & GL Reconciliation
Bridging the gap between the Planning System and the ERP (Oracle, SAP, NetSuite).
*   **Automated Tie-Outs:** Excel logic checking source GL balances against Planning balances, flagging discrepancies >$500.
*   **Integrity Rules:** Sign convention standardization (Revenue=Credit, Expense=Debit) and mathematical accuracy blocks.

### Module 3: Ad-Hoc Financial Reporting Suite
Executive-ready financial statements and visual analytic dashboards.
*   **Monthly Budget vs. Actual:** P&L outputs featuring 3-color variance heatmaps.
*   **Retail Impact Bridge:** A pseudo-waterfall analysis isolating the financial impact of *Volume vs. Price* shifting alongside qualitative risk callouts.
*   **Flash Reporting:** C-Suite 1-pagers featuring trendline charts, R/Y/G Traffic Light KPIs, and top decision logs.

## 🤝 Contributing

Contributions are welcome! If you'd like to extend the Python generators to include new reporting templates or advanced API integrations for the ERP data feed, please open an issue or submit a pull request.

## 📄 License
This project is licensed under the MIT License - see the LICENSE file for details.
