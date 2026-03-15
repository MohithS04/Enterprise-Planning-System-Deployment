#!/bin/bash

echo "=========================================================="
echo "🚀 Deploying Enterprise Planning System Modules..."
echo "=========================================================="

cd "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment"

echo ""
echo "📦 Module 1: Custom Planning Templates"
python3 scripts/generate_budget_template.py
python3 scripts/generate_rolling_forecast.py
python3 scripts/generate_dept_submission.py
python3 scripts/generate_consolidation.py

echo ""
echo "🔍 Module 2: Data Validation & GL Reconciliation Framework"
python3 scripts/generate_gl_recon.py
python3 scripts/generate_validation_checklist.py

echo ""
echo "📚 Module 3: User Support Guide & Training Materials"
echo "Markdown documents (User Support Guide, Quick Reference, Training Tracks) are already staged in 03_Training_and_Support/"

echo ""
echo "📊 Module 4: Ad-Hoc Financial Reporting Suite"
python3 scripts/generate_monthly_bva.py
python3 scripts/generate_retail_impact.py
python3 scripts/generate_change_log.py
python3 scripts/generate_flash_report.py

echo ""
echo "✅ Deployment Complete! All files are generated in their respective folders."
echo "=========================================================="
