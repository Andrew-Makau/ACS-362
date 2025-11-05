# Cash Flow Forecasting Deliverables (Group 2)

This folder contains scripts and a workflow to generate a Word report and PowerPoint deck with embedded charts for:
- Course: ACS 362A: Database Systems Group Project
- Instructor: Japheth Mursi, Ph.D.
- Team: 22-2207 – Makau Andrew Nzyoki; 22-2203 – Wanjihia Serena Wairati; 22-2245 – Dorothy Vugutsa; 22-2337 – Lewis Kagia; 23-2189 – Kitavi Christoph
- Context: Nairobi electronics retailer
- Horizon: 24 months
- Currency: KES
- Opening cash: KES 2,000,000
- CapEx: KES 500,000 in Month 3; KES 1,200,000 in Month 12

## What gets generated
- deliverables/Cash_Flow_Forecasting_Report_Group2.docx
- deliverables/Cash_Flow_Forecasting_Presentation_Group2.pptx
- deliverables/data/cash_series.csv
- deliverables/charts/*.png (embedded in both files)

## Run locally (optional)

```powershell
# Create and activate a virtual environment (Windows PowerShell)
python -m venv .venv
. .venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt

# Generate the report, deck, charts, and CSV
python scripts/generate_docs.py
```

Optional (macOS/Linux):

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scripts/generate_docs.py
```

## Run via GitHub Actions

1) Ensure Settings → Actions → General → Workflow permissions is “Read and write permissions.”
2) Commit this README, the workflow file, and requirements.txt to main.
3) Trigger the workflow:
   - Option A: Push any of the files above to main (auto-runs).
   - Option B: Actions tab → “Build and Commit Docs” → “Run workflow” → Branch: main → Run.
