#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib
# Use a non-interactive backend to avoid GUI/display issues in headless environments
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Inches, Pt
from pptx import Presentation
from pptx.util import Inches as PPTInches

# =========================
# Configuration (edit here)
# =========================
COURSE = "ACS 362A: Database Systems Group Project"
INSTRUCTOR = "Japheth Mursi, Ph.D."
TEAM = [
    "22-2207 – Makau Andrew Nzyoki",
    "22-2203 – Wanjihia Serena Wairati",
    "22-2245 – Dorothy Vugutsa",
    "22-2337 – Lewis Kagia",
    "23-2189 – Kitavi Christoph",
]
CONTEXT = "Nairobi-based consumer electronics retailer (store + online)"
CURRENCY = "KES"
HORIZON_MONTHS = 24

# Start month: first day of current month
START_MONTH = datetime(datetime.today().year, datetime.today().month, 1)

# Opening balances and CapEx schedule
OPENING_CASH = 2_000_000
CAPEX_SCHEDULE = {3: 500_000, 12: 1_200_000}  # month_number: amount

# Assumptions
COGS_PCT = 0.40
FIXED_OPEX = 1_800_000
VAR_OPEX_PCT = 0.08
DEPR_LIFE_MONTHS = 60
TAX_RATE = 0.30

# Scenarios
SCENARIOS = {
    "Base":     {"growth": 0.015, "dso": 45, "dpo": 30, "inv_days": 60},
    "Upside":   {"growth": 0.030, "dso": 35, "dpo": 35, "inv_days": 50},
    "Downside": {"growth": 0.005, "dso": 55, "dpo": 25, "inv_days": 70},
}

# Starting monthly revenue (Month 1)
START_REV = 5_000_000

# Output paths
ROOT = Path(__file__).resolve().parents[1]
DELIVERABLES = ROOT / "deliverables"
CHARTS_DIR = DELIVERABLES / "charts"
DATA_DIR = DELIVERABLES / "data"
DELIVERABLES.mkdir(parents=True, exist_ok=True)
CHARTS_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)

REPORT_PATH = DELIVERABLES / "Cash_Flow_Forecasting_Report_Group2.docx"
DECK_PATH = DELIVERABLES / "Cash_Flow_Forecasting_Presentation_Group2.pptx"
CSV_PATH = DATA_DIR / "cash_series.csv"


def build_schedule(growth: float, dso: int, dpo: int, inv_days: int) -> pd.DataFrame:
    months = pd.date_range(START_MONTH, periods=HORIZON_MONTHS, freq="MS")
    df = pd.DataFrame({"month": months})
    # Revenue
    rev = [START_REV]
    for _ in range(1, HORIZON_MONTHS):
        rev.append(rev[-1] * (1 + growth))
    df["revenue"] = rev
    # COGS, Gross Margin
    df["cogs"] = df["revenue"] * COGS_PCT
    df["gross_profit"] = df["revenue"] - df["cogs"]
    # Opex
    df["opex"] = FIXED_OPEX + VAR_OPEX_PCT * df["revenue"]
    df["ebitda"] = df["gross_profit"] - df["opex"]
    # CapEx and Depreciation
    df["capex"] = 0.0
    for m, amt in CAPEX_SCHEDULE.items():
        if 1 <= m <= HORIZON_MONTHS:
            df.loc[m - 1, "capex"] = amt
    # Depreciation: straight-line for each CapEx
    dep = np.zeros(HORIZON_MONTHS)
    for m, amt in CAPEX_SCHEDULE.items():
        start_idx = m - 1
        life = min(DEPR_LIFE_MONTHS, HORIZON_MONTHS - start_idx)
        if life > 0:
            dep[start_idx:start_idx + life] += amt / DEPR_LIFE_MONTHS
    df["depreciation"] = dep
    df["ebit"] = df["ebitda"] - df["depreciation"]
    # Taxes (only on positive EBIT for simplicity)
    df["tax"] = df["ebit"].clip(lower=0) * TAX_RATE
    df["net_income"] = df["ebit"] - df["tax"]
    # Working capital (monthly approximation)
    df["ar"] = df["revenue"] * (dso / 30.0)
    df["ap"] = df["cogs"] * (dpo / 30.0)
    df["inventory"] = df["cogs"] * (inv_days / 30.0)
    df["delta_ar"] = df["ar"].diff().fillna(df["ar"])
    df["delta_ap"] = df["ap"].diff().fillna(df["ap"])
    df["delta_inv"] = df["inventory"].diff().fillna(df["inventory"])
    df["delta_wc"] = df["delta_ar"] + df["delta_inv"] - df["delta_ap"]
    # Cash flow
    df["cfo"] = df["net_income"] + df["depreciation"] - df["delta_wc"]
    df["net_cash_flow"] = df["cfo"] - df["capex"]
    # Cash balance
    cash = [OPENING_CASH + df.loc[0, "net_cash_flow"]]
    for i in range(1, HORIZON_MONTHS):
        cash.append(cash[-1] + df.loc[i, "net_cash_flow"])
    df["closing_cash"] = cash
    return df


def build_all() -> dict[str, pd.DataFrame]:
    return {name: build_schedule(**cfg) for name, cfg in SCENARIOS.items()}


def save_data(dfs: dict[str, pd.DataFrame]) -> None:
    out = pd.DataFrame({"month": dfs["Base"]["month"]})
    for name in ["Base", "Upside", "Downside"]:
        out[f"{name}_closing_cash"] = dfs[name]["closing_cash"].round(2)
    out["Base_net_cash_flow"] = dfs["Base"]["net_cash_flow"].round(2)
    out.to_csv(CSV_PATH, index=False)


def chart_closing_cash_base(df: pd.DataFrame) -> Path:
    plt.figure(figsize=(10, 5))
    plt.plot(df["month"], df["closing_cash"], label="Closing Cash (Base)", color="#1f77b4", linewidth=2)
    plt.title("Closing Cash — Base Case")
    plt.xlabel("Month")
    plt.ylabel(f"Cash ({CURRENCY})")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    p = CHARTS_DIR / "closing_cash_base.png"
    plt.savefig(p, dpi=200)
    plt.close()
    return p


def chart_scenarios(dfs: dict[str, pd.DataFrame]) -> Path:
    plt.figure(figsize=(10, 5))
    for name, color in zip(["Base", "Upside", "Downside"], ["#1f77b4", "#2ca02c", "#d62728"]):
        plt.plot(dfs[name]["month"], dfs[name]["closing_cash"], label=name, linewidth=2, color=color)
    plt.title("Closing Cash — Base vs Upside vs Downside")
    plt.xlabel("Month")
    plt.ylabel(f"Cash ({CURRENCY})")
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    p = CHARTS_DIR / "closing_cash_scenarios.png"
    plt.savefig(p, dpi=200)
    plt.close()
    return p


def chart_wc_components(df: pd.DataFrame, month_idx: int = 11) -> Path:
    # ΔAR, ΔInventory, −ΔAP for Month 12 (index 11) by default
    m = min(max(month_idx, 0), len(df) - 1)
    components = {"ΔAR": df.loc[m, "delta_ar"], "ΔInventory": df.loc[m, "delta_inv"], "−ΔAP": -df.loc[m, "delta_ap"]}
    plt.figure(figsize=(8, 5))
    bars = plt.bar(list(components.keys()), list(components.values()), color=["#9467bd", "#8c564b", "#ff7f0e"])
    plt.title(f"Working Capital Components — Month {m+1}")
    plt.ylabel(f"Cash impact ({CURRENCY})")
    for b in bars:
        y = b.get_height()
        plt.text(b.get_x() + b.get_width()/2, y, f"{y:,.0f}", ha="center", va="bottom", fontsize=9)
    plt.tight_layout()
    p = CHARTS_DIR / "wc_components.png"
    plt.savefig(p, dpi=200)
    plt.close()
    return p


def build_report(paths: dict[str, Path]) -> None:
    doc = Document()
    doc.add_heading("Cash Flow Forecasting Model — Group 2", level=0)
    doc.add_paragraph(f"Course: {COURSE}")
    doc.add_paragraph(f"Instructor: {INSTRUCTOR}")
    doc.add_paragraph("Prepared by: Group 2 — " + "; ".join(TEAM))
    doc.add_paragraph(f"Date: {datetime.today():%B %d, %Y} • Version: v1.0")

    doc.add_heading("Executive Summary", level=1)
    doc.add_paragraph(
        "Objective: Build a transparent monthly cash flow forecasting model to project cash position, identify funding needs, and evaluate scenarios."
    )
    doc.add_paragraph(
        "Scope: 24‑month forecast in KES with Base / Upside / Downside scenarios; working capital (AR/AP/Inventory), CapEx/Depreciation, and optional debt."
    )
    doc.add_paragraph(f"Case Context: {CONTEXT}")

    doc.add_heading("Key Charts", level=1)
    doc.add_paragraph("Closing cash — Base case:")
    doc.add_picture(str(paths["base_line"]), width=Inches(6))
    doc.add_paragraph("Closing cash — Base vs Upside vs Downside:")
    doc.add_picture(str(paths["scenarios_line"]), width=Inches(6))
    doc.add_paragraph("Working capital components (ΔAR, ΔInventory, −ΔAP):")
    doc.add_picture(str(paths["wc_components"]), width=Inches(6))

    doc.add_heading("Methodology (summary)", level=1)
    doc.add_paragraph(
        "Revenue via monthly growth; COGS as % of revenue; Opex with fixed and variable components; working capital using DSO/DPO/InventoryDays; CapEx and straight‑line depreciation; simple taxes; cash flow = Net Income + Depreciation − ΔWC − CapEx."
    )

    doc.add_heading("Scenario Assumptions", level=1)
    for name, cfg in SCENARIOS.items():
        doc.add_paragraph(f"{name}: growth {cfg['growth']*100:.1f}%/mo, DSO {cfg['dso']}d, DPO {cfg['dpo']}d, Inventory {cfg['inv_days']}d")

    doc.add_heading("Notes & Recommendations", level=1)
    doc.add_paragraph(
        "Focus on reducing DSO, negotiating DPO, and phasing CapEx to smooth the cash trough. Update actuals monthly and refresh scenarios quarterly."
    )

    doc.save(str(REPORT_PATH))


def build_deck(paths: dict[str, Path]) -> None:
    prs = Presentation()
    # Title slide
    title_slide = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide)
    slide.shapes.title.text = "Cash Flow Forecasting Model (Group 2)"
    slide.placeholders[1].text = f"{COURSE}\nInstructor: {INSTRUCTOR}\nTeam: " + "; ".join(TEAM)

    def add_slide(title, bullets, notes=None):
        layout = prs.slide_layouts[1]
        s = prs.slides.add_slide(layout)
        s.shapes.title.text = title
        tf = s.shapes.placeholders[1].text_frame
        tf.clear()
        for i, b in enumerate(bullets):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = b
        if notes:
            s.notes_slide.notes_text_frame.text = notes

    add_slide("Agenda", ["Objective & Scope", "Approach & Assumptions", "Model Overview", "Results & Scenarios", "Sensitivities, Risks", "Recommendations"], "Keep to ~10–12 minutes; Q&A at the end.")
    add_slide("Objective & Scope", ["24‑month cash forecast in KES", "Revenue, margins, opex, working capital, CapEx", "Base / Upside / Downside scenarios"], "Why 24 months: shows seasonality; concise for class.")
    add_slide("Approach (Model Architecture)", ["Driver‑based; centralized Assumptions", "Schedules: Revenue, Opex, Working Capital, CapEx/Dep, Debt", "Consolidated Cash Flow; scenario switcher"], "No hardcoded numbers in formulas.")
    add_slide("Key Assumptions (Base)", ["Growth: 1.5% monthly", "COGS: 40% of revenue", "DSO/DPO/InvDays: 45 / 30 / 60", "Tax: 30%; Interest: 8% if debt"], None)

    # Chart slides
    for title, img in [("Results: Base Closing Cash", "base_line"), ("Scenario Comparison: Closing Cash", "scenarios_line"), ("Working Capital Components", "wc_components")]:
        layout = prs.slide_layouts[5]  # Title Only
        s = prs.slides.add_slide(layout)
        s.shapes.title.text = title
        s.shapes.add_picture(str(paths[img]), PPTInches(1), PPTInches(1.5), height=PPTInches(4.5))

    add_slide("Recommendations", ["Reduce DSO; negotiate DPO", "Phase CapEx after trough", "Maintain buffer via revolving facility", "Update actuals monthly; refresh scenarios quarterly"], None)
    prs.save(str(DECK_PATH))


def main():
    dfs = {name: build_schedule(**cfg) for name, cfg in SCENARIOS.items()}

    # Save scenario data
    out = pd.DataFrame({"month": dfs["Base"]["month"]})
    for name in ["Base", "Upside", "Downside"]:
        out[f"{name}_closing_cash"] = dfs[name]["closing_cash"].round(2)
    out["Base_net_cash_flow"] = dfs["Base"]["net_cash_flow"].round(2)
    out.to_csv(CSV_PATH, index=False)

    # Charts
    base_line = chart_closing_cash_base(dfs["Base"])
    scenarios_line = chart_scenarios(dfs)
    wc_components = chart_wc_components(dfs["Base"], month_idx=11)

    # Build docs
    paths = {"base_line": base_line, "scenarios_line": scenarios_line, "wc_components": wc_components}
    build_report(paths)
    build_deck(paths)
    print(f"Generated:\n- {REPORT_PATH}\n- {DECK_PATH}\n- {CSV_PATH}\nCharts in {CHARTS_DIR}")


if __name__ == "__main__":
    main()