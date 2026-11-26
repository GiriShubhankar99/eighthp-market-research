#!/usr/bin/env python3
"""
Eighth P Consultants — Real-Time Tier-2 City Expansion Automation (v2.0)
-------------------------------------------------------------------------
Researches any list of Indian cities (2 cities, 15 cities, or 60+ cities)
using Gemini API with live Google Search Grounding, and populates a complete
15-sheet Excel workbook preserving all formulas from the Eighth P template.

Upgrades in v2.0:
- Dynamic 10_Entry_Economics Profit formula (=IF(LEFT(C5,7)="Model 1",J5,...))
- Dynamic 15_Entry_Plan_90D generation for the actual Top 3 researched cities
- Strict UNAVAILABLE handling (missing data never silently becomes 0)
- Explicit SLBC/RERA data-gap & methodology logging in 13_Source_Register
- Dynamic RANK() & summary range resizing for N=2 to 60+ cities

Requirements:
    pip install google-genai openpyxl
"""

from __future__ import annotations

import argparse
import concurrent.futures
import copy
import datetime as dt
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.formula.translate import Translator
from google import genai
from google.genai import types


# ---------------------------------------------------------------------------
# 1. Gemini Live Web-Grounded Research Prompt & Schema
# ---------------------------------------------------------------------------

RESEARCH_PROMPT_TEMPLATE = """
You are a senior Indian real estate strategist, market researcher, and financial modeller
for Eighth P Consultants (a real estate sales & marketing mandate company).
Perform LIVE online research using Google Search for the Indian city: "{city_name}" (Date: {today}).

Find the most recent verified data from State RERA portals, NHB RESIDEX, SLBC district reports,
PropEquity, Knight Frank, CRISIL, 99acres, MagicBricks, and developer filings.

STRICT RESEARCH DISCIPLINE:
1. Never invent RERA registration numbers, project names, developer names, or contacts. Only include projects/developers found in live search results.
2. If a specific figure cannot be verified online, set it to null (do NOT invent or silently set to 0) and document the gap in `data_gap_note` (e.g. "City-level primary vs resale SLBC split unavailable; requires LDM request").
3. Mark coverage_status as "VERIFIED" only when backed by official RERA/NHB/PropEquity/Knight Frank sources found in search; otherwise mark "ESTIMATED".
4. Bound the city to its primary Development Authority planning area and exclude adjoining separate cities.
5. Include exact clickable source URLs in `sources` with period and methodology.

Return ONLY a single valid JSON object (no markdown fences, no commentary) matching this schema:
{{
  "city_name": "{city_name}",
  "state": "string",
  "zone": "East | West | North | South | Central | North-East | East-Central | South-East | Central-West",
  "goi_hra_class": "Class Y | Class X | Class Z",
  "planning_authority_boundary": "string",
  "excluded_adjoining_markets": "string",
  "ua_population_mn": number_or_null,
  "annual_primary_sales_units": integer_or_null,
  "avg_ticket_size_cr": number_or_null,
  "five_yr_price_growth_pct": number_or_null,
  "coverage_status": "VERIFIED | ESTIMATED | UNAVAILABLE",
  "data_gap_note": "string (document missing city-level SLBC/RERA data & exact LDM/authority request)",
  "demand_finance": {{
    "slbc_district_coverage": "string",
    "est_annual_sanctions_loans": integer_or_null,
    "gross_sanctioned_value_cr": number_or_null,
    "avg_ltv_ratio": number_or_null,
    "sanction_to_1st_disb_conversion": number_or_null,
    "primary_developer_purchase_share": number_or_null,
    "total_annual_disb_value_cr": number_or_null,
    "first_disb_share_of_total_disb": number_or_null,
    "mortgage_share_of_primary_sales": number_or_null,
    "home_loan_36m_yoy_cagr": number_or_null
  }},
  "supply_rera": {{
    "state_rera_portal_authority": "string (portal name + URL)",
    "active_rera_projects_count": integer_or_null,
    "stage1_approved_not_registered_units": integer_or_null,
    "stage2_registered_not_launched_units": integer_or_null,
    "stage3_launched_uc_total_units": integer_or_null,
    "stage3a_uc_sold_units": integer_or_null,
    "stage4_completed_unsold_units": integer_or_null,
    "stage5_stalled_delayed_units": integer_or_null,
    "non_rera_plotted_units": integer_or_null,
    "rera_disclosure_coverage": "VERIFIED | ESTIMATED | UNAVAILABLE"
  }},
  "demand_supply_balance": {{
    "opening_unsold_inventory_units": integer_or_null,
    "annual_new_launches_units": integer_or_null,
    "scheduled_completions_24m_units": integer_or_null,
    "mandate_need_velocity_score": number_or_null,
    "undersupplied_sweet_spot_segment": "string",
    "mandate_opportunity_assessment": "HIGH — ... | MEDIUM — ... | LOW — ..."
  }},
  "market_sizing_and_entry": {{
    "eligible_target_segment_share": number (decimal 0.45-0.65),
    "developer_outsourcing_propensity": number (decimal 0.18-0.40),
    "eighth_p_win_share_of_sam": number (decimal 0.08-0.30),
    "recommended_entry_model": "Model 1: Local Office | Model 2: Travelling Team -> Office | Model 3: Local JV Partner",
    "initial_bd_mobilisation_cost_cr": number,
    "time_to_first_revenue_months": number,
    "c3_sales_velocity_collections_score": number_or_null,
    "c4_low_competition_win_score": number_or_null,
    "c5_operating_fit_access_score": number_or_null,
    "c6_sustainable_demand_growth_score": number_or_null,
    "data_confidence_score": number (decimal 0.60-0.95),
    "recommended_wave_tier": "PRIORITY 1 — Enter | PRIORITY 2 — Monitor | PRIORITY 3 — Defer",
    "strategic_rationale": "string"
  }},
  "verified_projects": [
    {{
      "developer_name": "string",
      "project_name": "string",
      "micro_market": "string",
      "rera_reg_no": "string",
      "supply_stage": "3-Under Construction | 4-Ready with Unsold",
      "configuration": "string",
      "total_sanctioned_units": integer_or_null,
      "booked_sold_units": integer_or_null,
      "avg_ticket_size_cr": number_or_null,
      "expected_annual_sales_units": integer_or_null,
      "verification": "VERIFIED | ESTIMATED"
    }}
  ],
  "target_developers": [
    {{
      "developer_name": "string",
      "active_pipeline_projects": integer_or_null,
      "remaining_unsold_inventory_cr": number_or_null,
      "est_annual_primary_sales_cr": number_or_null,
      "eligible_project_sales_cr": number_or_null,
      "positioning_focus": "string",
      "internal_capability_gaps": "string",
      "recommended_engagement_model": "string",
      "acquisition_difficulty": "LOW | MEDIUM | HIGH",
      "outsourcing_probability": number,
      "win_probability": number
    }}
  ],
  "sources": [
    {{
      "module": "string",
      "authority": "string",
      "title": "string",
      "url": "https://...",
      "pub_date": "YYYY-MM-DD",
      "period": "string (e.g. FY25-26 / Q3 2026)",
      "methodology": "string",
      "confidence": number (decimal 0.65-0.98),
      "classification": "VERIFIED | ESTIMATED"
    }}
  ]
}}
"""


def _extract_json(raw_text: str) -> dict[str, Any]:
    """Strip optional markdown fences and parse JSON cleanly."""
    cleaned = raw_text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"No JSON object found in model output: {cleaned[:200]}")
    return json.loads(cleaned[start : end + 1])


def research_city(
    client: genai.Client,
    city_name: str,
    model_name: str = "gemini-2.5-flash",
    max_retries: int = 3,
) -> dict[str, Any]:
    """Run live Google Search-grounded Gemini research for one city."""
    today = dt.date.today().isoformat()
    prompt = RESEARCH_PROMPT_TEMPLATE.format(city_name=city_name, today=today)

    last_err: Exception | None = None
    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                    temperature=0.1,
                ),
            )
            data = _extract_json(response.text or "")

            grounding_urls: list[str] = []
            try:
                chunks = response.candidates[0].grounding_metadata.grounding_chunks or []
                for ch in chunks:
                    if getattr(ch, "web", None) and ch.web.uri:
                        grounding_urls.append(ch.web.uri)
            except Exception:
                pass
            data["_grounding_urls"] = list(dict.fromkeys(grounding_urls))
            return data
        except Exception as exc:
            last_err = exc
            wait_s = 2**attempt
            print(f"  [WARN] {city_name} attempt {attempt}/{max_retries} failed: {exc}. Retrying in {wait_s}s...")
            time.sleep(wait_s)

    raise RuntimeError(f"Failed researching {city_name} after {max_retries} attempts: {last_err}")


# ---------------------------------------------------------------------------
# 2. Scoring & Management Summary Helper (Matches Sheet 9/10/11 Logic)
# ---------------------------------------------------------------------------

def score_city_record(rec: dict[str, Any], net_fee_rate: float = 0.02, retainer_annual: float = 0.42) -> dict[str, Any]:
    """Compute Python-side summary metrics & attractiveness score for UI & Sheet 15."""
    units = rec.get("annual_primary_sales_units")
    ticket = rec.get("avg_ticket_size_cr")
    ms = rec.get("market_sizing_and_entry", {})

    if units is None or ticket is None:
        mkt_cr = 0.0
    else:
        mkt_cr = round(float(units) * float(ticket), 1)

    elig_share = float(ms.get("eligible_target_segment_share") or 0.55)
    outsource_p = float(ms.get("developer_outsourcing_propensity") or 0.25)
    win_share = float(ms.get("eighth_p_win_share_of_sam") or 0.15)

    elig_sales = mkt_cr * elig_share * outsource_p
    sam = (elig_sales * net_fee_rate) + ((elig_sales / 65.0) * retainer_annual)
    som_y3 = sam * win_share
    som_y2 = round(som_y3 * 0.70, 2)

    mandates_y2 = max(1.0, round(som_y2 / 1.65, 1))
    rec_model = ms.get("recommended_entry_model", "Model 1: Local Office")
    if rec_model.startswith("Model 2"):
        y2_cost = 0.48 + 0.40 * mandates_y2
        y2_profit = round(som_y2 * 0.92 - y2_cost, 2)
    elif rec_model.startswith("Model 3"):
        y2_cost = 0.30 + 0.18 * mandates_y2
        y2_profit = round(som_y2 * 0.72 - y2_cost, 2)
    else:
        y2_cost = 0.85 + 0.36 * mandates_y2
        y2_profit = round(som_y2 - y2_cost, 2)

    ebitda_margin = (y2_profit / som_y2) if som_y2 > 0 else 0.0
    bd_cost = float(ms.get("initial_bd_mobilisation_cost_cr") or 0.25)
    peak_cash = round(bd_cost + ((som_y2 - y2_profit) * (75.0 / 365.0)) + 0.25, 2)
    payback = round(12.0 + (peak_cash / max(0.15, y2_profit)) * 12.0, 1)

    c1 = min(100.0, round((som_y2 / 7.2) * 100.0, 1))
    c2 = min(100.0, round((y2_profit / 4.0) * 65.0 + (ebitda_margin * 35.0), 1))
    c3 = ms.get("c3_sales_velocity_collections_score")
    c4 = ms.get("c4_low_competition_win_score")
    c5 = ms.get("c5_operating_fit_access_score")
    c6 = ms.get("c6_sustainable_demand_growth_score")

    if None in (c3, c4, c5, c6) or units is None or ticket is None:
        composite = -1.0  # Identifiable incomplete/missing score (never silently 0)
    else:
        composite = round(
            c1 * 0.30 + c2 * 0.20 + float(c3) * 0.15 + float(c4) * 0.15 + float(c5) * 0.10 + float(c6) * 0.10,
            1,
        )

    return {
        "city_name": rec.get("city_name", ""),
        "state": rec.get("state", ""),
        "primary_mkt_cr": mkt_cr,
        "sam_cr": round(sam, 2),
        "som_y2_cr": som_y2,
        "y2_profit_cr": y2_profit,
        "ebitda_margin": ebitda_margin,
        "peak_cash_cr": peak_cash,
        "payback_mo": payback,
        "score": composite,
        "confidence": float(ms.get("data_confidence_score") or 0.80),
        "entry_model": rec_model,
        "wave_tier": ms.get("recommended_wave_tier", ""),
        "rationale": ms.get("strategic_rationale", ""),
        "data_gap_note": rec.get("data_gap_note", ""),
    }


# ---------------------------------------------------------------------------
# 3. Excel Workbook Population & Dynamic Formula Cloning
# ---------------------------------------------------------------------------

def _val_or_unavail(v: Any) -> Any:
    """Preserve explicit UNAVAILABLE marker for missing values instead of silent 0."""
    return "UNAVAILABLE" if v is None else v


def _copy_row_style(ws: openpyxl.worksheet.worksheet.Worksheet, src_row: int, dst_row: int, max_col: int) -> None:
    """Copy cell styles (font, fill, border, alignment, number_format) from src_row to dst_row."""
    for col in range(1, max_col + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)
        if src.has_style:
            dst.font = copy.copy(src.font)
            dst.fill = copy.copy(src.fill)
            dst.border = copy.copy(src.border)
            dst.alignment = copy.copy(src.alignment)
            dst.number_format = src.number_format


def _resize_range_bounds(formula: str, last_data_row: int) -> str:
    """Update template range bounds like $K$5:$K$19 or B5:B34 to end at last_data_row."""
    return re.sub(
        r"(\$?[A-Z]+\$?5:\$?[A-Z]+\$?)\d+",
        lambda m: f"{m.group(1)}{last_data_row}",
        formula,
    )


def _translate_formula(formula: str, src_row: int, dst_row: int, col_idx: int, last_data_row: int) -> str:
    """Translate row-relative Excel formula from src_row to dst_row and resize any 5..N range bounds."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    col_letter = openpyxl.utils.get_column_letter(col_idx)
    origin = f"{col_letter}{src_row}"
    target = f"{col_letter}{dst_row}"
    translated = Translator(formula, origin=origin).translate_formula(target)
    return _resize_range_bounds(translated, last_data_row)


def _prepare_city_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    n_cities: int,
    template_last_data_row: int,
    summary_label: str = "TOTAL",
    row5_overrides: dict[int, str] | None = None,
) -> None:
    """Resize a city-level sheet to n_cities rows, cloning formulas and updating the summary row."""
    max_col = ws.max_column
    template_summary_row = template_last_data_row + 1
    summary_formulas = [ws.cell(template_summary_row, c).value for c in range(1, max_col + 1)]
    summary_styles = [
        (
            copy.copy(ws.cell(template_summary_row, c).font),
            copy.copy(ws.cell(template_summary_row, c).fill),
            copy.copy(ws.cell(template_summary_row, c).border),
            copy.copy(ws.cell(template_summary_row, c).alignment),
            ws.cell(template_summary_row, c).number_format,
        )
        for c in range(1, max_col + 1)
    ]
    row5_formulas = [ws.cell(5, c).value for c in range(1, max_col + 1)]
    if row5_overrides:
        for col_idx, f_str in row5_overrides.items():
            row5_formulas[col_idx - 1] = f_str

    for r in range(5, ws.max_row + 5):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None

    last_data_row = 4 + n_cities

    for idx in range(n_cities):
        dst_row = 5 + idx
        if dst_row > 5:
            _copy_row_style(ws, 5, dst_row, max_col)
        for c in range(1, max_col + 1):
            val = row5_formulas[c - 1]
            if isinstance(val, str) and val.startswith("="):
                ws.cell(dst_row, c).value = _translate_formula(val, 5, dst_row, c, last_data_row)

    new_summary_row = 5 + n_cities
    for c in range(1, max_col + 1):
        cell = ws.cell(new_summary_row, c)
        f = summary_formulas[c - 1]
        if isinstance(f, str) and f.startswith("="):
            cell.value = _resize_range_bounds(f, last_data_row)
        else:
            cell.value = _resize_range_bounds(summary_label, last_data_row) if c == 1 else f
        font, fill, border, align, nfmt = summary_styles[c - 1]
        cell.font, cell.fill, cell.border, cell.alignment, cell.number_format = font, fill, border, align, nfmt


def _populate_entry_plan_90d(
    ws15: openpyxl.worksheet.worksheet.Worksheet,
    city_records: list[dict[str, Any]],
    ranked_summaries: list[dict[str, Any]],
) -> None:
    """Dynamically populate Sheet 15 (90-Day Entry Plan) for the actual Top 3 cities."""
    top_cities = [s for s in ranked_summaries if s["score"] >= 0][:3]
    if not top_cities:
        top_cities = ranked_summaries[: min(3, len(ranked_summaries))]

    top_names = [c["city_name"] for c in top_cities]
    ws15.cell(1, 1).value = f"RECOMMENDED 90-DAY MARKET ENTRY PLAN — TOP {len(top_names)} CITIES ({', '.join(n.upper() for n in top_names)})"

    rec_by_name = {r["city_name"]: r for r in city_records}
    rows_to_write: list[list[Any]] = []
    act_idx = 1
    per_city_budgets = [(5.5, 7.5, 17.0), (5.5, 7.5, 17.0), (5.0, 7.0, 15.0)]

    for i, c_sum in enumerate(top_cities):
        cname = c_sum["city_name"]
        rec = rec_by_name.get(cname, {})
        devs = [d.get("developer_name", "") for d in rec.get("target_developers", [])[:4] if d.get("developer_name")]
        dev_str = ", ".join(devs) if devs else f"Top Tier-A/B Developers in {cname}"
        anchor_str = ", ".join(devs[:2]) if len(devs) >= 2 else dev_str
        b1, b2, b3 = per_city_budgets[i] if i < len(per_city_budgets) else (5.0, 6.0, 14.0)
        sweet_spot = rec.get("demand_supply_balance", {}).get("undersupplied_sweet_spot_segment", "Mid-Premium 2/3 BHK")

        rows_to_write.append([
            f"ACT-{act_idx:02d}", f"{cname} (#{i+1})", "Days 1-30: Intelligence & BD", "Developer Outreach & RERA Audit",
            dev_str,
            f"Pitch institutional GTM, CRM & CP engine for {sweet_spot}; benchmark Bhubaneswar track record.",
            "1 Partner + 1 Regional BD Lead", b1,
            f"8 Founder/MD Meetings + 3 RERA Project Audits in {cname}",
            "GO if >=3 Tier-A/B builders share inventory sheets; NO-GO if net retained fee <2.0%.",
            "PENDING",
        ])
        act_idx += 1
        rows_to_write.append([
            f"ACT-{act_idx:02d}", f"{cname} (#{i+1})", "Days 31-60: Term Sheets & CP", "Mandate Structuring & CP Mapping",
            anchor_str,
            f"Structure 90-day GTM turnaround + INR 3.5L/mo retainer + net retained sales fee under {c_sum['entry_model']}.",
            f"Partner + {cname} City/Pod Lead", b2,
            f"2 Exclusive Mandate Term Sheets Signed + 150 Local CPs Mapped",
            "GO to full office/pod setup only after >=1 signed mandate with >=INR 60 Cr saleable inventory.",
            "PENDING",
        ])
        act_idx += 1
        rows_to_write.append([
            f"ACT-{act_idx:02d}", f"{cname} (#{i+1})", "Days 61-90: Mobilisation", "Site Pod Deployment & Launch",
            f"Signed {cname} Anchor Accounts",
            "Deploy 4-person site closing pod, integrate CRM lead routing, host 100+ broker CP Conclave.",
            f"1 City/Pod Lead + 4 Site Staff", b3,
            f"First 15 Verified Bookings Logged & First Retainer Invoiced",
            "GO to Mandate #3 expansion if Month-3 run-rate >=8 units/month and <60-day collections.",
            "PENDING",
        ])
        act_idx += 1

    # Clear old rows 5..20
    for r in range(5, ws15.max_row + 5):
        for c in range(1, 12):
            ws15.cell(r, c).value = None

    for idx, row_vals in enumerate(rows_to_write):
        r = 5 + idx
        if r > 5:
            _copy_row_style(ws15, 5, r, 11)
        for c, val in enumerate(row_vals, start=1):
            ws15.cell(r, c).value = val

    tot_row = 5 + len(rows_to_write)
    last_r = tot_row - 1
    ws15.cell(tot_row, 1).value = "TOTAL"
    ws15.cell(tot_row, 2).value = f"Top {len(top_names)} Expansion Cities"
    ws15.cell(tot_row, 3).value = "90-Day Rollout Window"
    ws15.cell(tot_row, 4).value = "BD, Contracting & Mobilisation"
    ws15.cell(tot_row, 5).value = "Priority Anchor Accounts"
    ws15.cell(tot_row, 6).value = f"Target: {2*len(top_names)}-{3*len(top_names)} Exclusive Mandates in 90 Days"
    ws15.cell(tot_row, 7).value = f"{len(top_names)} Leads + {4*len(top_names)} Site FTEs"
    ws15.cell(tot_row, 8).value = f"=SUM(H5:H{last_r})"
    ws15.cell(tot_row, 9).value = "Stage-Gated Capital Release"
    ws15.cell(tot_row, 10).value = "Milestone-Verified Payouts"
    ws15.cell(tot_row, 11).value = "ACTIVE"


def populate_workbook(
    template_path: Path,
    output_path: Path,
    city_records: list[dict[str, Any]],
) -> None:
    """Populate all 15 sheets of the Eighth P workbook for N cities."""
    wb = openpyxl.load_workbook(template_path)
    n = len(city_records)

    # Read Master Inputs fee rates for accurate ranking & header alignment
    ws2 = wb["2_Master_Inputs"]
    gross_fee = float(ws2.cell(10, 3).value or 0.02)
    cp_share = float(ws2.cell(11, 3).value or 0.0)
    net_fee = gross_fee * (1.0 - cp_share)
    retainer_annual = float(ws2.cell(13, 3).value or 0.035) * 12.0

    # Align Sheet 9 TAM headers with actual Master Inputs fee rates
    ws9 = wb["9_TAM_SAM_SOM"]
    ws9.cell(4, 4).value = f"Gross Mandate Fee TAM @ {gross_fee*100:.1f}% (INR Cr)"
    ws9.cell(4, 5).value = f"Net Retained Fee TAM @ {net_fee*100:.1f}% (INR Cr)"

    # Resize all city-indexed sheets
    _prepare_city_sheet(wb["3_City_Universe"], n, template_last_data_row=34, summary_label="TOTAL")
    for sname, lbl, overrides in [
        ("4_Demand_Finance", "TOTAL", None),
        ("5_Supply_Approvals", "TOTAL", None),
        ("6_Demand_Supply_Bal", "TOTAL", None),
        ("9_TAM_SAM_SOM", "TOTAL", None),
        (
            "10_Entry_Economics",
            "TOTAL",
            # Fix M5 to dynamically select Model 1 / Model 2 / Model 3 based on Column C!
            {13: '=IF(LEFT(C5,7)="Model 1",J5,IF(LEFT(C5,7)="Model 2",K5,L5))'},
        ),
        ("11_City_Ranking", "AVERAGE", None),
        ("12_Scenarios_Sens", "TOTAL", None),
        ("14_Mgmt_Dashboard", "=COUNT(A5:A19)", None),
    ]:
        _prepare_city_sheet(wb[sname], n, template_last_data_row=19, summary_label=lbl, row5_overrides=overrides)

    ws3 = wb["3_City_Universe"]
    ws4 = wb["4_Demand_Finance"]
    ws5 = wb["5_Supply_Approvals"]
    ws6 = wb["6_Demand_Supply_Bal"]
    ws10 = wb["10_Entry_Economics"]
    ws11 = wb["11_City_Ranking"]

    for idx, rec in enumerate(city_records):
        r = 5 + idx
        cid = f"CTY-{idx + 1:02d}"
        df = rec.get("demand_finance", {})
        sr = rec.get("supply_rera", {})
        ds = rec.get("demand_supply_balance", {})
        ms = rec.get("market_sizing_and_entry", {})

        # Sheet 3: Universe
        ws3.cell(r, 1).value = cid
        ws3.cell(r, 2).value = rec["city_name"]
        ws3.cell(r, 3).value = rec.get("state", "UNAVAILABLE")
        ws3.cell(r, 4).value = rec.get("zone", "UNAVAILABLE")
        ws3.cell(r, 5).value = f"Live AI Batch ({dt.date.today().isoformat()})"
        ws3.cell(r, 6).value = rec.get("goi_hra_class", "Class Y")
        ws3.cell(r, 7).value = rec.get("planning_authority_boundary", "UNAVAILABLE")
        ws3.cell(r, 8).value = rec.get("excluded_adjoining_markets", "UNAVAILABLE")
        ws3.cell(r, 9).value = _val_or_unavail(rec.get("ua_population_mn"))
        ws3.cell(r, 10).value = _val_or_unavail(rec.get("annual_primary_sales_units"))
        ws3.cell(r, 11).value = _val_or_unavail(rec.get("avg_ticket_size_cr"))
        ws3.cell(r, 13).value = _val_or_unavail(rec.get("five_yr_price_growth_pct"))
        ws3.cell(r, 14).value = rec.get("coverage_status", "ESTIMATED")

        # Sheet 4: Demand & Finance
        ws4.cell(r, 1).value = cid
        ws4.cell(r, 3).value = df.get("slbc_district_coverage", "UNAVAILABLE")
        ws4.cell(r, 4).value = _val_or_unavail(df.get("est_annual_sanctions_loans"))
        ws4.cell(r, 5).value = _val_or_unavail(df.get("gross_sanctioned_value_cr"))
        ws4.cell(r, 7).value = _val_or_unavail(df.get("avg_ltv_ratio"))
        ws4.cell(r, 8).value = _val_or_unavail(df.get("sanction_to_1st_disb_conversion"))
        ws4.cell(r, 10).value = _val_or_unavail(df.get("primary_developer_purchase_share"))
        ws4.cell(r, 12).value = _val_or_unavail(df.get("total_annual_disb_value_cr"))
        ws4.cell(r, 13).value = _val_or_unavail(df.get("first_disb_share_of_total_disb"))
        ws4.cell(r, 14).value = _val_or_unavail(df.get("mortgage_share_of_primary_sales"))
        ws4.cell(r, 17).value = _val_or_unavail(df.get("home_loan_36m_yoy_cagr"))

        # Sheet 5: Supply & RERA
        ws5.cell(r, 1).value = cid
        ws5.cell(r, 3).value = sr.get("state_rera_portal_authority", "UNAVAILABLE")
        ws5.cell(r, 4).value = _val_or_unavail(sr.get("active_rera_projects_count"))
        ws5.cell(r, 5).value = _val_or_unavail(sr.get("stage1_approved_not_registered_units"))
        ws5.cell(r, 6).value = _val_or_unavail(sr.get("stage2_registered_not_launched_units"))
        ws5.cell(r, 7).value = _val_or_unavail(sr.get("stage3_launched_uc_total_units"))
        ws5.cell(r, 8).value = _val_or_unavail(sr.get("stage3a_uc_sold_units"))
        ws5.cell(r, 10).value = _val_or_unavail(sr.get("stage4_completed_unsold_units"))
        ws5.cell(r, 11).value = _val_or_unavail(sr.get("stage5_stalled_delayed_units"))
        ws5.cell(r, 13).value = _val_or_unavail(sr.get("non_rera_plotted_units"))
        ws5.cell(r, 16).value = sr.get("rera_disclosure_coverage", "ESTIMATED")

        # Sheet 6: Demand-Supply Balance
        ws6.cell(r, 1).value = cid
        ws6.cell(r, 3).value = _val_or_unavail(ds.get("opening_unsold_inventory_units"))
        ws6.cell(r, 4).value = _val_or_unavail(ds.get("annual_new_launches_units"))
        ws6.cell(r, 14).value = _val_or_unavail(ds.get("scheduled_completions_24m_units"))
        ws6.cell(r, 15).value = _val_or_unavail(ds.get("mandate_need_velocity_score"))
        ws6.cell(r, 16).value = ds.get("undersupplied_sweet_spot_segment", "UNAVAILABLE")
        ws6.cell(r, 17).value = ds.get("mandate_opportunity_assessment", "UNAVAILABLE")

        # Sheet 9: TAM/SAM/SOM
        ws9.cell(r, 1).value = cid
        ws9.cell(r, 6).value = ms.get("eligible_target_segment_share", 0.55)
        ws9.cell(r, 7).value = ms.get("developer_outsourcing_propensity", 0.25)
        ws9.cell(r, 12).value = ms.get("eighth_p_win_share_of_sam", 0.15)

        # Sheet 10: Entry Economics
        ws10.cell(r, 1).value = cid
        ws10.cell(r, 3).value = ms.get("recommended_entry_model", "Model 1: Local Office")
        ws10.cell(r, 4).value = ms.get("initial_bd_mobilisation_cost_cr", 0.30)
        ws10.cell(r, 5).value = ms.get("time_to_first_revenue_months", 4.0)

        # Sheet 11: City Ranking (leave missing scores blank so COUNT<6 flags -1 instead of silent 0)
        ws11.cell(r, 1).value = cid
        ws11.cell(r, 6).value = ms.get("c3_sales_velocity_collections_score")
        ws11.cell(r, 7).value = ms.get("c4_low_competition_win_score")
        ws11.cell(r, 8).value = ms.get("c5_operating_fit_access_score")
        ws11.cell(r, 9).value = ms.get("c6_sustainable_demand_growth_score")
        ws11.cell(r, 13).value = ms.get("data_confidence_score", 0.80)
        ws11.cell(r, 15).value = ms.get("recommended_wave_tier", "PRIORITY 2 — Monitor")
        ws11.cell(r, 16).value = ms.get("strategic_rationale", "")

    # Sheet 7: Verified RERA Projects
    ws7 = wb["7_Project_Database"]
    all_projects = [(rec["city_name"], p) for rec in city_records for p in rec.get("verified_projects", [])]
    _prepare_city_sheet(ws7, max(1, len(all_projects)), template_last_data_row=34, summary_label="TOTAL")
    for idx, (cname, p) in enumerate(all_projects):
        r = 5 + idx
        ws7.cell(r, 1).value = f"PRJ-{idx + 1:03d}"
        ws7.cell(r, 2).value = cname
        ws7.cell(r, 3).value = p.get("developer_name", "UNAVAILABLE")
        ws7.cell(r, 4).value = p.get("project_name", "UNAVAILABLE")
        ws7.cell(r, 5).value = p.get("micro_market", "UNAVAILABLE")
        ws7.cell(r, 6).value = p.get("rera_reg_no", "UNAVAILABLE")
        ws7.cell(r, 7).value = p.get("supply_stage", "3-Under Construction")
        ws7.cell(r, 8).value = p.get("configuration", "UNAVAILABLE")
        ws7.cell(r, 9).value = _val_or_unavail(p.get("total_sanctioned_units"))
        ws7.cell(r, 10).value = _val_or_unavail(p.get("booked_sold_units"))
        ws7.cell(r, 12).value = _val_or_unavail(p.get("avg_ticket_size_cr"))
        ws7.cell(r, 14).value = _val_or_unavail(p.get("expected_annual_sales_units"))
        ws7.cell(r, 16).value = p.get("verification", "ESTIMATED")

    # Sheet 8: Target Developer Accounts
    ws8 = wb["8_Developer_Accounts"]
    all_devs = [(rec["city_name"], d) for rec in city_records for d in rec.get("target_developers", [])]
    _prepare_city_sheet(ws8, max(1, len(all_devs)), template_last_data_row=34, summary_label="TOTAL")
    for idx, (cname, d) in enumerate(all_devs):
        r = 5 + idx
        ws8.cell(r, 1).value = f"DEV-{idx + 1:03d}"
        ws8.cell(r, 2).value = cname
        ws8.cell(r, 3).value = d.get("developer_name", "UNAVAILABLE")
        ws8.cell(r, 4).value = _val_or_unavail(d.get("active_pipeline_projects"))
        ws8.cell(r, 5).value = _val_or_unavail(d.get("remaining_unsold_inventory_cr"))
        ws8.cell(r, 6).value = _val_or_unavail(d.get("est_annual_primary_sales_cr"))
        ws8.cell(r, 8).value = _val_or_unavail(d.get("eligible_project_sales_cr"))
        ws8.cell(r, 9).value = d.get("positioning_focus", "")
        ws8.cell(r, 10).value = d.get("internal_capability_gaps", "")
        ws8.cell(r, 11).value = d.get("recommended_engagement_model", "")
        ws8.cell(r, 12).value = d.get("acquisition_difficulty", "MEDIUM")
        ws8.cell(r, 13).value = _val_or_unavail(d.get("outsourcing_probability"))
        ws8.cell(r, 14).value = _val_or_unavail(d.get("win_probability"))

    # Sheet 13: Source Register + Explicit Data-Gap Documentation
    ws13 = wb["13_Source_Register"]
    for r in range(5, ws13.max_row + 1):
        for c in range(1, ws13.max_column + 1):
            ws13.cell(r, c).value = None
    src_idx = 1
    for rec in city_records:
        for s in rec.get("sources", []):
            r = 4 + src_idx
            ws13.cell(r, 1).value = f"SRC-{src_idx:03d}"
            ws13.cell(r, 2).value = s.get("module", "Market Research")
            ws13.cell(r, 3).value = s.get("authority", "")
            period_str = s.get("period", "")
            meth_str = s.get("methodology", "")
            title_extra = f" ({period_str})" if period_str else ""
            if meth_str:
                title_extra += f" — {meth_str}"
            ws13.cell(r, 4).value = f"[{rec['city_name']}] {s.get('title', '')}{title_extra}"
            ws13.cell(r, 5).value = s.get("url", "")
            ws13.cell(r, 6).value = s.get("pub_date", dt.date.today().isoformat())
            ws13.cell(r, 7).value = rec["city_name"]
            ws13.cell(r, 8).value = s.get("confidence", 0.85)
            ws13.cell(r, 9).value = s.get("classification", "ESTIMATED")
            src_idx += 1

        gap_note = rec.get("data_gap_note", "").strip()
        if gap_note:
            r = 4 + src_idx
            ws13.cell(r, 1).value = f"GAP-{src_idx:03d}"
            ws13.cell(r, 2).value = "Data Gap / LDM Request"
            ws13.cell(r, 3).value = "SLBC / State RERA"
            ws13.cell(r, 4).value = f"[{rec['city_name']}] {gap_note}"
            ws13.cell(r, 5).value = "N/A — Requires LDM/Authority Request"
            ws13.cell(r, 6).value = dt.date.today().isoformat()
            ws13.cell(r, 7).value = rec["city_name"]
            ws13.cell(r, 8).value = rec.get("market_sizing_and_entry", {}).get("data_confidence_score", 0.80)
            ws13.cell(r, 9).value = "DATA GAP"
            src_idx += 1

    # Sheet 15: Dynamic 90-Day Entry Plan for the Top 3 Researched Cities
    ranked = sorted(
        [score_city_record(rec, net_fee, retainer_annual) for rec in city_records],
        key=lambda x: x["score"],
        reverse=True,
    )
    _populate_entry_plan_90d(wb["15_Entry_Plan_90D"], city_records, ranked)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


# ---------------------------------------------------------------------------
# 4. CLI Entrypoint (Fresh Live Data by Default)
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Eighth P Real-Time City Research Automation")
    parser.add_argument("--template", required=True, help="Path to Eighth_P_Market_Expansion_Intelligence_Tool.xlsx")
    parser.add_argument("--cities", help="Comma-separated list of cities")
    parser.add_argument("--cities-file", help="Path to .txt file with one city name per line")
    parser.add_argument("--output", help="Output .xlsx path (default: timestamped filename)")
    parser.add_argument("--use-cache", action="store_true", help="Resume from today's checkpoint cache")
    parser.add_argument("--cache", default=f"eighth_p_cache_{dt.date.today().isoformat()}.json", help="Checkpoint JSON cache path")
    parser.add_argument("--workers", type=int, default=4, help="Parallel research workers (recommended 3-5)")
    parser.add_argument("--model", default="gemini-2.5-flash", help="Gemini model ID")
    args = parser.parse_args()

    cities: list[str] = []
    if args.cities:
        cities.extend([c.strip() for c in args.cities.split(",") if c.strip()])
    if args.cities_file:
        with open(args.cities_file, encoding="utf-8") as fh:
            cities.extend([line.strip() for line in fh if line.strip() and not line.startswith("#")])
    cities = list(dict.fromkeys(cities))

    if not cities:
        print("Error: Provide --cities or --cities-file.", file=sys.stderr)
        sys.exit(1)

    if not os.environ.get("GEMINI_API_KEY"):
        print("Error: GEMINI_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    cache_path = Path(args.cache)
    cache: dict[str, dict[str, Any]] = {}
    if args.use_cache and cache_path.exists():
        cache = json.loads(cache_path.read_text(encoding="utf-8"))

    pending = [c for c in cities if c not in cache]
    print(f"Total cities: {len(cities)} | Cached: {len(cities) - len(pending)} | Fetching fresh live: {len(pending)}")

    if pending:
        client = genai.Client()
        with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
            future_map = {pool.submit(research_city, client, city, args.model): city for city in pending}
            for fut in concurrent.futures.as_completed(future_map):
                city = future_map[fut]
                try:
                    result = fut.result()
                    cache[city] = result
                    cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
                    print(f"  [OK] {city} researched & saved ({len(result.get('sources', []))} sources)")
                except Exception as exc:
                    print(f"  [FAIL] {city}: {exc}", file=sys.stderr)

    ordered_records = [cache[c] for c in cities if c in cache]
    if not ordered_records:
        print("No city records available to write.", file=sys.stderr)
        sys.exit(1)

    ts = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out_path = Path(args.output) if args.output else Path(f"Eighth_P_Live_{len(ordered_records)}Cities_{ts}.xlsx")
    populate_workbook(Path(args.template), out_path, ordered_records)
    print(f"\nDone! Populated {len(ordered_records)}/{len(cities)} cities into: {out_path}")


if __name__ == "__main__":
    main()