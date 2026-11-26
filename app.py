#!/usr/bin/env python3
"""
Eighth P Consultants — Real-Time Tier-2 Market Expansion Web App (v2.1)
-----------------------------------------------------------------------
Deploy on Streamlit Community Cloud (zero terminal) or run locally:
    pip install streamlit google-genai openpyxl
    streamlit run app.py
"""

from __future__ import annotations

import concurrent.futures
import datetime as dt
import json
import os
import tempfile
from pathlib import Path
from typing import Any

import openpyxl
import streamlit as st
from google import genai

from eighth_p_auto_research import populate_workbook, research_city, score_city_record


st.set_page_config(
    page_title="Eighth P — Tier-2 Market Expansion AI",
    page_icon="🏙️",
    layout="wide",
)

st.title("Eighth P — Real-Time Tier-2 City Expansion Intelligence")
st.caption(
    "Paste 2 cities or 60+ cities, upload your Eighth P Excel template, and generate a "
    "brand-new 15-sheet workbook with fresh live Google Search-grounded data, rankings, and a 90-day entry plan."
)


def _get_saved_api_key() -> str:
    """Read Gemini API key from Streamlit Secrets or environment variable."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


@st.cache_data(ttl=3600, show_spinner=False)
def _list_available_models(api_key: str) -> list[str]:
    """Fetch live generateContent-capable Gemini models for the user's API key."""
    fallback = [
        "gemini-4-hydrogen",
        "gemini-3.8-flash",
        "gemini-3.5-flash-lite",
        "gemini-4-neon",
    ]
    if not api_key:
        return fallback
    try:
        client = genai.Client(api_key=api_key)
        preferred = []
        others = []
        for m in client.models.list():
            name = getattr(m, "name", "") or ""
            short = name.replace("models/", "")
            methods = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None) or []
            if "gemini" not in short.lower():
                continue
            if any(x in short.lower() for x in ("tts", "image", "audio", "live", "embedding", "robotics", "computer-use")):
                continue
            if methods and not any("generatecontent" in str(act).lower() for act in methods):
                continue
            if any(k in short.lower() for k in ("4-hydrogen", "3.8-flash", "3.5-flash-lite", "4-neon", "4-argon")):
                preferred.append(short)
            else:
                others.append(short)
        combined = list(dict.fromkeys(preferred + others))
        return combined if combined else fallback
    except Exception:
        return fallback


# ---------------------------------------------------------------------------
# Sidebar Configuration
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Configuration")
    saved_key = _get_saved_api_key()
    if saved_key:
        st.success("API Key connected automatically")
        with st.expander("Override API Key (optional)"):
            override_key = st.text_input("Custom Gemini API Key", type="password")
        api_key = override_key.strip() if override_key.strip() else saved_key
    else:
        api_key = st.text_input(
            "Gemini API Key",
            value="",
            type="password",
            help="Save GEMINI_API_KEY in Streamlit App Settings -> Secrets to never enter it again.",
        ).strip()

    available_models = _list_available_models(api_key)
    model_name = st.selectbox(
        "Gemini Model",
        options=available_models,
        index=0,
        help="Auto-populated from your API key's active Gemini models (gemini-4-hydrogen / gemini-3.8-flash recommended).",
    )
    workers = st.slider("Parallel Research Workers", min_value=1, max_value=8, value=4)
    resume_mode = st.checkbox(
        "Resume interrupted batch (use today's cache)",
        value=False,
        help="Leave OFF to always fetch 100% fresh live data every run. Turn ON only to resume a crashed 60-city run.",
    )

# ---------------------------------------------------------------------------
# Main Inputs: City List & Template Workbook
# ---------------------------------------------------------------------------
col_left, col_right = st.columns([3, 2], gap="large")

with col_left:
    st.subheader("1. Target City List")
    city_text = st.text_area(
        "Enter city names (comma-separated or one per line)",
        placeholder="Lucknow\nIndore\nRaipur\nNashik\nGuwahati",
        height=180,
    )
    cities_file = st.file_uploader(
        "Or upload a .txt / .csv city list (one city per line)",
        type=["txt", "csv"],
    )

with col_right:
    st.subheader("2. Excel Template")
    local_default = Path("Eighth_P_Market_Expansion_Intelligence_Tool.xlsx")
    uploaded_template = st.file_uploader(
        "Upload Eighth_P_Market_Expansion_Intelligence_Tool.xlsx",
        type=["xlsx"],
        help="Preserves all 15 sheets and formulas while resizing dynamically to your city count.",
    )
    if uploaded_template is None and local_default.exists():
        st.info(f"Using local template found in folder: `{local_default.name}`")

# Parse unique cities
cities: list[str] = []
if city_text.strip():
    for token in city_text.replace("\n", ",").split(","):
        name = token.strip()
        if name:
            cities.append(name)
if cities_file is not None:
    raw_lines = cities_file.read().decode("utf-8", errors="ignore").splitlines()
    for line in raw_lines:
        name = line.strip().strip(",")
        if name and not name.startswith("#"):
            cities.append(name)
cities = list(dict.fromkeys(cities))

st.markdown(f"**Cities queued:** `{len(cities)}`")

# ---------------------------------------------------------------------------
# Run Live Research & Build Brand-New Timestamped Workbook
# ---------------------------------------------------------------------------
run_btn = st.button("Start Fresh Live Research & Generate New Excel", type="primary", disabled=len(cities) == 0)

if run_btn:
    if not api_key:
        st.error("Please enter your Gemini API Key in the sidebar or save it in Streamlit Secrets.")
        st.stop()

    if uploaded_template is None and not local_default.exists():
        st.error("Please upload `Eighth_P_Market_Expansion_Intelligence_Tool.xlsx`.")
        st.stop()

    os.environ["GEMINI_API_KEY"] = api_key
    client = genai.Client(api_key=api_key)

    today_str = dt.date.today().isoformat()
    cache_path = Path(f"eighth_p_cache_{today_str}.json")
    cache: dict[str, dict[str, Any]] = {}
    if resume_mode and cache_path.exists():
        cache = json.loads(cache_path.read_text(encoding="utf-8"))

    pending = [c for c in cities if c not in cache]
    completed_count = len(cities) - len(pending)

    progress_bar = st.progress(completed_count / len(cities) if cities else 0)
    status_box = st.empty()
    log_area = st.empty()
    logs: list[str] = []

    if completed_count > 0:
        logs.append(f"Resumed {completed_count} cities from today's cache.")
        log_area.code("\n".join(logs[-12:]))

    if pending:
        status_box.info(f"Fetching fresh live data for {len(pending)} cities via Google Search ({model_name})...")
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
            future_map = {
                pool.submit(research_city, client, city, model_name): city
                for city in pending
            }
            for fut in concurrent.futures.as_completed(future_map):
                city = future_map[fut]
                completed_count += 1
                try:
                    result = fut.result()
                    cache[city] = result
                    cache_path.write_text(
                        json.dumps(cache, indent=2, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    src_count = len(result.get("sources", []))
                    logs.append(f"[OK] {city} — {src_count} sources ({result.get('coverage_status', 'ESTIMATED')})")
                except Exception as exc:
                    logs.append(f"[FAIL] {city} — {exc}")
                progress_bar.progress(completed_count / len(cities))
                log_area.code("\n".join(logs[-12:]))

    ordered_records = [cache[c] for c in cities if c in cache]
    if not ordered_records:
        status_box.error("No cities succeeded. Check your API key, selected model, or network connection.")
        st.stop()

    status_box.info("Building new 15-sheet Excel workbook with dynamic formulas & 90-day entry plan...")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        if uploaded_template is not None:
            tpl_path = tmp / "template.xlsx"
            tpl_path.write_bytes(uploaded_template.getvalue())
        else:
            tpl_path = local_default

        ts = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        out_name = f"Eighth_P_Live_{len(ordered_records)}Cities_{ts}.xlsx"
        out_path = tmp / out_name
        populate_workbook(tpl_path, out_path, ordered_records)
        xlsx_bytes = out_path.read_bytes()

        wb_out = openpyxl.load_workbook(out_path, data_only=False)
        ws15 = wb_out["15_Entry_Plan_90D"]
        plan_rows = []
        for r in range(5, ws15.max_row + 1):
            aid = ws15.cell(r, 1).value
            if aid and aid != "TOTAL":
                plan_rows.append(
                    {
                        "Action": aid,
                        "City": ws15.cell(r, 2).value,
                        "Phase": ws15.cell(r, 3).value,
                        "Priority Developers": ws15.cell(r, 5).value,
                        "Budget (INR L)": ws15.cell(r, 8).value,
                        "Milestone": ws15.cell(r, 9).value,
                        "Go / No-Go Gate": ws15.cell(r, 10).value,
                    }
                )

    status_box.success(f"Completed! {len(ordered_records)}/{len(cities)} cities populated into a fresh workbook.")

    st.download_button(
        label=f"Download {out_name}",
        data=xlsx_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    ranked = sorted([score_city_record(rec) for rec in ordered_records], key=lambda x: x["score"], reverse=True)
    tab_rec, tab_plan, tab_sources = st.tabs([
        "Management Recommendation & Ranking",
        "90-Day Entry Plan (Top 3 Cities)",
        "Sources & Data Gaps",
    ])

    with tab_rec:
        top5 = [c["city_name"] for c in ranked[:5]]
        next5 = [c["city_name"] for c in ranked[5:10]]
        m1, m2, m3 = st.columns(3)
        m1.metric("Priority 1 (Enter)", ", ".join(top5[:3]) + (f" +{len(top5)-3}" if len(top5) > 3 else ""))
        m2.metric("Priority 2 (Monitor)", ", ".join(next5[:3]) if next5 else "None")
        m3.metric("Total Y2 SOM (INR Cr)", f"{sum(c['som_y2_cr'] for c in ranked):.2f}")

        ranking_table = [
            {
                "Rank": idx + 1,
                "City": c["city_name"],
                "Score (0-100)": "INCOMPLETE" if c["score"] < 0 else c["score"],
                "Primary Mkt (INR Cr)": c["primary_mkt_cr"],
                "SAM (INR Cr)": c["sam_cr"],
                "Y2 SOM (INR Cr)": c["som_y2_cr"],
                "Y2 EBITDA (INR Cr)": c["y2_profit_cr"],
                "Payback (Mo)": c["payback_mo"],
                "Entry Model": c["entry_model"],
                "Confidence": f"{c['confidence']*100:.0f}%",
                "Rationale": c["rationale"],
            }
            for idx, c in enumerate(ranked)
        ]
        st.dataframe(ranking_table, use_container_width=True)

    with tab_plan:
        st.dataframe(plan_rows, use_container_width=True)

    with tab_sources:
        gap_rows = [
            {"City": rec["city_name"], "Coverage": rec.get("coverage_status"), "Data Gap / LDM Request": rec.get("data_gap_note", "None")}
            for rec in ordered_records
        ]
        st.subheader("Data Coverage & Gap Register")
        st.dataframe(gap_rows, use_container_width=True)

        src_rows = [
            {
                "City": rec["city_name"],
                "Authority": s.get("authority", ""),
                "Title": s.get("title", ""),
                "Period": s.get("period", ""),
                "URL": s.get("url", ""),
                "Status": s.get("classification", ""),
            }
            for rec in ordered_records
            for s in rec.get("sources", [])
        ]
        st.subheader("Verified Sources")
        st.dataframe(src_rows, use_container_width=True)
