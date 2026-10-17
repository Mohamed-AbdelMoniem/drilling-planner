import os
import streamlit as st
import pandas as pd
import docx
from docx.shared import Pt, RGBColor
from docxtpl import DocxTemplate
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy
from io import BytesIO
from pydantic import BaseModel, Field
from typing import List
from google import genai
from google.genai import types

# ==========================================
# 1. HYPER-DETAILED MULTI-AGENT SCHEMAS
# ==========================================

class DownholeEvent(BaseModel):
    date_or_year: str = Field(description="Date or year of the historical operation")
    depth_md_tvd: str = Field(description="Depth (MD/TVD) and hole/casing size")
    formation_or_zone: str = Field(description="Formation name or wellbore section")
    downhole_issue: str = Field(
        description="Strictly downhole wellbore issues (stuck, gas lock, losses, tight hole, fish, packer/plug issues, well control, casing damage). NO surface equipment or weather issues."
    )
    lesson_learned_mitigation: str = Field(description="Actionable technical mitigation for the upcoming operation")

class ClearanceAndSurveyCheck(BaseModel):
    depth_interval: str
    wellbore_restriction_or_dls: str = Field(description="Minimum ID restriction from sketch or high DLS from survey")
    planned_tool_or_bha_od: str = Field(description="Max OD of planned BHA, mill, packer, or completion string")
    risk_level: str = Field(description="HIGH, MEDIUM, or LOW")
    supervisor_action: str = Field(description="Required reaming, slow tripping, drift check, or centralizer action")

class PhaseGamePlan(BaseModel):
    phase_name: str = Field(description="e.g., Phase 2: RIH 6-1/8 Milling BHA & Mill Stuck Packer at 8,420 ft")
    depth_interval_and_bha: str = Field(description="Exact MD/TVD interval, BHA components with ODs/IDs/connections, and pipe tally notes")
    objective: str
    pre_job_and_safety_brief: List[str] = Field(description="PJSM topics, valve/manifold line-up checks, surface pressure test values, and barrier verification")
    step_by_step_execution: List[str] = Field(
        description="MINIMUM 8 to 15 ultra-detailed sequential rig-floor steps per phase. Include exact depths, tripping speeds, circulation rates, space-out, and tag procedures."
    )
    hydraulics_and_parameters: str = Field(description="Exact WOB, RPM, Torque limits, Pump Rate (BPM/GPM), Standpipe Pressure, Fluid weight/viscosity, and Max Overpull/Slack-off limits")
    wellbore_and_survey_hazards: List[str] = Field(description="Exact depths of high DLS, tight spots, liner tops, perforations, gas lock risks, or stuck pipe zones")
    if_then_contingencies: List[str] = Field(description="Concrete IF/THEN contingency actions for the rig supervisor if stuck, losses, high torque, or gas lock occurs")

class PrepItem(BaseModel):
    phase: str = Field(description="Operational phase this item is required for")
    category: str = Field(description="Tubulars, BHA/Milling/Fishing Tools, Crossovers & Handling Gear, Wellhead/BOP, Completion/Artificial Lift, Fluids/Chemicals, or 3rd-Party Services")
    item_description: str = Field(description="Exact tool, equipment, or fluid name and functional purpose")
    technical_specs_od_id_conn: str = Field(description="Exact Max OD, Min ID, Drift ID, Weight/Grade, Length, Fishneck OD, and Top/Bottom Connections")
    quantity_and_backup: str = Field(description="Exact primary quantity required + mandatory backup/spare units on location (e.g., 1 Primary + 1 Backup Mill, 100 bbls Hi-Vis Pill)")
    required_timing: str = Field(description="Exact mobilization lead time (e.g., 72 hrs prior to Spud, 48 hrs before Phase 2, Strapped on Catwalk prior to RIH)")
    verification_and_qaqc: str = Field(description="Rig-site supervisor QA/QC check: Drift with exact drift mandrel size, rabbiting, full tally, thread gauge,魚fishneck measurement, pressure test psi")
    responsible_party: str = Field(description="e.g., Day/Night Drilling Supervisor, Toolpusher, Mud Engineer, 3rd-Party Service Hand, Logistics")

class TimeEstimateRow(BaseModel):
    phase: str
    operation_step: str
    depth_interval: str
    estimated_hours: float
    operational_remarks: str

class MasterDrillingPlan(BaseModel):
    well_name: str
    operation_type: str = Field(description="e.g., Workover Milling & ESP Completion / Sidetrack Drilling")
    wellbore_summary_from_sketches: str = Field(description="Summary of casing strings, tubing, ID restrictions, TOL, PBTD, fish, packers from sketches")
    survey_critical_points: str = Field(description="Max inclination, max DLS and depth, kickoff point, keyseat/stuck risk zones")
    clearance_and_dls_audit: List[ClearanceAndSurveyCheck]
    historical_downhole_data: List[DownholeEvent]
    game_plan_phases: List[PhaseGamePlan]
    preparation_checklist: List[PrepItem] = Field(description="Comprehensive, granular preparation plan with at least 20 to 35 detailed items covering every tool, crossover, handling gear, fluid pill, and service across all phases")
    time_frame_schedule: List[TimeEstimateRow]


# ==========================================
# 2. BUILT-IN DEMO CASE (FOR INSTANT TESTING)
# ==========================================

DEMO_ENGINEERING_PROGRAM = """
WELL: PHIOPS-12 (Workover & Well Intervention Program)
OBJECTIVE: Kill well, pull existing 3-1/2" completion with ESP assembly suffering from gas lock and cable short, mill out stuck permanent packer at 8,420 ft MD, clean out fill to PBTD at 8,950 ft MD, and run new 3-1/2" completion with gas separator.

WELLBORE SKETCH SUMMARY:
- 9-5/8" Casing set at 7,800 ft MD (ID: 8.681", Drift: 8.525")
- 7" Liner hung at 7,600 ft to 9,050 ft MD (ID: 6.184", Drift: 6.059")
- Tight spot / casing deformation previously noted at 7,620 ft (top of 7" liner hanger)
- Existing Permanent Packer set at 8,420 ft MD (Min bore ID: 3.88")
- Perforations: 8,550 ft - 8,780 ft MD (Formation: Bahariya)

DIRECTIONAL SURVEY HIGHLIGHTS:
- Kickoff Point (KOP): 4,200 ft MD
- Max Inclination: 44.5 deg at 7,100 ft MD
- High Dogleg Severity (DLS): 6.8 deg/100ft between 6,350 ft and 6,500 ft MD (Keyseat & stuck risk)
- DLS across 7" Liner Top (7,600 ft): 4.9 deg/100ft

HISTORICAL DAILY REPORTS (OFFSET & OLD WORKOVERS):
- 2022 Drilling: Mud pump #2 liner failed (4 hrs surface downtime - IGNORE).
- 2022 Drilling: Encountered severe tight hole at 6,420 ft MD due to 6.8 deg/100ft DLS; worked pipe with 45 klbs overpull to free stuck BHA.
- 2023 Workover: Waiting on weather due to high wind (12 hrs - IGNORE).
- 2023 Workover: Partial mud losses (15 bbl/hr) into depleted Bahariya perforations at 8,600 ft MD while circulating 9.6 ppg brine; cured with CaCO3 pill.
- 2024 Intervention: Pump experienced repeated gas lock due to high GOR in Upper Bahariya zone. Wireline plug got stuck at 7,615 ft MD near liner top restriction.
"""


# ==========================================
# 3. FILE EXTRACTION ENGINE
# ==========================================

def extract_docx_text(uploaded_file) -> str:
    uploaded_file.seek(0)
    doc = docx.Document(uploaded_file)
    full_text = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            full_text.append(" | ".join(row_data))
    return "\n".join(full_text)

def extract_excel_text(uploaded_file) -> str:
    uploaded_file.seek(0)
    excel_data = pd.read_excel(uploaded_file, sheet_name=None)
    output = []
    for sheet_name, df in excel_data.items():
        output.append(f"--- SHEET: {sheet_name} ---")
        output.append(df.dropna(how='all').to_string(index=False))
    return "\n".join(output)


# ==========================================
# 4. WORD & EXCEL DELIVERABLE GENERATORS
# ==========================================

def copy_row_formatting(ws, source_row: int, target_row: int, max_col: int):
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        src = ws.cell(row=source_row, column=col)
        tgt = ws.cell(row=target_row, column=col)
        if src.has_style:
            tgt.font = copy(src.font)
            tgt.border = copy(src.border)
            tgt.fill = copy(src.fill)
            tgt.number_format = copy(src.number_format)
            tgt.alignment = copy(src.alignment)

def build_historical_word(plan: MasterDrillingPlan, corp_template=None) -> BytesIO:
    if corp_template:
        corp_template.seek(0)
        doc = DocxTemplate(corp_template)
        doc.render({
            "well_name": plan.well_name,
            "wellbore_summary": plan.wellbore_summary_from_sketches,
            "survey_summary": plan.survey_critical_points,
            "historical_downhole_data": [e.model_dump() for e in plan.historical_downhole_data],
        })
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    doc = docx.Document()
    doc.add_heading(f"Well Downhole Historical & Mechanical Analysis: {plan.well_name}", 0)
    
    doc.add_heading("1. Wellbore Schematic & Mechanical Status", level=1)
    doc.add_paragraph(plan.wellbore_summary_from_sketches)
    
    doc.add_heading("2. Directional Survey & Tortuosity Analysis", level=1)
    doc.add_paragraph(plan.survey_critical_points)

    doc.add_heading("3. BHA Clearance & Dogleg Severity (DLS) Audit", level=1)
    c_table = doc.add_table(rows=1, cols=5)
    c_table.style = 'Table Grid'
    c_headers = ["Depth Interval", "Restriction / DLS", "Planned Tool OD", "Risk", "Supervisor Action"]
    for i, h in enumerate(c_headers):
        cell = c_table.rows[0].cells[i]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
    for check in plan.clearance_and_dls_audit:
        row = c_table.add_row().cells
        row[0].text = check.depth_interval
        row[1].text = check.wellbore_restriction_or_
