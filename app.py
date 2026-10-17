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
# 1. STRUCTURED DATA SCHEMAS (AI OUTPUT)
# ==========================================

class DownholeEvent(BaseModel):
    date_or_year: str = Field(description="Date or year of the historical operation")
    depth_md_tvd: str = Field(description="Depth (MD/TVD) and hole/casing size")
    formation_or_zone: str = Field(description="Formation name or wellbore section")
    downhole_issue: str = Field(
        description="Strictly downhole wellbore issues (stuck, gas lock, losses, tight hole, fish, packer/plug issues, well control, casing damage). NO surface equipment or weather issues."
    )
    lesson_learned_mitigation: str = Field(description="Actionable technical mitigation for the upcoming operation")

class PhaseGamePlan(BaseModel):
    phase_name: str
    objective: str
    step_by_step_execution: List[str] = Field(description="Detailed sequential rig supervisor execution steps")
    wellbore_and_survey_hazards: List[str] = Field(description="Specific DLS, ID restrictions, stuck risks, gas lock risks, fish/packers from sketches & survey")
    critical_parameters: str = Field(description="WOB, RPM, Pump Rate, Mud Weight, Overpull limits, test pressures")

class PrepItem(BaseModel):
    phase: str
    category: str = Field(description="Tubulars, BHA/Tools, Wellhead/BOP, Fluids/Chemicals, or 3rd Party Services")
    item_description: str = Field(description="Exact tool, crossover, fishing tool, or material with OD/ID/Connection specs")
    required_timing: str = Field(description="e.g., On Location Prior to Rig Up, 48 hrs before Phase 2")
    verification_check: str = Field(description="Drift, pressure test, thread inspection, rabbiting, tallying")

class TimeEstimateRow(BaseModel):
    phase: str
    operation_step: str
    depth_interval: str
    estimated_hours: float
    operational_remarks: str

class MasterDrillingPlan(BaseModel):
    well_name: str
    wellbore_summary_from_sketches: str = Field(description="Detailed summary of casing strings, tubing, ID restrictions, TOL, PBTD, fish, packers from sketches")
    survey_critical_points: str = Field(description="Max inclination, max DLS and depth, kickoff point, keyseat/stuck risk zones")
    historical_downhole_data: List[DownholeEvent]
    game_plan_phases: List[PhaseGamePlan]
    preparation_checklist: List[PrepItem]
    time_frame_schedule: List[TimeEstimateRow]


# ==========================================
# 2. FILE EXTRACTION ENGINE
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
# 3. DUAL-MODE WORD & EXCEL GENERATORS
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
    doc.add_heading(f"Well Downhole Historical Analysis: {plan.well_name}", 0)
    
    doc.add_heading("1. Wellbore Schematic & Mechanical Status", level=1)
    doc.add_paragraph(plan.wellbore_summary_from_sketches)
    
    doc.add_heading("2. Directional Survey & Tortuosity Analysis", level=1)
    doc.add_paragraph(plan.survey_critical_points)
    
    doc.add_heading("3. Filtered Downhole Historical Events (Surface & Weather Excluded)", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["Date/Year", "Depth (MD/TVD)", "Zone/Section", "Downhole Issue", "Mitigation / Action"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
            
    for event in plan.historical_downhole_data:
        row = table.add_row().cells
        row[0].text = event.date_or_year
        row[1].text = event.depth_md_tvd
        row[2].text = event.formation_or_zone
        row[3].text = event.downhole_issue
        row[4].text = event.lesson_learned_mitigation
        
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def build_game_plan_word(plan: MasterDrillingPlan, corp_template=None) -> BytesIO:
    if corp_template:
        corp_template.seek(0)
        doc = DocxTemplate(corp_template)
        doc.render({
            "well_name": plan.well_name,
            "wellbore_summary": plan.wellbore_summary_from_sketches,
            "survey_summary": plan.survey_critical_points,
            "game_plan_phases": [p.model_dump() for p in plan.game_plan_phases],
        })
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    doc = docx.Document()
    doc.add_heading(f"Strategic Execution Game Plan: {plan.well_name}", 0)
    doc.add_paragraph(f"Wellbore Status: {plan.wellbore_summary_from_sketches}")
    doc.add_paragraph(f"Directional Hazards: {plan.survey_critical_points}")
    
    for idx, phase in enumerate(plan.game_plan_phases, 1):
        doc.add_heading(f"Phase {idx}: {phase.phase_name}", level=1)
        p_obj = doc.add_paragraph()
        r_obj = p_obj.add_run(f"Objective: {phase.objective}")
        r_obj.bold = True
        
        doc.add_heading("Rig Execution Steps:", level=2)
        for step in phase.step_by_step_execution:
            doc.add_paragraph(step, style='List Number')
            
        doc.add_heading("Wellbore, Sketch & Survey Hazards:", level=2)
        for hazard in phase.wellbore_and_survey_hazards:
            p_haz = doc.add_paragraph(style='List Bullet')
            run_haz = p_haz.add_run(hazard)
            if any(k in hazard.lower() for k in ["stuck", "gas lock", "loss", "dls", "fish", "tight"]):
                run_haz.font.color.rgb = RGBColor(192, 0, 0)
                run_haz.bold = True
                
        doc.add_heading("Critical Operational Parameters:", level=2)
        doc.add_paragraph(phase.critical_parameters)
        
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def build_prep_plan_word(plan: MasterDrillingPlan) -> BytesIO:
    doc = docx.Document()
    doc.add_heading(f"Operational Preparation & Readiness Plan: {plan.well_name}", 0)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ["Phase", "Category", "Equipment / Specification", "Required Timing", "Site Verification"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
            
    for item in plan.preparation_checklist:
        row = table.add_row().cells
        row[0].text = item.phase
        row[1].text = item.category
        row[2].text = item.item_description
        row[3].text = item.required_timing
        row[4].text = item.verification_check
        
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def build_prep_excel(plan: MasterDrillingPlan, corp_template=None, start_row: int = 7) -> BytesIO:
    if corp_template:
        corp_template.seek(0)
        wb = openpyxl.load_workbook(corp_template)
        ws = wb.active
        ws["C3"] = plan.well_name
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Preparation_Plan"
        start_row = 4
        ws["A1"] = f"OPERATIONAL PREPARATION PLAN - {plan.well_name}"
        ws["A1"].font = Font(size=14, bold=True, color="1F4E78")
        headers = ["Phase", "Category", "Equipment / Material Specification", "Required Timing", "Site Verification Check", "Status"]
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for col_idx, text in enumerate(headers, 1):
            c = ws.cell(row=3, column=col_idx, value=text)
            c.fill = header_fill
            c.font = header_font

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for idx, item in enumerate(plan.preparation_checklist):
        r = start_row + idx
        if corp_template and idx > 0:
            copy_row_formatting(ws, source_row=start_row, target_row=r, max_col=6)
            
        vals = [item.phase, item.category, item.item_description, item.required_timing, item.verification_check, "Pending"]
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            if not corp_template:
                cell.border = thin_border
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    if not corp_template:
        widths = [20, 18, 45, 22, 30, 15]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def build_time_frame_excel(plan: MasterDrillingPlan, corp_template=None, start_row: int = 7) -> BytesIO:
    if corp_template:
        corp_template.seek(0)
        wb = openpyxl.load_workbook(corp_template)
        ws = wb.active
        ws["C3"] = plan.well_name
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Time_Frame_Schedule"
        start_row = 4
        ws["A1"] = f"EXPECTED OPERATIONAL TIME FRAME - {plan.well_name}"
        ws["A1"].font = Font(size=14, bold=True, color="1F4E78")
        headers = ["Step #", "Phase", "Operation Description", "Depth Interval", "Est. Hours", "Est. Days", "Cum. Hours", "Cum. Days", "Operational Remarks"]
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for col_idx, text in enumerate(headers, 1):
            c = ws.cell(row=3, column=col_idx, value=text)
            c.fill = header_fill
            c.font = header_font

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    risk_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    for idx, step in enumerate(plan.time_frame_schedule):
        r = start_row + idx
        if corp_template and idx > 0:
            copy_row_formatting(ws, source_row=start_row, target_row=r, max_col=9)

        ws.cell(row=r, column=1, value=idx + 1)
        ws.cell(row=r, column=2, value=step.phase)
        ws.cell(row=r, column=3, value=step.operation_step)
        ws.cell(row=r, column=4, value=step.depth_interval)
        ws.cell(row=r, column=5, value=step.estimated_hours)
        ws.cell(row=r, column=6, value=f"=ROUND(E{r}/24, 2)")
        ws.cell(row=r, column=7, value=f"=E{r}" if idx == 0 else f"=G{r-1}+E{r}")
        ws.cell(row=r, column=8, value=f"=ROUND(G{r}/24, 2)")
        ws.cell(row=r, column=9, value=step.operational_remarks)

        is_risk = any(k in (step.operation_step + " " + step.operational_remarks).lower() for k in ["stuck", "gas lock", "milling", "fishing", "loss", "dls"])
        for c_idx in range(1, 10):
            cell = ws.cell(row=r, column=c_idx)
            if not corp_template:
                cell.border = thin_border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if is_risk:
                    cell.fill = risk_fill

    if not corp_template:
        widths = [8, 20, 42, 16, 12, 12, 14, 14, 38]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ==========================================
# 4. STREAMLIT DASHBOARD UI
# ==========================================

st.set_page_config(page_title="Automated Drilling & Workover Strategic Planner", layout="wide")
st.title("Automated Drilling & Workover Strategic Planner")
st.caption("Upload your Engineering Program, Well History, Survey, Sketches, and Reference Plans to automatically generate all 4 deliverables.")

with st.sidebar:
    st.header("1. AI Engine Configuration")
    saved_key = ""
    try:
        saved_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        pass
        
    api_key = st.text_input("Gemini API Key", value=saved_key, type="password", help="Enter your Google AI Studio Gemini API Key")
    model_choice = st.selectbox("Select Gemini Model", ["gemini-2.5-pro", "gemini-2.5-flash"], index=0)
    
    st.divider()
    st.header("2. Optional Corporate Branding")
    st.caption("Leave empty to auto-generate formatted files, or upload your company templates.")
    corp_hist_tpl = st.file_uploader("Corporate History Word Template (.docx)", type=["docx"])
    corp_gp_tpl = st.file_uploader("Corporate Game Plan Word Template (.docx)", type=["docx"])
    corp_prep_tpl = st.file_uploader("Corporate Prep Plan Excel Template (.xlsx)", type=["xlsx"])
    corp_time_tpl = st.file_uploader("Corporate Time Frame Excel Template (.xlsx)", type=["xlsx"])

col1, col2 = st.columns(2)
