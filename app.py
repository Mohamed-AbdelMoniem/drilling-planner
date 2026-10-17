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
# 1. MULTI-AGENT STRUCTURED OUTPUT SCHEMAS
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
    operation_type: str = Field(description="e.g., Workover Milling & ESP Completion / Sidetrack Drilling")
    wellbore_summary_from_sketches: str = Field(description="Summary of casing strings, tubing, ID restrictions, TOL, PBTD, fish, packers from sketches")
    survey_critical_points: str = Field(description="Max inclination, max DLS and depth, kickoff point, keyseat/stuck risk zones")
    clearance_and_dls_audit: List[ClearanceAndSurveyCheck]
    historical_downhole_data: List[DownholeEvent]
    game_plan_phases: List[PhaseGamePlan]
    preparation_checklist: List[PrepItem]
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
        row[1].text = check.wellbore_restriction_or_dls
        row[2].text = check.planned_tool_or_bha_od
        row[3].text = check.risk_level
        row[4].text = check.supervisor_action
    
    doc.add_heading("4. Filtered Downhole Historical Events (Surface & Weather Excluded)", level=1)
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
    doc.add_paragraph(f"Operation: {plan.operation_type}")
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
# 5. STREAMLIT MULTI-AGENT DASHBOARD UI
# ==========================================

st.set_page_config(page_title="RigOps AI — Drilling & Workover Strategic Planner", layout="wide")
st.title("🛢️ RigOps AI: Automated Drilling & Workover Strategic Planner")
st.caption("Multi-Agent Well Engineering & Rig Readiness System — Processes Programs, Well History, Surveys, and Sketches into 4 Field Deliverables.")

with st.sidebar:
    st.header("1. AI Engine Configuration")
    saved_key = ""
    try:
        saved_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        pass
        
    api_key = st.text_input("Gemini API Key", value=saved_key, type="password", help="Enter your Google AI Studio Gemini API Key")
    
    # Updated with active Gemini 3.x & 4 Argon models + automatic fallback list
    selected_model = st.selectbox(
        "Select Gemini Model", 
        [
            "gemini-3.8-flash",
            "gemini-3.1-pro-preview",
            "gemini-3-flash-preview",
            "gemini-3.5-flash",
            "gemini-3.1-flash-lite",
            "gemini-4-argon",
            "Custom Model ID..."
        ], 
        index=0
    )
    if selected_model == "Custom Model ID...":
        model_choice = st.text_input("Enter Custom Gemini Model ID", value="gemini-3.1-pro-preview")
    else:
        model_choice = selected_model
    
    st.divider()
    st.header("2. Optional Corporate Branding")
    st.caption("Leave empty to auto-generate formatted files, or upload your company templates.")
    corp_hist_tpl = st.file_uploader("Corporate History Word Template (.docx)", type=["docx"])
    corp_gp_tpl = st.file_uploader("Corporate Game Plan Word Template (.docx)", type=["docx"])
    corp_prep_tpl = st.file_uploader("Corporate Prep Plan Excel Template (.xlsx)", type=["xlsx"])
    corp_time_tpl = st.file_uploader("Corporate Time Frame Excel Template (.xlsx)", type=["xlsx"])

demo_col1, demo_col2 = st.columns([3, 1])
with demo_col1:
    st.info("💡 **Evaluating or testing on mobile?** Upload your well package below, or toggle **Demo Mode** to test the Multi-Agent system immediately with a pre-loaded workover & milling well package.")
with demo_col2:
    use_demo_mode = st.toggle("⚡ Enable Demo Well Package", value=False)

if not use_demo_mode:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Primary Well Inputs")
        eng_prog = st.file_uploader("1. Engineering Program (.docx) [Required]", type=["docx"])
        hist_files = st.file_uploader("2. Historical Well Files (.xlsx, .pdf)", type=["xlsx", "pdf"], accept_multiple_files=True)
        survey_file = st.file_uploader("3. Directional Survey (.xlsx, .pdf)", type=["xlsx", "pdf"])

    with col2:
        st.subheader("Sketches & Reference Plans")
        sketch_files = st.file_uploader("4. Drilling & Last Workover Sketches (.xlsx, .pdf)", type=["xlsx", "pdf"], accept_multiple_files=True)
        ref_files = st.file_uploader("5. Similar Game Plan & Prep Plan (.xlsx)", type=["xlsx"], accept_multiple_files=True)
else:
    st.success("✅ **Demo Well Package Loaded (Well PHIOPS-12):** Workover ESP pull, stuck packer milling at 8,420 ft, directional survey with 6.8°/100ft DLS, and historical reports containing surface noise + downhole losses/gas lock.")
    with st.expander("Preview Loaded Demo Well Package Data"):
        st.code(DEMO_ENGINEERING_PROGRAM)

if st.button("🚀 Run Multi-Agent Strategic Planner & Generate All 4 Deliverables", type="primary", use_container_width=True):
    if not api_key:
        st.error("Please enter your Gemini API Key in the left sidebar.")
    elif not use_demo_mode and not eng_prog:
        st.error("Please upload the Engineering Program (.docx) or enable Demo Mode to proceed.")
    else:
        with st.spinner(f"Executing Multi-Agent Pipeline using {model_choice}..."):
            try:
                client = genai.Client(api_key=api_key)
                contents = []

                system_prompt = """
                You are an elite Multi-Agent Drilling & Workover Supervision System consisting of 4 specialized agents:
                
                - AGENT 1 (DOWNHOLE HISTORICAL FILTER):
                  Extract ONLY wellbore/downhole events tied to depths and formations (stuck pipe/tools, tight spots, mud losses, kicks, gas lock, fishing, milling, packer/plug issues, casing damage).
                  STRICTLY EXCLUDE surface equipment issues (mud pumps, top drive, generators, rig mechanic downtime) and weather delays.
                  Always use exact English industry terminology ("stuck", "gas lock", "tight hole", "lost circulation") and omit financial/cost figures.
                
                - AGENT 2 (WELLBORE GEOMETRY & SURVEY AUDITOR):
                  Cross-check planned BHA, milling tools, and completion ODs against casing/liner drift IDs and restrictions from the sketches.
                  Identify high Dogleg Severity (DLS) intervals from the Directional Survey and prescribe reaming/tripping precautions.
                
                - AGENT 3 (STRATEGIC GAME PLAN & READINESS PLANNER):
                  Convert the Engineering Program into a phase-by-phase rig-floor execution guide integrating restrictions, fish tops, and survey hazards.
                  Generate a complete equipment & material Preparation Checklist with required timing and site verification checks (drifting, tallying, pressure testing).
                
                - AGENT 4 (OPERATIONAL TIME ESTIMATOR):
                  Break down the operation into sequential steps with realistic hour estimates based on tripping speeds, circulation, milling/drilling rates, and rig-up/rig-down times.
                """

                if use_demo_mode:
                    contents.append(DEMO_ENGINEERING_PROGRAM)
                else:
                    contents.append(f"=== 1. ENGINEERING PROGRAM ===\n{extract_docx_text(eng_prog)}")

                    file_groups = [
                        ("2. HISTORICAL WELL FILE", hist_files),
                        ("3. DIRECTIONAL SURVEY", [survey_file] if survey_file else []),
                        ("4. WELLBORE & WORKOVER SKETCH", sketch_files),
                        ("5. REFERENCE GAME PLAN & PREP PLAN", ref_files),
                    ]

                    for group_title, flist in file_groups:
                        for f in flist:
                            if f is not None:
                                if f.name.lower().endswith(".xlsx"):
                                    contents.append(f"=== {group_title} ({f.name}) ===\n{extract_excel_text(f)}")
                                elif f.name.lower().endswith(".pdf"):
                                    f.seek(0)
                                    pdf_bytes = f.read()
                                    contents.append(f"=== {group_title} PDF ({f.name}) ===")
                                    contents.append(types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"))

                # Automatic fallback loop across active Gemini 3.x / 4 models if one is unavailable on the user's tier
                fallback_models = [model_choice, "gemini-3.8-flash", "gemini-3.1-pro-preview", "gemini-3-flash-preview", "gemini-3.5-flash"]
                seen = set()
                ordered_models = [m for m in fallback_models if not (m in seen or seen.add(m))]
                
                last_err = None
                response = None
                used_model = model_choice
                for m_name in ordered_models:
                    try:
                        response = client.models.generate_content(
                            model=m_name,
                            contents=contents,
                            config=types.GenerateContentConfig(
                                system_instruction=system_prompt,
                                response_mime_type="application/json",
                                response_schema=MasterDrillingPlan,
                                temperature=0.1,
                            ),
                        )
                        used_model = m_name
                        break
                    except Exception as loop_err:
                        last_err = loop_err
                        if "404" in str(loop_err) or "NOT_FOUND" in str(loop_err):
                            continue
                        else:
                            raise loop_err

                if response is None:
                    raise last_err

                st.session_state["plan_result"] = MasterDrillingPlan.model_validate_json(response.text)
                st.success(f"Successfully generated all 4 deliverables for Well: {st.session_state['plan_result'].well_name} (Model: {used_model})")

            except Exception as e:
                st.error(f"Error during processing: {e}")


# ==========================================
# 6. EXECUTIVE KPI SUMMARY & DOWNLOAD CENTER
# ==========================================

if "plan_result" in st.session_state:
    plan: MasterDrillingPlan = st.session_state["plan_result"]
    
    st.divider()
    
    total_hours = sum(step.estimated_hours for step in plan.time_frame_schedule)
    total_days = round(total_hours / 24.0, 1)
    high_risk_count = sum(1 for c in plan.clearance_and_dls_audit if "HIGH" in c.risk_level.upper())
    
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Well Name", plan.well_name)
    k2.metric("Total Expected Time", f"{total_hours:.1f} hrs ({total_days} Days)")
    k3.metric("Filtered Downhole Events", f"{len(plan.historical_downhole_data)} Events")
    k4.metric("High-Risk Clearance/DLS Alerts", f"{high_risk_count} Critical Zones")

    st.subheader(f"📥 Download Generated Deliverables — {plan.well_name}")
    
    dl1, dl2, dl3, dl4, dl5 = st.columns(5)
    with dl1:
        st.download_button(
            "📄 1. Historical Data (.docx)",
            data=build_historical_word(plan, corp_hist_tpl),
            file_name=f"{plan.well_name}_1_Historical_Data.docx",
            use_container_width=True
        )
    with dl2:
        st.download_button(
            "📘 2. Game Plan (.docx)",
            data=build_game_plan_word(plan, corp_gp_tpl),
            file_name=f"{plan.well_name}_2_Game_Plan.docx",
            use_container_width=True
        )
    with dl3:
        st.download_button(
            "📊 3A. Prep Plan (.xlsx)",
            data=build_prep_excel(plan, corp_prep_tpl),
            file_name=f"{plan.well_name}_3_Preparation_Plan.xlsx",
            use_container_width=True
        )
    with dl4:
        st.download_button(
            "📝 3B. Prep Plan (.docx)",
            data=build_prep_plan_word(plan),
            file_name=f"{plan.well_name}_3_Preparation_Plan.docx",
            use_container_width=True
        )
    with dl5:
        st.download_button(
            "⏱️ 4. Time Frame (.xlsx)",
            data=build_time_frame_excel(plan, corp_time_tpl),
            file_name=f"{plan.well_name}_4_Expected_Time_Frame.xlsx",
            use_container_width=True
        )

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "1. Downhole History (Filtered)", 
        "2. Execution Game Plan", 
        "3. Preparation Plan", 
        "4. Expected Time Frame",
        "5. BHA Clearance & DLS Audit"
    ])
    
    with tab1:
        st.markdown(f"**Wellbore Summary (From Sketches):** {plan.wellbore_summary_from_sketches}")
        st.markdown(f"**Directional Survey Critical Points:** {plan.survey_critical_points}")
        st.dataframe(pd.DataFrame([e.model_dump() for e in plan.historical_downhole_data]), use_container_width=True)
        
    with tab2:
        for p in plan.game_plan_phases:
            with st.expander(f"Phase: {p.phase_name} — {p.objective}", expanded=True):
                st.markdown("**Execution Steps:**")
                for s in p.step_by_step_execution:
                    st.markdown(f"- {s}")
                st.markdown("**Wellbore & Survey Hazards:**")
                for hazard in p.wellbore_and_survey_hazards:
                    st.markdown(f"- ⚠️ **{hazard}**")
                st.info(f"**Critical Parameters:** {p.critical_parameters}")
                
    with tab3:
        st.dataframe(pd.DataFrame([i.model_dump() for i in plan.preparation_checklist]), use_container_width=True)
        
    with tab4:
        df_time = pd.DataFrame([t.model_dump() for t in plan.time_frame_schedule])
        df_time["Est. Days"] = (df_time["estimated_hours"] / 24.0).round(2)
        df_time["Cum. Hours"] = df_time["estimated_hours"].cumsum()
        df_time["Cum. Days"] = (df_time["Cum. Hours"] / 24.0).round(2)
        st.dataframe(df_time, use_container_width=True)
        
        st.markdown("#### Phase Duration Breakdown (Hours)")
        phase_chart = df_time.groupby("phase", as_index=False)["estimated_hours"].sum()
        st.bar_chart(phase_chart.set_index("phase"))

    with tab5:
        st.markdown("#### Automated BHA OD vs. Wellbore ID & Dogleg Severity Audit")
        st.dataframe(pd.DataFrame([c.model_dump() for c in plan.clearance_and_dls_audit]), use_container_width=True)
