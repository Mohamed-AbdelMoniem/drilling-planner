# 🛢️ RigOps AI: Automated Drilling & Workover Strategic Planner

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://drilling-planner-jqh7bnhxmuktjtqbmwrvej.streamlit.app/)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue)
![Gemini 2.5 Pro](https://img.shields.io/badge/AI_Engine-Gemini_2.5_Pro-orange)

**Live Application:** [Launch RigOps AI Strategic Planner](https://drilling-planner-jqh7bnhxmuktjtqbmwrvej.streamlit.app/)

---

## 🎯 Executive Summary
**RigOps AI** is a Multi-Agent Well Engineering and Rig Supervision system designed to automate well preparation, downhole risk auditing, and operational game planning. 

Instead of spending days manually cross-referencing offset daily reports, directional surveys, and scanned wellbore schematics, drilling and workover supervisors can upload their raw well package (or use the built-in **1-Click Demo Mode**) to automatically generate **4 field-ready deliverables** in Word (`.docx`) and dynamic Excel (`.xlsx`) formats.

---

## 🧠 Multi-Agent System Architecture

```text
[Uploaded Well Package]
  ├── 1. Engineering Program (.docx)        --> Operational Objectives & Target Depths
  ├── 2. Historical Well Files (.xlsx/.pdf)  --> Offset Well Performance & Legacy Logs
  ├── 3. Directional Survey (.xlsx/.pdf)     --> Inclination, Azimuth, and DLS Doglegs
  ├── 4. Wellbore & Workover Sketches (PDF)  --> Visual Casing, Tubing, TOL, Fish & Packers
  └── 5. Reference Templates (.xlsx)         --> Baseline Preparation & Rig Readiness Formats
                 │
                 ▼
     [Gemini 2.5 Multimodal Multi-Agent Pipeline]
                 │
                 ├── Agent 1 (Downhole Filter): Strips surface/weather NPT; isolates stuck, gas lock, losses & kicks
                 ├── Agent 2 (Geometry Auditor): Audits BHA OD vs. Casing/Liner drift IDs & high DLS intervals
                 ├── Agent 3 (Game Plan & Prep): Builds phase-by-phase execution steps & tangible readiness checklists
                 └── Agent 4 (Time Estimator): Calculates operational hours & injects live Excel formulas
                 │
                 ▼
[4 Downloadable Deliverables: Formatted .docx & Formula-Driven .xlsx]
