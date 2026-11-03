"""Build the input workbooks for allocate.py:
   project_template.xlsx  – a realistic example project (façade panel installation)
   worked_example.xlsx    – the 3-task worked example used on the slides
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

DIR = os.path.dirname(os.path.abspath(__file__)) + os.sep
F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor="14213D")
INPUT = Font(name=F, color="0000FF")
BASE, BOLD = Font(name=F), Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, size=14)
NOTE = Font(name=F, italic=True, color="5B6577")
WRAP = Alignment(wrap_text=True, vertical="top")

CRIT = [("C1", "Task duration"), ("C2", "Cost"), ("C3", "Safety risk exposure"),
        ("C4", "Physical demand and fatigue"), ("C5", "Cognitive workload"),
        ("C6", "Task complexity"), ("C7", "Standardisation and supervisability")]
WEIGHTS = [0.15, 0.10, 0.30, 0.15, 0.10, 0.10, 0.10]
RUBRIC = [
    ("C1", "Task duration", "Computed: min(1, duration ÷ reference duration)", "", "", ""),
    ("C2", "Cost", "Computed: min(1, cost rate × duration ÷ reference cost)", "", "", ""),
    ("C3", "Safety risk exposure", "Rating 1–5 from the site risk assessment (likelihood × severity) for the people involved in this option",
     "Unlikely, minor injury at most", "Possible injury needing treatment", "Likely harm, or serious injury possible"),
    ("C4", "Physical demand and fatigue", "Rating 1–5 for the workers involved (load, posture, duration)",
     "Light, short, neutral posture", "Moderate loads or repeated awkward postures", "Heavy handling, prolonged or overhead work"),
    ("C5", "Cognitive workload", "Rating 1–5 for the workers involved (attention, decisions, supervision)",
     "Routine, little attention", "Regular checking or coordination", "Continuous attention, frequent decisions or supervising a robot"),
    ("C6", "Task complexity", "Rating 1–5 of the task done this way (steps, precision, coordination)",
     "Few simple steps, wide tolerance", "Several steps or moderate precision", "Many steps, tight tolerance or close human–robot coordination"),
    ("C7", "Poor fit with task structure", "Rating 1–5: how badly the option suits how repetitive and easy to monitor the task is",
     "Good fit (robot on repetitive, monitored work)", "Partial fit", "Poor fit (robot on non-standard work, worker on highly repetitive work)"),
]


def header(ws, cols, widths, height=32):
    for j, h in enumerate(cols, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font, c.fill, c.alignment = H, HF, Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = widths[j - 1]
    ws.row_dimensions[1].height = height


def build(out, ex):
    wb = openpyxl.Workbook()

    # README
    ws = wb.active
    ws.title = "README"
    lines = [(ex["title"], TITLE), ("", BASE)] + [(l, NOTE if l.startswith("Example") else BASE) for l in ex["intro"]] + [
        ("", BASE),
        ("Who provides what", BOLD),
        ("Experts, once (questionnaire): the criterion weights (Part C, BWM), the hard rules (Part E) and the scoring rubric (Part D).", BASE),
        ("Each project: the crew and rules (Settings), the tasks (Tasks) and the data for every feasible task–option pair (Options).", BASE),
        ("", BASE),
        ("Steps for a new project", BOLD),
        ("1. Weights: enter the BWM weights of C1–C7 (the same for every project).", BASE),
        ("2. Settings: crew, cost rates, λ, safety rules. Keep the two reference values the same across projects so that scores stay comparable.", BASE),
        ("3. Tasks: one row per task with its predecessors. Per panel = Yes repeats the task for every panel (Settings › Panels); "
         "After previous panel names a task of the previous panel that must finish first. Task type and zone are for reference.", BASE),
        ("4. Options: one row per feasible task–option pair (W, R or HR) with duration, worker exposure and ratings 1–5 for C3–C7 "
         "(see Scoring rubric). No row = that option cannot do the task.", BASE),
        ("5. Run: python allocate.py <file>.xlsx, or load the file in the web page opened with python app.py.", BASE),
        ("", BASE),
        ("How scores are built", BOLD),
        ("C1 = min(1, duration ÷ reference duration).  C2 = min(1, cost rate × duration ÷ reference cost).  C3–C7 = (rating − 1) ÷ 4.", BASE),
        ("C(i,k) = Σ w(j) · s(i,j,k).  Objective: minimise Σ C(i,k)·x(i,k) + λ · makespan (minutes).", BASE),
        ("", BASE),
        ("Blue text = a value you enter. Dark header = do not edit. Durations in multiples of 5 minutes keep the Timeline sheet exact.", BASE),
    ]
    for i, (txt, f) in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=txt).font = f
    ws.column_dimensions["A"].width = 150

    # Settings
    ws = wb.create_sheet("Settings")
    header(ws, ["Setting", "Value", "Unit / allowed values", "Meaning"], [24, 12, 26, 80])
    rows = [
        ("Workers", ex["nW"], "whole number ≥ 1", "Number of workers available"),
        ("Robots", ex["nR"], "whole number ≥ 0", "Number of robots available"),
        ("Cost rate W", ex["rates"][0], ex["cost_unit"] + " per minute", "One worker working alone"),
        ("Cost rate R", ex["rates"][1], ex["cost_unit"] + " per minute", "One robot working alone (hire, energy, set-up)"),
        ("Cost rate HR", ex["rates"][2], ex["cost_unit"] + " per minute", "One worker and one robot working together"),
        ("Reference duration", ex["ref"][0], "minutes", "Duration that scores 1 on C1; keep the same across projects"),
        ("Reference cost", ex["ref"][1], ex["cost_unit"], "Cost that scores 1 on C2; keep the same across projects"),
        ("Lambda", ex["lam"], "score units per minute", "Weight on makespan in the objective; 0 ignores makespan"),
        ("Zone rule", ex["zone"], "Yes / No", "Yes: a worker alone and a robot alone never work at the same time"),
        ("Exposure cap", ex["cap"], "minutes, blank = off", "Upper limit on total worker minutes in a hazard zone"),
        ("Time limit", 30, "seconds", "Maximum CP-SAT solving time"),
        ("Panels", ex.get("panels", 1), "whole number ≥ 1", "How many times the per-panel tasks are repeated (Tasks › Per panel = Yes)"),
    ]
    for i, (k, v, u, m) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=k).font = BOLD
        ws.cell(row=i, column=2, value=v).font = INPUT
        ws.cell(row=i, column=3, value=u).font = BASE
        ws.cell(row=i, column=4, value=m).font = BASE
    for cell, txt in ex.get("setting_notes", {}).items():
        ws[cell].comment = Comment(txt, "template")
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B10")

    # Weights
    ws = wb.create_sheet("Weights")
    header(ws, ["Code", "Criterion", "Weight (from BWM)", "Normalised"], [8, 38, 18, 14])
    for i, ((code, name), w) in enumerate(zip(CRIT, WEIGHTS), start=2):
        ws.cell(row=i, column=1, value=code).font = BOLD
        ws.cell(row=i, column=2, value=name).font = BASE
        ws.cell(row=i, column=3, value=w).font = INPUT
        n = ws.cell(row=i, column=4, value=f"=C{i}/SUM($C$2:$C$8)")
        n.font, n.number_format = BASE, "0.000"
    ws.cell(row=9, column=2, value="Total").font = BOLD
    ws.cell(row=9, column=3, value="=SUM(C2:C8)").font = BOLD
    ws.cell(row=9, column=4, value="=SUM(D2:D8)").font = BOLD
    ws.cell(row=9, column=4).number_format = "0.000"
    ws.cell(row=11, column=1, value="Example weights; replace with the aggregated BWM weights from the questionnaire.").font = NOTE

    # Scoring rubric
    ws = wb.create_sheet("Scoring rubric")
    header(ws, ["Code", "Criterion", "How it is scored", "1 = best", "3", "5 = worst"], [8, 26, 44, 34, 34, 38])
    for i, row in enumerate(RUBRIC, start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.alignment = (BOLD if j == 1 else BASE), WRAP
    ws.cell(row=10, column=1, value="Draft anchors, to be confirmed by the experts (questionnaire Part D). "
            "2 and 4 lie between the anchors; averages of several raters may be decimals.").font = NOTE

    # Tasks
    ws = wb.create_sheet("Tasks")
    header(ws, ["Task ID", "Task name", "Predecessors", "Per panel", "After previous panel", "Task type", "Zone", "Notes"],
           [9, 30, 14, 10, 14, 26, 14, 58])
    for i, row in enumerate(ex["tasks"], start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = INPUT if j <= 7 else BASE
            c.alignment = WRAP
    dvp = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dvp)
    dvp.add("D2:D500")
    ws["D2"].comment = Comment("Yes: this task is done once for every panel (Settings › Panels). IDs become P1A, P1B, …", "template")
    ws["E2"].comment = Comment("Optional: a per-panel task of the PREVIOUS panel that must finish first, e.g. P4 (its brackets fixed).", "template")
    ws["C2"].comment = Comment("Comma-separated IDs of tasks that must finish first, e.g. T1, T4", "template")
    ws.freeze_panes = "C2"

    # Options
    ws = wb.create_sheet("Options")
    cols = ["Task ID", "Option", "Duration (min)", "Worker exposure (min)", "C3 Safety risk (1–5)", "C4 Physical load (1–5)",
            "C5 Cognitive load (1–5)", "C6 Complexity (1–5)", "C7 Poor fit (1–5)", "Notes"]
    header(ws, cols, [9, 8, 11, 12, 12, 12, 12, 12, 11, 70], 44)
    for i, row in enumerate(ex["options"], start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = INPUT if j <= 9 else BASE
            if j == 10:
                c.alignment = WRAP
    dvo = DataValidation(type="list", formula1='"W,R,HR"', allow_blank=False)
    ws.add_data_validation(dvo)
    dvo.add("B2:B2000")
    dvr = DataValidation(type="decimal", operator="between", formula1="1", formula2="5", showErrorMessage=True,
                         errorTitle="Rating out of range", error="Ratings are between 1 and 5.")
    ws.add_data_validation(dvr)
    dvr.add("E2:I2000")
    ws["B2"].comment = Comment("W = worker alone, R = robot alone, HR = one worker and one robot together. "
                               "Add a row only for options that can do the task.", "template")
    ws.freeze_panes = "C2"

    # Assumptions
    if ex.get("assumptions"):
        ws = wb.create_sheet("Assumptions")
        header(ws, ["Item", "Assumption used in this example", "Replace with"], [26, 90, 48])
        for i, row in enumerate(ex["assumptions"], start=2):
            for j, v in enumerate(row, start=1):
                c = ws.cell(row=i, column=j, value=v)
                c.font, c.alignment = (BOLD if j == 1 else BASE), WRAP

    wb.save(out)
    print("saved", out)


# ============================================================================ realistic example
P = "Level 2 façade, bays A–B"
realistic = {
    "title": "Example project: installation of prefabricated façade panels (Level 2)",
    "intro": [
        "Storey-height prefabricated façade panels (about 300 kg each) are installed on the Level 2 slab edge; Settings › Panels sets how many (2 in this example).",
        "Crew: two workers and one collaborative mobile robot with a lifting-assist arm and a drilling end-effector.",
        "Each panel needs anchor drilling, lifting and positioning, alignment and levelling, bracket fixing and joint sealing; "
        "setting out comes first and a final inspection last.",
        "Example values are plausible assumptions for demonstration, not site measurements. See the Assumptions sheet.",
    ],
    "nW": 2, "nR": 1, "rates": [0.55, 0.75, 1.30], "cost_unit": "GBP", "ref": [60, 40],
    "lam": 0.02, "zone": "Yes", "cap": None,
    "setting_notes": {
        "B4": "About £33/h all-in labour rate; assumption.",
        "B5": "About £45/h robot hire, energy and set-up; assumption.",
        "B8": "Reference values fixed for all façade projects of this kind: 60 min and £40 score 1 on C1 and C2.",
    },
    "panels": 2,
    "tasks": [
        ("T1", "Set out bracket positions", "", "No", "", "TT8 Measuring and layout", "Slab edge", "Mark anchor and bracket positions for all panels"),
        ("P1", "Drill anchors", "T1", "Yes", "", "TT5 Drilling and cutting", "Panel bay edge", "Six M12 anchor holes into the slab edge"),
        ("P2", "Lift and position", "P1", "Yes", "P4", "TT1 Heavy lifting", "Panel bay edge", "≈300 kg: no manual lifting. Waits for the previous panel's brackets (shared lifting route)"),
        ("P3", "Align and level", "P2", "Yes", "", "TT3 Precise alignment", "Panel bay edge", "±3 mm tolerance; needs human judgement"),
        ("P4", "Fix brackets", "P3", "Yes", "", "TT4 Fastening", "Panel bay edge", "Torque-controlled bolting"),
        ("P5", "Seal joints", "P4", "Yes", "P5", "TT6 Joining and sealing", "Panel bay edge", "Waits for the previous panel's sealing (shared joint)"),
        ("T2", "Final inspection", "P5", "No", "", "TT9 Inspection", "All bays", "After every panel is sealed: check alignment, fixings and seals"),
    ],
    "options": [],
    "assumptions": [
        ("Panels", "Two storey-height prefabricated façade panels, about 300 kg each, installed from the slab edge.", "Panel schedule and weights from the drawings"),
        ("Robot", "One collaborative mobile robot with a lifting-assist arm and interchangeable drilling, bolting and sealing tools.", "The robot actually planned, with its datasheet"),
        ("Durations", "Rounded to 5 minutes; estimated, not measured. Robot alone is faster on repetitive drilling and bolting, slower on sealing.", "Work-study data, robot trials or supplier figures"),
        ("Worker exposure", "Minutes a worker spends at the open slab edge or inside the robot's working area. Robot alone still needs a worker for set-up and supervision.", "Method statement and risk assessment"),
        ("Ratings C3–C7", "One person's ratings with the scoring rubric, for illustration.", "Average of the project team's ratings"),
        ("Infeasible options", "No manual lifting of 300 kg panels (no W for P2); the robot alone cannot align to ±3 mm (no R for P3); "
                               "a team adds nothing to setting out and inspection (no HR for T1, T2).", "Feasibility matrix agreed with the site team"),
        ("Costs", "Labour about £33/h, robot about £45/h, team the sum; rounded to £ per minute.", "Project cost rates"),
        ("Weights", "Example weights from the worked example.", "Aggregated BWM weights from the questionnaire"),
    ],
}
# Options: (task, option, duration, exposure, C3, C4, C5, C6, C7, note)
spec = {
    "set":   [("W", 40, 10, 2, 2, 3, 2, 3, "Tape and laser level"), ("R", 25, 5, 1, 1, 2, 2, 1, "Robot layout with total station")],
    "drill": [("W", 45, 30, 3, 4, 2, 2, 4, "Dust, vibration and repetitive drilling at the edge"),
              ("R", 30, 5, 2, 1, 2, 2, 1, "Drilling end-effector; worker supervises from behind the barrier"),
              ("HR", 30, 15, 2, 2, 3, 3, 2, "Worker guides the robot to each mark")],
    "lift":  [("R", 30, 10, 3, 1, 3, 3, 2, "Lifting-assist arm; banksman keeps clear of the load"),
              ("HR", 25, 25, 3, 2, 4, 4, 2, "Worker steadies and guides the suspended panel")],
    "align": [("W", 35, 20, 2, 3, 4, 4, 2, "Shims and levelling by hand while temporarily propped"),
              ("HR", 25, 20, 2, 2, 4, 4, 3, "Robot holds the panel; worker adjusts")],
    "fix":   [("W", 40, 25, 2, 3, 2, 2, 4, "Repetitive torque bolting"),
              ("R", 30, 5, 2, 1, 2, 2, 1, "Bolting end-effector"),
              ("HR", 25, 15, 2, 2, 3, 3, 2, "Worker places brackets, robot torques")],
    "seal":  [("W", 30, 20, 2, 3, 3, 3, 3, "Sealant gun along the joints"),
              ("R", 35, 5, 1, 1, 2, 3, 2, "Sealing end-effector; slower around corners"),
              ("HR", 25, 10, 2, 2, 3, 3, 2, "")],
    "insp":  [("W", 30, 10, 1, 2, 4, 2, 2, "Visual check and measurements"), ("R", 20, 5, 1, 1, 2, 2, 2, "3D scan of the installed panels")],
}
kind = {"T1": "set", "P1": "drill", "P2": "lift", "P3": "align", "P4": "fix", "P5": "seal", "T2": "insp"}
for tid, _, *_ in realistic["tasks"]:
    for (k, d, e, c3, c4, c5, c6, c7, note) in spec[kind[tid]]:
        realistic["options"].append((tid, k, d, e, c3, c4, c5, c6, c7, note))

# ============================================================================ worked example (slides)
worked = {
    "title": "Worked example from the WP5 slides (3 tasks)",
    "intro": ["One worker, one robot; reproduces scenarios A–D on the slides.",
              "Example values are illustrative ratings (rating = 1 + 4 × score), not measured data."],
    "nW": 1, "nR": 1, "rates": [0.6, 1.0, 1.6], "cost_unit": "cost units", "ref": [60, 56],
    "lam": 0.03, "zone": "No", "cap": None,
    "tasks": [("T1", "Lift panel", "", "No", "", "TT1 Heavy lifting", "", ""), ("T2", "Align panel", "T1", "No", "", "TT3 Precise alignment", "", ""),
              ("T3", "Fix bolts", "T1", "No", "", "TT4 Fastening", "", "")],
    "options": [
        ("T1", "W", 60, 60, 4.2, 4.6, 2.6, 3.0, 3.0, ""), ("T1", "R", 40, 5, 2.2, 1.4, 1.8, 2.6, 1.8, ""),
        ("T1", "HR", 35, 20, 2.6, 2.2, 3.0, 3.4, 2.2, ""), ("T2", "W", 30, 10, 1.8, 2.2, 3.4, 3.8, 1.8, ""),
        ("T2", "HR", 25, 10, 2.2, 1.8, 3.0, 3.4, 2.2, "No R row: the robot alone cannot align"),
        ("T3", "W", 50, 20, 1.8, 3.0, 2.2, 2.2, 3.4, ""), ("T3", "R", 40, 5, 1.8, 1.4, 1.8, 2.2, 1.4, ""),
        ("T3", "HR", 30, 15, 2.2, 1.8, 2.6, 3.0, 2.2, ""),
    ],
}

build(DIR + "project_template.xlsx", realistic)
build(DIR + "worked_example.xlsx", worked)
