"""Builds Kanban_Tracker.xlsb: team task list, self-filling Kanban board, team summary, archive and macros."""
import datetime as dt
import io
import subprocess
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

HERE = Path(__file__).resolve().parent
OUT = HERE / "Kanban_Tracker.xlsb"
VBA_TEMPLATE = HERE / "vba" / "vba_template.xlsb"      # Excel-made workbook holding the compiled macros

CAP = 1000                                # task rows, and card slots per board column
TEAM_CAP, PROJECT_CAP = 200, 50
STATUSES = ["Backlog", "To Do", "In Progress", "Review", "Done", "On Hold"]
WIP_LIMITS = [None, 15, 10, 5, None, None]
GROUP = ["blue", "blue", "rose", "rose", "rose", "purple"]
DONE, WORKING = 5, 3                      # status positions: finished, and "working on now"
PRIORITIES = ["High", "Medium", "Low"]
TEAM = ["Stuart", "Alex", "Jordan", "Sam"]
PROJECTS = ["Reporting", "Forecasting", "Systems", "Admin"]

HEAD = {"blue": "4A6FA5", "rose": "D48C9E", "purple": "9478D0"}    # Tasks / Lists headers
DARK = {"blue": "4A6FA5", "rose": "C0657D", "purple": "8063C4"}    # ticked text, Lists status cells
PALE = {"blue": "DCE5F3", "rose": "F6DEE4", "purple": "E7E0F6"}    # ticked cells, summary counts
CARD = {"High": "F4CCD6", "Medium": "FBF3D6", "Low": "DCE5F3", "None": "EDEDED"}
NAVY, WHITE, BLACK = "002060", "FFFFFF", "000000"                  # Summary / Board headers (Stuart's restyle)
ROSE, CREAM, PICK, GRID, TEXT, MUTED, RED = "C0657D", "FBF3D6", "3A2A6E", "D9D9D9", "262626", "8C8C8C", "C00000"
ZERO = "C8C8C8"                                     # zero counts: light grey, no fill
VIEWS = ["Team member", "Project"]                   # Summary VIEW BY options

# Tasks sheet: header row 6, task rows 7..LAST; Calc rows line up with Tasks rows
HDR, FIRST = 6, 7
LAST = FIRST + CAP - 1
TASK_COLS = [("A", "ID", 8), ("B", "TASK", 34), ("C", "OWNER", 16), ("D", "STATUS", 14), ("E", "PRIORITY", 14),
             ("F", "PROJECT", 16), ("G", "START DATE", 14), ("H", "DUE DATE", 14), ("I", "DONE DATE", 14),
             ("J", "NOTES", 34)]
TICK_COLS = ["L", "M", "N", "O", "P", "Q"]
BOARD_COLS = ["B", "D", "F", "H", "J", "L"]
MATCH_COLS = ["P", "Q", "R", "S", "T", "U"]          # on Calc: Tasks row behind card k of each status
CARD_TOP = 12                                        # first card row on Board; each card is 3 rows + 1 spacer
SUM_TOP = 14                                         # Summary: UNASSIGNED row, names from the row below
ARCHIVE_COLS = [("A", "ARCHIVED ON", 14), ("B", "ID", 8)] + [(chr(ord(c) + 1), t, w) for c, t, w in TASK_COLS[1:]]

# Buttons: (sheet file, macro, label, from (col, colOff, row, rowOff), to (...)); zero-based cells, EMU offsets
BUTTONS = [
    ("sheet2.xml", "ResetBoardFilters", "RESET FILTERS", (7, 46182, 5, 11545), (7, 1212272, 7, 23090)),
    ("sheet3.xml", "ArchiveDoneTasks", "ARCHIVE DONE TASKS", (15, 38100, 1, 38100), (16, 742950, 3, 266700)),
]

CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)


def font(size=10, bold=False, color=TEXT, italic=False, name="Arial"):
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def box(color=GRID):
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def put(ws, ref, value=None, *, f=None, bg=None, border=None, wrap=False, fmt=None, unlocked=False):
    """Write a centred cell with optional font, fill, border, number format; unlocked = typeable when protected."""
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = f or font()
    c.alignment = CENTER_WRAP if wrap else CENTER
    if bg:
        c.fill = fill(bg)
    if border:
        c.border = border
    if fmt:
        c.number_format = fmt
    if unlocked:
        c.protection = Protection(locked=False)
    return c


def header(ws, ref, value, colour):
    return put(ws, ref, value, f=font(bold=True, color=WHITE), bg=colour, border=box(WHITE), wrap=True)


def navy(ws, ref, value, size=10):
    """Navy header with a thin black box (Summary and Board)."""
    return put(ws, ref, value, f=font(size, bold=True, color=WHITE), bg=NAVY, border=box(BLACK), wrap=True)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def outline(ws, rng):
    """Thin black outline round a block of cells."""
    (c1, r1), (c2, r2) = ((ws[x].column, ws[x].row) for x in rng.split(":"))
    side = Side(style="thin", color=BLACK)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = Border(left=side if c == c1 else None, right=side if c == c2 else None,
                                          top=side if r == r1 else None, bottom=side if r == r2 else None)


def banner(ws, date_cells, title_cells, title, plain=False):
    """Date block beside a title block, rows 2-4: cream/rose, or white with a black outline (plain)."""
    ws.merge_cells(date_cells)
    ws.merge_cells(title_cells)
    d, t = date_cells.split(":")[0], title_cells.split(":")[0]
    put(ws, d, "=TODAY()", f=font(11, bold=True, italic=True, color=BLACK if plain else "3A3A3A"),
        bg=WHITE if plain else CREAM, fmt="dddd d mmmm yyyy", wrap=True)
    put(ws, t, title, f=font(28, bold=True, color=BLACK if plain else WHITE, name="Rockwell"), bg=WHITE if plain else ROSE)
    if plain:
        outline(ws, date_cells)
        outline(ws, title_cells)
    ws.row_dimensions[1].height = 6
    for r in (2, 3, 4):
        ws.row_dimensions[r].height = 24
    ws.row_dimensions[5].height = 6


def add_rule(ws, rng, formula, **fmt):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], **fmt))


def dropdown(ws, source, rng, title, msg):
    dv = DataValidation(type="list", formula1=source, allow_blank=True, showErrorMessage=True, errorTitle=title, error=msg)
    ws.add_data_validation(dv)
    dv.add(rng)


def sample_tasks():
    """Example rows: every status, an overdue task, an unassigned task and a Done task old enough to leave the board."""
    today = dt.date.today()
    d = lambda n: today + dt.timedelta(days=n)
    return [
        ("Month-end management pack", "Stuart", "In Progress", "High", "Reporting", d(-6), d(3), None, "Waiting on two cost-centre files"),
        ("Refresh Q4 sales forecast", "Alex", "To Do", "High", "Forecasting", d(-2), d(7), None, None),
        ("Clean up supplier master data", "Jordan", "Backlog", "Low", "Systems", None, None, None, "After year end"),
        ("Board slides: pricing update", "Stuart", "Review", "Medium", "Reporting", d(-10), d(1), None, "Alex to review"),
        ("Automate weekly KPI email", "Sam", "In Progress", "Medium", "Systems", d(-8), d(-2), None, "Overdue example"),
        ("Budget template for 2027", "Alex", "Backlog", "Medium", "Forecasting", None, d(30), None, None),
        ("Reconcile intercompany balances", "Jordan", "In Progress", "High", "Reporting", d(-1), d(5), None, None),
        ("Team training: Power Query", "Sam", "To Do", "Low", "Admin", None, d(21), None, None),
        ("Update holiday rota", "Stuart", "Done", "Low", "Admin", d(-12), d(-5), d(-6), None),
        ("Fix broken links in cost model", "Alex", "Done", "Medium", "Systems", d(-9), d(-3), d(-4), None),
        ("Supplier price review", "Jordan", "On Hold", "Medium", "Forecasting", d(-3), d(10), None, "Waiting on supplier data"),
        ("Year-end timetable", None, "Backlog", "High", "Reporting", None, d(14), None, "Needs an owner"),
        ("Stock count plan", "Alex", "In Progress", "Medium", "Admin", d(-4), d(12), None, None),
        ("Archive 2024 files", "Jordan", "Done", "Low", "Admin", d(-50), d(-42), d(-40), "Done 40 days ago: off the board"),
        ("Credit control catch-up", "Sam", "Review", "High", "Reporting", d(-5), d(2), None, None),
    ]


def build_lists(ws):
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 16, "B": 12, "C": 2, "D": 14, "E": 2, "F": 18, "G": 2, "H": 18, "I": 2, "J": 70})
    ws.row_dimensions[1].height = 32
    for ref, text in {"A1": "STATUS", "B1": "WIP LIMIT", "D1": "PRIORITY", "F1": "TEAM MEMBERS",
                      "H1": "PROJECTS", "J1": "NOTES"}.items():
        header(ws, ref, text, HEAD["blue"])
    for i, (status, limit) in enumerate(zip(STATUSES, WIP_LIMITS), start=2):
        put(ws, f"A{i}", status, f=font(bold=True, color=WHITE), bg=DARK[GROUP[i - 2]], border=box(WHITE), unlocked=True)
        put(ws, f"B{i}", limit, border=box(), unlocked=True)
    for i, p in enumerate(PRIORITIES, start=2):
        put(ws, f"D{i}", p, border=box(), unlocked=True)
    for i in range(2, TEAM_CAP + 2):
        put(ws, f"F{i}", TEAM[i - 2] if i - 2 < len(TEAM) else None, border=box(), unlocked=True)
    for i in range(2, PROJECT_CAP + 2):
        put(ws, f"H{i}", PROJECTS[i - 2] if i - 2 < len(PROJECTS) else None, border=box(), unlocked=True)
    notes = ["TEAM MEMBERS and PROJECTS: top down, no gaps. Up to 200 people and 50 projects.",
             "New names reach every dropdown, the Board filters and Summary straight away.",
             "WIP: work in progress. WIP LIMIT: most tickets the whole team should have in that column. Blank = no limit.",
             "A Board column header turns red when the whole team goes over its WIP LIMIT.",
             "STATUS: rename freely but keep the order. Done stays 5th: finished-task logic and ARCHIVE use it."]
    for i, note in enumerate(notes, start=2):
        put(ws, f"J{i}", note, f=font(italic=True, color=MUTED))
    ws.freeze_panes = "A2"


def build_tasks(ws):
    ws.sheet_view.showGridLines = False
    widths(ws, {**{c: w for c, _, w in TASK_COLS}, "K": 2, **{c: 11 for c in TICK_COLS}})
    banner(ws, "A2:B4", "C2:O4", "TASKS")              # P2:Q4 holds the ARCHIVE button
    ws.row_dimensions[HDR].height = 32
    for col, text, _ in TASK_COLS:
        header(ws, f"{col}{HDR}", text, HEAD["blue"])
    for i, col in enumerate(TICK_COLS):
        header(ws, f"{col}{HDR}", f"=UPPER(Lists!$A${i + 2})", HEAD[GROUP[i]])

    rows = sample_tasks()
    for r in range(FIRST, LAST + 1):
        ws.row_dimensions[r].height = 18
        put(ws, f"A{r}", f'=IF(B{r}="","",ROW()-{HDR})', f=font(color=MUTED), border=box())
        values = rows[r - FIRST] if r - FIRST < len(rows) else (None,) * 9
        for (col, _, _), v in zip(TASK_COLS[1:], values):
            put(ws, f"{col}{r}", v, border=box(), fmt="d mmm yy" if col in "GHI" else None, unlocked=True)
        for j, col in enumerate(TICK_COLS, start=1):
            put(ws, f"{col}{r}", f'=IF($B{r}="","",IF(OR(AND(Calc!$A{r}<={DONE},{j}<=Calc!$A{r}),Calc!$A{r}={j}),"☑","☐"))',
                f=font(12, color="A6A6A6"), border=box())

    body = lambda col: f"{col}{FIRST}:{col}{LAST}"
    add_rule(ws, f"B{FIRST}:B{LAST}", f"Calc!$F{FIRST}=1", font=Font(color=RED, bold=True))
    add_rule(ws, f"B{FIRST}:J{LAST}", f"Calc!$A{FIRST}={DONE}", font=Font(color=MUTED))
    for i, col in enumerate(TICK_COLS):
        add_rule(ws, body(col), f'{col}{FIRST}="☑"', fill=fill(PALE[GROUP[i]]), font=Font(color=DARK[GROUP[i]], bold=True))

    dropdown(ws, "=TeamList", body("C"), "Owner", "Pick a name from the list. Add people on the Lists sheet.")
    dropdown(ws, "=StatusList", body("D"), "Status", "Pick a status from the list.")
    dropdown(ws, "=PriorityList", body("E"), "Priority", "Pick High, Medium or Low.")
    dropdown(ws, "=ProjectList", body("F"), "Project", "Pick a project. Add more on the Lists sheet.")
    dates = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                           showErrorMessage=True, errorTitle="Date", error="Enter a date, e.g. 20/10/2026")
    ws.add_data_validation(dates)
    dates.add(f"G{FIRST}:I{LAST}")
    ws.freeze_panes = f"A{FIRST}"


def build_calc(ws):
    """Workings: per-task helpers on the same rows as Tasks, status counts, filter lists, card lookups."""
    widths(ws, {c: 14 for c in "ABCDEFGJKMNPQRSTUXY"})
    widths(ws, {"H": 2, "I": 2, "L": 2, "O": 2, "V": 2, "W": 2})
    ws.merge_cells("A1:G1")
    put(ws, "A1", "WORKINGS: these formulas drive the Board and Summary. Nothing to type here.", f=font(bold=True, color=MUTED))
    heads = {"A": "STATUS NUMBER", "B": "SORT ORDER", "C": "ON BOARD (1 = YES)", "D": "CARD NUMBER",
             "E": "CARD KEY", "F": "OVERDUE (1 = YES)", "G": "DAYS UNTIL DUE",
             "J": "STATUS", "K": "CARDS ON BOARD", "M": "WIP LIMIT", "N": "WHOLE TEAM TASKS",
             "X": "TEAM SELECTION LIST", "Y": "PROJECT LIST"}
    heads.update({mc: f"TASKS ROW: {STATUSES[i].upper()}" for i, mc in enumerate(MATCH_COLS)})
    for col, text in heads.items():
        header(ws, f"{col}{HDR}", text, "7F7F7F")
    ws.row_dimensions[HDR].height = 32

    rng = lambda col: f"${col}${FIRST}:${col}${LAST}"
    for r in range(FIRST, LAST + 1):
        task, owner, status, prio, proj, due, done = (f"Tasks!${c}{r}" for c in "BCDEFHI")
        open_ = f"A{r}>0,A{r}<>{DONE}"
        ws[f"A{r}"] = f'=IF({task}="",0,IF({status}="",1,IFERROR(MATCH({status},StatusList,0),0)))'
        ws[f"B{r}"] = f'=IF(A{r}=0,"",IFERROR(MATCH({prio},PriorityList,0),4)*100000+IF(ISNUMBER({due}),{due},99999))'
        ws[f"C{r}"] = (f'=IF(A{r}=0,0,IF(AND(OR(TeamSelection="All",{owner}=TeamSelection),'
                       f'OR(ProjectSelection="All",{proj}=ProjectSelection),OR(A{r}<>{DONE},NOT(ISNUMBER({done})),'
                       f'TODAY()-N({done})<=HideAfterDays)),1,0))')
        ws[f"D{r}"] = (f'=IF(C{r}=1,COUNTIFS({rng("A")},A{r},{rng("C")},1,{rng("B")},"<"&B{r})'
                       f'+COUNTIFS($A${FIRST}:A{r},A{r},$C${FIRST}:C{r},1,$B${FIRST}:B{r},B{r}),"")')
        ws[f"E{r}"] = f'=IF(D{r}="","",A{r}&"|"&D{r})'
        ws[f"F{r}"] = f'=IF(AND({open_},ISNUMBER({due})),IF({due}<TODAY(),1,0),0)'
        ws[f"G{r}"] = f'=IF(AND({open_},ISNUMBER({due})),{due}-TODAY(),"")'
        card = r - FIRST + 1
        for i, mc in enumerate(MATCH_COLS, start=1):
            ws[f"{mc}{r}"] = f'=IF({card}>$K${FIRST + i - 1},"",MATCH("{i}|{card}",{rng("E")},0))'
        for col in "ABCDEFG" + "".join(MATCH_COLS):
            ws[f"{col}{r}"].font = font(color=MUTED)
            ws[f"{col}{r}"].alignment = CENTER

    for i in range(1, len(STATUSES) + 1):
        r = FIRST + i - 1
        put(ws, f"J{r}", f"=Lists!$A${i + 1}", border=box())
        put(ws, f"K{r}", f'=COUNTIFS({rng("A")},{i},{rng("C")},1)', border=box())
        put(ws, f"M{r}", f'=IF(Lists!$B${i + 1}="","",Lists!$B${i + 1})', border=box())
        put(ws, f"N{r}", f'=COUNTIF({rng("A")},{i})', border=box())

    for col, src, cap in (("X", "F", TEAM_CAP), ("Y", "H", PROJECT_CAP)):
        put(ws, f"{col}{FIRST}", "All", border=box())
        for i in range(cap):
            put(ws, f"{col}{FIRST + 1 + i}", f'=IF(Lists!${src}${2 + i}="","",Lists!${src}${2 + i})', border=box())
    ws.freeze_panes = f"A{FIRST}"


def card_rows(k):
    top = CARD_TOP + (k - 1) * 4
    return top, top + 1, top + 2


def build_board(ws):
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 4, **{c: 30 for c in BOARD_COLS}, **{c: 2 for c in "CEGIK"}})
    banner(ws, "B2:B4", "D2:L4", "KANBAN BOARD")

    filters = [("B", "TEAM SELECTION", "All", "=TeamSelectList"), ("D", "PROJECT", "All", "=ProjectSelectList"),
               ("F", "HIDE DONE AFTER (DAYS)", 30, None)]
    for col, label, default, source in filters:
        navy(ws, f"{col}6", label)
        put(ws, f"{col}7", default, f=font(11, bold=True, color=PICK), bg=WHITE, border=box(BLACK), unlocked=True)
        if source:
            dropdown(ws, source, f"{col}7", label.title(), "Pick All or a name from the list.")
    ws.merge_cells("I6:L6")                             # H6:H7 holds the RESET FILTERS button
    ws.merge_cells("I7:L7")
    put(ws, "I6", "CARD COLOUR = PRIORITY:  HIGH  ·  MEDIUM  ·  LOW", f=font(9, bold=True, color=MUTED))
    put(ws, "I7", "RED TEXT = OVERDUE  ·  RED HEADER = OVER WIP LIMIT", f=font(9, bold=True, color=RED))
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[7].height = 22
    ws.row_dimensions[8].height = 8
    ws.row_dimensions[9].height = 26
    ws.row_dimensions[10].height = 20
    ws.row_dimensions[11].height = 6

    for i, col in enumerate(BOARD_COLS):
        r = FIRST + i
        navy(ws, f"{col}9", f'=UPPER(Calc!$J${r})&" ("&Calc!$K${r}&")"', size=12)
        put(ws, f"{col}10", f'=IF(Calc!$M${r}="","No WIP limit","WIP limit "&Calc!$M${r}&"  ·  team has "&Calc!$N${r})',
            f=font(9, color=BLACK), bg=WHITE, border=box(BLACK))
        add_rule(ws, f"{col}9:{col}10", f"AND(ISNUMBER(Calc!$M${r}),Calc!$N${r}>Calc!$M${r})",
                 fill=fill(RED), font=Font(color=WHITE, bold=True), stopIfTrue=True)

    tasks = lambda c: f"Tasks!${c}${FIRST}:${c}${LAST}"
    for k in range(1, CAP + 1):
        r1, r2, r3 = card_rows(k)
        ws.row_dimensions[r1].height = 30
        ws.row_dimensions[r2].height = 18
        ws.row_dimensions[r3].height = 18
        ws.row_dimensions[r3 + 1].height = 6
        for r in (r1, r2, r3):
            put(ws, f"A{r}", k, f=font(7, color=GRID))
        for col, mc in zip(BOARD_COLS, MATCH_COLS):
            row = f"Calc!${mc}${FIRST + k - 1}"
            at = lambda c: f"INDEX({tasks(c)},{row})"
            put(ws, f"{col}{r1}", f'=IF({row}="","","#"&{at("A")}&"  "&{at("B")})', f=font(bold=True), wrap=True)
            put(ws, f"{col}{r2}", f'=IF({row}="","",IF({at("C")}="","Unassigned",{at("C")})&IF({at("F")}="","","  |  "&{at("F")}))',
                f=font(9, color="404040"))
            put(ws, f"{col}{r3}", f'=IF({row}="","",IF({at("E")}="","No priority",{at("E")})&"  |  "&IF({at("H")}="","No due date",'
                                  f'IF(INDEX(Calc!$F${FIRST}:$F${LAST},{row})=1,"OVERDUE ","Due ")&TEXT({at("H")},"d mmm")))',
                f=font(9, color="404040"))

    # cards are coloured only when a task fills the slot; empty slots and 6pt spacer rows (no card number) stay white
    end = card_rows(CAP)[2]
    for col, mc in zip(BOARD_COLS, MATCH_COLS):
        area = f"{col}{CARD_TOP}:{col}{end}"
        row = f"INDEX(Calc!${mc}${FIRST}:${mc}${LAST},$A{CARD_TOP})"
        has_card = f"ISNUMBER($A{CARD_TOP})"
        add_rule(ws, area, f"AND({has_card},IFERROR(INDEX(Calc!$F${FIRST}:$F${LAST},{row})=1,FALSE))", font=Font(color=RED, bold=True))
        for p in PRIORITIES:
            add_rule(ws, area, f'AND({has_card},IFERROR(INDEX({tasks("E")},{row})="{p}",FALSE))', fill=fill(CARD[p]), stopIfTrue=True)
        add_rule(ws, area, f"AND({has_card},IFERROR(ISNUMBER({row}),FALSE))", fill=fill(CARD["None"]), stopIfTrue=True)
    ws.freeze_panes = f"A{CARD_TOP}"


def build_summary(ws):
    """One table that switches between team members and projects (VIEW BY dropdown); same columns either way."""
    ws.sheet_view.showGridLines = False
    status_cols = ["C", "D", "E", "F", "G", "H"]
    extra = ["I", "J", "K", "L"]
    widths(ws, {"A": 2, "B": 20, **{c: 11 for c in status_cols + extra}, "M": 70})
    banner(ws, "B2:C4", "D2:M4", "TEAM SUMMARY", plain=True)
    for r, h in {5: 21, 6: 41, 7: 34, 8: 29, 9: 22, 10: 5, 11: 32, 13: 5}.items():
        ws.row_dimensions[r].height = h

    T = lambda c: f"Tasks!${c}${FIRST}:${c}${LAST}"
    C = lambda c: f"Calc!${c}${FIRST}:${c}${LAST}"
    kpis = [("TOTAL TASKS", f'=COUNTIF({T("B")},"?*")', NAVY),
            ("OPEN", f'=COUNTIFS({C("A")},">0",{C("A")},"<>{DONE}")', NAVY),
            ("OVERDUE", f'=SUM({C("F")})', RED),
            ("DUE IN 7 DAYS", f'=COUNTIFS({C("G")},">=0",{C("G")},"<=7")', NAVY),
            ("DONE IN LAST 30 DAYS", f'=COUNTIFS({C("A")},{DONE},{T("I")},">="&(TODAY()-30))'
                                     f'+COUNTIF(Archive!$J:$J,">="&(TODAY()-30))', NAVY)]
    for col, (label, formula, colour) in zip(["D", "F", "H", "J", "L"], kpis):
        navy(ws, f"{col}6", label)
        put(ws, f"{col}7", formula, f=font(20, bold=True, color=ZERO), bg=PALE["blue"], border=box(BLACK))
        add_rule(ws, f"{col}7", f"N({col}7)>0", font=Font(color=colour, bold=True))

    navy(ws, "B9", "VIEW BY")
    ws.merge_cells("C9:D9")
    put(ws, "C9", VIEWS[0], f=font(11, bold=True, color=PICK), bg=WHITE, unlocked=True)
    outline(ws, "C9:D9")
    dropdown(ws, f'"{",".join(VIEWS)}"', "C9", "View by", "Pick Team member or Project.")
    is_project = f'SummaryView="{VIEWS[1]}"'

    navy(ws, "B11", f'=IF({is_project},"PROJECT","TEAM MEMBER")')
    for i, col in enumerate(status_cols):
        navy(ws, f"{col}11", f"=UPPER(Lists!$A${i + 2})")
    for col, text in zip(extra, ["OPEN", "OVERDUE", "DUE IN 7 DAYS", "TOTAL"]):
        navy(ws, f"{col}11", text)
    navy(ws, "M11", f'="WORKING ON NOW ("&UPPER(Lists!$A${WORKING + 1})&")"')

    bottom = SUM_TOP + TEAM_CAP
    open_cols = [c for i, c in enumerate(status_cols, start=1) if i != DONE]

    def fill_row(r, who, crit):
        """crit: COUNTIFS criteria pair picking this row's tasks."""
        for i, col in enumerate(status_cols, start=1):
            ws[f"{col}{r}"] = f'=IF({who}="","",COUNTIFS({crit},{C("A")},{i}))'
        ws[f"I{r}"] = f'=IF({who}="","",{"+".join(f"{c}{r}" for c in open_cols)})'
        ws[f"J{r}"] = f'=IF({who}="","",COUNTIFS({crit},{C("F")},1))'
        ws[f"K{r}"] = f'=IF({who}="","",COUNTIFS({crit},{C("G")},">=0",{C("G")},"<=7"))'
        ws[f"L{r}"] = f'=IF({who}="","",SUM(C{r}:H{r}))'

    label = f'{T("B")}&IF({is_project}," ("&IF({T("C")}="","Unassigned",{T("C")})&")","")'   # project view adds the owner
    working = lambda match: f'_xlfn.TEXTJOIN("  |  ",TRUE,IF(({match})*({T("B")}<>"")*({C("A")}={WORKING}),{label},""))'

    put(ws, f"B{SUM_TOP}", f'=IF({is_project},"NO PROJECT","UNASSIGNED")', f=font(bold=True, color=MUTED), border=box())
    fill_row(SUM_TOP, f"$B{SUM_TOP}", f'{T("B")},"?*",SummaryField,""')
    ws[f"M{SUM_TOP}"] = ArrayFormula(f"M{SUM_TOP}", "=" + working('SummaryField=""'))
    for r in range(SUM_TOP + 1, bottom + 1):
        n = r - SUM_TOP + 1
        ws[f"B{r}"] = f'=IF({is_project},IF(Lists!$H${n}="","",Lists!$H${n}),IF(Lists!$F${n}="","",Lists!$F${n}))'
        fill_row(r, f"$B{r}", f"SummaryField,$B{r}")
        ws[f"M{r}"] = ArrayFormula(f"M{r}", f'=IF($B{r}="","",{working(f"SummaryField=$B{r}")})')
    for r in range(SUM_TOP, bottom + 1):
        ws.row_dimensions[r].height = 18
        for col in ["B"] + status_cols + extra + ["M"]:
            c = ws[f"{col}{r}"]
            c.font = font(bold=(col == "B"), color=MUTED if (r == SUM_TOP and col == "B") else (ZERO if col in status_cols + extra else TEXT))
            c.alignment = CENTER

    total = 12
    edge = Side(style="thin", color=GRID)
    put(ws, f"B{total}", "TOTAL", f=font(bold=True), bg=PALE["blue"])
    for col in status_cols + extra:
        put(ws, f"{col}{total}", f"=SUM({col}{SUM_TOP}:{col}{bottom})", f=font(bold=True, color=ZERO), bg=PALE["blue"])
    put(ws, f"M{total}", f'=IF(L{total}=COUNTIF({T("B")},"?*"),"CHECK OK: all "&L{total}&" tasks counted",'
                         f'"CHECK: totals do not match. A name on Tasks is missing from Lists")',
        f=font(italic=True, color=MUTED), bg=PALE["blue"])
    for col in ["B"] + status_cols + extra + ["M"]:
        ws[f"{col}{total}"].border = Border(left=edge, right=edge, bottom=edge)
    add_rule(ws, f"M{total}", f'LEFT(M{total},8)<>"CHECK OK"', font=Font(color=RED, bold=True))

    # zeros stay light grey with no fill; non-zero counts get their colour (first matching font rule wins)
    add_rule(ws, f"B{SUM_TOP}:M{bottom}", f'OR($B{SUM_TOP}<>"",ROW()={SUM_TOP})', border=box())
    for i, col in enumerate(status_cols):
        add_rule(ws, f"{col}{SUM_TOP}:{col}{bottom}", f"N({col}{SUM_TOP})>0", fill=fill(PALE[GROUP[i]]),
                 font=Font(color=DARK[GROUP[i]], bold=True))
    add_rule(ws, f"J{total}:J{bottom}", f"N(J{total})>0", font=Font(color=RED, bold=True))
    add_rule(ws, f"C{total}:L{bottom}", f"N(C{total})>0", font=Font(color=TEXT, bold=True))
    ws.freeze_panes = f"A{SUM_TOP}"


def build_archive(ws):
    """Hidden store for archived Done tasks; the ARCHIVE macro appends to the next free row."""
    ws.sheet_state = "hidden"
    ws.sheet_view.showGridLines = False
    widths(ws, {c: w for c, _, w in ARCHIVE_COLS})
    ws.row_dimensions[1].height = 32
    for col, text, _ in ARCHIVE_COLS:
        navy(ws, f"{col}1", text)
    ws.freeze_panes = "A2"


def build_how_to(ws):
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 2, "B": 26, "C": 110})
    ws.merge_cells("B2:C4")
    put(ws, "B2", "HOW TO USE", f=font(28, bold=True, color=WHITE, name="Rockwell"), bg=ROSE)
    for r in (2, 3, 4):
        ws.row_dimensions[r].height = 24
    sections = [
        ("SHEETS", [
            ("SUMMARY", "Whole-team view. VIEW BY switches between team members and projects: tasks at each stage, overdue, due in 7 days, and what is being worked on now. The workbook opens here."),
            ("BOARD", "The Kanban board. One card per task, in the column for its status. Builds itself from Tasks."),
            ("TASKS", "The task list. The only place you type tasks. Ticks show how far each task has got."),
            ("LISTS", "Team names, projects, status names, priorities and WIP limits. Every dropdown reads from here."),
            ("CALC", "Workings behind the Board and Summary. Nothing to type here."),
            ("ARCHIVE", "Hidden. Archived Done tasks, newest at the bottom. To look: right-click any tab > Unhide > Archive.")]),
        ("GETTING STARTED", [
            ("1. TEAM", "Type your team's names on Lists, column F, and projects in column H. Top down, no gaps."),
            ("2. SAMPLE TASKS", "Clear the 15 example rows on Tasks, columns B to J. ID and ticks are formulas and clear themselves."),
            ("3. ADD A TASK", "Type it in the next empty row on Tasks. Pick OWNER, STATUS, PRIORITY and PROJECT from the dropdowns. Room for 1,000 tasks."),
            ("4. MOVE A CARD", "Change the task's STATUS. The ticks fill in and the card jumps to that column on the Board."),
            ("5. FINISH A TASK", "Set STATUS to Done and type the DONE DATE. The card leaves the Board after 30 days (Board, HIDE DONE AFTER)."),
            ("6. ONE PERSON", "Pick a name in TEAM SELECTION on the Board to see only their cards. PROJECT works the same way."),
            ("7. RESET FILTERS", "The RESET FILTERS button on the Board puts TEAM SELECTION and PROJECT back to All."),
            ("8. ARCHIVE", "The ARCHIVE DONE TASKS button on Tasks asks Yes/No, then moves every Done task to the hidden Archive sheet and closes up the gaps so the list has room again. IDs renumber afterwards.")]),
        ("READING IT", [
            ("TICKS", "☑ marks each stage a task has reached: Backlog, To Do, In Progress, Review, Done. On Hold ticks only its own box."),
            ("BLANK STATUS", "Counts as Backlog, so a new task always appears on the Board."),
            ("WIP", "Work in progress: tasks being worked on right now. WIP LIMIT: most cards the whole team should have in a column."),
            ("CARD COLOUR", "Priority: rose = High, cream = Medium, blue = Low, grey = none. Empty slots stay blank."),
            ("RED TEXT", "Past its due date and not Done. The task name on Tasks turns red too."),
            ("RED HEADER", "The whole team has more tasks in that column than its WIP LIMIT (Lists, column B)."),
            ("CARD ORDER", "High priority first, then soonest due date."),
            ("ENOUGH CARDS", "Every column has a slot for every task, so the Board never runs out."),
            ("ZERO COUNTS", "Show in light grey with no colour, so the real numbers stand out.")]),
        ("PROTECTION", [
            ("ON OPEN", "A macro protects every sheet (no password), sets 110% zoom and opens on Summary. Only input cells take typing: formulas are locked."),
            ("UNPROTECT", "Review > Unprotect Sheet. No password needed."),
            ("MACROS OFF", "Everything still works except the buttons, and sheets stay unprotected.")]),
    ]
    r = 6
    for title, rows in sections:
        put(ws, f"B{r}", title, f=font(bold=True, color=WHITE), bg=HEAD["blue"])
        put(ws, f"C{r}", None, bg=HEAD["blue"])
        r += 1
        for label, text in rows:
            put(ws, f"B{r}", label, f=font(bold=True), border=box())
            put(ws, f"C{r}", text, border=box(), wrap=True)
            ws.row_dimensions[r].height = 30
            r += 1
        r += 1


def build_workbook():
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    board = wb.create_sheet("Board")
    tasks = wb.create_sheet("Tasks")
    lists = wb.create_sheet("Lists")
    how = wb.create_sheet("How to use")
    calc = wb.create_sheet("Calc")
    archive = wb.create_sheet("Archive")

    names = {
        "StatusList": f"Lists!$A$2:$A${len(STATUSES) + 1}",
        "PriorityList": f"Lists!$D$2:$D${len(PRIORITIES) + 1}",
        "TeamList": f'OFFSET(Lists!$F$2,0,0,MAX(1,COUNTIF(Lists!$F$2:$F${TEAM_CAP + 1},"?*")),1)',
        "ProjectList": f'OFFSET(Lists!$H$2,0,0,MAX(1,COUNTIF(Lists!$H$2:$H${PROJECT_CAP + 1},"?*")),1)',
        "TeamSelectList": f'OFFSET(Calc!$X${FIRST},0,0,COUNTIF(Calc!$X${FIRST}:$X${FIRST + TEAM_CAP},"?*"),1)',
        "ProjectSelectList": f'OFFSET(Calc!$Y${FIRST},0,0,COUNTIF(Calc!$Y${FIRST}:$Y${FIRST + PROJECT_CAP},"?*"),1)',
        "TeamSelection": "Board!$B$7",
        "ProjectSelection": "Board!$D$7",
        "HideAfterDays": "Board!$F$7",
        "SummaryView": "Summary!$C$9",
        "SummaryField": f'CHOOSE(IF(SummaryView="{VIEWS[1]}",2,1),Tasks!$C${FIRST}:$C${LAST},Tasks!$F${FIRST}:$F${LAST})',
    }
    for name, ref in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)

    build_lists(lists)
    build_calc(calc)
    build_tasks(tasks)
    build_board(board)
    build_summary(summary)
    build_how_to(how)
    build_archive(archive)

    tabs = [(summary, NAVY), (board, ROSE), (tasks, HEAD["blue"]), (lists, HEAD["blue"]),
            (how, HEAD["purple"]), (calc, "A6A6A6"), (archive, "A6A6A6")]
    for i, (ws, tab) in enumerate(tabs, start=1):
        ws.sheet_properties.tabColor = tab
        ws.sheet_properties.codeName = f"Sheet{i}"
        ws.sheet_view.zoomScale = 110
    wb.code_name = "ThisWorkbook"
    wb.active = 0
    return wb


def _button_xml(i, macro, text, frm, to):
    pos = lambda tag, p: (f"<xdr:{tag}><xdr:col>{p[0]}</xdr:col><xdr:colOff>{p[1]}</xdr:colOff>"
                          f"<xdr:row>{p[2]}</xdr:row><xdr:rowOff>{p[3]}</xdr:rowOff></xdr:{tag}>")
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<xdr:twoCellAnchor editAs="absolute">{pos("from", frm)}{pos("to", to)}'
            f'<xdr:sp macro="[0]!{macro}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{i + 2}" name="{text.title()}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{BLACK}"/></a:solidFill></a:ln></xdr:spPr>'
            f'<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-GB" sz="1000" b="1"><a:solidFill><a:srgbClr val="{WHITE}"/></a:solidFill>'
            f'<a:latin typeface="Arial"/></a:rPr><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp>'
            f'<xdr:clientData fPrintsWithSheet="0"/></xdr:twoCellAnchor></xdr:wsDr>').encode()


def add_buttons(xlsx):
    """Add macro buttons as drawing shapes (openpyxl cannot write shapes)."""
    src = xlsx.read_bytes()
    with zipfile.ZipFile(io.BytesIO(src)) as zin, zipfile.ZipFile(xlsx, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for item in zin.infolist():
            data = zin.read(item.filename)
            for i, (sheet, *_rest) in enumerate(BUTTONS, start=1):
                if item.filename == f"xl/worksheets/{sheet}":
                    data = data.replace(b"</worksheet>", b'<drawing xmlns:r="http://schemas.openxmlformats.org/'
                                        b'officeDocument/2006/relationships" r:id="rIdBtn"/></worksheet>')
            if item.filename == "[Content_Types].xml":
                parts = b"".join(f'<Override PartName="/xl/drawings/button{i}.xml" ContentType='
                                 f'"application/vnd.openxmlformats-officedocument.drawing+xml"/>'.encode()
                                 for i in range(1, len(BUTTONS) + 1))
                data = data.replace(b"</Types>", parts + b"</Types>")
            zout.writestr(item, data)
        for i, (sheet, macro, text, frm, to) in enumerate(BUTTONS, start=1):
            rels = f"xl/worksheets/_rels/{sheet}.rels"
            assert rels not in names, f"{rels} already exists"
            zout.writestr(rels, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns='
                                f'"http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBtn" '
                                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                                f'Target="../drawings/button{i}.xml"/></Relationships>')
            zout.writestr(f"xl/drawings/button{i}.xml", _button_xml(i, macro, text, frm, to))


def excel_save_as_xlsb(src, dest):
    """Have Excel (macOS) open the macro-free .xlsx and save it as Excel binary."""
    script = f'''
    with timeout of 300 seconds
    tell application "Microsoft Excel"
        open POSIX file "{src}"
        set wb to active workbook
        calculate full
        save workbook as wb filename (POSIX file "{dest}" as text) file format Excel binary file format
        close wb saving no
    end tell
    end timeout'''
    subprocess.run(["osascript", "-e", script], check=True)


def add_vba(xlsb):
    """Copy the Excel-made macros (vbaProject.bin) from the template into the .xlsb package."""
    with zipfile.ZipFile(VBA_TEMPLATE) as z:
        vba = z.read("xl/vbaProject.bin")
    src = xlsb.read_bytes()
    with zipfile.ZipFile(io.BytesIO(src)) as zin, zipfile.ZipFile(xlsb, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", b'<Override PartName="/xl/vbaProject.bin" '
                                    b'ContentType="application/vnd.ms-office.vbaProject"/></Types>')
            elif item.filename == "xl/_rels/workbook.bin.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/'
                                    b'office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
            zout.writestr(item, data)
        zout.writestr("xl/vbaProject.bin", vba)


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as tmp:
        xlsx = Path(tmp) / "Kanban_Tracker.xlsx"
        build_workbook().save(xlsx)
        add_buttons(xlsx)
        OUT.unlink(missing_ok=True)
        excel_save_as_xlsb(xlsx, OUT)
        add_vba(OUT)
    print(f"saved {OUT}")
