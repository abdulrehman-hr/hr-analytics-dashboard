#!/usr/bin/env python3
"""Build HR Analytics Dashboard workbook."""
import random
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference

random.seed(42)
OUT = "/home/hatch/workspace/hr-projects/hr-analytics-dashboard/hr-analytics-dashboard.xlsx"

# ---------- theme ----------
NAVY = "1F4E79"
BLUE = "2E75B6"
LBLUE = "DDEBF7"
GREY = "F2F2F2"
DGREY = "595959"
WHITE = "FFFFFF"
GREEN = "548235"
AMBER = "BF8F00"
RED = "C00000"

hdr_font = Font(name="Calibri", bold=True, color=WHITE, size=11)
hdr_fill = PatternFill("solid", fgColor=NAVY)
title_font = Font(name="Calibri", bold=True, color=NAVY, size=16)
sec_font = Font(name="Calibri", bold=True, color=NAVY, size=12)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

wb = Workbook()

# =====================================================================
# SHEET 1: Read Me
# =====================================================================
ws = wb.active
ws.title = "Read Me"
ws.sheet_properties.tabColor = NAVY
ws.column_dimensions["A"].width = 110
readme_lines = [
    ("HR Analytics Dashboard", title_font),
    ("Built by Abdul Rehman — Senior HR Executive (1000+ hires)", Font(name="Calibri", italic=True, color=DGREY, size=11)),
    ("", None),
    ("What this workbook does", sec_font),
    ("Turns raw hiring and headcount data into the three metrics every HR leader watches:", None),
    ("1. Hiring Funnel — where candidates drop off, and which stage leaks the most.", None),
    ("2. Time to Hire — how long hiring really takes, by department.", None),
    ("3. Turnover — monthly and annualized separation rates, plus tenure risk.", None),
    ("", None),
    ("How to use it", sec_font),
    ("1. Replace the SAMPLE data on the 'Hiring Data' and 'Headcount' sheets with your own.", None),
    ("   Keep the column order and date format (YYYY-MM-DD). Leave a cell blank when a", None),
    ("   stage did not happen (e.g. no Offer Date for rejected candidates).", None),
    ("2. Everything else recalculates automatically — Funnel, Time to Hire, Turnover,", None),
    ("   and the Dashboard charts all run on formulas. No macros, no plugins.", None),
    ("3. Review the Dashboard sheet for the headline KPIs each month.", None),
    ("", None),
    ("Sheets", sec_font),
    ("Hiring Data   — one row per candidate, all stages and dates (sample: 75 candidates).", None),
    ("Headcount     — monthly opening headcount, joiners and leavers by department.", None),
    ("Funnel        — stage-by-stage conversion, calculated from Hiring Data.", None),
    ("Time to Hire  — average days per hiring phase, by department.", None),
    ("Turnover      — monthly turnover %, annualized rate, and tenure mix.", None),
    ("Dashboard     — KPI cards plus charts. Print this sheet for leadership reviews.", None),
    ("", None),
    ("Notes on the sample data", sec_font),
    ("The data is fictional and models a ~200-person BPO/consulting company in Pakistan", None),
    ("over 12 months (Oct 2025 – Sep 2026). It is realistic on purpose: Customer Support", None),
    ("hires fast with higher churn; specialist roles take longer to fill.", None),
]
for i, (txt, font) in enumerate(readme_lines, start=1):
    c = ws.cell(row=i, column=1, value=txt)
    c.alignment = left
    if font:
        c.font = font

# =====================================================================
# SHEET 2: Hiring Data (sample)
# =====================================================================
ws = wb.create_sheet("Hiring Data")
ws.sheet_properties.tabColor = BLUE
headers = ["Candidate ID", "Name", "Role", "Department", "Source",
           "Applied Date", "Screened Date", "Interview Date", "Offer Date",
           "Join Date", "Final Status", "Last Stage"]
widths = [14, 22, 28, 20, 14, 14, 14, 14, 14, 14, 14, 16]
for j, (h, wdt) in enumerate(zip(headers, widths), start=1):
    c = ws.cell(row=1, column=j, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
    ws.column_dimensions[get_column_letter(j)].width = wdt

first_names = ["Ahmed", "Fatima", "Bilal", "Ayesha", "Usman", "Sana", "Ali", "Mahnoor",
               "Hamza", "Iqra", "Danish", "Hira", "Fahad", "Zainab", "Imran", "Sadia",
               "Kamran", "Rabia", "Shahid", "Nimra", "Tariq", "Asma", "Nabeel", "Kiran"]
last_names = ["Khan", "Ahmed", "Malik", "Shah", "Butt", "Raza", "Hussain", "Farooq",
              "Iqbal", "Qureshi", "Sheikh", "Mahmood", "Akhtar", "Javed", "Nawaz"]
dept_roles = {
    "Customer Support": ["Customer Support Executive", "Senior CSR", "Team Lead - Support"],
    "Human Resources": ["HR Executive", "Talent Acquisition Specialist"],
    "Information Technology": ["Software Engineer", "IT Support Officer"],
    "Finance": ["Accounts Officer", "Payroll Executive"],
    "Operations": ["Operations Executive", "Quality Analyst"],
}
sources = ["LinkedIn", "Indeed", "Rozee", "Referral", "Walk-in"]
dept_weights = ["Customer Support"] * 45 + ["Human Resources"] * 8 + ["Information Technology"] * 15 + ["Finance"] * 7 + ["Operations"] * 25

def rdate(d1, d2):
    return d1 + timedelta(days=random.randint(0, (d2 - d1).days))

row = 2
for i in range(1, 76):
    dept = random.choice(dept_weights)
    role = random.choice(dept_roles[dept])
    name = f"{random.choice(first_names)} {random.choice(last_names)}"
    src = random.choice(sources)
    applied = rdate(date(2025, 11, 1), date(2026, 9, 15))
    # funnel progression probabilities
    r = random.random()
    screened = applied + timedelta(days=random.randint(1, 5)) if r < 0.80 else None
    interviewed = (screened + timedelta(days=random.randint(3, 12))) if screened and random.random() < 0.62 else None
    offered = (interviewed + timedelta(days=random.randint(2, 8))) if interviewed and random.random() < 0.55 else None
    if offered:
        if random.random() < 0.82:
            status, joined = "Hired", offered + timedelta(days=random.randint(7, 25))
            last_stage = "Hired"
        else:
            status, joined, last_stage = "Withdrawn", None, "Offered"
    elif interviewed:
        status, joined, last_stage = ("Rejected", None, "Interviewed") if random.random() < 0.9 else ("In Process", None, "Interviewed")
    elif screened:
        status, joined, last_stage = ("Rejected", None, "Screened") if random.random() < 0.9 else ("In Process", None, "Screened")
    else:
        status, joined, last_stage = ("Rejected", None, "Applied") if random.random() < 0.92 else ("In Process", None, "Applied")
    vals = [f"CAND-{i:04d}", name, role, dept, src, applied, screened, interviewed,
            offered, joined, status, last_stage]
    for j, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=j, value=v)
        c.border = border; c.alignment = center
        if isinstance(v, date):
            c.number_format = "YYYY-MM-DD"
    row += 1
ws.auto_filter.ref = f"A1:L{row-1}"
ws.freeze_panes = "A2"
N = row - 1  # 75

# =====================================================================
# SHEET 3: Headcount (sample, 12 months x 5 depts)
# =====================================================================
ws = wb.create_sheet("Headcount")
ws.sheet_properties.tabColor = "548235"
headers = ["Month", "Department", "Opening HC", "Joiners", "Leavers", "Closing HC"]
widths = [14, 20, 14, 12, 12, 14]
for j, (h, wdt) in enumerate(zip(headers, widths), start=1):
    c = ws.cell(row=1, column=j, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
    ws.column_dimensions[get_column_letter(j)].width = wdt
months = ["Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26",
          "Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26"]
base_hc = {"Customer Support": 118, "Human Resources": 8, "Information Technology": 24,
           "Finance": 10, "Operations": 38}
row = 2
for m in months:
    for dept, hc in base_hc.items():
        joiners = random.randint(2, 9) if dept == "Customer Support" else random.randint(0, 2)
        leavers = random.randint(2, 8) if dept == "Customer Support" else random.randint(0, 2)
        closing = hc + joiners - leavers
        for j, v in enumerate([m, dept, hc, joiners, leavers, closing], start=1):
            c = ws.cell(row=row, column=j, value=v)
            c.border = border; c.alignment = center
        base_hc[dept] = closing
        row += 1
ws.freeze_panes = "A2"
M = row - 1  # 60

# =====================================================================
# SHEET 4: Funnel (formulas)
# =====================================================================
ws = wb.create_sheet("Funnel")
ws.sheet_properties.tabColor = "BF8F00"
for j, (h, wdt) in enumerate(zip(["Stage", "Candidates", "% of Applied", "Conversion from Previous"],
                                 [26, 16, 16, 26]), start=1):
    c = ws.cell(row=1, column=j, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
    ws.column_dimensions[get_column_letter(j)].width = wdt
stages = ["Applied", "Screened", "Interviewed", "Offered", "Hired"]
# column L (12) = Last Stage, F (6) = Applied Date, K (11) = Final Status
for i, st in enumerate(stages):
    r = i + 2
    ws.cell(row=r, column=1, value=st).border = border
    if st == "Applied":
        f = f"=COUNTA('Hiring Data'!F2:F{N+1})"
    elif st == "Hired":
        f = f"=COUNTIF('Hiring Data'!K2:K{N+1},\"Hired\")"
    else:
        # reached stage = last stage at/after this stage, or hired
        order = {"Applied": 0, "Screened": 1, "Interviewed": 2, "Offered": 3, "Hired": 4}
        idx = order[st]
        later = '","'.join([s for s, o in order.items() if o >= idx])
        f = f"=SUM(COUNTIF('Hiring Data'!L2:L{N+1},{{\"{later}\"}}))"
    c = ws.cell(row=r, column=2, value=f); c.border = border; c.alignment = center
    c = ws.cell(row=r, column=3, value=f"=IF($B$2=0,0,B{r}/$B$2)")
    c.number_format = "0.0%"; c.border = border; c.alignment = center
    c = ws.cell(row=r, column=4, value=(f"=IF(B{r-1}=0,0,B{r}/B{r-1})" if i else "—"))
    if i: c.number_format = "0.0%"
    c.border = border; c.alignment = center
    ws.cell(row=r, column=1).alignment = left
# insight row
ws.cell(row=8, column=1, value="Biggest drop-off:").font = sec_font
ws.cell(row=9, column=1,
        value='=INDEX(A3:A6,MATCH(MIN(D3:D6),D3:D6,0)) & " stage (" & TEXT(MIN(D3:D6),"0.0%") & " conversion)"')
ws.merge_cells("A9:D9")

# =====================================================================
# SHEET 5: Time to Hire (formulas)
# =====================================================================
ws = wb.create_sheet("Time to Hire")
ws.sheet_properties.tabColor = "2E75B6"
for j, (h, wdt) in enumerate(zip(["Department", "Avg Days: Applied → Offer",
                                  "Avg Days: Offer → Join", "Avg Days: Applied → Join", "Hires"],
                                 [22, 22, 22, 22, 12]), start=1):
    c = ws.cell(row=1, column=j, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
    ws.column_dimensions[get_column_letter(j)].width = wdt
depts = ["Customer Support", "Human Resources", "Information Technology", "Finance", "Operations"]
for i, dept in enumerate(depts):
    r = i + 2
    ws.cell(row=r, column=1, value=dept).border = border
    rng_d, rng_i, rng_k, rng_j = f"'Hiring Data'!D2:D{N+1}", f"'Hiring Data'!I2:I{N+1}", \
                                 f"'Hiring Data'!K2:K{N+1}", f"'Hiring Data'!J2:J{N+1}"
    # avg applied->offer for hired in dept
    ws.cell(row=r, column=2, value=f"=IFERROR(AVERAGEIFS({rng_i},{rng_d},\"{dept}\",{rng_k},\"Hired\")-AVERAGEIFS('Hiring Data'!F2:F{N+1},{rng_d},\"{dept}\",{rng_k},\"Hired\"),\"\")")
    ws.cell(row=r, column=3, value=f"=IFERROR(AVERAGEIFS({rng_j},{rng_d},\"{dept}\",{rng_k},\"Hired\")-AVERAGEIFS({rng_i},{rng_d},\"{dept}\",{rng_k},\"Hired\"),\"\")")
    ws.cell(row=r, column=4, value=f"=IFERROR(B{r}+C{r},\"\")")
    ws.cell(row=r, column=5, value=f"=COUNTIFS({rng_d},\"{dept}\",{rng_k},\"Hired\")")
    for j in range(1, 6):
        c = ws.cell(row=r, column=j); c.border = border; c.alignment = center
        if j in (2, 3, 4): c.number_format = "0.0"
    ws.cell(row=r, column=1).alignment = left
r = len(depts) + 2
ws.cell(row=r, column=1, value="Company Average").font = Font(bold=True)
for j, col in [(2, "B"), (3, "C"), (4, "D")]:
    c = ws.cell(row=r, column=j, value=f"=AVERAGE({col}2:{col}{r-1})")
    c.number_format = "0.0"; c.font = Font(bold=True); c.border = border; c.alignment = center
c = ws.cell(row=r, column=5, value=f"=SUM(E2:E{r-1})")
c.font = Font(bold=True); c.border = border; c.alignment = center
ws.cell(row=r, column=1).border = border

# =====================================================================
# SHEET 6: Turnover (formulas)
# =====================================================================
ws = wb.create_sheet("Turnover")
ws.sheet_properties.tabColor = "C00000"
for j, (h, wdt) in enumerate(zip(["Month", "Avg Headcount", "Separations",
                                  "Monthly Turnover %", "Annualized Turnover %"],
                                 [14, 16, 14, 20, 22]), start=1):
    c = ws.cell(row=1, column=j, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
    ws.column_dimensions[get_column_letter(j)].width = wdt
for i, m in enumerate(months):
    r = i + 2
    ws.cell(row=r, column=1, value=m).border = border
    # avg HC = (sum opening + sum closing) / 2 across departments for the month
    ws.cell(row=r, column=2, value=f"=(SUMIFS(Headcount!C2:C{M+1},Headcount!A2:A{M+1},\"{m}\")+SUMIFS(Headcount!F2:F{M+1},Headcount!A2:A{M+1},\"{m}\"))/2")
    ws.cell(row=r, column=3, value=f"=SUMIFS(Headcount!E2:E{M+1},Headcount!A2:A{M+1},\"{m}\")")
    ws.cell(row=r, column=4, value=f"=IF(B{r}=0,0,C{r}/B{r})")
    ws.cell(row=r, column=5, value=f"=D{r}*12")
    for j in range(1, 6):
        c = ws.cell(row=r, column=j); c.border = border; c.alignment = center
    ws.cell(row=r, column=4).number_format = "0.0%"
    ws.cell(row=r, column=5).number_format = "0.0%"
    ws.cell(row=r, column=2).number_format = "0"
r = len(months) + 2
ws.cell(row=r, column=1, value="12-Month Average").font = Font(bold=True)
for j, col in [(4, "D"), (5, "E")]:
    c = ws.cell(row=r, column=j, value=f"=AVERAGE({col}2:{col}{r-1})")
    c.number_format = "0.0%"; c.font = Font(bold=True); c.border = border; c.alignment = center
ws.cell(row=r, column=1).border = border
# by-department annual turnover
r0 = r + 2
ws.cell(row=r0, column=1, value="Annualized Turnover by Department").font = sec_font
for j, h in enumerate(["Department", "Avg Headcount", "Separations (12m)", "Annual Turnover %"], start=1):
    c = ws.cell(row=r0 + 1, column=j, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
for i, dept in enumerate(depts):
    r = r0 + 2 + i
    ws.cell(row=r, column=1, value=dept).border = border
    ws.cell(row=r, column=2, value=f"=(SUMIFS(Headcount!C2:C{M+1},Headcount!B2:B{M+1},\"{dept}\")+SUMIFS(Headcount!F2:F{M+1},Headcount!B2:B{M+1},\"{dept}\"))/2")
    ws.cell(row=r, column=3, value=f"=SUMIFS(Headcount!E2:E{M+1},Headcount!B2:B{M+1},\"{dept}\")")
    ws.cell(row=r, column=4, value=f"=IF(B{r}=0,0,C{r}/B{r})")
    for j in range(1, 5):
        c = ws.cell(row=r, column=j); c.border = border; c.alignment = center
    ws.cell(row=r, column=4).number_format = "0.0%"
    ws.cell(row=r, column=1).alignment = left
    ws.cell(row=r, column=2).number_format = "0"

# =====================================================================
# SHEET 7: Dashboard
# =====================================================================
ws = wb.create_sheet("Dashboard")
ws.sheet_properties.tabColor = "1F4E79"
ws.column_dimensions["A"].width = 30
for col in "BCDE":
    ws.column_dimensions[col].width = 20
ws.cell(row=1, column=1, value="HR Analytics Dashboard").font = title_font
ws.cell(row=2, column=1, value="Oct 2025 – Sep 2026  •  Sample data — replace with your own on Hiring Data / Headcount").font = Font(italic=True, color=DGREY, size=10)
# KPI cards
kpis = [
    ("Total Hires (12m)", f"=COUNTIF('Hiring Data'!K2:K{N+1},\"Hired\")", "0"),
    ("Avg Time to Hire (days)", "='Time to Hire'!D7", "0.0"),
    ("Avg Monthly Turnover", f"=AVERAGE(Turnover!D2:D13)", "0.0%"),
    ("Offer Acceptance Rate", f"=IFERROR(COUNTIF('Hiring Data'!K2:K{N+1},\"Hired\")/COUNTIF('Hiring Data'!L2:L{N+1},\"Offered\")+COUNTIF('Hiring Data'!L2:L{N+1},\"Hired\"),0)", "0.0%"),
]
for i, (label, formula, fmt) in enumerate(kpis):
    r = 4 + i * 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    c = ws.cell(row=r, column=1, value=label)
    c.font = Font(bold=True, color=WHITE, size=11); c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = center
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=2)
    c = ws.cell(row=r + 1, column=1, value=formula)
    c.font = Font(bold=True, color=NAVY, size=18); c.alignment = center
    c.number_format = fmt
    c.fill = PatternFill("solid", fgColor=LBLUE)
    c.border = border

# --- Chart 1: Funnel (horizontal bar) ---
ch1 = BarChart()
ch1.type = "bar"; ch1.style = 10
ch1.title = "Hiring Funnel — Candidates by Stage"
ch1.y_axis.title = "Stage"; ch1.x_axis.title = "Candidates"
data = Reference(wb["Funnel"], min_col=2, min_row=1, max_row=6)
cats = Reference(wb["Funnel"], min_col=1, min_row=2, max_row=6)
ch1.add_data(data, titles_from_data=True)
ch1.set_categories(cats)
ch1.shape = 4
ws.add_chart(ch1, "D4")

# --- Chart 2: Time to hire by department (column) ---
ch2 = BarChart()
ch2.type = "col"; ch2.style = 10
ch2.title = "Avg Time to Hire by Department (days)"
ch2.y_axis.title = "Days"
data = Reference(wb["Time to Hire"], min_col=4, min_row=1, max_row=6)
cats = Reference(wb["Time to Hire"], min_col=1, min_row=2, max_row=6)
ch2.add_data(data, titles_from_data=True)
ch2.set_categories(cats)
ws.add_chart(ch2, "D19")

# --- Chart 3: Turnover trend (line) ---
ch3 = LineChart()
ch3.style = 10
ch3.title = "Monthly Turnover % Trend"
ch3.y_axis.title = "%"; ch3.y_axis.number_format = "0%"
ch3.x_axis.title = "Month"
data = Reference(wb["Turnover"], min_col=4, min_row=1, max_row=13)
cats = Reference(wb["Turnover"], min_col=1, min_row=2, max_row=13)
ch3.add_data(data, titles_from_data=True)
ch3.set_categories(cats)
ws.add_chart(ch3, "D34")

# print setup
ws.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(OUT)
print("written:", OUT)
