# HR Analytics Dashboard

A practical Excel dashboard that turns raw hiring and headcount data into the
three metrics every HR leader actually watches: hiring funnel, time to hire,
and turnover.
Built by **Abdul Rehman** — Senior HR Executive, 1000+ hires.

## Why these three metrics?

Most HR reporting stops at "we hired X people." That tells leadership nothing
about *how* hiring is going. These three answer the questions that matter:

- **Hiring Funnel** — where do candidates drop off? If 80% vanish after the
  first interview, the problem is your interview process, not your sourcing.
- **Time to Hire** — how long does hiring really take, by department? What gets
  measured gets shortened.
- **Turnover** — what is your monthly separation rate, and is it getting better
  or worse? Annualized, so it is comparable across companies.

## What's inside

| File | What it is |
|---|---|
| `hr-analytics-dashboard.xlsx` | The full workbook: sample data, formula-driven analysis sheets, and a chart dashboard |
| `build-dashboard.py` | The Python script that generates the workbook (openpyxl) — tweak it to match your own data model |

### Sheets in the workbook

| Sheet | What it does |
|---|---|
| Read Me | Instructions and how to swap in your own data |
| Hiring Data | One row per candidate: role, department, source, every stage date, final status (75 sample rows) |
| Headcount | Monthly opening headcount, joiners and leavers by department (12 months) |
| Funnel | Stage-by-stage counts, % of applied, and conversion from previous stage — all formulas |
| Time to Hire | Average days per hiring phase (applied to offer, offer to join), by department — all formulas |
| Turnover | Monthly turnover %, annualized rate, and turnover by department — all formulas |
| Dashboard | KPI cards plus three charts: funnel, time-to-hire by department, turnover trend |

## How to use it

1. **Replace the sample data.** Overwrite the rows on `Hiring Data` and
   `Headcount` with your own. Keep the column order and use YYYY-MM-DD dates.
   Leave a cell blank when a stage did not happen (for example, no Offer Date
   for a rejected candidate) — the formulas handle blanks.
2. **Do nothing else.** Funnel, Time to Hire, Turnover and the Dashboard all
   recalculate automatically. No macros, no plugins, no Power Query.
3. **Review monthly.** Open the Dashboard sheet in your leadership review. The
   biggest drop-off stage is flagged automatically on the Funnel sheet.

## Reading the numbers

- **Funnel conversion below 50% at any stage** deserves a closer look. Screened
  to interviewed below 50% usually means weak screening criteria; interviewed
  to offered below 50% usually means interviewers are not aligned.
- **Time to hire** is measured in calendar days from applied date to join date.
  Anything consistently above 45 days for non-specialist roles is a red flag —
  good candidates do not wait.
- **Turnover** is separations divided by average headcount for the month,
  annualized (x12) for comparability. Healthy for most Pakistani service
  companies is under 20% annualized; Customer Support often runs higher.

## Customize it

- Add your own departments and roles to the sample data — the department tables
  pick up new values automatically if you extend the lists on each analysis sheet.
- Change the date range by adding more monthly rows to `Headcount`.
- The chart styles are standard Excel charts, so they survive being copied
  into PowerPoint for board decks.

## About the sample data

The data is fictional and models a ~200-person BPO/consulting company in
Pakistan over 12 months (Oct 2025 – Sep 2026). It is realistic on purpose:
Customer Support hires fast with higher churn; specialist roles take longer
to fill. Replace it with your real numbers before presenting anywhere.

## About

Built from real-world hiring experience: 1000+ hires across consulting,
logistics, facilities management, and customer operations in Pakistan and
the UAE.

- LinkedIn: https://www.linkedin.com/in/abdul-rehman-131912225/
- Portfolio: https://abdulrehman-hr.github.io
