# Production-Surveillance-
Production surveillance dashboard analyzing 50 wells across 3 fields. Features XLOOKUP lookups, SUMIFS aggregation, conditional formatting alerts, and nested decision logic. Includes KPIs, charts, and technical reporting.

Project Overview

A professional production surveillance system** built in Excel that monitors 50 oil wells across 3 geographic fields (Alpha, Bravo, Charlie). The system automatically analyzes well performance, flags problems, and recommends engineering actions for a Production Superintendent.

Purpose:Transform raw production data into actionable engineering insights using advanced Excel functions, data validation, and conditional logic.

Real world Application: This reflects daily workflows at major oil & gas operators like Equinor, Shell, and Chevron where Production Surveillance Engineers review overnight well data and recommend actions before 7 AM operations meetings.

Key Features

1.Automated Lookup System
- XLOOKUP/INDEX MATCH formulas retrieve pump efficiency ratings
- Eliminates manual reference table lookups
- 50 wells processed in milliseconds

 2. Multi-Field Summarization
- SUMIFS calculates production totals by field
- Breaks down oil, gas, and water by geographic area
- Field-level KPIs for management reporting

3. Intelligent KPI Dashboard
- COUNTIFS counts wells by status (Active/Shut-in/Monitor)
- AVERAGEIFS calculates field-wide pressure, temperature, water cut
- 12 key performance indicators auto-calculated

 4. Visual Alert System
- Conditional formatting highlights problematic wells:
  - 🔴 Red: Oil < 1,500 bbl/day (low production)
  - 🟠 Orange: Water > 1,000 bbl/day (high water)
  - 🟢 Green: Oil > 3,500 bbl/day (excellent performance)
  - 🔵 Blue: Pressure > 3,100 psi (strong pressure)
- Problem wells visible at a glance; speeds decision-making

5. Automated Decision Matrix
- Nested IF statements recommend actions:
  - Workover Candidate (low oil)
  - Water Shutoff Study (high water)
  - Reservoir Review (low pressure)
  - Maintain Current Operations (high performers)
  - Restart Evaluation (shut-in wells)
  - Monitor (normal operations)

 6. Executive Dashboard
- 12 KPI cards for superintendent briefing
- 4 charts showing field trends
- One page technical report with recommendations

Dataset Description

Well Data (50 wells, 3 fields)

Fields:
- Alpha: 16 wells
- Bravo: 17 wells
- Charlie: 17 wells

Parameters per Well:
| Parameter | Range | Unit |

| Oil Production | 800–4,500 | bbl/day |
| Gas Production | 150–750 | Mscf/day |
| Water Production | 400–1,500 | bbl/day |
| Reservoir Pressure | 2,200–3,300 | psi |
| Temperature | 170–195 | °F |
| Pump Type | ESP, Rod Pump, Gas Lift | 
| Well Status | Active, Monitor, Shut-in | 

Data Quality: Realistic production patterns based on North Sea field analogues (Volve Field data).

 Lookup Tables

Pump Efficiency Reference:
| Pump Type | Efficiency |
| ESP | 94% |
| Rod Pump | 88% |
| Gas Lift | 91% |


 Excel Functions Used
Lookup Functions
- XLOOKUP / INDEX-MATCH: Retrieve pump efficiency for each well
  
  =IFERROR(INDEX($N$2:$N$4,MATCH(H2,$M$2:$M$4,0)),"N/A")
  

Aggregation Functions
- SUMIFS: Total production by field
  ```excel
  =SUMIFS(C:C,B:B,"Alpha")
  ```
- **COUNTIFS**: Count wells by status/field
  ```excel
  =COUNTIFS(I:I,"Active")
  ```
- **AVERAGEIFS**: Average pressure/water cut by field
  ```excel
  =AVERAGEIFS(F:F,B:B,"Bravo")
  ```

### Decision Functions
- **Nested IF**: Recommend engineering actions
  ```excel
  =IF(C2<1500,"Workover Candidate",
      IF(E2>1200,"Water Shutoff Study",
         IF(F2<2400,"Reservoir Review",
            IF(AND(I2="Active",C2>3500),"Maintain Current Operations",
               "Monitor"))))
  
- AND/OR: Multiple condition logic

 Calculation Functions
- Water Cut %: `=ROUND((E2/(C2+E2))*100,1)`

Dashboard Sheets

 Sheet 1: "Production Data"
Purpose:Raw data + calculated columns

Columns:
- A: Well ID
- B: Field
- C: Oil (bbl/day)
- D: Gas (Mscf/day)
- E: Water (bbl/day)
- F: Pressure (psi)
- G: Temperature (°F)
- H: Pump Type
- I: Status
- J: **Pump Efficiency** (XLOOKUP formula)
- K: **Water Cut %** (calculated)
- L: **Recommended Action** (nested IF formula)

Formatting:
- Conditional formatting rules applied to columns C, E, F
- Headers: Dark blue background, white text
- 50 data rows

 Sheet 2: "Field Summary"
Purpose: Aggregate production by field

Rows:
| Field | Oil Total | Gas Total | Water Total | Avg Pressure | Active Wells | Shut-in Wells |
|---|---|---|---|---|---|---|
| Alpha | (SUMIFS) | (SUMIFS) | (SUMIFS) | (AVERAGEIFS) | (COUNTIFS) | (COUNTIFS) |
| Bravo | ... | ... | ... | ... | ... | ... |
| Charlie | ... | ... | ... | ... | ... | ... |

Formula Pattern:
```excel
=SUMIFS(Column_to_Sum, Field_Column, "Field_Name")
```

---

### Sheet 3: "KPIs & Summary"
Purpose: Executive-level key performance indicators

Metrics:
1. Total Oil Production (all wells, all fields)
2. Average Oil Rate per well
3. Highest Producing Well (name + rate)
4. Lowest Producing Well (name + rate)
5. Number of Active Wells
6. Number of Shut-in Wells
7. Average Pressure (field-wide)
8. Average Water Cut (field-wide)
9. Total Gas Production
10. Total Water Production
11. Average Temperature
12. Fleet Efficiency Average

Formula Examples:
```excel
=SUM('Production Data'!C2:C51)                    [Total Oil]
=COUNTIF('Production Data'!I2:I51,"Active")      [Active Well Count]
=AVERAGE('Production Data'!K2:K51)                [Avg Water Cut]
```

---

 Sheet 4: "Alert Wells"
Purpose:Flag wells requiring immediate attention

Logic: Displays only wells that meet at least one alert condition:
- Oil < 1,500 bbl/day
- Water > 1,200 bbl/day
- Pressure < 2,400 psi
- Status = Shut-in

Columns:
- Well ID
- Field
- Oil Production
- Water Cut %
- Pressure
- Status
- Recommended Action

Rows:Auto-populated using IF logic; only problem wells appear

 Charts 

 Chart 1: Oil Production by Field
- Type: Column Chart
- X-axis:Field (Alpha, Bravo, Charlie)
- Y-axis: Total Oil (bbl/day)
- Data Source: Field Summary sheet SUMIFS formulas
- Insight: Shows which field produces most oil

 Chart 2: Water Production by Field
- Type: Column Chart
- X-axis: Field
- Y-axis: Total Water (bbl/day)
- Insight: Shows water disposal requirements by field

Chart 3: Well Status Distribution
- Type:Pie Chart
- Categories: Active, Monitor, Shut-in
- Data Source: COUNTIF formulas from KPIs sheet
- Insight: Shows operational health snapshot

Chart 4: Pump Type Distribution
- Type:Bar Chart
- Categories: ESP, Rod Pump, Gas Lift
- Data Source: COUNTIF by pump type
- Insight: Shows equipment composition

Technical Report

File:Production_Surveillance_Report.md

Contents:
1. Executive Summary 
2. Field Performance Overview (with table)
3. Wells Requiring Attention (prioritized by issue)
4. Green Flag Wells (performing well)
5. Shut-in Evaluation (restart opportunities)
6. Key Trends & Analysis (4-5 observations)
7. Priority Recommendations (immediate, short-term, medium-term)
8. Financial Impact Analysis
9. Conclusion & Next Steps

Audience: Production Superintendent

Tone:Professional, data driven, actionable

Key Metrics Called Out:
- Total production: 139,847 bbl/day
- Average water cut: 34.8%
- Active wells: 42 (84% availability)
- Problem wells: 8 (immediate attention)
- Financial opportunity: $170M/year (with interventions)

