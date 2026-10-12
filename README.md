# Microsoft Excel

A practical, tool-by-tool guide to Excel for data analysis, office reporting, and entry-level analyst roles.

## Table of Contents
1. [Goals and Setup](#1-goals-and-setup)
2. [Excel Interface](#2-excel-interface)
3. [Data Entry and Formatting](#3-data-entry-and-formatting)
4. [Formulas and References](#4-formulas-and-references)
5. [Core Functions](#5-core-functions)
6. [Data Cleaning](#6-data-cleaning)
7. [Sort, Filter, Tables, and Validation](#7-sort-filter-tables-and-validation)
8. [Lookup Functions](#8-lookup-functions)
9. [Logical, Text, Date, and Dynamic Array Functions](#9-logical-text-date-and-dynamic-array-functions)
10. [Analysis Tools](#10-analysis-tools)
11. [Charts](#11-charts)
12. [Dashboards](#12-dashboards)
13. [Power Query](#13-power-query)
14. [Power Pivot and DAX](#14-power-pivot-and-dax)
15. [Macros and VBA](#15-macros-and-vba)
16. [Quality Checks and Protection](#16-quality-checks-and-protection)
17. [Keyboard Shortcuts](#17-keyboard-shortcuts)
18. [8-Week Study Plan](#18-8-week-study-plan)
19. [Portfolio Projects](#19-portfolio-projects)
20. [Practice Data](#20-practice-data)
21. [Job-Readiness Checklist](#21-job-readiness-checklist)
22. [Learning Resources](#22-learning-resources)

---

## 1. Goals and Setup

By completing this roadmap, aim to:
- Enter and organize data accurately.
- Write and troubleshoot formulas.
- Clean datasets and check data quality.
- Summarize information with PivotTables.
- Create charts and interactive dashboards.
- Import and transform data with Power Query.
- Understand when Power Pivot, DAX, and VBA are useful.
- Explain findings clearly.

**Software:** Microsoft Excel for Microsoft 365 or a recent desktop version is recommended. Feature availability varies by version, platform, and licence. Some Power Query, Power Pivot, and VBA features require desktop Excel. Save standard workbooks as `.xlsx`; use `.xlsm` for macro-enabled workbooks.

Suggested folders:
```text
excel-learning/
├── 01-practice-data/
├── 02-formulas/
├── 03-pivot-tables/
├── 04-dashboards/
├── 05-power-query/
└── 06-portfolio-projects/
```

## 2. Excel Interface

Learn these terms and actions:

| Item | Meaning / use |
|---|---|
| Workbook | An Excel file |
| Worksheet | A sheet inside a workbook |
| Cell / range | One cell (`A1`) / group of cells (`A2:D20`) |
| Formula bar | View or edit cell content |
| Name Box | Navigate to a cell or named range |
| Ribbon | Tabs containing Excel commands |
| Freeze Panes | Keep headers visible while scrolling |

Practise creating, renaming, copying, moving, and deleting sheets; Find and Replace; saving and exporting; print settings; and understanding `.xlsx`, `.csv`, and `.xlsm`.

**Practice:** Create `Raw_Data`, `Clean_Data`, `Analysis`, and `Dashboard` worksheets.

## 3. Data Entry and Formatting

Learn:
- Entering text, numbers, dates, percentages, currency, and times.
- AutoFill, Flash Fill, and Format Painter.
- Font, fill, borders, alignment, wrapping, and number formats.
- Column widths, row heights, and Cell Styles.
- Converting a range to an Excel Table.

Remember: formatting changes display, not necessarily the underlying value.

**Practice:** Create a sales sheet with Date, Order ID, Product, Region, Quantity, Unit Price, and Revenue.

## 4. Formulas and References

All formulas start with `=`.

| Operator | Example | Meaning |
|---|---|---|
| `+` | `=A2+B2` | Add |
| `-` | `=A2-B2` | Subtract |
| `*` | `=A2*B2` | Multiply |
| `/` | `=A2/B2` | Divide |
| `&` | `=A2&" "&B2` | Join text |

References:
- Relative: `A1` changes when copied.
- Absolute: `$A$1` keeps row and column fixed.
- Mixed: `$A1` or `A$1` fixes one dimension.
- Cross-sheet: `=Sheet2!A1`

If quantity is in `E2` and unit price in `F2`:
```excel
=E2*F2
```

If the tax rate is in `$B$1`:
```excel
=E2*F2*(1+$B$1)
```

Common errors:
- `#N/A`: lookup did not find a match.
- `#VALUE!`: inappropriate value or data type.
- `#REF!`: invalid reference.
- `#DIV/0!`: division by zero.
- `#NAME?`: unrecognized function or name.
- `#####`: often a narrow column or an un-displayable date/time.

Check types, references, spelling, and brackets before hiding an error with `IFERROR`.

## 5. Core Functions

| Function | Purpose | Example |
|---|---|---|
| `SUM` | Add values | `=SUM(E2:E100)` |
| `AVERAGE` | Mean | `=AVERAGE(E2:E100)` |
| `MIN` / `MAX` | Lowest / highest | `=MAX(E2:E100)` |
| `COUNT` | Count numeric cells | `=COUNT(E2:E100)` |
| `COUNTA` | Count non-empty cells | `=COUNTA(A2:A100)` |
| `ROUND` | Round numbers | `=ROUND(E2,2)` |
| `SUBTOTAL` | Summarize filtered lists | `=SUBTOTAL(9,E2:E100)` |
| `AGGREGATE` | Calculation with optional exclusions | `=AGGREGATE(9,6,E2:E100)` |

**Practice:** Calculate total revenue, average sale, maximum sale, and number of orders. Verify totals independently.

## 6. Data Cleaning

Keep an untouched copy of the original data. Learn these tools:

- **Remove Duplicates:** remove confirmed duplicate records using chosen key columns.
- **Text to Columns:** split data by delimiter or fixed width.
- **Flash Fill:** infer a pattern from examples.
- **Find and Replace:** standardize inconsistent labels.
- **TRIM:** remove extra ordinary spaces.
- **CLEAN:** remove many non-printing characters.
- **PROPER / UPPER / LOWER:** standardize capitalization.
- **VALUE:** convert numeric text where appropriate.
- **Data Validation:** prevent invalid entries.
- **Power Query:** repeat cleaning steps after refresh.

Examples:
```excel
=TRIM(A2)
=CLEAN(A2)
=PROPER(A2)
=VALUE(A2)
=SUBSTITUTE(A2,"-","")
```

Check blank cells, duplicate IDs, inconsistent dates, spelling differences, numbers stored as text, and suspicious values. Investigate before deleting unusual rows.

**Practice:** Clean customer records containing extra spaces, inconsistent city names, duplicate IDs, and mixed date formats.

## 7. Sort, Filter, Tables, and Validation

### Sort and Filter
- Sort alphabetically, numerically, by date, or by multiple columns.
- Filter by text, number, date, or colour.
- Learn Advanced Filter for criteria-based filtering.

### Excel Tables
Create with **Insert → Table** or `Ctrl+T`.

Benefits include structured references, automatic expansion, filters, formula fill-down, and convenient sources for PivotTables and charts.

Example structured reference:
```excel
=[@Quantity]*[@[Unit Price]]
```

### Data Validation
Use **Data → Data Validation** for dropdown lists, numeric limits, date ranges, text length, or custom rules.

**Practice:** Create a status dropdown: `New`, `In Progress`, `Complete`.

## 8. Lookup Functions

| Function | Best use |
|---|---|
| `XLOOKUP` | Flexible lookup in modern Excel |
| `VLOOKUP` | Legacy vertical lookup |
| `INDEX` + `MATCH` | Flexible lookup, including older Excel |
| `XMATCH` | Find a matching item's position in supported versions |

Example: find a product name for the ID in `A2`, using IDs in `H2:H100` and names in `I2:I100`:
```excel
=XLOOKUP(A2,$H$2:$H$100,$I$2:$I$100,"Not found")
```

Legacy example:
```excel
=VLOOKUP(A2,$H$2:$J$100,2,FALSE)
```

`FALSE` requests an exact match. VLOOKUP searches the first column of the range and returns a value to its right.

Good habits:
- Ensure keys use consistent data types.
- Check duplicate keys and missing matches.
- Use exact match unless approximate match is intentional.
- Use `IFNA` or `IFERROR` only when the fallback makes sense.

**Practice:** Use a product master table to add product names and categories to transactions.

## 9. Logical, Text, Date, and Dynamic Array Functions

### Logical functions
```excel
=IF(E2>=1000,"High","Low")
=IFS(E2>=10000,"High",E2>=5000,"Medium",TRUE,"Low")
=AND(E2>=1000,F2="West")
=OR(F2="West",F2="East")
=IFERROR(A2/B2,0)
=IFNA(XLOOKUP(A2,H:H,I:I),"Missing")
```

### Conditional aggregation
```excel
=SUMIF(C:C,"Laptop",G:G)
=SUMIFS(G:G,C:C,"Laptop",D:D,"West")
=COUNTIF(D:D,"West")
=COUNTIFS(D:D,"West",E:E,">=10")
=AVERAGEIF(C:C,"Laptop",G:G)
```
In these examples, `C` is Product, `D` is Region, `E` is Quantity, and `G` is Revenue. Adapt the columns to your dataset.

### Text functions
- `LEFT`, `RIGHT`, `MID`: extract text.
- `LEN`: count characters.
- `TRIM`, `CLEAN`: clean text.
- `FIND`, `SEARCH`: locate text.
- `SUBSTITUTE`, `REPLACE`: replace text.
- `TEXTJOIN`, `CONCAT`: combine text.
- `TEXT`: format a number or date as text.

Example:
```excel
=TEXTJOIN(" ",TRUE,A2,B2)
```

### Date and time functions
Learn `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `EOMONTH`, `EDATE`, `NETWORKDAYS`, and `WORKDAY`.

```excel
=NETWORKDAYS(A2,B2)
```
This counts workdays between two dates using Excel's default weekend rules. Provide a holiday range if needed.

### Dynamic array functions
In supported versions, learn `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, and `SEQUENCE`.

```excel
=FILTER(A2:G100,D2:D100="West","No results")
=SORT(A2:G100,7,-1)
=UNIQUE(C2:C100)
```
Spill ranges need empty cells beside and below the formula.

## 10. Analysis Tools

### PivotTables
Use **Insert → PivotTable** to summarize data.

Learn to:
1. Select a clean source table.
2. Place fields in Rows, Columns, Values, and Filters.
3. Change Sum to Count or Average when appropriate.
4. Group dates into months or quarters.
5. Sort by values and show percentages of totals.
6. Refresh after source data changes.
7. Add slicers or a timeline.

### Other tools
- **What-If Analysis:** Goal Seek, Scenario Manager, Data Tables.
- **Analysis ToolPak:** descriptive statistics, correlation, regression, and other analysis (may require enabling the add-in).
- **Forecast Sheet:** basic time-series forecasts when data is suitable.
- **Solver:** optimize an objective under constraints; availability depends on version and add-in settings.

**Practice:** Create a PivotTable showing revenue by month, region, and product category, with a region slicer.

## 11. Charts

| Chart | Use |
|---|---|
| Bar / Column | Compare categories |
| Line | Trends over time |
| Stacked bar / column | Compare totals and composition |
| Scatter | Explore relationships between numeric variables |
| Histogram | Inspect numeric distributions |
| Pie / Doughnut | A small number of simple parts of a whole |
| Combo | Compare series with different patterns |

Chart checklist:
- Use a meaningful title and label units.
- Choose a suitable scale.
- Keep colours consistent and readable.
- Avoid unnecessary 3D effects.
- Use data labels only when helpful.
- Do not imply causation from correlation.

**Practice:** Make a monthly revenue line chart and a top-five-products bar chart.

## 12. Dashboards

A dashboard should answer business questions quickly.

Suggested layout:
1. Title and reporting period.
2. KPI cards: revenue, orders, average order value, growth.
3. Trend chart: revenue over time.
4. Comparison chart: region or category.
5. Detail table: top products or exceptions.
6. Slicers for date, region, and category.
7. Notes on data source, refresh date, and metric definitions.

Build process:
1. Define the audience and questions.
2. Clean and validate data.
3. Create PivotTables or summary formulas.
4. Create charts.
5. Connect slicers to relevant PivotTables.
6. Align objects and standardize formatting.
7. Test filters, refresh data, and reconcile totals.
8. Document assumptions and limitations.

Example metrics:
- Revenue = sum of sales value.
- Orders = count of unique order IDs, if available.
- Average order value = revenue / number of orders.
- Growth % = `(current period - previous period) / previous period`.

Be careful: counting transaction rows is not the same as counting unique orders if an order has multiple items.

## 13. Power Query

Power Query imports, cleans, combines, and refreshes data using repeatable steps.

Learn to:
- Import Excel, CSV, text files, and supported data sources.
- Promote headers and set data types.
- Keep or remove rows.
- Remove duplicates and replace values.
- Split, merge, fill, and group columns.
- Merge queries (join tables).
- Append queries (stack rows).
- Unpivot columns.
- Load to a worksheet or data model.

Workflow:
1. Select **Data → Get Data**.
2. Choose a source.
3. Open Power Query Editor.
4. Apply and rename transformation steps.
5. Check types and row counts.
6. Choose **Close & Load**.
7. Use **Refresh All** when source data changes.

**Merge** joins tables using matching keys. **Append** stacks rows from similar tables.

**Practice:** Combine monthly CSV files, standardize column names, remove confirmed duplicates, and load one clean table.

## 14. Power Pivot and DAX

Useful for related tables and larger analytical models. Availability depends on Excel edition and platform.

Learn:
- Fact tables for transactions.
- Dimension tables for dates, products, customers, and regions.
- One-to-many relationships using unique dimension keys.
- Measures and filter context.
- Avoiding ambiguous or unnecessary relationships.

Example DAX measure for a model with a `Sales` table:
```dax
Total Revenue := SUM(Sales[Revenue])
```

If the model contains `Sales[OrderID]`:
```dax
Order Count := DISTINCTCOUNT(Sales[OrderID])
```

**Practice:** Create Sales, Products, and Calendar tables, relate them, and report revenue by month and category.

## 15. Macros and VBA

VBA can automate repetitive tasks in desktop Excel. Learn it after core formulas and Power Query unless your target role requires macros.

Learning order:
1. Record a simple macro.
2. Save as `.xlsm`.
3. Inspect the VBA editor.
4. Learn variables, ranges, worksheets, loops, and conditions.
5. Add error handling and test on a copy.
6. Document the macro.

Example:
```vb
Sub FormatHeader()
    With Worksheets("Sales").Range("A1:G1")
        .Font.Bold = True
        .AutoFilter
    End With
End Sub
```
This assumes a worksheet named `Sales` and header range `A1:G1`. Adjust and test on a copy.

**Security:** Enable macros only from trusted sources. Never run unknown VBA on sensitive workbooks.

## 16. Quality Checks and Protection

Learn worksheet/workbook protection, comments or notes, version history, and appropriate sharing settings. Worksheet protection helps prevent accidental edits; it is not strong encryption or a replacement for access controls.

Before sharing a report:
- [ ] Confirm source files and reporting period.
- [ ] Check row counts before and after cleaning.
- [ ] Check duplicate keys and missing values.
- [ ] Validate date, currency, and number formats.
- [ ] Reconcile totals to the source.
- [ ] Test formulas and edge cases.
- [ ] Refresh queries and PivotTables.
- [ ] Test dashboard filters and slicers.
- [ ] Check for formula errors and broken references.
- [ ] Label units, definitions, and assumptions.
- [ ] Remove confidential information from shared copies.

## 17. Keyboard Shortcuts

Common Windows Excel shortcuts; availability can vary by keyboard layout and platform.

| Shortcut | Action |
|---|---|
| `Ctrl+C` / `Ctrl+V` | Copy / paste |
| `Ctrl+X` | Cut |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo |
| `Ctrl+S` | Save |
| `Ctrl+1` | Format Cells |
| `Ctrl+T` | Create a Table |
| `Ctrl+Shift+L` | Toggle filters |
| `Alt+=` | AutoSum |
| `F2` | Edit active cell |
| `Ctrl+Arrow` | Move to edge of data region |
| `Ctrl+Shift+Arrow` | Select to edge of data region |
| `Ctrl+;` | Insert current date |
| `Ctrl+Page Up / Page Down` | Switch worksheets |
| `F4` | Repeat an action or toggle reference types while editing a formula, depending on context |

**Practice:** Build a small report using keyboard shortcuts whenever practical.

## 18. 8-Week Study Plan

Assume 1–2 hours per day, five or six days each week. Adjust to your starting level.

| Week | Focus | Deliverable |
|---|---|---|
| 1 | Interface, data entry, formatting, basic formulas | Student marks sheet |
| 2 | References, sorting, filtering, Tables | Formatted sales table |
| 3 | Logical formulas, conditional aggregation, text/date functions | Formula workbook |
| 4 | Lookups and data cleaning | Clean customer/product data |
| 5 | PivotTables, PivotCharts, conditional formatting | Regional sales report |
| 6 | Charts, dynamic arrays, analysis tools | Sales analysis report |
| 7 | Power Query and dashboard design | Refreshable dashboard |
| 8 | Portfolio project, validation, presentation | Project and README |

Daily routine:
- 20 minutes: learn a concept.
- 40 minutes: hands-on practice.
- 20 minutes: solve a problem without a tutorial.
- Extra time: review mistakes and explain findings.

## 19. Portfolio Projects

Complete at least three projects. Include the workbook, a README, key findings, and limitations.

### Project 1: Retail Sales Analysis
Suggested fields: Order Date, Order ID, Product, Category, Region, Quantity, Unit Price, Discount, Revenue.

Tasks:
- Validate types and investigate duplicates.
- Calculate revenue and discount.
- Analyse monthly and regional performance.
- Create PivotTables and charts.
- Write three evidence-based findings.

### Project 2: Employee / HR Summary
Suggested fields: Employee ID, Department, Join Date, Status, Location.

Tasks:
- Check missing or duplicate employee IDs.
- Summarize headcount by department and status.
- Analyse joining trends.
- Protect personal and confidential employee information.

### Project 3: Interactive Management Dashboard
Tasks:
- Define KPIs and metric definitions.
- Build clean source tables and summaries.
- Add charts and slicers.
- Verify each KPI against source data.
- Document refresh instructions and assumptions.

### Project README template
```markdown
# Project Name

## Objective
What business question does the project answer?

## Tools
Excel formulas, PivotTables, charts, and Power Query (if used).

## Dataset
Describe the source, date range, and fields. Do not include confidential data.

## Steps
1. Data validation and cleaning
2. Analysis and calculations
3. Visualization and dashboard
4. Quality checks

## Key Findings
- Finding 1
- Finding 2
- Finding 3

## Limitations
Explain missing data, assumptions, and limitations.

## How to Use
Explain how to open the workbook and refresh data, if applicable.
```

## 20. Practice Data

Create a table like this, then expand it to 100–500 rows.

| Date | Order ID | Product | Region | Quantity | Unit Price |
|---|---|---|---|---:|---:|
| 2026-01-03 | ORD-1001 | Keyboard | West | 2 | 800 |
| 2026-01-04 | ORD-1002 | Mouse | North | 3 | 450 |
| 2026-01-06 | ORD-1003 | Monitor | West | 1 | 9000 |
| 2026-01-07 | ORD-1004 | Keyboard | South | 4 | 800 |
| 2026-01-09 | ORD-1005 | Mouse | East | 2 | 450 |

If the data is an Excel Table, add a Revenue column:
```excel
=[@Quantity]*[@[Unit Price]]
```

Questions:
1. What is total revenue?
2. Which product generated the most revenue?
3. Which region had the highest revenue?
4. What was the average order value?
5. How does revenue change by month?

This is only a tiny sample. Use more varied records before drawing conclusions.

## 21. Job-Readiness Checklist

### Core Excel
- [ ] Enter and format data accurately.
- [ ] Use relative, absolute, and mixed references.
- [ ] Write and troubleshoot common formulas.
- [ ] Use sorting, filtering, Tables, and validation.
- [ ] Clean data safely.

### Analysis and reporting
- [ ] Use lookups and conditional aggregation.
- [ ] Create PivotTables and PivotCharts.
- [ ] Choose appropriate charts.
- [ ] Build a usable dashboard.
- [ ] Reconcile totals and explain findings.

### Advanced tools
- [ ] Import and transform data using Power Query.
- [ ] Understand when Power Pivot and DAX are useful.
- [ ] Understand the purpose and risks of VBA.
- [ ] Document refresh steps and assumptions.

### Job preparation
- [ ] Finish three portfolio projects.
- [ ] Explain each project in 2–3 minutes.
- [ ] Practise formula, lookup, PivotTable, and data-cleaning exercises.
- [ ] Tailor your CV to the job description.
- [ ] For data analyst roles, continue with SQL and a BI tool.

Excel alone does not guarantee employment. Employers may also assess communication, domain knowledge, accuracy, problem-solving, and other role-specific skills.

## 22. Learning Resources

- **Microsoft Excel Support:** https://support.microsoft.com/en-us/excel
- **GCFGlobal Excel Tutorial:** https://edu.gcfglobal.org/en/excel/
- **Kaggle Datasets:** https://www.kaggle.com/datasets

Use official documentation to confirm feature availability and syntax for your Excel version.

## Recommended Next Skills for Data Analytics

After building a solid Excel foundation:
1. **SQL** — query, filter, join, and aggregate data.
2. **Power BI** — create interactive reports and dashboards.
3. **Statistics** — understand averages, variation, distributions, and comparisons.
4. **Python** — automate analysis and work with larger datasets.

You can study basic statistics and analytical thinking alongside Excel.

## Final Advice

Learn one feature at a time, practise it on data, and use it in a project. Do not memorize formulas without understanding the question they answer. A professional spreadsheet is accurate, understandable, refreshable, and documented—not just visually attractive.
