# 🔄 Excel Power Query — NHS A&E Data Transformation

## 📊 Project Overview

This independent Excel practice project demonstrates how Microsoft Power Query can be used to import, combine, clean, transform, and refresh NHS England Accident & Emergency (A&E) data.

Using 12 monthly NHS England A&E CSV files, I developed a repeatable data preparation workflow to consolidate provider-level records and prepare regional attendance summaries.

The objective was to practise folder-based imports, appending data, merging lookup tables, grouping records, unpivoting columns, assigning data types, and refreshing queries.

**Key Questions:**

- How can multiple monthly CSV files be combined into one dataset?
- How can provider-level data be summarised by NHS England region?
- How can a lookup table add region codes and names?
- How can attendance categories be converted from columns into rows?
- How can Power Query automate repeatable data preparation?

**Project Type:** Independent Excel learning and portfolio practice exercise.

---

## 🛠️ Tools & Technologies

| Tool / Feature | Purpose |
|---|---|
| Microsoft Excel | Data storage and reviewing transformed results |
| Power Query | Importing, cleaning and transforming datasets |
| From Folder | Importing multiple monthly CSV files |
| Combine Files | Appending monthly files into one table |
| Group By | Aggregating regional attendance figures |
| Merge Queries | Joining regional summaries with a lookup table |
| Left Outer Join | Preserving summary records while adding matching region information |
| Unpivot Columns | Restructuring attendance categories into rows |
| Data Types | Ensuring dates, text and counts have appropriate formats |
| Refresh All | Updating loaded results from configured sources |

---

## 📁 Dataset Information

**Source:** NHS England — A&E Attendances and Emergency Admissions

**Reporting Period:** 12 monthly reporting files

**Geographical Coverage:** England

**Data Level:** Monthly NHS provider-level records

**Monthly Source Files:** 12 CSV files

**Combined Provider-Level Records:** 2,387

**Regional Summary Records:** 84

**Unpivoted Attendance Records:** 252

**NHS England Regions:** 7

**Official Source:**

https://www.england.nhs.uk/statistics/statistical-work-areas/ae-waiting-times-and-activity/

The source files contain A&E attendance information, provider identifiers, NHS organisational groupings, and waiting-time measures.

The monthly CSV files are stored in the `monthly-data/` folder.

---

## 📂 Workbook Structure

| Worksheet / Query | Description |
|---|---|
| `NHS_AE_Combined_2023_24` | Combined provider-level NHS A&E data |
| `NHS_AE_Regional_Summary` | Monthly regional attendance totals with lookup information |
| `NHS_AE_Unpivoted_Attendances` | Attendance categories transformed into a long-format table |
| `Region_Lookup` | Regional lookup data used to support the merge |
| `tblRegionLookup` | Power Query connection to the seven-region lookup table |

The workbook separates source data, aggregated summaries, lookup information, and transformed attendance records.

---

## 🧹 Power Query Data Cleaning & Transformation

### 1. Importing Monthly CSV Files

Imported 12 NHS England A&E CSV files using Excel's **Data → Get Data → From File → From Folder** functionality.

**Purpose:** Create a repeatable process for importing monthly data without manually copying individual files.

### 2. Combining Monthly Files

Used Power Query's Combine Files functionality to consolidate the monthly source files into one query.

**Result:** 2,387 combined records loaded into Excel.

**Purpose:** Prepare one consolidated dataset for analysis across multiple reporting periods.

### 3. Data Type Validation

Reviewed and assigned suitable data types to relevant columns.

| Data Category | Power Query Data Type |
|---|---|
| Reporting dates | Date |
| Organisation identifiers | Text |
| Region codes | Text |
| Region names | Text |
| Attendance counts | Whole Number |

**Purpose:** Prevent calculation and transformation errors caused by incorrect data types.

### 4. Grouping Regional Attendance Data

Used **Group By → Advanced** to aggregate attendance records by:

- `Parent Org`
- `Reporting_Date`

Created three aggregated columns:

| Output Column | Aggregation |
|---|---|
| `Total_Type1_Attendances` | Sum of Type 1 attendances |
| `Total_Type2_Attendances` | Sum of Type 2 attendances |
| `Total_Other_Attendances` | Sum of other A&E attendances |

**Power Query M example from the transformation:**

```powerquery
Table.Group(
    Source,
    {"Parent Org", "Reporting_Date"},
    {
        {"Total_Type1_Attendances",
            each List.Sum([#"A&E attendances Type 1"]),
            type nullable number},
        {"Total_Type2_Attendances",
            each List.Sum([#"A&E attendances Type 2"]),
            type nullable number},
        {"Total_Other_Attendances",
            each List.Sum([#"A&E attendances Other A&E Department"]),
            type nullable number}
    }
)
```

**Result:** 84 regional-month records.

**Purpose:** Convert detailed provider-level attendance data into monthly regional summaries.

### 5. Cleaning Regional Lookup Data

Prepared a lookup table containing seven NHS England regional groupings.

The lookup table includes:

- `Parent Org`
- `Region_Code`
- `Region_Name`

Applied text cleaning operations to improve consistency before merging.

**Power Query M example:**

```powerquery
Table.TransformColumns(
    Source,
    {{"Parent Org", Text.Trim, type text}}
)
```

**Purpose:** Remove unnecessary leading and trailing spaces from organisation names.

### 6. Merging Queries

Merged `NHS_AE_Regional_Summary` with `tblRegionLookup` using the `Parent Org` column.

**Join Type:** Left Outer

**Matching Column:** `Parent Org`

Expanded the following lookup fields:

- `Region_Code`
- `Region_Name`

**Power Query M example:**

```powerquery
Table.ExpandTableColumn(
    #"Merged Queries",
    "tblRegionLookup",
    {"Region_Code", "Region_Name"},
    {"Region_Code", "Region_Name"}
)
```

**Purpose:** Enrich regional attendance summaries with consistent region identifiers and readable region names.

After cleaning the matching text fields, the regional lookup columns showed valid values across the summary records.

### 7. Unpivoting Attendance Columns

Created the `NHS_AE_Unpivoted_Attendances` query by referencing the regional summary.

Preserved the identifier columns:

- `Parent Org`
- `Reporting_Date`
- `Region_Code`
- `Region_Name`

Unpivoted the three attendance columns into two fields:

- `Attendance_Type`
- `Attendance_Count`

**Power Query M example:**

```powerquery
Table.UnpivotOtherColumns(
    Source,
    {
        "Parent Org",
        "Reporting_Date",
        "Region_Code",
        "Region_Name"
    },
    "Attendance_Type",
    "Attendance_Count"
)
```

**Result:** 252 records.

**Calculation:**

```text
84 regional-month records × 3 attendance categories
= 252 unpivoted records
```

**Purpose:** Convert wide-format attendance data into a long-format structure suitable for filtering, PivotTables and comparative analysis.

### 8. Refreshing Power Query Data

Used Excel's **Data → Refresh All** functionality to refresh the queries.

The workflow retains the transformations so they can be reapplied when the configured source data is refreshed.

The `tblRegionLookup` query was configured as **Connection Only** to avoid loading an unnecessary duplicate table into the workbook.

**Note:** Refresh depends on the source file paths remaining accessible. Refreshing after moving the workbook to another computer has not been independently verified.

---

## 📊 Data Transformation Results

| Query | Loaded Records | Purpose |
|---|---:|---|
| `NHS_AE_Combined_2023_24` | 2,387 | Consolidated provider-level data |
| `NHS_AE_Regional_Summary` | 84 | Regional attendance aggregation |
| `tblRegionLookup` | 7 lookup records | Regional lookup connection |
| `NHS_AE_Unpivoted_Attendances` | 252 | Long-format attendance dataset |

### Regional Lookup

The seven regions used in the lookup are:

| Region Code | Region Name |
|---|---|
| LON | London |
| SE | South East |
| SW | South West |
| MID | Midlands |
| E | East of England |
| NW | North West |
| NEY | North East and Yorkshire |

---

## 🔍 Key Findings

### Finding 1 — Consolidated Monthly Data

Power Query combined 12 monthly NHS England A&E CSV files into a single provider-level dataset containing 2,387 records.

This reduced the need for repetitive manual imports and established a reusable data preparation workflow.

### Finding 2 — Regional Aggregation Simplified the Dataset

Grouping provider-level records by region and reporting month produced 84 regional summary records.

The lookup merge added region codes and names, allowing the data to be identified consistently across regional analyses.

### Finding 3 — Unpivoting Improved the Data Structure

Unpivoting three attendance categories transformed 84 regional-month records into 252 long-format records.

This structure makes it easier to compare attendance categories using Excel PivotTables, filters and charts.

---

## 🎯 Skills Demonstrated

- Importing CSV files into Power Query
- Combining monthly files from a folder
- Appending data into one consolidated dataset
- Validating and changing column data types
- Cleaning text fields
- Grouping records using multiple criteria
- Aggregating numerical values using Sum
- Creating and maintaining lookup tables
- Merging queries using a Left Outer Join
- Expanding merged table columns
- Unpivoting attendance columns
- Renaming transformed columns
- Creating connection-only queries
- Refreshing Power Query data
- Preparing structured datasets for further analysis

---

## 🖼️ Excel Practice Preview

Screenshots of the completed Power Query workflow should be stored in the `screenshots/` folder.

Recommended screenshots:

| Screenshot | Content |
|---|---|
| `combined_monthly_data.png` | Combined monthly NHS A&E records |
| `regional_summary.png` | Grouped regional attendance data |
| `merged_region_lookup.png` | Regional summary with merged region codes and names |
| `unpivoted_attendances.png` | Long-format attendance dataset |
| `power_query_connections.png` | Excel Queries & Connections showing loaded queries |

Once these screenshots are saved, they can be embedded in the README using Markdown image links.

---

## 📂 Project Files

| File / Folder | Description |
|---|---|
| `power_query_practice.xlsx` | Excel workbook containing the Power Query transformations |
| `monthly-data/` | Twelve monthly NHS England CSV source files |
| `screenshots/` | Screenshots of the Power Query workflow |
| `README.md` | Project documentation |

---

## 📚 Data Source & References

**NHS England — A&E Attendances and Emergency Admissions**

https://www.england.nhs.uk/statistics/statistical-work-areas/ae-waiting-times-and-activity/

The data is publicly available from NHS England and is used here for independent learning and portfolio development.

---

## ✅ Project Summary

This project demonstrates an end-to-end Excel Power Query workflow for transforming multiple monthly NHS England A&E files into structured analytical datasets.

The exercise covers importing, combining, cleaning, grouping, merging, unpivoting, and refreshing data, providing a foundation for repeatable reporting and further Excel analysis.

**Repository:** [Excel Practice Portfolio](https://github.com/Alirazasadaq6/excel-practice)