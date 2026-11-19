# NHS England A&E Dataset

## 📊 Dataset Overview

This folder contains the original monthly NHS England Accident & Emergency (A&E) Attendances and Emergency Admissions datasets used for the Excel A&E Performance Dashboard practice project.

## 📁 Dataset Information

| Detail | Description |
|---|---|
| Dataset | A&E Attendances and Emergency Admissions |
| Publisher | NHS England |
| Reporting Period | October 2023 – September 2026 |
| Coverage | England |
| Data Frequency | Monthly |
| Data Level | NHS provider |
| File Format | CSV |
| Number of Files | 36 |

## 🔗 Official Data Source

[NHS England – A&E Waiting Times and Activity](https://www.england.nhs.uk/statistics/statistical-work-areas/ae-waiting-times-and-activity/)

## 🧹 Data Preparation

The monthly CSV files were imported and combined using Microsoft Excel Power Query.

The preparation process included:

- Combining monthly datasets into one table.
- Validating data types.
- Removing national total rows to prevent double-counting.
- Creating calculated fields for performance analysis.
- Preparing data for PivotTables and dashboard visualizations.

## 📂 Folder Structure

```text
data/
├── README.md
└── NHS_AE_Source_Files/
    └── 36 monthly CSV files
```

## 🎯 Dataset Usage

These files support analysis of:

- Total A&E attendances
- Four-hour performance
- Twelve-hour waits
- Regional performance differences
- NHS provider performance

The processed data and dashboard are available in the `10-dashboard-exercise` folder.

## ⚠️ Disclaimer

This dataset is published by NHS England. The analysis is an independent Excel learning exercise and is not an official NHS England report.