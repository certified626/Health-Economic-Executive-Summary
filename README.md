# 📊 Healthcare Economics & Operational Analytics (Excel)

[![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Power Query](https://img.shields.io/badge/Power_Query-00758F?style=flat&logo=powerbi&logoColor=white)](https://powerbi.microsoft.com/)
[![Pivot Tables](https://img.shields.io/badge/Pivot_Tables-Data_Model-green.svg)](https://support.microsoft.com/en-us/excel)
#systeyaokothznation
## 📌 Project Overview
This project presents an end-to-end **Health Economics & Clinical Analytics Dashboard** built entirely in **Microsoft Excel**. It evaluates patient billing, clinical length of stay (LOS), socioeconomic income dynamics, and insurance coverage across four regional healthcare facilities in Kenya (`Facility A`, `Facility B`, `Facility C`, and `Facility D`).

The project demonstrates advanced spreadsheet modeling: from raw data auditing and ETL transformation using **Power Query & Excel Formulas**, to data modeling with **Pivot Tables & Data Model Medians**, and final interactive **Executive KPI Dashboarding**.

---

## 🎯 Business & Clinical Objectives
1. **Financial Burden Assessment:** Quantify average billed treatment costs against patient monthly income levels.
2. **Bed & Resource Utilization:** Determine which diagnoses drive the highest inpatient length of stay (LOS).
3. **Facility Operational Benchmarking:** Benchmark cost efficiency and patient throughput across regional hospital branches.
4. **Payer Mix Dynamics:** Evaluate coverage distribution across public schemes (*NHIF, SHA*) versus *Private Insurance* and unbilled records.

---

## 📊 Executive KPIs Summary

| Metric | Cleaned KPI Value | Analytical Takeaway |
| :--- | :--- | :--- |
| **Total Recorded Encounters** | **200** | Full patient visit volume across 4 facilities. |
| **Total Valid Billed Revenue** | **KES 5,575,211.00** | Excludes placeholder outlier values (`9,999,999`). |
| **Average Treatment Cost** | **KES 39,822.94** | Mean billing cost per patient (Median: **KES 39,107.50**). |
| **Average Length of Stay (LOS)** | **9.61 Days** | Inpatient bed occupancy average (Median: **10.0 Days**). |
| **Average Monthly Income** | **KES 112,327.08** | Mean patient income (Median: **KES 99,399.00**). |
| **Cost-to-Income Ratio** | **~35.5%** | An average visit consumes over one-third of reported monthly income. |

---

## 🔍 In-Depth Analytical Findings

### 1. Resource Utilization & Cost by Diagnosis
* **Diabetes** represents the highest average treatment cost at **KES 43,967.05**, followed by **Asthma** at **KES 38,715.05**.
* **Pneumonia** incurs the longest bed occupancy (**10.57 days** average LOS), creating peak bed utilization demands.

| Diagnosis | Patients | Mean Cost (KES) | Median Cost (KES) | Total Cost (KES) | Mean LOS (Days) |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Asthma** | 40 | 38,715.05 | 36,579.50 | 851,731.00 | 9.13 |
| **Diabetes** | 31 | **43,967.05** | **42,644.00** | **923,308.00** | 8.84 |
| **Hypertension**| 32 | 32,767.50 | 23,251.50 | 655,350.00 | 9.67 |
| **Malaria** | 32 | 37,271.06 | 34,636.50 | 670,879.00 | 8.97 |
| **Pneumonia** | 30 | 36,697.00 | 32,861.50 | 880,728.00 | **10.57** |

---

### 2. Regional Facility Performance Benchmark

| Facility | Billed Encounters | Total Revenue (KES) | Avg. Cost / Patient (KES) | Avg. Length of Stay (Days) |
| :--- | :---: | :---: | :---: | :---: |
| **Facility A** | 28 | 1,204,444.00 | **43,015.86** | **8.32** |
| **Facility B** | 40 | **1,610,336.00** | 40,258.40 | 9.48 |
| **Facility C** | 40 | 1,507,295.00 | 37,682.38 | **11.03** |
| **Facility D** | 32 | 1,253,136.00 | 39,160.50 | 10.19 |

* **Facility B** is the top revenue contributor (**KES 1.61M**).
* **Facility A** shows the highest cost per patient (**KES 43,015.86**) despite the lowest LOS (**8.32 days**), reflecting intensive care per day.
* **Facility C** has the longest patient stays (**11.03 days**).

---

### 3. Payer & Coverage Dynamics
* **NHIF:** 37.3% of reported insured patients (Mean Cost: KES 43,659.44).
* **Private Insurance:** 36.5% of reported insured patients (Mean Cost: KES 37,084.04).
* **Social Health Authority (SHA):** 26.2% of reported insured patients (Mean Cost: KES 39,793.92).
* **Uninsured / Unreported:** **37.0%** of total records lacked documented insurance coverage.

---

## 🛠️ Excel Data Cleaning & Modeling Methodology

### 1. Data Cleaning & Standardization
* **Currency Formatting:** Handled mixed string formats (e.g., `'45,000'`) using Text-to-Columns and `TRIM()` / `SUBSTITUTE()`.
* **Outlier Removal:** Filtered dummy placeholder entries (`KES 9,999,999.00`) and negative length of stay values (`-5 days`).
* **Categorical Normalization:** Cleaned inconsistent gender entries (`F`, `Female`, `female`, `M`, `Male`, `MALE`) into standard values.
* **Missing Value Treatment:** Implemented subgroup median imputation using dynamic Excel formulas:
  ```excel
  =IF(ISBLANK(H2), MEDIAN(FILTER($H$2:$H$201, $J$2:$J$201=J2)), H2)
  ```

### 2. Excel Dashboard Architecture
* **Data Model Pivot Tables:** Added raw data to the Excel Data Model to calculate exact medians and dynamic aggregations.
* **Interactive Slicers:** Connected timeline slicers and multi-attribute filters (`Facility`, `Insurance`, `Diagnosis`) to all Pivot Tables.
* **Executive Presentation:** Applied gridline removal, card-style KPI callouts, and clean formatting for executive readability.

---

## 📁 Repository Structure
```bash
├── data/
│   ├── Health_Economics_Data_Raw.xlsx          # Original uncleaned dataset
│   └── Health_Economics_Data_Cleaned.xlsx      # Cleaned data model
├── dashboards/
│   └── Healthcare_Economics_Executive_Dashboard.xlsx # Final Excel interactive dashboard
├── docs/
│   └── Executive_Summary_Report.pdf            # PDF version of executive report
└── README.md                                   # Project documentation
```

---

## 👤 Author
* **Odhiambo Okoth**
* Data Analytics & Insights
* [LinkedIn Profile](www.linkedin.com/in/odhiambookoth) | [GitHub Profile](https://github.com/certified626)
