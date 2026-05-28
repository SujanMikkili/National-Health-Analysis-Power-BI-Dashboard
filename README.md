# 🏥 National Health Analysis — Power BI Dashboard

![Power BI](https://img.shields.io/badge/Power%20BI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Healthcare-blue?style=for-the-badge)
![Records](https://img.shields.io/badge/Records-55%2C500-orange?style=for-the-badge)
![DAX](https://img.shields.io/badge/DAX-Measures%20%26%20Columns-red?style=for-the-badge)
![Power Query](https://img.shields.io/badge/Power%20Query-M%20Language-purple?style=for-the-badge)

> An end-to-end healthcare analytics dashboard built in Power BI, exploring patient admissions, billing patterns, medical conditions, and insurance distribution across 10 hospitals in the United States.

---

## 📌 Project Overview

This project analyzes a **55,500-record national healthcare dataset** to uncover insights around patient demographics, hospital billing, medical conditions, medication usage, and insurance provider distribution. The interactive Power BI report (`National_Health_Analysis.pbix`) enables stakeholders to drill down across multiple dimensions and make data-driven decisions.

---

## 📂 Repository Structure

```
National-Health-Analysis/
│
├── data/
│   └── Healthcare_Analysis_Dataset.xlsx   # Raw dataset (55,500 patient records)
│
├── report/
│   └── National_Health_Analysis.pbix      # Power BI Dashboard file
│
├── screenshots/
│   └── *.png                              # Dashboard preview images
│
└── README.md                              # Project documentation
```

---

## 📊 Dataset Description

**File:** `Healthcare_Analysis_Dataset.xlsx`  
**Sheet:** `Data` (55,500 rows) + `Data Dictionary`  
**Source:** Synthetic national healthcare data

### Columns

| Column | Type | Description |
|---|---|---|
| `Patient ID` | Text | Unique identifier for each patient |
| `Age` | Number | Patient age in years |
| `Gender` | Text | Male / Female / Non-binary |
| `Blood Type` | Text | Blood group (e.g., A+, B-, O+) |
| `Medical Condition` | Text | Primary diagnosis |
| `Date of Admission` | Date | Hospital admission date |
| `Doctor` | Text | Attending doctor (not necessarily a specialist) |
| `Hospital` | Text | Name of the hospital |
| `Insurance Provider` | Text | Medicare / UnitedHealthCare / Aetna / Cigna |
| `Billing Amount` | Decimal | Total treatment cost in USD |
| `Room Number` | Number | Assigned room during stay |
| `Admission Type` | Text | Elective / Urgent / Emergency |
| `Discharge Date` | Date | Date of discharge |
| `Medication` | Text | General medication prescribed |
| `Test Results` | Text | Normal / Abnormal / Inconclusive |
| `Hospital Latitude` | Decimal | Geographic latitude of hospital |
| `Hospital Longitude` | Decimal | Geographic longitude of hospital |

### Key Dataset Stats

- **Total Records:** 55,500 patients
- **Hospitals:** 10 unique hospitals
- **Medical Conditions:** Hypertension, Cancer, Asthma, Diabetes, Obesity, Arthritis
- **Insurance Providers:** Medicare, UnitedHealthCare, Aetna, Cigna
- **Admission Types:** Elective, Urgent, Emergency
- **Medications:** Ibuprofen, Lipitor, Penicillin, Paracetamol, Aspirin
- **Billing Range:** ~$0 to ~$52,764
- **Gender Split:** Female (50%), Male (40%), Non-binary (10%)

---

## 🔄 Power Query — Data Transformation Steps

The following transformations were applied in **Power Query (M Language)** before loading data into the model:

### 1. Change Data Types
```
Date of Admission → Date
Discharge Date    → Date
Billing Amount    → Decimal Number
Age               → Whole Number
Room Number       → Whole Number
```

### 2. Add Length of Stay Column
```m
= Table.AddColumn(#"Changed Type", "Length of Stay",
    each Duration.Days([Discharge Date] - [Date of Admission]),
    Int64.Type)
```

### 3. Filter Out Negative Billing Amounts
```m
= Table.SelectRows(#"Added LOS", each [Billing Amount] >= 0)
```

### 4. Add Age Group Column
```m
= Table.AddColumn(#"Filtered Rows", "Age Group", each
    if [Age] < 18 then "Minor (< 18)"
    else if [Age] < 35 then "Young Adult (18–34)"
    else if [Age] < 55 then "Middle Age (35–54)"
    else if [Age] < 70 then "Senior (55–69)"
    else "Elderly (70+)", type text)
```

### 5. Trim & Clean Text Columns
```m
= Table.TransformColumns(#"Added Age Group", {
    {"Hospital", Text.Trim},
    {"Doctor", Text.Trim},
    {"Medical Condition", Text.Trim}
  })
```

---

## 📐 DAX Measures

All DAX measures are organized in a dedicated **`_Measures`** table inside the Power BI model.

### 🔢 Core KPIs

```dax
-- Total Patients
Total Patients = COUNTROWS(Data)

-- Total Billing Revenue
Total Billing = SUM(Data[Billing Amount])

-- Average Billing per Patient
Avg Billing = AVERAGE(Data[Billing Amount])

-- Average Length of Stay
Avg Length of Stay =
AVERAGE(Data[Length of Stay])
```

### 🏥 Admission Analysis

```dax
-- Emergency Admissions Count
Emergency Admissions =
CALCULATE(COUNTROWS(Data), Data[Admission Type] = "Emergency")

-- Emergency Admission %
Emergency % =
DIVIDE([Emergency Admissions], [Total Patients], 0)

-- Elective Admissions %
Elective % =
DIVIDE(
    CALCULATE(COUNTROWS(Data), Data[Admission Type] = "Elective"),
    [Total Patients], 0
)
```

### 💰 Billing Analysis

```dax
-- Billing by Condition
Billing by Condition =
CALCULATE(SUM(Data[Billing Amount]),
    ALLEXCEPT(Data, Data[Medical Condition]))

-- Max Billing Amount
Max Billing =
MAXX(Data, Data[Billing Amount])

-- Billing Above Average Flag
High Billing Flag =
IF(Data[Billing Amount] > [Avg Billing], "Above Avg", "Below Avg")
```

### 🧪 Test Results Analysis

```dax
-- Abnormal Test Results Count
Abnormal Results =
CALCULATE(COUNTROWS(Data), Data[Test Results] = "Abnormal")

-- Abnormal Test Result %
Abnormal % =
DIVIDE([Abnormal Results], [Total Patients], 0)

-- Normal Results %
Normal % =
DIVIDE(
    CALCULATE(COUNTROWS(Data), Data[Test Results] = "Normal"),
    [Total Patients], 0
)
```

### 📅 Time Intelligence

```dax
-- Admissions Month over Month Change
MoM Admissions Change =
VAR CurrentMonth = [Total Patients]
VAR PrevMonth =
    CALCULATE([Total Patients],
        DATEADD('Date'[Date], -1, MONTH))
RETURN
DIVIDE(CurrentMonth - PrevMonth, PrevMonth, 0)

-- Year-to-Date Billing
YTD Billing =
TOTALYTD(SUM(Data[Billing Amount]), 'Date'[Date])
```

### 👥 Demographics

```dax
-- Average Patient Age
Avg Age = AVERAGE(Data[Age])

-- % Female Patients
Female % =
DIVIDE(
    CALCULATE(COUNTROWS(Data), Data[Gender] = "Female"),
    [Total Patients], 0
)
```

---

## 🗓️ Date Table (DAX)

A dedicated **Date Table** was created for time intelligence:

```dax
Date Table =
ADDCOLUMNS(
    CALENDAR(MIN(Data[Date of Admission]), MAX(Data[Discharge Date])),
    "Year",        YEAR([Date]),
    "Month",       FORMAT([Date], "MMMM"),
    "Month No",    MONTH([Date]),
    "Quarter",     "Q" & QUARTER([Date]),
    "Week No",     WEEKNUM([Date]),
    "Day of Week", FORMAT([Date], "dddd"),
    "Is Weekend",  IF(WEEKDAY([Date], 2) > 5, TRUE, FALSE)
)
```

---

## 📈 Dashboard Highlights

The Power BI report (`National_Health_Analysis.pbix`) includes:

- **Patient Demographics** — Age group distribution, gender breakdown, blood type analysis
- **Medical Conditions** — Frequency and billing by condition across hospitals
- **Billing Analysis** — Average and total billing by hospital, insurance provider, and admission type
- **Admission Trends** — Elective vs. Urgent vs. Emergency patterns over time (MoM)
- **Medication Usage** — Medication distribution across conditions and patient groups
- **Test Results** — Normal vs. Abnormal vs. Inconclusive breakdown
- **Geographic View** — Hospital locations mapped using latitude/longitude
- **Insurance Insights** — Patient volume and billing by insurance provider
- **Length of Stay** — Average stay duration by condition and admission type

---

## 📸 Dashboard Preview

![Dashboard Overview](screenshots/overview.png)

> 💡 Download the `.pbix` file and open in **Power BI Desktop** (free) to explore the full interactive dashboard.

---

## 🛠️ Tools & Technologies

| Tool | Purpose |
|---|---|
| Microsoft Power BI Desktop | Dashboard development and visualization |
| Microsoft Excel | Raw data storage and data dictionary |
| Power Query (M Language) | Data cleaning and transformation |
| DAX | Calculated columns, measures, and KPIs |
| Power BI Map Visual | Geographic hospital mapping |

---

## ⚙️ How to Use

### Prerequisites
- [Power BI Desktop](https://powerbi.microsoft.com/en-us/desktop/) (free download)
- Microsoft Excel (to view the raw dataset)

### Steps

1. **Clone or download this repository**
   ```bash
   git clone https://github.com/YOUR_USERNAME/National-Health-Analysis.git
   cd National-Health-Analysis
   ```

2. **Open the dashboard**
   - Launch **Power BI Desktop**
   - Go to `File → Open report`
   - Select `report/National_Health_Analysis.pbix`

3. **Refresh data (if needed)**
   - Go to `Home → Transform data`
   - Update the file path in Power Query to point to your local copy of `data/Healthcare_Analysis_Dataset.xlsx`
   - Click `Close & Apply`

4. **Explore the report**
   - Use slicers to filter by hospital, condition, insurance provider, or time period
   - Hover over visuals for tooltips and drill-through options

---

## 🔍 Key Insights

- **Hypertension** is the most frequently diagnosed condition across all hospitals
- **Medicare** covers the largest patient share among all insurance providers
- **Elective admissions** dominate overall, but Emergency cases drive the highest average billing
- **Billing amounts** vary significantly across hospitals and admission types
- Patients aged **55–70** account for the highest volume of admissions
- **Abnormal test results** are disproportionately linked to Emergency admissions

---

## ⚠️ Data Quality Notes

- Some `Billing Amount` values are **negative** — filtered out in Power Query as likely data entry errors
- `Doctor` field may not reflect a specialist assignment
- `Medication` field may not be condition-specific
- Dataset is **synthetic** — intended for analytics practice and portfolio use only

---

## 🚀 Future Enhancements

- [ ] Publish to **Power BI Service** for live web sharing
- [ ] Add **Row-Level Security (RLS)** for hospital-level access control
- [ ] Build a **predictive model** for admission type classification (Python/ML)
- [ ] Integrate **real hospital data** for validation
- [ ] Add **What-If parameters** for billing scenario analysis
- [ ] Create a **mobile-optimized** layout in Power BI

---

## 🙋 Author

**[Mikkili Sujan]**  
📧 [sujan.mikkili2812@gmail.com]  
🔗 [LinkedIn Profile](https://www.linkedin.com/in/sujanmikkili/)  
🐙 [GitHub Profile](https://github.com/SujanMikkili)

---

## 📄 License

This project uses a **synthetic dataset** for educational and portfolio purposes. No real patient data is included. Feel free to use this project as a reference or portfolio piece.

---

*If you found this useful, consider giving the repo a ⭐ on GitHub!*
