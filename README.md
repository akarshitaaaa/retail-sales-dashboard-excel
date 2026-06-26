# retail-sales-dashboard-excel
# 📊 Retail Sales Analysis Dashboard — Microsoft Excel

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Data Analytics](https://img.shields.io/badge/Domain-Data%20Analytics-0078D4?style=for-the-badge)
![Dashboard](https://img.shields.io/badge/Type-Business%20Dashboard-F4A261?style=for-the-badge)
![Business Intelligence](https://img.shields.io/badge/BI-Business%20Intelligence-6A0DAD?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=for-the-badge)

> An interactive Excel dashboard built to analyze the Vrinda Store's 2022 annual sales data — uncovering actionable insights across sales channels, customer demographics, order fulfillment, and regional performance.

---

## 📋 Table of Contents

1. [Project Overview](#-project-overview)
2. [Dashboard Preview](#-dashboard-preview)
3. [Dataset](#-dataset)
4. [Dashboard Features](#-dashboard-features)
5. [Data Preparation](#-data-preparation)
6. [Excel Features Used](#-excel-features-used)
7. [Excel Functions Worksheet](#-excel-functions-worksheet)
8. [Business Insights](#-business-insights)
9. [Repository Structure](#-repository-structure)
10. [Skills Demonstrated](#-skills-demonstrated)
11. [Key Learnings](#-key-learnings)
12. [Future Improvements](#-future-improvements)
13. [Resume Description](#-resume-description)
14. [Interview Talking Points](#-interview-talking-points)
15. [Conclusion](#-conclusion)

---

## 🎯 Project Overview

| Attribute | Details |
|---|---|
| **Project Title** | Retail Sales Analysis Dashboard using Microsoft Excel |
| **Project Type** | Data Analytics · Business Intelligence · Excel Dashboard |
| **Dataset** | Vrinda Store Annual Sales Data (2022) |
| **Tool Used** | Microsoft Excel |
| **Objective** | Analyze retail sales performance across channels, demographics, order status, and geographies to enable data-driven business decisions |

This project simulates a real-world retail analytics scenario. The goal was to process raw transactional sales data, engineer relevant features, build dynamic Pivot-based charts, and assemble them into a single interactive dashboard — all within Microsoft Excel.

---

## 🖼 Dashboard Preview

> **Main Dashboard**

![Dashboard Preview](Dashboard.png)

> **Excel Functions Worksheet**

![Excel Functions](Excel%20Functions.png)

---

## 📂 Dataset

**Dataset Name:** Vrinda Store Annual Sales Dataset (2022)

The dataset contains transactional retail records across the full calendar year 2022. Key fields include:

| Field | Description |
|---|---|
| `Order ID` | Unique identifier for each order |
| `Customer ID` | Unique identifier for each customer |
| `Gender` | Customer gender (Male / Female) |
| `Age` | Customer age |
| `Age Group` | Derived category (Senior / Adult / Teenager) |
| `Date` | Order date |
| `Month` | Derived from Date (helper column) |
| `Order Status` | Delivered, Cancelled, Returned, Refunded |
| `Channel` | Sales platform (Amazon, Myntra, Flipkart, etc.) |
| `SKU` | Stock Keeping Unit identifier |
| `Category` | Product category |
| `Size` | Product size |
| `Quantity` | Units ordered |
| `Currency` | Transaction currency |
| `Amount` | Order value |
| `Ship City` | Destination city |
| `Ship State` | Destination state |
| `Postal Code` | Delivery postal code |

---

## 📊 Dashboard Features

The dashboard contains **6 interactive charts** powered by Pivot Tables and Pivot Charts, along with **3 slicers** for dynamic filtering.

### 📈 Charts

| # | Chart | Purpose |
|---|---|---|
| 1 | **Orders vs Sales** | Compares order volume against revenue on a monthly basis |
| 2 | **Sales: Men vs Women** | Breaks down revenue contribution by gender |
| 3 | **Order Status Distribution** | Visualizes proportion of delivered, cancelled, returned, and refunded orders |
| 4 | **Orders by Sales Channel** | Shows which platforms (Amazon, Myntra, Flipkart, etc.) drive the most orders |
| 5 | **Top 5 States by Sales** | Ranks the five highest revenue-generating states |
| 6 | **Orders: Age Group vs Gender** | Cross-analyzes order counts by age group and gender |

### 🎛 Interactive Slicers

The dashboard supports dynamic filtering via three slicers:

- **Month** — Filter performance by individual months
- **Channel** — Isolate data for specific sales platforms
- **Category** — Drill into product categories

All charts update simultaneously when a slicer selection is applied.

---

## 🧹 Data Preparation

The following preprocessing steps were applied to prepare the raw dataset for analysis:

- **Month Column** — A helper column was derived from the `Date` field to enable monthly aggregation in Pivot Tables
- **Age Group Column** — A derived column was created using nested `IF` functions to categorize customers into: `Senior`, `Adult`, and `Teenager`
- **Sorting & Filtering** — Data was sorted and filtered to ensure logical ordering and remove irrelevant views during exploration
- **Conditional Formatting** — Applied to the `Amount` column to visually highlight high- and low-value transactions using a color scale
- **Pivot Table Preparation** — Data was structured in a flat, tabular format to ensure compatibility with Excel Pivot Tables and Pivot Charts

---

## ⚙️ Excel Features Used

| Feature | Application in This Project |
|---|---|
| **Pivot Tables** | Aggregated sales, orders, and customer data across multiple dimensions (channel, gender, state, month, age group) |
| **Pivot Charts** | Converted Pivot Table summaries into dynamic, linked bar, pie, and line charts |
| **Slicers** | Enabled one-click interactive filtering across all connected Pivot Tables and charts simultaneously |
| **IF Function** | Used to create the `Age Group` helper column based on customer age thresholds |
| **SUMIFS** | Calculated conditional revenue totals — e.g., channel-specific or gender-specific sales figures |
| **COUNTIFS** | Counted orders meeting multiple criteria — e.g., delivered orders by a specific gender |
| **VLOOKUP** | Mapped SKU identifiers to brand/product information across sheets |
| **Conditional Formatting** | Applied gradient color scales to the `Amount` column for quick visual identification of order value bands |
| **Date Functions** | Extracted month information from order dates to support time-series analysis |
| **Sorting & Filtering** | Organized data by date, state, or channel for exploratory analysis and validation |

---

## 🧮 Excel Functions Worksheet

A separate worksheet titled **"Excel Functions"** is included in the workbook to demonstrate practical use of advanced Excel formulas on the actual dataset.

| Function | Use Case |
|---|---|
| **SUMIFS** | Calculates total revenue for individual sales channels — enabling channel-wise performance benchmarking |
| **COUNTIFS** | Counts the number of successfully delivered orders filtered by customer gender — supporting demographic fulfillment analysis |
| **VLOOKUP** | Performs SKU-to-brand mapping by looking up product identifiers across a reference table — demonstrating cross-sheet data retrieval |

This worksheet serves as a structured reference for formula logic alongside the main dashboard.

---

## 💡 Business Insights

The following insights were derived from the dashboard:

- **Women outperform men in sales contribution** — Female customers account for a larger share of total revenue compared to male customers
- **Amazon leads all sales channels** — Amazon generates the highest order volume among all listed platforms
- **Maharashtra is the top revenue state** — It ranks first among the Top 5 States by Sales on the dashboard
- **Adults place the most orders** — The Adult age group (typically 25–49) drives the largest portion of orders across both genders
- **Majority of orders are delivered successfully** — The Order Status chart confirms that most orders reach completion without returns, cancellations, or refunds
- **Monthly sales show seasonal variation** — Order volumes and revenue fluctuate across months, indicating potential seasonal demand patterns

> *Note: Exact percentages and revenue figures reflect what is visible in the dashboard charts and have not been manually estimated.*

---

## 🗂 Repository Structure

```
Retail-Sales-Analysis-Excel/
│
├── Vrinda Store Data Analysis.xlsx   # Main workbook with raw data, pivot tables, and dashboard
├── Dashboard.png                     # Screenshot of the completed Excel dashboard
├── Excel Functions.png               # Screenshot of the Excel Functions worksheet
├── README.md                         # Project documentation (this file)
│
└── Images/                           # Supporting visuals
```

---

## 🛠 Skills Demonstrated

- Data Cleaning & Preparation
- Exploratory Data Analysis (EDA)
- Business Analytics
- Dashboard Development
- Pivot Table Analysis
- Data Visualization
- Excel Reporting
- Business Insight Generation
- Spreadsheet Modeling

---

## 📚 Key Learnings

- Understood how to structure raw transactional data for Pivot Table compatibility, including the importance of flat table formats and consistent field naming
- Gained hands-on experience building multi-chart interactive dashboards using Pivot Charts and Slicers in Excel
- Learned how to engineer helper columns (Month, Age Group) using Date Functions and nested IF logic to enable more granular analysis
- Developed proficiency with SUMIFS and COUNTIFS for multi-criteria conditional aggregations that go beyond simple SUM and COUNT formulas
- Practiced applying VLOOKUP for cross-sheet lookups to enrich transactional data with reference information
- Understood how Conditional Formatting enhances data readability by surfacing patterns in numerical columns without requiring manual review
- Gained clarity on translating visual dashboard outputs into business language — converting chart patterns into actionable insights for stakeholders
- Appreciated the importance of designing dashboards with the end user in mind — ensuring slicers, labels, and layout support intuitive self-service analysis

---

## 🚀 Future Improvements

| Enhancement | Description |
|---|---|
| **Dynamic KPI Cards** | Add summary metric tiles (Total Revenue, Total Orders, Avg Order Value) at the top of the dashboard |
| **Additional Interactive Filters** | Introduce slicers for Category, Size, or Ship State for deeper drill-down capability |
| **Power BI Version** | Migrate the dashboard to Power BI for richer visuals, drill-through pages, and automated refresh |
| **SQL Integration** | Store the dataset in a relational database and query it via SQL to practice end-to-end data pipelines |
| **Automated Data Refresh** | Connect the workbook to a live data source and schedule automated refresh cycles |
| **Predictive Sales Analysis** | Apply Excel's Forecast Sheet or regression analysis to project future monthly sales trends |

---

## 📄 Resume Description

> **Built an interactive retail sales analytics dashboard in Microsoft Excel using the Vrinda Store 2022 dataset (~10K+ transactions); engineered helper columns using IF and Date functions; developed 6 Pivot Chart visualizations with dynamic slicers; and extracted business insights on sales channel performance, customer demographics, order fulfillment, and regional revenue distribution.**

---

## 🎤 Interview Talking Points

Use this section to structure your response to **"Walk me through your project"** during interviews.

---

### ⏱ Project Overview (60-second pitch)

> "I built an interactive retail sales analytics dashboard in Microsoft Excel using annual transaction data from Vrinda Store for 2022. The goal was to help business stakeholders understand sales performance across different dimensions — channels, demographics, geographies, and order outcomes — without needing a dedicated BI tool. I cleaned and prepared the data, engineered helper columns, built Pivot Tables, and assembled a single-page dashboard with 6 charts and 3 interactive slicers."

---

### ❓ Business Problem

Retail businesses collect large volumes of transactional data but often struggle to extract timely insights from it. The business needed answers to questions like: *Which sales channels are driving the most revenue? Which customer segments order the most? Which states should be prioritized for logistics investment?* This project addressed those questions through a self-service Excel dashboard.

---

### 📦 Dataset

The dataset is the **Vrinda Store Annual Sales Data for 2022**, containing order-level records with fields covering customer demographics (Gender, Age), order metadata (Channel, Status, SKU, Category), financials (Amount, Quantity), and geography (Ship City, Ship State). Two helper columns — Month and Age Group — were engineered during preparation.

---

### 📊 Dashboard

The dashboard contains 6 Pivot Charts:
- Orders vs Sales (monthly trend)
- Sales by Gender
- Order Status breakdown
- Orders by Sales Channel
- Top 5 States by Revenue
- Orders by Age Group and Gender

Three slicers — Month, Channel, and Category — allow users to filter all charts simultaneously for ad hoc analysis.

---

### ⚙️ Excel Features Used

Pivot Tables, Pivot Charts, and Slicers form the core of the dashboard. I used IF functions to create the Age Group column, Date Functions for the Month column, SUMIFS and COUNTIFS for metric validation in the Excel Functions worksheet, VLOOKUP for SKU-to-brand mapping, and Conditional Formatting on the Amount column to visually distinguish transaction value bands.

---

### 💡 Business Insights

Key findings from the dashboard include: women drive more sales than men; Amazon is the dominant channel; Maharashtra generates the highest revenue among all states; adults place the most orders; and most orders are fulfilled successfully with minimal returns or cancellations.

---

### 🧗 Challenges Faced

One challenge was correctly structuring the data for Pivot Table compatibility — particularly ensuring no blank rows, consistent data types, and proper column headers. Another was connecting multiple Pivot Charts to the same set of slicers, which required verifying that each chart's Pivot Cache was appropriately linked. Translating chart observations into concise business language was also a skill that required deliberate practice.

---

### 🎓 Key Learnings

This project reinforced the full workflow of a data analyst: from raw data preparation and feature engineering, through exploratory analysis and visualization, to business insight communication. I developed a stronger command of Excel's analytical features and learned how to structure a dashboard that a non-technical stakeholder can navigate independently.

---

## ✅ Conclusion

This project demonstrates a complete, end-to-end data analytics workflow — from raw data ingestion and preparation to interactive visualization and business insight generation — using only Microsoft Excel. It reflects practical skills that are directly applicable to analyst roles at data-driven organizations, and serves as a strong portfolio piece for roles in business analytics, reporting, and BI.

---

> 📌 *If you found this project useful, please consider giving it a ⭐ on GitHub!*

---

**Author:** Akarshita Sharma
