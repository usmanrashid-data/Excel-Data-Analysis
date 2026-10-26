# 📊 Excel Data Analysis & Business Intelligence Portfolio

Welcome to my Excel portfolio repository! This project showcases practical data analysis work completed as part of the **LeepTalent Data Technician Bootcamp**. It highlights data cleaning, advanced spreadsheet functions, multi-variable filtering, dynamic PivotTables, and professional data visualization using retail, sales, and student datasets.

---

## 🛠️ Key Skills & Functions Demonstrated

* **Mathematical & Statistical Formulas:** Applied `SUM`, `SUMIF`, `AVERAGE`, and `AVERAGEIF` to aggregate financial totals and calculate commission metrics.
* **Date & Time Functions:** Extracted and manipulated temporal data using `DATE`, `MONTH`, and `YEAR` formulas.
* **Lookup & Logic Functions:** Utilized `VLOOKUP` for cross-table matching and the `SWITCH` function for dynamic conditional categorization (e.g., grouping sales volumes into "High", "Medium", and "Low" tiers).
* **Data Organization:** Filtered and sorted multi-column datasets (sorting by age, transaction ID, and numerical ranges).
* **PivotTables & Aggregations:** Built multi-level PivotTables summarising sales volume, profitability, and customer demographics across international markets and regions.
* **Data Visualization:** Developed line charts, stacked bar charts, box plots, and pie charts to translate raw spreadsheets into clear business insights.

---

## 📁 Project Modules & Highlights

### 1. Retail Sales & Commission Analysis (`retail_sales_dataset.xlsx`)
* **Core Tasks:** Applied multi-level sorting on customer age and transaction IDs. Calculated total and average commission figures in target cells (`P10` and `P11`) using `SUM` and `AVERAGE`.
* **Business Application:** Cleaned and evaluated customer demographics and transaction amounts to identify top-performing retail product categories (Clothing, Electronics, and Beauty).

### 2. Regional Sales & Performance Lab (`Bike_Sales_Pivot_Lab.xlsx` & `Bike_Sales_Visualizations_Lab.xlsx`)
* **Core Tasks:** Constructed interactive PivotTables grouping order quantities and profit margins by **Country**, **Age Group**, and **Customer Gender**.
* **Key Findings:** 
  * Identified that the **United States** and **Australia** generated the highest overall revenue ($30.8M and $25.4M respectively).
  * Uncovered market gaps, such as zero sales generated from customers under 25 in North America and specific European markets.

### 3. County Sales Volume Categorisation (`Task 2 Lab`)
* **Core Tasks:** Summarised county-level sales data for products like Laptops, Printers, and Smartphones using PivotTables.
* **Logic Application:** Implemented the `SWITCH` function to automatically categorise sales volumes:
  ```excel
  =SWITCH(TRUE, C2 > 600, "High", C2 >= 300, "Medium", "Low")
