# 🛒 Flipkart E-Commerce Sales Analysis

> End-to-end sales analysis of 5,000 Flipkart orders (Jan 2023 – Aug 2026) using **Python, NumPy, Pandas and Matplotlib** — from raw Excel star-schema data to cleaned dataset, KPIs, charts and business insights.

---

## 📌 Overview

This project analyses an e-commerce dataset stored as a **star schema** (1 fact table + 3 dimension tables) in an Excel workbook. The script cleans and merges the data, calculates key business metrics, generates 7 visualisations and prints a text report with actionable business insights.

**Business questions answered**

- How much revenue and profit is the business making, and what is the overall margin?
- Which categories and products drive sales — and which are actually *profitable*?
- How do sales trend month over month?
- Which zones and customer segments contribute the most?
- How many orders get cancelled or returned?
- Do Flipkart Plus members spend more than Regular members?
- What is the repeat-purchase rate?

---

## 📂 Dataset

Source file: `Flipkart.xlsx`

| Sheet | Rows | Description |
|---|---|---|
| `Fact Orders` | 5,000 | Order-level data: quantity, unit price, discount, sales, profit, ship mode, payment mode, delivery status |
| `Dim Customer` | 1,200 | Name, gender, segment, membership type, city, state, zone |
| `Dim Product` | 450 | Product name, category, sub-category, brand, MRP, rating |
| `Dim Date` | 1,339 | Calendar attributes (year, quarter, month, day type) |

> ⚠️ The dataset is for learning/portfolio purposes. Customer names, emails and phone numbers are not real people's data.


---

## 🧹 Data Cleaning & Feature Engineering

1. Trimmed whitespace in all text columns
2. Removed duplicate rows
3. Converted numeric columns with `pd.to_numeric` (invalid values → `NaN`)
4. Removed rows with `Quantity <= 0` or `Unit Price <= 0`
5. Merged **Fact Orders** ⟶ **Dim Customer** ⟶ **Dim Product** (left joins) and dropped unmatched rows
6. Created new features:
   - `OrderMonth`, `OrderYear`, `OrderMonthName`
   - `ShippingDays` = Ship Date − Order Date
   - `ProfitMarginPct` = Profit ÷ Sales × 100
7. **Cancelled orders are excluded** from all revenue analysis (4,770 of 5,000 rows used)

---

## 📊 Key Results

| Metric | Value |
|---|---|
| Total Sales | **₹12.64 Cr** (₹126,421,929) |
| Total Profit | **₹1.31 Cr** (₹13,077,583) |
| Overall Profit Margin | **10.3%** |
| Total Orders | 2,867 |
| Units Sold | 9,258 |
| Average Order Value | ≈ ₹44,096 |
| Average Discount | 17.4% |
| Repeat Purchase Rate | 75.4% (823 of 1,092 customers) |

### 🔍 Insights

1. **Two categories dominate revenue** — Electronics (44.9%) and Appliances (40.4%) together make up ~85% of sales, but they have *low margins* (7.9% and 9.8%).
2. **High-margin categories are small** — Beauty & Personal Care (30.1%) and Fashion (24.2%) are the most profitable per rupee but contribute only ~6% of sales combined. Grocery has the weakest margin (7.1%).
3. **Top products** — Whirlpool Max, LG Lite and Realme Lite 8 lead by revenue.
4. **Seasonality** — Sales are volatile month to month. Peak month: **Oct 2025** (≈ ₹51.5L); lowest: **Apr 2023** (≈ ₹12.4L). Several October peaks (festive season) are visible.
5. **Zones** — Northeast leads (≈ ₹3.42 Cr), while West and Central are the weakest zones — an expansion opportunity.
6. **Segments** — Consumer segment contributes ~61% of sales; Corporate and Small Business follow.
7. **Fulfilment** — 82.1% delivered, 8.1% in shipping, **5.1% returned, 4.7% cancelled** (~10% of orders don't convert into final revenue).
8. **Membership** — Flipkart Plus members spend more per customer (≈ ₹1.21L) than Regular members (≈ ₹1.13L).
9. **Payments** — UPI is the most-used payment mode (1,012 orders), followed by Credit Card and Cash on Delivery.

---

## 🛠️ Tech Stack

- **Python 3.9+**
- **Pandas** – data cleaning, merging, aggregation
- **NumPy** – numerical calculations
- **Matplotlib** – charts
- **openpyxl** – reading the Excel workbook

---

## 📁 Project Structure

```
flipkart sales analysis/
├── E-Commerce Sales Analysis.py     # main analysis script
├── Flipkart.xlsx                    # raw data (4 sheets)
├── data/
│   └── flipkart data clean.csv      # cleaned & merged dataset
├── charts/
│   ├── 1 monthly sales trend.png
│   ├── 2 category sales share.png
│   ├── 3 top products.png
│   ├── 4 category profit margin.png
│   ├── 5 delivery status.png
│   ├── 6 sales by zone.png
│   └── 7 customer segment sales.png
└── README.md
```

---

## ▶️ How to Run

```bash
# 1. Clone the repo
git clone https://github.com/nithish-0706/Flipkart-E-Commerce-Sales-Analysis.git
cd flipkart-sales-analysis

# 2. Install dependencies
pip install pandas numpy matplotlib openpyxl

# 3. Run the analysis
python E-Commerce Sales Analysis.py
```
---


## 👤 Author

**Nithish**
📧 nithishjrp@gmail.com · 🔗 [LinkedIn](https://www.linkedin.com/in/nithish-jrp/) · 💻 [GitHub](https://github.com/nithish-0706)

---

## 📄 License

This project is open-sourced under the [MIT License](LICENSE).
