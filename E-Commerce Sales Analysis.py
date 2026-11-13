import numpy as np
import pandas as pd
import matplotlib

# Matplotlib backend
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import os


# ============================================================
# SETTINGS
# ============================================================

pd.set_option("display.width", 120)
pd.set_option("display.max_columns", None)

plt.rcParams["figure.dpi"] = 120


# ============================================================
# FILE PATHS
# ============================================================

SRC_PATH = r"C:\Users\Nithish\Downloads\Python\Ecommerce\Flipkart.xlsx"

# Create output folders automatically
DATA_DIR = r"C:\Users\Nithish\Downloads\Python\Ecommerce\data"
CHART_DIR = r"C:\Users\Nithish\Downloads\Python\Ecommerce\charts"

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(CHART_DIR, exist_ok=True)

CLEAN_PATH = os.path.join(DATA_DIR, "Flipkart data clean.csv")


# ============================================================
# CHECK EXCEL FILE
# ============================================================

if not os.path.exists(SRC_PATH):
    raise FileNotFoundError(
        f"\nExcel file not found!\n"
        f"Please check this path:\n{SRC_PATH}"
    )


print("\nExcel file found successfully!")


# ============================================================
# REPORT FUNCTION
# ============================================================

report_lines = []


def log(line=""):
    print(line)
    report_lines.append(line)


# ============================================================
# 1. LOAD ALL SHEETS
# ============================================================

log("=" * 70)
log("FLIPKART E-COMMERCE SALES ANALYSIS - INSIGHTS REPORT")
log("=" * 70)

log("\nLoading Excel file...")


xl = pd.ExcelFile(SRC_PATH)

# Show available sheets
log(f"\nAvailable sheets: {xl.sheet_names}")


# Required sheets
required_sheets = [
    "Fact Orders",
    "Dim Customer",
    "Dim Product",
    "Dim Date"
]


# Check sheets
missing_sheets = [
    sheet for sheet in required_sheets
    if sheet not in xl.sheet_names
]


if missing_sheets:
    raise ValueError(
        f"\nThese sheets are missing from Excel:\n{missing_sheets}\n"
        f"\nAvailable sheets are:\n{xl.sheet_names}"
    )


# Read sheets
fact = xl.parse("Fact Orders")
dim_customer = xl.parse("Dim Customer")
dim_product = xl.parse("Dim Product")
dim_date = xl.parse("Dim Date")


log(f"\nFact Orders:   {fact.shape[0]} rows, {fact.shape[1]} cols")
log(f"Dim Customer:  {dim_customer.shape[0]} rows")
log(f"Dim Product:   {dim_product.shape[0]} rows")
log(f"Dim Date:      {dim_date.shape[0]} rows")


# ============================================================
# 2. DATA CLEANING & MERGING
# ============================================================

log("\n" + "-" * 70)
log("STEP 1: DATA CLEANING & MERGING")
log("-" * 70)


before = len(fact)


# ------------------------------------------------------------
# Trim whitespace
# ------------------------------------------------------------

for d in (fact, dim_customer, dim_product):
    for col in d.select_dtypes(include="str").columns:
        d[col] = d[col].astype(str).str.strip()


# ------------------------------------------------------------
# Remove duplicate rows
# ------------------------------------------------------------

dup_rows = fact.duplicated().sum()

fact = fact.drop_duplicates()


# ------------------------------------------------------------
# Check required numeric columns
# ------------------------------------------------------------

required_fact_columns = [
    "Quantity",
    "Unit Price (INR)",
    "Sales (INR)",
    "Profit (INR)",
    "Discount (%)",
    "Order ID",
    "Customer ID",
    "Product ID",
    "Order Date",
    "Ship Date",
    "Delivery Status"
]


missing_columns = [
    col for col in required_fact_columns
    if col not in fact.columns
]


if missing_columns:

    raise ValueError(
        f"\nMissing columns in Fact Orders:\n{missing_columns}"
    )


# ------------------------------------------------------------
# Convert numeric columns
# ------------------------------------------------------------

numeric_columns = [
    "Quantity",
    "Unit Price (INR)",
    "Sales (INR)",
    "Profit (INR)",
    "Discount (%)"
]


for col in numeric_columns:

    fact[col] = pd.to_numeric(
        fact[col],
        errors="coerce"
    )


# ------------------------------------------------------------
# Remove invalid quantity / price
# ------------------------------------------------------------

bad_qty = (fact["Quantity"] <= 0).sum()

bad_price = (fact["Unit Price (INR)"] <= 0).sum()


fact = fact[
    (fact["Quantity"] > 0)
    &
    (fact["Unit Price (INR)"] > 0)
].copy()


log(f"Duplicate rows removed: {dup_rows}")

log(
    f"Rows with invalid quantity (<=0) removed: {bad_qty}"
)

log(
    f"Rows with invalid unit price (<=0) removed: {bad_price}"
)


# ============================================================
# MERGE FACT + DIMENSION TABLES
# ============================================================

log("\nMerging Fact Orders with Customer and Product tables...")


master = (
    fact
    .merge(
        dim_customer,
        on="Customer ID",
        how="left"
    )
    .merge(
        dim_product,
        on="Product ID",
        how="left"
    )
)


# ------------------------------------------------------------
# Check unmatched records
# ------------------------------------------------------------

if "Customer Name" in master.columns:

    unmatched_cust = master["Customer Name"].isnull().sum()

else:

    unmatched_cust = 0


if "Product Name" in master.columns:

    unmatched_prod = master["Product Name"].isnull().sum()

else:

    unmatched_prod = 0


log(
    f"Rows with no matching customer record: {unmatched_cust}"
)

log(
    f"Rows with no matching product record: {unmatched_prod}"
)


# Remove unmatched records
if "Customer Name" in master.columns and "Product Name" in master.columns:

    master = master.dropna(
        subset=["Customer Name", "Product Name"]
    )


# ============================================================
# FEATURE ENGINEERING
# ============================================================

master["Order Date"] = pd.to_datetime(
    master["Order Date"],
    errors="coerce"
)


master["Ship Date"] = pd.to_datetime(
    master["Ship Date"],
    errors="coerce"
)


master["OrderMonth"] = (
    master["Order Date"]
    .dt.to_period("M")
    .astype(str)
)


master["OrderYear"] = (
    master["Order Date"]
    .dt.year
)


master["OrderMonthName"] = (
    master["Order Date"]
    .dt.strftime("%b")
)


master["ShippingDays"] = (
    master["Ship Date"]
    - master["Order Date"]
).dt.days


# Profit margin
master["ProfitMarginPct"] = np.round(
    np.where(
        master["Sales (INR)"] != 0,
        master["Profit (INR)"]
        / master["Sales (INR)"]
        * 100,
        0
    ),
    2
)


after = len(master)


log(
    f"\nFinal merged & cleaned rows: "
    f"{after} (started with {before})"
)


# ============================================================
# SAVE CLEAN DATA
# ============================================================

master.to_csv(
    CLEAN_PATH,
    index=False
)


log(
    f"Cleaned master dataset saved -> {CLEAN_PATH}"
)


# ============================================================
# REMOVE CANCELLED ORDERS
# ============================================================

sales = master[
    master["Delivery Status"] != "Cancelled"
].copy()


if len(master) > 0:

    percentage = round(
        len(sales) / len(master) * 100,
        1
    )

else:

    percentage = 0


log(
    f"\nRows used for revenue analysis "
    f"(excluding Cancelled): {len(sales)} "
    f"({percentage}% of clean data)"
)


# ============================================================
# 3. OVERALL SALES PERFORMANCE
# ============================================================

log("\n" + "-" * 70)
log("STEP 2: OVERALL SALES PERFORMANCE")
log("-" * 70)


total_sales = sales["Sales (INR)"].sum()

total_profit = sales["Profit (INR)"].sum()


if total_sales != 0:

    overall_margin = (
        total_profit
        / total_sales
        * 100
    )

else:

    overall_margin = 0


total_orders = sales["Order ID"].nunique()

total_units = sales["Quantity"].sum()


if total_orders != 0:

    aov = total_sales / total_orders

else:

    aov = 0


avg_discount = sales["Discount (%)"].mean()


log(
    f"Total Sales:               ₹{total_sales:,.0f}"
)

log(
    f"Total Profit:              ₹{total_profit:,.0f}"
)

log(
    f"Overall Profit Margin:     {overall_margin:.1f}%"
)

log(
    f"Total Orders:              {total_orders:,}"
)

log(
    f"Total Units Sold:          {total_units:,}"
)

log(
    f"Average Order Value(AOV):  ₹{aov:,.2f}"
)

log(
    f"Average Discount Given:    {avg_discount:.1f}%"
)


# ============================================================
# DELIVERY STATUS
# ============================================================

status_counts = (
    master
    .groupby("Delivery Status")["Order ID"]
    .nunique()
)


status_pct = (
    status_counts
    / status_counts.sum()
    * 100
).round(1)


log(
    "\nDelivery Status Breakdown "
    "(unique orders):\n"
    +
    pd.concat(
        [
            status_counts,
            status_pct
        ],
        axis=1,
        keys=["Orders", "%"]
    ).to_string()
)


# ============================================================
# 4. PRODUCT & CATEGORY PERFORMANCE
# ============================================================

log("\n" + "-" * 70)
log("STEP 3: PRODUCT & CATEGORY PERFORMANCE")
log("-" * 70)


# ------------------------------------------------------------
# Top Products
# ------------------------------------------------------------

top_products = (
    sales
    .groupby("Product Name")
    .agg(
        UnitsSold=("Quantity", "sum"),
        Sales=("Sales (INR)", "sum"),
        Profit=("Profit (INR)", "sum"),
        Orders=("Order ID", "nunique")
    )
    .sort_values(
        "Sales",
        ascending=False
    )
    .head(10)
)


log(
    "\nTop 10 Best-Selling Products (by Sales):"
)

log(
    top_products.round(0).to_string()
)


# ------------------------------------------------------------
# Most profitable products
# ------------------------------------------------------------

most_profitable_products = (
    sales
    .groupby("Product Name")["Profit (INR)"]
    .sum()
    .sort_values(
        ascending=False
    )
    .head(5)
)


log(
    "\nTop 5 Most Profitable Products:"
)

log(
    most_profitable_products.round(0).to_string()
)


# ------------------------------------------------------------
# Category performance
# ------------------------------------------------------------

category_perf = (
    sales
    .groupby("Category")
    .agg(
        Sales=("Sales (INR)", "sum"),
        Profit=("Profit (INR)", "sum"),
        UnitsSold=("Quantity", "sum"),
        Orders=("Order ID", "nunique"),
        AvgDiscount=("Discount (%)", "mean")
    )
    .sort_values(
        "Sales",
        ascending=False
    )
)


category_perf["MarginPct"] = (
    category_perf["Profit"]
    /
    category_perf["Sales"]
    * 100
).round(1)


category_perf["SalesSharePct"] = (
    category_perf["Sales"]
    /
    category_perf["Sales"].sum()
    * 100
).round(1)


log(
    "\nCategory-wise Performance:"
)

log(
    category_perf.round(1).to_string()
)


# ------------------------------------------------------------
# Sub-category
# ------------------------------------------------------------

subcat_perf = (
    sales
    .groupby("Sub-Category")
    .agg(
        Sales=("Sales (INR)", "sum"),
        Profit=("Profit (INR)", "sum")
    )
    .sort_values(
        "Sales",
        ascending=False
    )
    .head(10)
)


log(
    "\nTop 10 Sub-Categories by Sales:"
)

log(
    subcat_perf.round(0).to_string()
)


# ------------------------------------------------------------
# Category insights
# ------------------------------------------------------------

most_profitable_category = (
    category_perf["MarginPct"].idxmax()
)

least_profitable_category = (
    category_perf["MarginPct"].idxmin()
)

top_category_by_sales = (
    category_perf.index[0]
)


# ============================================================
# 5. MONTHLY SALES TREND
# ============================================================

log("\n" + "-" * 70)
log("STEP 4: MONTHLY SALES TREND")
log("-" * 70)


monthly = (
    sales
    .groupby("OrderMonth")
    .agg(
        Sales=("Sales (INR)", "sum"),
        Profit=("Profit (INR)", "sum"),
        Orders=("Order ID", "nunique")
    )
    .sort_index()
)


monthly["MoM_Growth_%"] = (
    monthly["Sales"]
    .pct_change()
    .mul(100)
    .round(1)
)


log(
    monthly.round(0).to_string()
)


best_month = (
    monthly["Sales"].idxmax()
)

worst_month = (
    monthly["Sales"].idxmin()
)


# ============================================================
# 6. CUSTOMER BEHAVIOR
# ============================================================

log("\n" + "-" * 70)
log("STEP 5: CUSTOMER BEHAVIOR ANALYSIS")
log("-" * 70)


# ------------------------------------------------------------
# Customer Segment
# ------------------------------------------------------------

segment_perf = (
    sales
    .groupby("Segment")
    .agg(
        Sales=("Sales (INR)", "sum"),
        Orders=("Order ID", "nunique"),
        Customers=("Customer ID", "nunique")
    )
)


segment_perf["SalesPerCustomer"] = (
    segment_perf["Sales"]
    /
    segment_perf["Customers"]
).round(0)


log(
    "\nSales by Customer Segment:"
)

log(
    segment_perf.round(0).to_string()
)


# ------------------------------------------------------------
# Membership
# ------------------------------------------------------------

membership_perf = (
    sales
    .groupby("Membership Type")
    .agg(
        Sales=("Sales (INR)", "sum"),
        Orders=("Order ID", "nunique"),
        Customers=("Customer ID", "nunique"),
        AvgDiscount=("Discount (%)", "mean")
    )
)


membership_perf["SalesPerCustomer"] = (
    membership_perf["Sales"]
    /
    membership_perf["Customers"]
).round(0)


log(
    "\nSales by Membership Type:"
)

log(
    membership_perf.round(1).to_string()
)


# ------------------------------------------------------------
# Repeat customers
# ------------------------------------------------------------

order_counts_per_cust = (
    sales
    .groupby("Customer ID")["Order ID"]
    .nunique()
)


repeat_customers = (
    order_counts_per_cust > 1
).sum()


one_time_customers = (
    order_counts_per_cust == 1
).sum()


total_customers = (
    repeat_customers
    + one_time_customers
)


if total_customers > 0:

    repeat_rate = (
        repeat_customers
        /
        total_customers
        * 100
    )

else:

    repeat_rate = 0


log(
    f"\nRepeat customers: "
    f"{repeat_customers} | "
    f"One-time customers: "
    f"{one_time_customers}"
)


log(
    f"Repeat purchase rate: "
    f"{repeat_rate:.1f}%"
)


# ------------------------------------------------------------
# Top Customers
# ------------------------------------------------------------

top_customers = (
    sales
    .groupby(
        [
            "Customer ID",
            "Customer Name"
        ]
    )["Sales (INR)"]
    .sum()
    .sort_values(
        ascending=False
    )
    .head(10)
)


log(
    "\nTop 10 Customers by Spend:"
)

log(
    top_customers.round(0).to_string()
)


# ------------------------------------------------------------
# Zone
# ------------------------------------------------------------

zone_perf = (
    sales
    .groupby("Zone")["Sales (INR)"]
    .sum()
    .sort_values(
        ascending=False
    )
)


log(
    "\nSales by Zone:"
)

log(
    zone_perf.round(0).to_string()
)


# ------------------------------------------------------------
# Payment
# ------------------------------------------------------------

payment_perf = (
    sales
    .groupby("Payment Mode")["Order ID"]
    .nunique()
    .sort_values(
        ascending=False
    )
)


log(
    "\nOrders by Payment Mode:"
)

log(
    payment_perf.to_string()
)


# ------------------------------------------------------------
# Shipping
# ------------------------------------------------------------

shipmode_perf = (
    sales
    .groupby("Ship Mode")
    .agg(
        Orders=("Order ID", "nunique"),
        AvgShippingDays=("ShippingDays", "mean")
    )
    .sort_values(
        "Orders",
        ascending=False
    )
)


log(
    "\nOrders & Avg Shipping Days by Ship Mode:"
)

log(
    shipmode_perf.round(1).to_string()
)


# ============================================================
# 7. VISUALIZATIONS
# ============================================================

log("\n" + "-" * 70)
log("STEP 6: GENERATING CHARTS")
log("-" * 70)


def fmt_inr(ax, axis="y"):

    fmt = mticker.FuncFormatter(
        lambda x, _: (
            f"₹{x / 100000:,.1f}L"
            if abs(x) >= 100000
            else f"₹{x:,.0f}"
        )
    )

    if axis == "y":

        ax.yaxis.set_major_formatter(fmt)

    else:

        ax.xaxis.set_major_formatter(fmt)


# ------------------------------------------------------------
# 1. Monthly Sales Trend
# ------------------------------------------------------------

fig, ax = plt.subplots(
    figsize=(11, 5)
)


ax.plot(
    monthly.index,
    monthly["Sales"],
    marker="o",
    linewidth=2
)


ax.set_title(
    "Monthly Sales Trend",
    fontsize=13,
    fontweight="bold"
)


ax.set_xlabel("Month")

ax.set_ylabel("Sales")


fmt_inr(ax)


plt.xticks(
    rotation=60,
    fontsize=8
)


ax.grid(
    alpha=0.3
)


plt.tight_layout()


plt.savefig(
    os.path.join(
        CHART_DIR,
        "01_monthly_sales_trend.png"
    )
)


plt.close()


# ------------------------------------------------------------
# 2. Category Sales Share
# ------------------------------------------------------------

fig, ax = plt.subplots(
    figsize=(7.5, 7.5)
)


ax.pie(
    category_perf["Sales"],
    labels=category_perf.index,
    autopct="%1.1f%%",
    startangle=90
)


ax.set_title(
    "Sales Share by Category",
    fontsize=13,
    fontweight="bold"
)


plt.tight_layout()


plt.savefig(
    os.path.join(
        CHART_DIR,
        "02_category_sales_share.png"
    )
)


plt.close()


# ------------------------------------------------------------
# 3. Top 10 Products
# ------------------------------------------------------------

fig, ax = plt.subplots(
    figsize=(10, 6)
)


tp_sorted = (
    top_products
    .sort_values("Sales")
)


ax.barh(
    tp_sorted.index,
    tp_sorted["Sales"]
)


ax.set_title(
    "Top 10 Best-Selling Products by Sales",
    fontsize=13,
    fontweight="bold"
)


ax.set_xlabel("Sales")


fmt_inr(
    ax,
    axis="x"
)


plt.tight_layout()


plt.savefig(
    os.path.join(
        CHART_DIR,
        "03_top_products.png"
    )
)


plt.close()


# ------------------------------------------------------------
# 4. Category Profit Margin
# ------------------------------------------------------------

fig, ax = plt.subplots(
    figsize=(9, 5)
)


cat_margin_sorted = (
    category_perf["MarginPct"]
    .sort_values()
)


ax.barh(
    cat_margin_sorted.index,
    cat_margin_sorted
)


ax.set_title(
    "Profit Margin % by Category",
    fontsize=13,
    fontweight="bold"
)


ax.set_xlabel(
    "Profit Margin (%)"
)


ax.axvline(
    0,
    linewidth=0.8
)


plt.tight_layout()


plt.savefig(
    os.path.join(
        CHART_DIR,
        "04_category_profit_margin.png"
    )
)


plt.close()


# ------------------------------------------------------------
# 5. Delivery Status
# ------------------------------------------------------------

fig, ax = plt.subplots(
    figsize=(7, 5)
)


status_counts.sort_values(
    ascending=False
).plot(
    kind="bar",
    ax=ax
)


ax.set_title(
    "Delivery Status Breakdown",
    fontsize=13,
    fontweight="bold"
)


ax.set_ylabel(
    "Number of Orders"
)


plt.xticks(
    rotation=20
)


plt.tight_layout()


plt.savefig(
    os.path.join(
        CHART_DIR,
        "05_delivery_status.png"
    )
)


plt.close()


# ------------------------------------------------------------
# 6. Sales by Zone
# ------------------------------------------------------------

fig, ax = plt.subplots(
    figsize=(8, 5)
)


zone_perf.plot(
    kind="bar",
    ax=ax
)


ax.set_title(
    "Sales by Zone",
    fontsize=13,
    fontweight="bold"
)


fmt_inr(ax)


plt.xticks(
    rotation=0
)


plt.tight_layout()


plt.savefig(
    os.path.join(
        CHART_DIR,
        "06_sales_by_zone.png"
    )
)


plt.close()


# ------------------------------------------------------------
# 7. Customer Segment
# ------------------------------------------------------------

fig, ax = plt.subplots(
    figsize=(7, 5)
)


segment_perf["Sales"].sort_values().plot(
    kind="barh",
    ax=ax
)


ax.set_title(
    "Sales by Customer Segment",
    fontsize=13,
    fontweight="bold"
)


fmt_inr(
    ax,
    axis="x"
)


plt.tight_layout()


plt.savefig(
    os.path.join(
        CHART_DIR,
        "07_customer_segment_sales.png"
    )
)


plt.close()


log(
    f"\n7 charts saved to: {CHART_DIR}"
)


# ============================================================
# 8. BUSINESS INSIGHTS
# ============================================================

log("\n" + "=" * 70)
log("KEY BUSINESS INSIGHTS")
log("=" * 70)


cancel_rate = (
    status_counts.get("Cancelled", 0)
    /
    status_counts.sum()
    * 100
)


return_rate = (
    status_counts.get("Returned", 0)
    /
    status_counts.sum()
    * 100
)


top_product_name = (
    top_products.index[0]
)


top_payment = (
    payment_perf.index[0]
)


top_zone = (
    zone_perf.index[0]
)


plus_vs_regular = (
    membership_perf["SalesPerCustomer"]
)


plus_sales = plus_vs_regular.get(
    "Flipkart Plus",
    0
)


regular_sales = plus_vs_regular.get(
    "Regular",
    0
)


if plus_sales > regular_sales:

    membership_message = (
        "membership clearly drives higher spend, "
        "supporting continued investment in the Plus program."
    )

else:

    membership_message = (
        "membership status shows little spend advantage here, "
        "worth re-examining Plus program perks."
    )


insights = [

    f"1. Total sales stood at ₹{total_sales:,.0f} "
    f"with ₹{total_profit:,.0f} profit "
    f"({overall_margin:.1f}% overall margin) "
    f"across {total_orders:,} orders "
    f"(AOV ≈ ₹{aov:,.0f}).",


    f"2. '{top_category_by_sales}' leads in raw sales "
    f"({category_perf.loc[top_category_by_sales, 'SalesSharePct']}% share), "
    f"but '{most_profitable_category}' has the healthiest margin "
    f"({category_perf.loc[most_profitable_category, 'MarginPct']}%) "
    f"while '{least_profitable_category}' is the weakest "
    f"({category_perf.loc[least_profitable_category, 'MarginPct']}%).",


    f"3. '{top_product_name}' is the top-selling "
    f"product by revenue — a strong candidate for "
    f"featured placement and stock priority.",


    f"4. {best_month} was the strongest sales month "
    f"and {worst_month} the weakest.",


    f"5. Repeat purchase rate is "
    f"{repeat_rate:.1f}% "
    f"({repeat_customers} of "
    f"{total_customers} customers ordered more than once).",


    f"6. Flipkart Plus members spend "
    f"₹{plus_sales:,.0f} per customer vs "
    f"₹{regular_sales:,.0f} for Regular members — "
    f"{membership_message}",


    f"7. {cancel_rate:.1f}% of orders are cancelled "
    f"and {return_rate:.1f}% are returned — "
    f"combined that's "
    f"{cancel_rate + return_rate:.1f}% "
    f"of demand not converting to final revenue.",


    f"8. '{top_payment}' is the most used "
    f"payment method.",


    f"9. '{top_zone}' zone generates "
    f"the highest sales.",


    f"10. Average discount given is "
    f"{avg_discount:.1f}%. "
    f"Since {least_profitable_category} shows "
    f"the lowest margin, review its discounting."
]


# Print insights
for line in insights:

    log(
        f"\n{line}"
    )


# ============================================================
# SAVE REPORT
# ============================================================

# ============================================================
# SAVE REPORT
# ============================================================

log("\n" + "=" * 70)
log("END OF REPORT")
log("=" * 70)

REPORT_PATH = os.path.join(
    os.path.dirname(SRC_PATH),
    "insights_report.txt"
)

with open(REPORT_PATH, "w", encoding="utf-8") as f:
    f.write("\n".join(report_lines))

print(f"\nFull report saved -> {REPORT_PATH}")
print("\nAnalysis completed successfully!")
