"""
E-commerce Sales & Customer Analytics  |  SQL + Python + Excel + Power BI
=========================================================================
Dataset : Option A (recommended) - real data, e.g. Kaggle "Brazilian E-Commerce
          (Olist)" or "Online Retail II". Put your CSV in data/ and adapt the
          LOAD section (columns needed: order_id, customer_id, order_date,
          product, category, quantity, unit_price).
          Option B - the script generates a realistic synthetic dataset
          automatically if data/orders_raw.csv is missing.

Run     : python ecommerce_analysis.py
Outputs : outputs/ (charts, Excel workbook, CSVs for Power BI, SQLite DB)
"""

import os
import sqlite3
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
SEED = 7
DATA_PATH = "data/orders_raw.csv"
OUT_DIR = "outputs"
os.makedirs("data", exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)


# --------------------------------------------------------------------------
# 1. SYNTHETIC DATA GENERATOR (only used if no real data is supplied)
# --------------------------------------------------------------------------
def generate_orders(n_customers=3000, seed=SEED):
    rng = np.random.default_rng(seed)
    catalog = {
        "Electronics": [("Headphones", 59), ("Smartwatch", 149), ("Power Bank", 29),
                        ("Bluetooth Speaker", 45), ("Webcam", 55)],
        "Fashion": [("T-Shirt", 18), ("Jeans", 48), ("Sneakers", 79),
                    ("Jacket", 99), ("Backpack", 39)],
        "Home & Kitchen": [("Air Fryer", 89), ("Blender", 49), ("Cookware Set", 120),
                           ("Vacuum Cleaner", 140), ("Desk Lamp", 25)],
        "Beauty": [("Face Serum", 28), ("Perfume", 65), ("Hair Dryer", 42),
                   ("Skincare Kit", 55)],
        "Books & Stationery": [("Novel", 12), ("Planner", 15), ("Notebook Set", 9),
                               ("Textbook", 45)],
    }
    cat_weights = [0.25, 0.30, 0.2, 0.15, 0.10]
    cities = ["Delhi", "Mumbai", "Bengaluru", "Lucknow", "Hyderabad", "Pune",
              "Kolkata", "Chennai", "Jaipur", "Prayagraj"]
    start, end = pd.Timestamp("2024-01-01"), pd.Timestamp("2025-12-31")
    days = (end - start).days

    rows, order_id = [], 100000
    for cid in range(1, n_customers + 1):
        signup = start + pd.Timedelta(days=int(rng.integers(0, days - 30)))
        city = rng.choice(cities)
        n_orders = max(1, int(rng.poisson(1.6)))          # most buy 1-3 times
        order_date = signup
        for _ in range(n_orders):
            # purchase gaps; Nov-Dec festive boost handled below
            order_date += pd.Timedelta(days=int(rng.exponential(60)))
            if order_date > end:
                break
            if order_date.month in (10, 11, 12) or rng.random() < 0.85:
                pass
            else:
                continue
            order_id += 1
            for _ in range(int(rng.integers(1, 4))):       # 1-3 line items
                cat = rng.choice(list(catalog), p=cat_weights)
                prod, price = catalog[cat][rng.integers(len(catalog[cat]))]
                rows.append({
                    "order_id": order_id, "customer_id": f"CUST{cid:05d}",
                    "order_date": order_date.date(), "city": city,
                    "category": cat, "product": prod,
                    "quantity": int(rng.integers(1, 3)),
                    "unit_price": round(price * rng.uniform(0.9, 1.1), 2),
                    "discount_pct": float(rng.choice([0, 0, 5, 10, 20],
                                                     p=[.4, .2, .15, .15, .1])),
                    "order_status": rng.choice(["Delivered", "Returned", "Cancelled"],
                                               p=[.89, .07, .04]),
                })
    df = pd.DataFrame(rows)
    # Inject realistic dirt so the cleaning step is meaningful
    dirty = rng.choice(df.index, 60, replace=False)
    df.loc[dirty[:20], "city"] = df.loc[dirty[:20], "city"].str.upper()
    df.loc[dirty[20:40], "quantity"] = np.nan
    df = pd.concat([df, df.sample(40, random_state=SEED)])   # duplicate rows
    return df.sample(frac=1, random_state=SEED).reset_index(drop=True)


# --------------------------------------------------------------------------
# 2. LOAD
# --------------------------------------------------------------------------
if os.path.exists(DATA_PATH):
    raw = pd.read_csv(DATA_PATH)
    print(f"Loaded {DATA_PATH}: {raw.shape}")
else:
    raw = generate_orders()
    raw.to_csv(DATA_PATH, index=False)
    print(f"[!] No data found -> generated synthetic data {raw.shape}")

# --------------------------------------------------------------------------
# 3. CLEANING  (document every step; interviewers ask about this)
# --------------------------------------------------------------------------
df = raw.copy()
n0 = len(df)
df = df.drop_duplicates()
print(f"Removed {n0 - len(df)} duplicate rows")

df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")
df = df.dropna(subset=["order_date", "customer_id", "unit_price"])
df["quantity"] = df["quantity"].fillna(df["quantity"].median()).astype(int)
df["city"] = df["city"].str.strip().str.title()
df["category"] = df["category"].str.strip()
df = df[(df["quantity"] > 0) & (df["unit_price"] > 0)]

# Revenue = quantity x price after discount
df["line_revenue"] = (df["quantity"] * df["unit_price"]
                      * (1 - df["discount_pct"] / 100)).round(2)
df["year_month"] = df["order_date"].dt.to_period("M").astype(str)
df["year"] = df["order_date"].dt.year
df["month"] = df["order_date"].dt.month
df["weekday"] = df["order_date"].dt.day_name()

# Business rule: only delivered orders count as revenue
sales = df[df["order_status"] == "Delivered"].copy()
print(f"Clean rows: {len(df)} | Delivered rows used for revenue: {len(sales)}")

# --------------------------------------------------------------------------
# 4. SQL ANALYSIS
# --------------------------------------------------------------------------
conn = sqlite3.connect(os.path.join(OUT_DIR, "ecommerce.db"))
df.assign(order_date=df["order_date"].dt.strftime("%Y-%m-%d")).to_sql(
    "orders", conn, if_exists="replace", index=False)

queries = {
    "KPI summary": """
        SELECT ROUND(SUM(line_revenue), 0)                      AS total_revenue,
               COUNT(DISTINCT order_id)                         AS total_orders,
               COUNT(DISTINCT customer_id)                      AS total_customers,
               ROUND(SUM(line_revenue) / COUNT(DISTINCT order_id), 2) AS avg_order_value
        FROM orders WHERE order_status = 'Delivered';""",

    "Monthly revenue with MoM growth (window function)": """
        WITH monthly AS (
            SELECT year_month, ROUND(SUM(line_revenue), 0) AS revenue
            FROM orders WHERE order_status = 'Delivered'
            GROUP BY year_month)
        SELECT year_month, revenue,
               LAG(revenue) OVER (ORDER BY year_month) AS prev_month,
               ROUND(100.0 * (revenue - LAG(revenue) OVER (ORDER BY year_month))
                     / LAG(revenue) OVER (ORDER BY year_month), 1) AS mom_growth_pct
        FROM monthly ORDER BY year_month LIMIT 12;""",

    "Category performance + share of revenue": """
        SELECT category,
               ROUND(SUM(line_revenue), 0) AS revenue,
               ROUND(100.0 * SUM(line_revenue) /
                     (SELECT SUM(line_revenue) FROM orders
                      WHERE order_status = 'Delivered'), 1) AS revenue_share_pct
        FROM orders WHERE order_status = 'Delivered'
        GROUP BY category ORDER BY revenue DESC;""",

    "Top 2 products per category (RANK)": """
        SELECT * FROM (
            SELECT category, product, ROUND(SUM(line_revenue), 0) AS revenue,
                   RANK() OVER (PARTITION BY category
                                ORDER BY SUM(line_revenue) DESC) AS rnk
            FROM orders WHERE order_status = 'Delivered'
            GROUP BY category, product)
        WHERE rnk <= 2 ORDER BY category, rnk;""",

    "Return and cancellation rate by category": """
        SELECT category,
               ROUND(100.0 * SUM(order_status = 'Returned')  / COUNT(*), 1) AS return_pct,
               ROUND(100.0 * SUM(order_status = 'Cancelled') / COUNT(*), 1) AS cancel_pct
        FROM orders GROUP BY category ORDER BY return_pct DESC;""",

    "Repeat customer rate": """
        WITH per_customer AS (
            SELECT customer_id, COUNT(DISTINCT order_id) AS n_orders
            FROM orders WHERE order_status = 'Delivered' GROUP BY customer_id)
        SELECT COUNT(*) AS customers,
               SUM(n_orders > 1) AS repeat_customers,
               ROUND(100.0 * SUM(n_orders > 1) / COUNT(*), 1) AS repeat_rate_pct
        FROM per_customer;""",

    "Revenue by city (top 5)": """
        SELECT city, ROUND(SUM(line_revenue), 0) AS revenue,
               COUNT(DISTINCT customer_id) AS customers
        FROM orders WHERE order_status = 'Delivered'
        GROUP BY city ORDER BY revenue DESC LIMIT 5;""",

    "Does discounting pay off?": """
        SELECT discount_pct, COUNT(*) AS line_items,
               ROUND(AVG(quantity), 2) AS avg_qty,
               ROUND(SUM(line_revenue), 0) AS revenue
        FROM orders WHERE order_status = 'Delivered'
        GROUP BY discount_pct ORDER BY discount_pct;""",
}
print("\n================ SQL INSIGHTS ================")
for title, q in queries.items():
    print(f"\n--- {title}")
    print(pd.read_sql(q, conn).to_string(index=False))
conn.close()

# --------------------------------------------------------------------------
# 5. CUSTOMER SEGMENTATION: RFM
#    Recency   = days since last purchase (lower is better)
#    Frequency = number of orders
#    Monetary  = total spend
# --------------------------------------------------------------------------
snapshot = sales["order_date"].max() + pd.Timedelta(days=1)
rfm = (sales.groupby("customer_id")
       .agg(Recency=("order_date", lambda s: (snapshot - s.max()).days),
            Frequency=("order_id", "nunique"),
            Monetary=("line_revenue", "sum"))
       .reset_index())

# Scores 1-5 (5 = best). rank(method="first") avoids duplicate-edge errors.
rfm["R"] = pd.qcut(rfm["Recency"].rank(method="first"), 5, labels=[5, 4, 3, 2, 1]).astype(int)
rfm["F"] = pd.qcut(rfm["Frequency"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
rfm["M"] = pd.qcut(rfm["Monetary"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)


def segment(r):
    if r.R >= 4 and r.F >= 4:
        return "Champions"
    if r.R >= 3 and r.F >= 3:
        return "Loyal Customers"
    if r.R >= 4 and r.F <= 2:
        return "New / Promising"
    if r.R <= 2 and r.F >= 4:
        return "At Risk (high value)"
    if r.R <= 2:
        return "Hibernating / Lost"
    return "Needs Attention"


rfm["Segment"] = rfm.apply(segment, axis=1)
rfm["Monetary"] = rfm["Monetary"].round(2)
seg = (rfm.groupby("Segment")
       .agg(Customers=("customer_id", "count"),
            Revenue=("Monetary", "sum"),
            Avg_Orders=("Frequency", "mean"),
            Avg_Recency_Days=("Recency", "mean"))
       .round(1).sort_values("Revenue", ascending=False))
seg["Revenue_Share_%"] = (100 * seg["Revenue"] / seg["Revenue"].sum()).round(1)
print("\n================ RFM SEGMENTS ================")
print(seg.to_string())

# --------------------------------------------------------------------------
# 6. COHORT RETENTION (month of first purchase vs months since)
# --------------------------------------------------------------------------
sales["order_month"] = sales["order_date"].dt.to_period("M")
sales["cohort"] = sales.groupby("customer_id")["order_date"].transform("min").dt.to_period("M")
sales["period"] = (sales["order_month"] - sales["cohort"]).apply(lambda x: x.n)
cohort = (sales.groupby(["cohort", "period"])["customer_id"].nunique()
          .unstack(fill_value=0))
retention = cohort.div(cohort[0], axis=0).round(3)
retention.index = retention.index.astype(str)
print("\nAverage retention (all cohorts) month 1 / 2 / 3:",
      retention[[1, 2, 3]].mean().mul(100).round(1).tolist(), "%")

# --------------------------------------------------------------------------
# 7. CHARTS
# --------------------------------------------------------------------------
monthly = sales.groupby("year_month")["line_revenue"].sum()
fig, ax = plt.subplots(2, 2, figsize=(15, 10))

monthly.plot(ax=ax[0, 0], marker="o", color="#337ab7")
ax[0, 0].set_title("Monthly revenue"); ax[0, 0].tick_params(axis="x", rotation=60)

(sales.groupby("category")["line_revenue"].sum().sort_values()
 .plot.barh(ax=ax[0, 1], color="#5cb85c"))
ax[0, 1].set_title("Revenue by category")

seg["Customers"].sort_values().plot.barh(ax=ax[1, 0], color="#f0ad4e")
ax[1, 0].set_title("Customers per RFM segment")

im = ax[1, 1].imshow(retention.iloc[:, :9].values, cmap="Blues", aspect="auto")
ax[1, 1].set_title("Cohort retention (months since first purchase)")
ax[1, 1].set_xlabel("Months since first order"); ax[1, 1].set_ylabel("Cohort")
ax[1, 1].set_yticks(range(0, len(retention), 3))
ax[1, 1].set_yticklabels(retention.index[::3])
plt.colorbar(im, ax=ax[1, 1])
plt.tight_layout()
plt.savefig(f"{OUT_DIR}/ecommerce_dashboard.png", dpi=130)
plt.close()

# --------------------------------------------------------------------------
# 8. EXPORTS: Power BI + Excel
# --------------------------------------------------------------------------
sales.drop(columns=["order_month", "cohort", "period"]).to_csv(
    f"{OUT_DIR}/powerbi_sales_fact.csv", index=False)
rfm.to_csv(f"{OUT_DIR}/powerbi_rfm_customers.csv", index=False)
retention.to_csv(f"{OUT_DIR}/powerbi_cohort_retention.csv")

with pd.ExcelWriter(f"{OUT_DIR}/ecommerce_report.xlsx") as xl:
    monthly.rename("revenue").to_frame().to_excel(xl, sheet_name="Monthly Revenue")
    (sales.groupby("category")["line_revenue"].sum().sort_values(ascending=False)
     .rename("revenue").to_frame().to_excel(xl, sheet_name="Category"))
    seg.to_excel(xl, sheet_name="RFM Segments")
    retention.to_excel(xl, sheet_name="Cohort Retention")

print(f"\nDone. Files saved in ./{OUT_DIR}/")
