"""
B2B BookStore - Sales Analysis
Reads the cleaned data (book_sales_clean.xlsx, produced by book_sales_analysis.ipynb)
and prints each analysis with a short brief.

Each "# %%" line starts a cell. In VS Code, click "Run Cell" above a cell (or press Shift+Enter).
Run the "Setup" cell first: it loads the data and defines the helpers every other cell uses.
You can also run the whole file:  venv\\Scripts\\python analysis.py
"""

# %% Setup - run this cell first
import calendar

import pandas as pd

pd.set_option("display.width", 140)
pd.set_option("display.max_columns", None)
pd.set_option("display.float_format", "{:,.2f}".format)

df = pd.read_excel("book_sales_clean.xlsx")
df["GrossSale"] = df["Quantity"] * df["UnitPrice"]
df["DiscountAmount"] = df["GrossSale"] * df["DiscountRate"]

sales = df[~df["IsReturn"]]
returns = df[df["IsReturn"]]


def section(title, brief):
    print()
    print("=" * 80)
    print(title)
    print("-" * 80)
    print(brief.strip())
    print("-" * 80)


def share(series):
    """Return a series as a % share of its total."""
    return (series / series.sum() * 100).round(1)


# %% 1. Overall KPIs
kpis = pd.Series({
    "Net sales (after discounts & returns)": df["NetSale"].sum(),
    "Gross sales (before discount, sales only)": sales["GrossSale"].sum(),
    "Discounts given (sales only)": sales["DiscountAmount"].sum(),
    "Returns value (net)": returns["NetSale"].sum(),
    "Units sold": sales["Quantity"].sum(),
    "Units returned": -returns["Quantity"].sum(),
    "Sales invoices": sales["InvoiceNo"].nunique(),
    "Active customers": df["CustomerID"].nunique(),
    "Books (products)": df["ProductID"].nunique(),
    "Average order value (net)": sales.groupby("InvoiceNo")["NetSale"].sum().mean(),
    "Average discount rate %": sales["DiscountRate"].mean() * 100,
    "Return rate % (of sales value)": -returns["NetSale"].sum() / sales["NetSale"].sum() * 100,
})
section("1. Overall KPIs", """
Over 2024-2025 the business made ~130.2M in net sales from 120 customers and 74 books.
Discounts are large: ~35.7M, about 21% of gross sales, so pricing policy matters more than returns.
Returns are small (0.57% of sales value). The average sales invoice is worth ~22.2K.
""")
print(kpis.to_string())


# %% 2. Sales by year
by_year = df.groupby(df["InvoiceDate"].dt.year).agg(
    NetSales=("NetSale", "sum"),
    Units=("Quantity", "sum"),
    Invoices=("InvoiceNo", "nunique"),
    Customers=("CustomerID", "nunique"),
)
by_year["NetSales growth %"] = (by_year["NetSales"].pct_change() * 100).round(1)
section("2. Sales by year", """
Net sales grew 7.5% from 2024 to 2025 (62.7M -> 67.4M).
The growth came from more orders (+18.6% invoices), while the average order got smaller
(~23.3K -> ~21.2K per sales invoice). 3 customers bought in 2024 but not in 2025.
""")
print(by_year)


# %% 3. Monthly trend and seasonality
monthly = df.pivot_table(index=df["InvoiceDate"].dt.month, columns=df["InvoiceDate"].dt.year,
                         values="NetSale", aggfunc="sum")
monthly.index = [calendar.month_name[m] for m in monthly.index]
monthly["Average"] = monthly.mean(axis=1)
monthly["% of year"] = share(monthly["Average"])
section("3. Monthly trend and seasonality", """
Sales are strongly seasonal: August and September (back to school) bring ~30% of the year's sales,
and Q4 is the second peak (~30%). June-July is the quietest period.
In 2025 the first half grew strongly (e.g. April +61%), but August fell ~21% and October ~23% vs 2024.
""")
print(monthly)


# %% 4. Customer type
by_ctype = df.groupby("CustomerType").agg(
    Customers=("CustomerID", "nunique"),
    NetSales=("NetSale", "sum"),
    Units=("Quantity", "sum"),
    AvgDiscountPct=("DiscountRate", lambda s: s.mean() * 100),
)
by_ctype["Share %"] = share(by_ctype["NetSales"])
by_ctype["NetSales per customer"] = by_ctype["NetSales"] / by_ctype["Customers"]
by_ctype["Avg order value"] = sales.groupby(["CustomerType", "InvoiceNo"])["NetSale"].sum().groupby("CustomerType").mean()
by_ctype = by_ctype.sort_values("NetSales", ascending=False)
section("4. Sales by customer type", """
Schools drive the business: 30 customers (25%) bring 62% of net sales, with the largest orders (~51K each)
and low discounts (~15%). Bookstores are the largest group (50) but only 19% of sales.
Online retailers get the highest discount (~34%); libraries are the smallest segment (4%).
""")
print(by_ctype)


# %% 5. Region and city
by_region = df.groupby("Region").agg(Customers=("CustomerID", "nunique"), NetSales=("NetSale", "sum"))
by_region["Share %"] = share(by_region["NetSales"])
by_region["NetSales per customer"] = by_region["NetSales"] / by_region["Customers"]
by_region = by_region.sort_values("NetSales", ascending=False)

by_city = df.groupby(["City", "Region"]).agg(Customers=("CustomerID", "nunique"), NetSales=("NetSale", "sum"))
by_city["Share %"] = share(by_city["NetSales"])
by_city = by_city.sort_values("NetSales", ascending=False)
section("5. Sales by region and city", """
Marmara is the largest region (40.5%), but Aegean customers spend the most each (~1.78M per customer).
Three cities - Istanbul, Izmir and Ankara - make up ~81% of all sales.
Mediterranean is the smallest region (7.8%) and has the lowest sales per customer.
""")
print(by_region)
print()
print(by_city)


# %% 6. Category
by_cat = df.groupby("Category").agg(
    Books=("ProductID", "nunique"),
    NetSales=("NetSale", "sum"),
    Units=("Quantity", "sum"),
    AvgPrice=("UnitPrice", "mean"),
)
by_cat["Share %"] = share(by_cat["NetSales"])
cat_year = df.pivot_table(index="Category", columns=df["InvoiceDate"].dt.year, values="NetSale", aggfunc="sum")
by_cat["Growth %"] = ((cat_year.iloc[:, -1] / cat_year.iloc[:, 0] - 1) * 100).round(1)
by_cat = by_cat.sort_values("NetSales", ascending=False)
section("6. Sales by category", """
Education is the largest category (32%), driven almost entirely by schools (48% of school spending).
Business (+21.7%) and Fiction (+17.3%) grew fastest from 2024 to 2025; Science was flat.
Each customer type buys a different mix: bookstores and online retailers favour Fiction and Business,
libraries favour Science and History.
""")
print(by_cat)
print()
print("Category mix by customer type (% of each customer type's net sales):")
print((df.pivot_table(index="Category", columns="CustomerType", values="NetSale", aggfunc="sum")
       .pipe(lambda t: t / t.sum() * 100)).round(1))


# %% 7. Publisher
by_pub = df.groupby("Publisher").agg(Books=("ProductID", "nunique"), NetSales=("NetSale", "sum"),
                                     Units=("Quantity", "sum"))
by_pub["Share %"] = share(by_pub["NetSales"])
by_pub = by_pub.sort_values("NetSales", ascending=False)
section("7. Sales by publisher", """
Sales are spread fairly evenly across publishers. The top three (Cinar, Ada, Atlas) hold ~73% combined,
in line with the number of books they publish. No single publisher dominates.
""")
print(by_pub)


# %% 8. Top books and customers
top_books = (df.groupby(["BookTitle", "Category"])
             .agg(NetSales=("NetSale", "sum"), Units=("Quantity", "sum"))
             .sort_values("NetSales", ascending=False))
top_books["Share %"] = share(top_books["NetSales"])

top_customers = (df.groupby(["CustomerName", "CustomerType", "City"])
                 .agg(NetSales=("NetSale", "sum"), Invoices=("InvoiceNo", "nunique"))
                 .sort_values("NetSales", ascending=False))
top_customers["Share %"] = share(top_customers["NetSales"])
top10_cust_share = top_customers["Share %"].head(10).sum()
top20pct_cust_share = top_customers["Share %"].head(round(len(top_customers) * 0.2)).sum()
section("8. Top books and customers", """
The best sellers are all school workbooks (Education and Language Learning); the top title,
Grade 6 Turkish Workbook, is 4.1% of sales. The weakest books are mostly Business titles (~0.5% each).
Customers are concentrated: the top 10 (9 of them schools) bring 35% of sales, and the top 20% bring 59%.
""")
print("Top 10 books:")
print(top_books.head(10))
print()
print("Bottom 5 books:")
print(top_books.tail(5))
print()
print("Top 10 customers:")
print(top_customers.head(10))
print()
print(f"Top 10 customers' share of net sales:   {top10_cust_share:.1f}%")
print(f"Top 20% of customers' share of net sales: {top20pct_cust_share:.1f}%")


# %% 9. Returns
def return_rate(group_col):
    t = pd.DataFrame({
        "Sales": sales.groupby(group_col)["NetSale"].sum(),
        "Returns": -returns.groupby(group_col)["NetSale"].sum(),
    }).fillna(0)
    t["Return rate %"] = (t["Returns"] / t["Sales"] * 100).round(2)
    return t.sort_values("Return rate %", ascending=False)

section("9. Returns", """
Returns are low and widespread: 361 lines across 101 of the 120 customers, ~0.6% of sales value.
Return rates are similar for all customer types (0.55-0.63%).
Education has the highest return rate (0.80%), Science the lowest (0.35%).
""")
print(f"Return lines: {len(returns)}  |  Return invoices: {returns['InvoiceNo'].nunique()}  |  "
      f"Customers who returned: {returns['CustomerID'].nunique()}")
print()
print(return_rate("CustomerType"))
print()
print(return_rate("Category"))


# %% 10. Discounts
sales_disc = sales.assign(DiscountBand=pd.cut(sales["DiscountRate"], [0, 0.10, 0.20, 0.30, 0.40],
                                              labels=["0-10%", "10-20%", "20-30%", "30-40%"],
                                              include_lowest=True))
by_band = sales_disc.groupby("DiscountBand", observed=True).agg(
    Lines=("NetSale", "size"),
    AvgQuantity=("Quantity", "mean"),
    NetSales=("NetSale", "sum"),
    Discounts=("DiscountAmount", "sum"),
)
by_band["Share of net sales %"] = share(by_band["NetSales"])
disc_by_ctype = sales.groupby("CustomerType").agg(
    AvgDiscountPct=("DiscountRate", lambda s: s.mean() * 100),
    Discounts=("DiscountAmount", "sum"),
    GrossSales=("GrossSale", "sum"),
)
disc_by_ctype["Discount % of gross"] = disc_by_ctype["Discounts"] / disc_by_ctype["GrossSales"] * 100
corr = sales["DiscountRate"].corr(sales["Quantity"])
section("10. Discounts", """
Each customer type gets a fairly fixed discount: online retailers ~34%, bookstores ~30%, schools ~15%,
libraries ~10%. The 10-20% band holds 62% of net sales because that is where schools' large orders sit.
Higher discounts do not come with bigger orders (correlation -0.28): discount follows customer type, not volume.
""")
print(by_band)
print()
print(disc_by_ctype.sort_values("AvgDiscountPct", ascending=False))
print()
print(f"Correlation between discount rate and quantity per line: {corr:.3f}")
