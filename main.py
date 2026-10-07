from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


SHOW_DETAILS = False

project_folder = Path(__file__).resolve().parent
file_path = project_folder / "data" / "Online Retail.xlsx"
cache_path = file_path.with_suffix(".pkl")
output_folder = project_folder / "output"

output_folder.mkdir(exist_ok=True)

if (
    cache_path.exists()
    and cache_path.stat().st_mtime >= file_path.stat().st_mtime
):
    print("Загружаю сохранённую копию...")
    df = pd.read_pickle(cache_path)
else:
    print("Читаю Excel и сохраняю копию...")
    df = pd.read_excel(file_path, engine="openpyxl")
    df.to_pickle(cache_path)

print("Данные загружены!")

if SHOW_DETAILS:
    print("\nКоличество строк:", df.shape[0])
    print("Количество столбцов:", df.shape[1])

    print("\nНазвания столбцов:")
    print(df.columns.tolist())

    print("\nПервые пять строк:")
    print(df.head().to_string(index=False))

    print("\nТипы данных:")
    df.info(show_counts=True)

    print("\nПропуски по столбцам:")
    print(df.isna().sum())

    print("\nПолные повторы строк:", df.duplicated().sum())

    print("\nПериод данных:")
    print("Начало:", df["InvoiceDate"].min())
    print("Конец:", df["InvoiceDate"].max())

    negative_quantity = df["Quantity"] < 0
    print("\nСтрок с отрицательным количеством:", negative_quantity.sum())

    zero_price_rows = df[df["UnitPrice"] == 0]
    print("\nСтроки с нулевой ценой:")
    print(zero_price_rows.head().to_string(index=False))
    print("Количество строк:", zero_price_rows.shape[0])

    negative_price_rows = df[df["UnitPrice"] < 0]
    print("\nСтроки с отрицательной ценой:")
    print(negative_price_rows.head().to_string(index=False))
    print("Количество строк:", negative_price_rows.shape[0])


is_cancelled = df["InvoiceNo"].astype(str).str.startswith("C")

sales = df[
    (df["Quantity"] > 0)
    & (df["UnitPrice"] > 0)
    & (~is_cancelled)
].copy()

cancelled = df[
    is_cancelled
    & (df["Quantity"] < 0)
    & (df["UnitPrice"] > 0)
].copy()

sales["LineAmount"] = sales["Quantity"] * sales["UnitPrice"]
cancelled["LineAmount"] = cancelled["Quantity"] * cancelled["UnitPrice"]

total_sales = sales["LineAmount"].sum()
orders_count = sales["InvoiceNo"].nunique()
average_order = total_sales / orders_count

cancelled_amount = -cancelled["LineAmount"].sum()
net_sales = total_sales - cancelled_amount

transactions = pd.concat([sales, cancelled])

sales["Month"] = sales["InvoiceDate"].dt.to_period("M")
monthly_sales = sales.groupby("Month")["LineAmount"].sum()

country_sales = (
    sales.groupby("Country")["LineAmount"]
    .sum()
    .sort_values(ascending=False)
)

uk_share = country_sales.get("United Kingdom", 0) / total_sales

product_summary = (
    sales.groupby("StockCode")
    .agg(
        Product=("Description", "first"),
        Sales=("LineAmount", "sum"),
        Units=("Quantity", "sum"),
        Orders=("InvoiceNo", "nunique"),
    )
    .sort_values("Sales", ascending=False)
)

# Доставка и ручные записи не входят в товарный рейтинг.
excluded_codes = ["DOT", "POST", "M"]

product_transactions = transactions[
    ~transactions["StockCode"].astype(str).isin(excluded_codes)
]

net_products = (
    product_transactions.groupby("StockCode")
    .agg(
        Product=("Description", "first"),
        NetSales=("LineAmount", "sum"),
        NetUnits=("Quantity", "sum"),
    )
    .sort_values("NetSales", ascending=False)
)


# Повторы сохраняем в основном расчёте и отдельно оцениваем их влияние.
sales_no_duplicates = sales.drop_duplicates()
sales_without_duplicates = sales_no_duplicates["LineAmount"].sum()
duplicate_sales_difference = total_sales - sales_without_duplicates

# Проверяем крупные продажи вместе с отрицательными операциями.
codes_to_check = ["23843", "23166"]

unusual_rows = df[
    df["StockCode"].astype(str).isin(codes_to_check)
].copy()

unusual_rows["LineAmount"] = (
    unusual_rows["Quantity"] * unusual_rows["UnitPrice"]
)

unusual_rows = unusual_rows.sort_values(
    "Quantity",
    key=abs,
    ascending=False,
)


print("\nПоказатели до вычета отмен")
print(f"Сумма продаж: {total_sales:,.2f} GBP")
print("Количество заказов:", orders_count)
print(f"Средний чек до вычета отмен: {average_order:,.2f} GBP")
print(f"Доля Великобритании до вычета отмен: {uk_share:.2%}")

print("\nПродажи с учётом отмен")
print(f"Сумма отмен: {cancelled_amount:,.2f} GBP")
print(f"После вычета отмен: {net_sales:,.2f} GBP")

print("\nВлияние повторов продаж")
print("Повторных строк:", sales.shape[0] - sales_no_duplicates.shape[0])
print(f"Сумма без повторов: {sales_without_duplicates:,.2f} GBP")
print(f"Разница: {duplicate_sales_difference:,.2f} GBP")
print(f"Разница в процентах: {duplicate_sales_difference / total_sales:.2%}")

print("\nТоп-10 товаров после вычета отмен")
print(net_products.head(10).round(2).to_string())

if SHOW_DETAILS:
    print("\nСтрок в таблице продаж:", sales.shape[0])
    print(sales[["Quantity", "UnitPrice", "LineAmount"]].head())

    print("\nПродажи по месяцам до вычета отмен:")
    print(monthly_sales.round(2))

    print("\nТоп-10 стран до вычета отмен:")
    print(country_sales.head(10).round(2))

    print("\nТоп-15 позиций до вычета отмен:")
    print(product_summary.head(15).round(2).to_string())

    print("\nКрупные операции по выбранным товарам:")
    print(
        unusual_rows[
            [
                "InvoiceNo",
                "StockCode",
                "Quantity",
                "UnitPrice",
                "LineAmount",
                "InvoiceDate",
                "CustomerID",
            ]
        ].head(12).to_string(index=False)
    )


# Декабрь 2011 неполный: данные заканчиваются 9 декабря.
monthly_sales_full_months = monthly_sales[
    monthly_sales.index < "2011-12"
]

fig_monthly, ax_monthly = plt.subplots(figsize=(10, 5))

ax_monthly.plot(
    monthly_sales_full_months.index.astype(str),
    monthly_sales_full_months.values / 1000,
    marker="o",
)

ax_monthly.set_title("Продажи по месяцам, до вычета возвратов и отмен")
ax_monthly.set_xlabel("Месяц")
ax_monthly.set_ylabel("Сумма продаж, тыс. GBP")
ax_monthly.tick_params(axis="x", rotation=45)
ax_monthly.grid(axis="y", alpha=0.3)

fig_monthly.tight_layout()
fig_monthly.savefig(output_folder / "monthly_sales.png", dpi=150)


# Долю Великобритании показываем отдельно, здесь сравниваем другие страны.
international_sales = country_sales.drop(
    "United Kingdom", errors="ignore"
)
top_countries = international_sales.head(10).sort_values()

fig_countries, ax_countries = plt.subplots(figsize=(10, 6))

ax_countries.barh(
    top_countries.index,
    top_countries.values / 1000,
)

ax_countries.set_title("Топ-10 стран вне Великобритании по сумме продаж")
ax_countries.set_xlabel(
    "Продажи до вычета возвратов и отмен, тыс. GBP"
)
ax_countries.set_ylabel("Страна")
ax_countries.grid(axis="x", alpha=0.3)
ax_countries.set_axisbelow(True)

fig_countries.tight_layout()
fig_countries.savefig(output_folder / "top_countries.png", dpi=150)


top_products = net_products.head(10).sort_values("NetSales")

fig_products, ax_products = plt.subplots(figsize=(12, 7))

ax_products.barh(
    top_products["Product"],
    top_products["NetSales"] / 1000,
)

ax_products.set_title("Топ-10 товаров после вычета отмен")
ax_products.set_xlabel("Сумма после вычета отмен, тыс. GBP")
ax_products.grid(axis="x", alpha=0.3)
ax_products.set_axisbelow(True)

fig_products.tight_layout()
fig_products.savefig(output_folder / "top_products.png", dpi=150)


summary = pd.DataFrame({
    "Metric": [
        "Продажи до вычета отмен, GBP",
        "Сумма отмен, GBP",
        "Продажи после вычета отмен, GBP",
        "Количество заказов до учёта отмен",
        "Средний чек до вычета отмен, GBP",
        "Доля Великобритании до вычета отмен, %",
        "Разница от удаления повторов продаж, GBP",
    ],
    "Value": [
        total_sales,
        cancelled_amount,
        net_sales,
        orders_count,
        average_order,
        uk_share * 100,
        duplicate_sales_difference,
    ],
})

summary.to_csv(
    output_folder / "summary.csv",
    index=False,
    encoding="utf-8-sig",
)

monthly_sales.to_csv(
    output_folder / "monthly_sales.csv",
    header=["GrossSales_GBP"],
    encoding="utf-8-sig",
)

country_sales.to_csv(
    output_folder / "country_sales.csv",
    header=["GrossSales_GBP"],
    encoding="utf-8-sig",
)

net_products.to_csv(
    output_folder / "products_after_cancellations.csv",
    encoding="utf-8-sig",
)

print("\nРезультаты сохранены в:", output_folder)

plt.show()