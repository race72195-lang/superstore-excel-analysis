# Superstore Sales Analysis in Excel

An Excel workbook that answers one business question: **where does the store make and lose money?**
The workbook is built with a Python script, but everything inside it is regular Excel: KPIs and summary tables are live formulas (`SUMIFS`, `SUM`, `AVERAGE`), so they update if the data changes.

## Dataset

[Superstore sales data](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) (Kaggle): order lines of a retail store with date, segment, region, category, sales, quantity, discount and profit.

## What is inside the workbook

| Sheet | Content |
|---|---|
| **Dashboard** | KPI tiles (sales, profit, margin, orders, average discount) and 3 charts: monthly sales, profit by sub-category, profit margin by discount band |
| **Analysis** | Summary tables by category, region, segment, year, sub-category, month and discount band (all formulas) |
| **Data** | Cleaned data with helper columns: year, year-month, shipping days, profit margin |

## Key findings

Scope: 9,993 order lines, 5,009 orders.

- **Overall:** sales 2,296,919, profit 286,409, profit margin **12.5%**.
- **Discounts are the main leak.** Margin is 29.5% with no discount and 11.6% at 11-20%. Above 20% every band loses money: -10.1% (21-30%), -24.8% (31-50%) and -119.2% (51% and more).
- **Loss-making sub-categories:** Tables (-17,725), Bookcases (-3,473) and Supplies (-1,189). The most profitable are Copiers (55,618), Phones (44,516) and Accessories (41,937).
- **Regions:** West has the best margin (14.9%) and Central the worst (7.9%), even though Central sells more than South.

## Recommendations

- Cap discounts at about 20%, especially on Tables and Bookcases: discounts above 20% lose money in every band.
- Review pricing and costs for Tables, the biggest loss maker.
- Look at why Central earns a much lower margin than the other regions (discount policy and product mix are the first things to check).

## How to run

```bash
pip install -r requirements.txt
python get_data.py          # downloads the dataset into data/
python build_workbook.py    # creates superstore_analysis.xlsx
```

Open `superstore_analysis.xlsx` in Excel. The formulas calculate when the file opens.

## Project structure

```
get_data.py          downloads the data from Kaggle
build_workbook.py    cleans the data and builds the Excel workbook
requirements.txt
```

---
Alizhan Islam | [(https://www.linkedin.com/in/alizhan-islam-960790362)] | alizhan1809@gmail.com