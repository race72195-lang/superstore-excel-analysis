"""
Builds an Excel workbook for the Superstore sales analysis.

All numbers in the analysis sheets are LIVE Excel formulas (SUMIFS, COUNTIFS,
SUMPRODUCT) that point to the Data sheet, so they update when the data changes.

Sheets:
    Dashboard  - KPI tiles and charts
    Analysis   - summary tables used by the charts (formulas)
    Data       - the cleaned data with helper columns (formulas)

Usage:  python build_workbook.py
"""

from pathlib import Path

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

BASE = Path(__file__).resolve().parent
CSV = BASE / "data" / "superstore.csv"
OUT = BASE / "superstore_analysis.xlsx"

DISCOUNT_BANDS = [("0%", 0, 0), ("1-10%", 0.0001, 0.10), ("11-20%", 0.1001, 0.20),
                  ("21-30%", 0.2001, 0.30), ("31-50%", 0.3001, 0.50), ("51%+", 0.5001, 1.0)]

BLUE, ORANGE, INK, MUTED = "#2a78d6", "#eb6834", "#0b0b0b", "#52514e"


def load():
    """Read the CSV (Kaggle file is usually latin-1) and keep the columns we need."""
    for enc in ("utf-8-sig", "latin-1"):
        try:
            df = pd.read_csv(CSV, encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    df.columns = [c.strip() for c in df.columns]
    need = ["Order ID", "Order Date", "Ship Date", "Segment", "Region", "Category",
            "Sub-Category", "Sales", "Quantity", "Discount", "Profit"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise SystemExit(f"Columns not found in the CSV: {missing}\nFound: {list(df.columns)}")
    df = df[need].copy()
    df["Order Date"] = pd.to_datetime(df["Order Date"], dayfirst=False, errors="coerce")
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], dayfirst=False, errors="coerce")
    before = len(df)
    df = df.dropna(subset=["Order Date", "Sales", "Profit"]).drop_duplicates()
    print(f"Rows: {before} -> {len(df)} after dropping blanks/duplicates")
    return df.sort_values("Order Date").reset_index(drop=True)


def main():
    if not CSV.exists():
        raise SystemExit(f"File not found: {CSV}\nRun:  python get_data.py")
    df = load()
    n = len(df)
    last = n + 1  # last Excel row of data

    wb = xlsxwriter.Workbook(OUT)
    f_head = wb.add_format({"bold": True, "bg_color": BLUE, "font_color": "white", "border": 0})
    f_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    f_money = wb.add_format({"num_format": "#,##0"})
    f_pct = wb.add_format({"num_format": "0.0%"})
    f_title = wb.add_format({"bold": True, "font_size": 16, "font_color": INK})
    f_sub = wb.add_format({"font_color": MUTED, "italic": True})
    f_h2 = wb.add_format({"bold": True, "font_size": 12, "font_color": INK})
    f_kpi_label = wb.add_format({"font_color": MUTED, "font_size": 10})
    f_kpi_money = wb.add_format({"bold": True, "font_size": 18, "num_format": "#,##0", "font_color": BLUE})
    f_kpi_pct = wb.add_format({"bold": True, "font_size": 18, "num_format": "0.0%", "font_color": BLUE})

    dash = wb.add_worksheet("Dashboard")
    ana = wb.add_worksheet("Analysis")
    data = wb.add_worksheet("Data")

    # ---------------- Data sheet ----------------
    df["First Line"] = (~df["Order ID"].duplicated()).astype(int)  # 1 on the first row of each order
    cols = list(df.columns[:11]) + ["Year", "Year-Month", "Ship Days", "Profit Margin", "First Line"]
    for j, c in enumerate(cols):
        data.write(0, j, c, f_head)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        r = i + 1  # Excel row number
        data.write(i, 0, row[0])
        data.write_datetime(i, 1, row[1].to_pydatetime(), f_date)
        if pd.notna(row[2]):
            data.write_datetime(i, 2, row[2].to_pydatetime(), f_date)
        for j in range(3, 7):
            data.write(i, j, row[j])
        data.write_number(i, 7, row[7])
        data.write_number(i, 8, row[8])
        data.write_number(i, 9, row[9])
        data.write_number(i, 10, row[10])
        data.write_formula(i, 11, f"=YEAR(B{r})", None, row[1].year)
        data.write_formula(i, 12, f'=TEXT(B{r},"yyyy-mm")', None, row[1].strftime("%Y-%m"))
        ship = (row[2] - row[1]).days if pd.notna(row[2]) else ""
        data.write_formula(i, 13, f'=IF(C{r}="","",C{r}-B{r})', None, ship)
        data.write_formula(i, 14, f"=IF(H{r}=0,0,K{r}/H{r})", f_pct, row[10] / row[7] if row[7] else 0)
        data.write_number(i, 15, row[11])
    data.freeze_panes(1, 0)
    data.autofilter(0, 0, n, len(cols) - 1)
    data.set_column(0, 0, 16)
    data.set_column(1, 2, 12)
    data.set_column(3, 6, 14)
    data.set_column(7, 15, 12)

    def rng(col):
        return f"Data!${col}$2:${col}${last}"

    SALES, QTY, DISC, PROFIT = rng("H"), rng("I"), rng("J"), rng("K")
    SEG, REG, CAT, SUB, YEAR, YM = rng("D"), rng("E"), rng("F"), rng("G"), rng("L"), rng("M")

    # ---------------- Analysis sheet ----------------
    ana.write(0, 0, "Summary tables (all cells are formulas pointing to the Data sheet)", f_title)
    ana.set_column(0, 0, 22)
    ana.set_column(1, 5, 14)
    ana.set_column(7, 12, 14)

    def table(top, left, title, labels, key_range, first_header, with_margin=True):
        """labels down the side; Sales, Profit, Margin (and Orders) by SUMIFS."""
        ana.write(top, left, title, f_h2)
        heads = [first_header, "Sales", "Profit", "Margin"]
        for k, h in enumerate(heads):
            ana.write(top + 1, left + k, h, f_head)
        for i, lab in enumerate(labels):
            r = top + 2 + i
            lc = xl_col_to_name(left)
            sc = xl_col_to_name(left + 1)
            pc = xl_col_to_name(left + 2)
            ana.write(r, left, lab)
            ana.write_formula(r, left + 1, f'=SUMIFS({SALES},{key_range},{lc}{r+1})', f_money)
            ana.write_formula(r, left + 2, f'=SUMIFS({PROFIT},{key_range},{lc}{r+1})', f_money)
            ana.write_formula(r, left + 3, f'=IF({sc}{r+1}=0,0,{pc}{r+1}/{sc}{r+1})', f_pct)
        return top + 2, top + 1 + len(labels)  # first, last data row (0-based)

    cats = sorted(df["Category"].unique())
    regs = sorted(df["Region"].unique())
    segs = sorted(df["Segment"].unique())
    years = sorted(df["Order Date"].dt.year.unique())
    subs = sorted(df["Sub-Category"].unique())
    months = sorted(df["Order Date"].dt.strftime("%Y-%m").unique())

    t_cat = table(2, 0, "By category", cats, CAT, "Category")
    t_reg = table(2 + len(cats) + 4, 0, "By region", regs, REG, "Region")
    t_seg = table(2 + len(cats) + len(regs) + 8, 0, "By segment", segs, SEG, "Segment")
    t_year = table(2 + len(cats) + len(regs) + len(segs) + 12, 0, "By year", [int(y) for y in years], YEAR, "Year")
    t_sub = table(2, 7, "By sub-category", subs, SUB, "Sub-Category")
    t_month = table(2 + len(subs) + 4, 7, "By month", months, YM, "Year-Month")

    # Discount bands (SUMIFS with >= and <= on the Discount column)
    top = 2 + len(subs) + len(months) + 8
    ana.write(top, 7, "Discount bands: do discounts destroy profit?", f_h2)
    for k, h in enumerate(["Discount band", "Sales", "Profit", "Margin"]):
        ana.write(top + 1, 7 + k, h, f_head)
    for i, (lab, lo, hi) in enumerate(DISCOUNT_BANDS):
        r = top + 2 + i
        ana.write(r, 7, lab)
        crit = f'{DISC},">={lo}",{DISC},"<={hi}"'
        ana.write_formula(r, 8, f"=SUMIFS({SALES},{crit})", f_money)
        ana.write_formula(r, 9, f"=SUMIFS({PROFIT},{crit})", f_money)
        ana.write_formula(r, 10, f"=IF(I{r+1}=0,0,J{r+1}/I{r+1})", f_pct)
    t_disc = (top + 2, top + 1 + len(DISCOUNT_BANDS))

    # ---------------- Dashboard ----------------
    dash.hide_gridlines(2)
    dash.set_column(0, 0, 2)
    dash.set_column(1, 12, 13)
    dash.write(0, 1, "Superstore Sales Analysis", f_title)
    dash.write(1, 1, f"{n:,} order lines, {months[0]} to {months[-1]}. KPIs and tables are live Excel formulas.", f_sub)

    kpis = [
        ("Total sales", f"=SUM({SALES})", f_kpi_money),
        ("Total profit", f"=SUM({PROFIT})", f_kpi_money),
        ("Profit margin", f"=SUM({PROFIT})/SUM({SALES})", f_kpi_pct),
        ("Orders", f"=SUM({rng('P')})", f_kpi_money),
        ("Average discount", f"=AVERAGE({DISC})", f_kpi_pct),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        col = 1 + i * 2
        dash.write(3, col, label, f_kpi_label)
        dash.write_formula(4, col, formula, fmt)

    def chart(kind, title, cats_ref, vals_ref, color, name, sub=None):
        ch = wb.add_chart({"type": kind, **({"subtype": sub} if sub else {})})
        ch.add_series({"name": name, "categories": cats_ref, "values": vals_ref,
                       "fill": {"color": color}, "line": {"color": color, "width": 2.25},
                       "gap": 60 if kind != "line" else None, "smooth": False})
        ch.set_title({"name": title, "name_font": {"size": 12, "bold": True, "color": INK}})
        ch.set_legend({"none": True})
        ch.set_y_axis({"major_gridlines": {"visible": True, "line": {"color": "#e6e5e1"}},
                       "line": {"none": True}, "num_font": {"color": MUTED}})
        ch.set_x_axis({"num_font": {"color": MUTED}, "line": {"color": "#c9c8c3"}})
        ch.set_size({"width": 520, "height": 300})
        return ch

    def ref(sheet, col, r1, r2):
        c = xl_col_to_name(col)
        return f"={sheet}!${c}${r1+1}:${c}${r2+1}"

    m1, m2 = t_month
    ch1 = chart("line", "Monthly sales", ref("Analysis", 7, m1, m2), ref("Analysis", 8, m1, m2), BLUE, "Sales")
    ch1.set_x_axis({"num_font": {"color": MUTED, "rotation": -45}, "interval_unit": 3})
    ch1.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#e6e5e1"}},
                    "line": {"none": True}, "num_font": {"color": MUTED}})
    dash.insert_chart("B7", ch1)

    s1, s2 = t_sub
    ch2 = chart("bar", "Profit by sub-category", ref("Analysis", 7, s1, s2), ref("Analysis", 9, s1, s2), BLUE, "Profit")
    ch2.set_size({"width": 520, "height": 420})
    ch2.set_y_axis({"reverse": True, "num_font": {"color": MUTED, "size": 9}, "label_position": "low"})
    ch2.set_x_axis({"major_gridlines": {"visible": True, "line": {"color": "#e6e5e1"}},
                    "num_font": {"color": MUTED}})
    dash.insert_chart("H7", ch2)

    d1, d2 = t_disc
    ch3 = chart("column", "Profit margin by discount band", ref("Analysis", 7, d1, d2), ref("Analysis", 10, d1, d2), ORANGE, "Margin")
    ch3.set_x_axis({"label_position": "low", "num_font": {"color": MUTED}, "line": {"color": "#c9c8c3"}})
    ch3.set_y_axis({"num_format": "0%", "major_gridlines": {"visible": True, "line": {"color": "#e6e5e1"}},
                    "line": {"none": True}, "num_font": {"color": MUTED}})
    dash.insert_chart("B23", ch3)

    # ---------------- print key numbers for the README ----------------
    tot_s, tot_p = df["Sales"].sum(), df["Profit"].sum()
    print("\nKEY NUMBERS (paste these when you fill the README)")
    print(f"Sales {tot_s:,.0f} | Profit {tot_p:,.0f} | Margin {tot_p/tot_s:.1%} | Orders {df['Order ID'].nunique():,}")
    by_sub = df.groupby("Sub-Category")["Profit"].sum().sort_values()
    print("Worst sub-categories by profit:", {k: round(v) for k, v in by_sub.head(3).items()})
    print("Best sub-categories by profit:", {k: round(v) for k, v in by_sub.tail(3).items()})
    for lab, lo, hi in DISCOUNT_BANDS:
        sub = df[(df["Discount"] >= lo) & (df["Discount"] <= hi)]
        if len(sub):
            print(f"Discount {lab:>7}: sales {sub['Sales'].sum():>10,.0f}  profit {sub['Profit'].sum():>10,.0f}  margin {sub['Profit'].sum()/sub['Sales'].sum():.1%}")
    by_reg = df.groupby("Region")[["Sales", "Profit"]].sum()
    by_reg["Margin"] = by_reg["Profit"] / by_reg["Sales"]
    print(by_reg.round(3).to_string())

    wb.close()
    print("\nSaved:", OUT)
    print("Open it in Excel. Formulas are calculated when the file opens.")


if __name__ == "__main__":
    main()
