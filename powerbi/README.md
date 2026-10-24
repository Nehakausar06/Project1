# Power BI dashboard setup

This folder contains the Power BI handoff for the Python/Pandas sales analysis. First run the Python project to generate `outputs/cleaned_sales.csv`. That file—not the raw Excel worksheet—is the recommended, consistently cleaned data source.

**The supplied workbook identifies these as fictional practice records and specifies no currency.** Do not present example numbers as actual business results or assume dollars/rupees; use the values computed from the supplied dataset and label them as source units.

## Create the report

1. Install and open **Microsoft Power BI Desktop**.
2. Choose **Get data → Text/CSV** and load the project's `outputs/cleaned_sales.csv`. In Power Query, rename the query `Sales` and choose **Close & Apply**.
3. Alternatively, in Power Query choose **New source → Blank query → Advanced Editor**, paste the contents of [`Sales_Data_Source.m`](Sales_Data_Source.m), replace the sample file path with the full path to the generated `outputs/cleaned_sales.csv`, choose **Done**, name the query `Sales`, and **Close & Apply**.
4. In **Model view**, check that `Sales[Order Date]` is a date column. Consult `outputs/data_quality_report.csv` for missing dates, profit, sales, and unusual quantities.
5. Choose **Modeling → New table** and paste [`Calendar.dax`](Calendar.dax).
6. Create a one-to-many, single-direction relationship from `Calendar[Date]` (one) to `Sales[Order Date]` (many). Select the calendar, choose **Table tools → Mark as date table**, and select `Calendar[Date]`. Sort `Calendar[Year-Month]` by `Calendar[Year-Month Sort]`, and `Calendar[Month]` by `Calendar[Month Number]`.
7. Create the measures **one at a time** using the definitions in [`Sales_Measures.dax`](Sales_Measures.dax). Format sales, profit, and average order value as decimal numbers; format margin and growth measures as percentages.
8. Optionally choose **View → Themes → Browse for themes** and apply [`CodSoft_Sales_Theme.json`](CodSoft_Sales_Theme.json).
9. Create the report pages below, add slicers, and test filter interactions. In Power BI Desktop, save the completed report as `Sales_Dashboard.pbix`.

## Suggested report pages

### Sales overview

- KPI cards: `Total Sales`, `Total Profit`, `Total Orders`, `Units Sold`, `Total Customers`, `Average Order Value`, and `Profit Margin`.
- Monthly `Total Sales` line chart using `Calendar[Year-Month]`; add Year, Quarter, and Month as a drill-down hierarchy.
- Bar chart for sales by region or category.
- Top-5-products-by-sales bar chart. Apply a **Top N = 5** visual-level filter on Product by `Total Sales`.
- Date, Region, Category, Product, and Customer slicers.

### Product and category performance

- Compare category `Total Sales` and `Total Profit`.
- Compare the rankings of products by `Total Sales` and `Total Profit`; strong sales do not necessarily mean strong profit.
- Add `Profit Margin` and `Units Sold` to visual tooltips. Preserve negative profit as a reported loss, not a missing value.

### Regional and customer performance

- Compare `Total Sales`, `Total Profit`, `Total Orders`, and `Profit Margin` across regions.
- Customer bar chart with `Total Orders` and `Repeat Customers` tooltips.
- Customer table showing sales, profit, distinct orders, and repeat-customer status.
- Optional product/customer detail page with the corresponding field in **Drill-through filters**, order-level detail, and a back button.

Use readable labels, accessible theme colors, descriptive titles, and informative tooltips. Add page navigation and a consistent layout. Drill-through detail pages and custom tooltips are optional enhancements.

## Validate before submission

- Compare KPI cards to `outputs/kpi_summary.csv` and the KPI cards and cleaned transaction detail on the single `Dashboard` sheet in `outputs/Sales_Analysis.xlsx`.
- Confirm that order counts use distinct order IDs, while separate product lines within an order are retained.
- Confirm that unknown customers are excluded from identified-customer counts.
- Review missing profits, undated sales, negative sales, losses, and zero/negative quantities before interpreting them.
- Check the date relationship and use a comparison period before interpreting month-over-month or year-over-year growth.
- Test each slicer against every KPI and chart, including the date range.
- Test Year → Quarter → Month drill-down and optional product/customer drill-through.
- Check blank visuals, axis scales, titles, readability, clipped labels, and filter interactions.
- Save the report as `Sales_Dashboard.pbix`, reopen it, and export a screenshot of each report page for the internship submission.

The Python and Excel analysis still runs and produces reports without Power BI Desktop. An actual `.pbix` report and report screenshots must be created and saved from Power BI Desktop.

The generated Excel report, `outputs/Sales_Analysis.xlsx`, is also available as a separate workbook with **one worksheet containing the entire dashboard and all analysis tables**.
