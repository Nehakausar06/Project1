# Sales analysis: findings and recommendations

> Source amounts are reported in the dataset's units. The source does not specify a currency.

## Data and methodology

- Cleaned records analyzed: **1,498**.
- Order counts use distinct Order IDs; valid multiple line items are preserved.
- Identical records after text normalization are removed, but suspicious returns or values are not automatically deleted.
- Rows with missing or invalid dates contribute to available sales totals but are excluded from monthly, quarterly, and yearly analyses.
- Missing Sales, Profit, and Quantity values are kept unknown rather than filled with zero.
- Profit margin is calculated from total known profit divided by total known sales.

## Findings

- **Highest-sales product:** Laptop - 19,873,914.74 (43.4% of net sales).
- **Highest-sales category:** Electronics - 37,289,913.68 (81.4% of net sales).
- **Highest-sales region:** South - 10,524,765.85 (23.0% of net sales).
- **Highest-sales month:** January 2025 — 2,925,063.89.
- **Change in the latest comparable month:** -21.4%. This compares with the immediately preceding calendar month, not the same month last year.
- **Repeat customers:** 285 of 288 identified customers (99.0%) placed orders on more than one distinct Order ID.

## Data-quality issues to review

- **4 records have missing or invalid dates and are excluded from time-based charts.**
- **1 record has negative sales; confirm whether it is a return or credit.**
- **2 records have zero or negative quantities; confirm return conventions.**
- **1 record has missing or invalid profit.**

## Recommended next steps

- Compare repeat-customer revenue and order frequency with first-time customers before choosing retention incentives.
- Review availability, margins, and customer demand for **Laptop**, the highest-sales product; prioritize profitable growth rather than revenue alone.
- Check whether **Electronics**, the highest reported total-profit category, also performs well on sales and margin; protect profitable categories when adjusting the product mix.
- Ask the source owner to confirm currency, negative-quantity conventions, negative sales, and missing profit before finalizing business or financial decisions.
