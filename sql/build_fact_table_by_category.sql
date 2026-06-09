-- build_fact_table_by_category.sql | FP&A Variance v2 (segment = PRODUCT CATEGORY)
-- Run from project root:  duckdb < sql/build_fact_table_by_category.sql
-- Derives product categories from item descriptions via keyword rules.

CREATE OR REPLACE TABLE raw_sales AS
SELECT * FROM read_csv_auto('data/raw/online_retail_II.csv', header = true);

-- CLEAN + CATEGORIZE
CREATE OR REPLACE TABLE clean_sales AS
WITH base AS (
    SELECT Invoice, StockCode, Description, Quantity, Price,
           CAST(InvoiceDate AS TIMESTAMP) AS invoice_ts,
           Quantity * Price AS line_revenue,
           UPPER(COALESCE(Description, '')) AS d
    FROM raw_sales
    WHERE Invoice NOT LIKE 'C%' AND Quantity > 0 AND Price > 0
      AND Description IS NOT NULL
      -- exclude NON-PRODUCT lines (precise match so 'CARRIAGE CLOCK' etc. survive)
      AND UPPER(Description) NOT IN (
          'POSTAGE','DOTCOM POSTAGE','MANUAL','AMAZON FEE','CARRIAGE',
          'NEXT DAY CARRIAGE','BANK CHARGES')
      AND UPPER(Description) NOT LIKE '%GIFT VOUCHER%'
)
SELECT *,
    CASE
        WHEN d LIKE '%CHRISTMAS%' OR d LIKE '%BUNTING%' OR d LIKE '%PARTY%'
             OR d LIKE '%PAPER CHAIN%' OR d LIKE '%ADVENT%'              THEN 'Party & Seasonal'
        WHEN d LIKE '%BOX%' OR d LIKE '%CRATE%' OR d LIKE '% TIN%'
             OR d LIKE '%STORAGE%' OR d LIKE '%CABINET%' OR d LIKE '%DRAWER%'
             OR d LIKE '%NESTING%'                                       THEN 'Storage'
        WHEN d LIKE '%T-LIGHT%' OR d LIKE '%TLIGHT%' OR d LIKE '% LIGHT%'
             OR d LIKE '%LIGHTS%' OR d LIKE '%LANTERN%' OR d LIKE '%CANDLE%' THEN 'Lighting'
        WHEN d LIKE '%BAG%' OR d LIKE '%SHOPPER%' OR d LIKE '%SATCHEL%'    THEN 'Bags'
        WHEN d LIKE '%MUG%' OR d LIKE '%BOWL%' OR d LIKE '%PLATE%' OR d LIKE '%JAR%'
             OR d LIKE '%CAKE%' OR d LIKE '%BOTTLE%' OR d LIKE '%CUP%'
             OR d LIKE '%TEAPOT%' OR d LIKE '%TEA SET%' OR d LIKE '%CUTLERY%'
             OR d LIKE '%SPOON%' OR d LIKE '%BAKING%' OR d LIKE '%POPCORN%'
             OR d LIKE '%RECIPE%'                                        THEN 'Kitchen & Dining'
        WHEN d LIKE '%CARD%' OR d LIKE '%NOTEBOOK%' OR d LIKE '%PEN%'
             OR d LIKE '%PENCIL%' OR d LIKE '%CRAFT%' OR d LIKE '%PAPER%'
             OR d LIKE '%MEMOBOARD%' OR d LIKE '%CHALK%' OR d LIKE '%RIBBON%' THEN 'Stationery & Craft'
        WHEN d LIKE '%HEART%' OR d LIKE '%ORNAMENT%' OR d LIKE '%FRAME%'
             OR d LIKE '%WICKER%' OR d LIKE '%CLOCK%' OR d LIKE '%CUSHION%'
             OR d LIKE '%HOOK%' OR d LIKE '%DECORATION%' OR d LIKE '%DOORMAT%'
             OR d LIKE '%MIRROR%' OR d LIKE '%SIGN%' OR d LIKE '%BLACK BOARD%' THEN 'Home & Decor'
        ELSE 'Other'
    END AS category
FROM base;

-- FACT: aggregate by Year x Month x Category
CREATE OR REPLACE TABLE monthly_fact AS
SELECT EXTRACT(year FROM invoice_ts) AS year,
       EXTRACT(month FROM invoice_ts) AS month,
       category,
       SUM(Quantity) AS units,
       ROUND(SUM(line_revenue), 2) AS revenue,
       ROUND(SUM(line_revenue)/SUM(Quantity), 4) AS avg_price
FROM clean_sales
GROUP BY 1,2,3
ORDER BY 1,2,3;

-- Coverage check
SELECT category,
       ROUND(SUM(revenue),0) AS revenue,
       ROUND(100.0*SUM(revenue)/SUM(SUM(revenue)) OVER (),1) AS pct
FROM monthly_fact GROUP BY category ORDER BY revenue DESC;

-- EXPORT (separate file — v1 monthly_fact.csv untouched)
COPY monthly_fact TO 'data/processed/monthly_fact_by_category.csv' (HEADER, DELIMITER ',');
