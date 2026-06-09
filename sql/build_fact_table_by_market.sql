-- ============================================================================
--  build_fact_table_by_market.sql  |  FP&A Variance -- pipeline (segment = MARKET)
--  Engine: DuckDB.  Run from project root:  duckdb < sql/build_fact_table_by_market.sql
--
--  Turns raw Online Retail II transactions into one tidy monthly fact table,
--  segmented by market (UK + 4 material countries + "Other"). All finance logic
--  (forecast, variance bridge) lives in the Excel model by design.
-- ============================================================================

-- 1) RAW: read the real Online Retail II CSV (download separately; see README).
CREATE OR REPLACE TABLE raw_sales AS
SELECT * FROM read_csv_auto('data/raw/online_retail_II.csv', header = true);

-- 2) CLEAN + SEGMENT: drop junk, bucket each row into a market segment.
--    - Cancellations ('C' invoices), non-positive qty, and bad prices are removed.
--    - Missing Customer ID is KEPT -- revenue does not need it.
CREATE OR REPLACE TABLE clean_sales AS
SELECT
    r.Invoice,
    r.StockCode,
    CASE r.Country
        WHEN 'United Kingdom' THEN 'United Kingdom'
        WHEN 'EIRE'           THEN 'EIRE'
        WHEN 'Netherlands'    THEN 'Netherlands'
        WHEN 'Germany'        THEN 'Germany'
        WHEN 'France'         THEN 'France'
        ELSE 'Other'
    END                                     AS category,   -- "category" = market segment
    r.Quantity,
    CAST(r.InvoiceDate AS TIMESTAMP)        AS invoice_ts,
    r.Price,
    r.Quantity * r.Price                    AS line_revenue
FROM raw_sales r
WHERE r.Invoice NOT LIKE 'C%'
  AND r.Quantity > 0
  AND r.Price   > 0;

-- 3) FACT: aggregate to Year x Month x Segment.
CREATE OR REPLACE TABLE monthly_fact AS
SELECT
    EXTRACT(year  FROM invoice_ts)          AS year,
    EXTRACT(month FROM invoice_ts)          AS month,
    category,
    SUM(Quantity)                           AS units,
    ROUND(SUM(line_revenue), 2)             AS revenue,
    ROUND(SUM(line_revenue) / SUM(Quantity), 4) AS avg_price
FROM clean_sales
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3;

-- 4) EXPORT the clean fact table for Excel.
COPY monthly_fact TO 'data/processed/monthly_fact.csv' (HEADER, DELIMITER ',');

-- 5) SANITY CHECKS (Stage 2/4)
SELECT 'rows kept'   AS metric, COUNT(*) AS value FROM clean_sales
UNION ALL SELECT 'fact rows', (SELECT COUNT(*) FROM monthly_fact)
UNION ALL SELECT 'segments',  (SELECT COUNT(DISTINCT category) FROM monthly_fact);

-- Reconciliation: these two must match (to the penny, modulo rounding).
SELECT ROUND(SUM(line_revenue), 2) AS clean_total FROM clean_sales;
SELECT ROUND(SUM(revenue), 2)      AS fact_total  FROM monthly_fact;
