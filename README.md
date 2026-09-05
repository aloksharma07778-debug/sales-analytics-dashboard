# Sales, Customer & Inventory Analytics — BI Project

A complete, placement-ready business intelligence project: raw data → ETL →
star-schema model → DAX measures → interactive dashboard.

## Start here
1. Open **Project_Documentation.docx** — the full write-up (methodology, data
   model, DAX, insights, and a ready-to-use resume bullet).
2. Open **dashboard/Dashboard.html** in any browser — a live, filterable
   4-page dashboard (Executive, Sales Analysis, Customer Insights, Inventory).
   No install needed, works offline.
3. Load **data/cleaned/Cleaned_Sales_Data.xlsx** into Power BI to rebuild the
   original Power BI version using the steps in the documentation.
4. Paste formulas from **dax/DAX_Measures.txt** into Power BI's measure editor.

## Folder structure
```
data/raw/            original uploaded files, untouched
data/cleaned/         cleaned & merged workbook (ETL output)
scripts/etl.py        Python ETL reproducing the Power Query steps
scripts/build_doc.js  generates Project_Documentation.docx
dax/DAX_Measures.txt  all DAX formulas
dashboard/Dashboard.html   interactive offline dashboard
Project_Documentation.docx  full project write-up
```
