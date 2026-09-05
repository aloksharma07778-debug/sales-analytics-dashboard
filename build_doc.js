const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat } = require("docx");
const fs = require("fs");

const NAVY = "0F2438";
const TEAL = "1C8C6E";
const GREY = "5C6B78";
const LINE = "E3E8EB";

function h1(text){
  return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing:{before:280, after:120} });
}
function h2(text){
  return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing:{before:220, after:100} });
}
function body(text, opts={}){
  return new Paragraph({ children:[ new TextRun({text, ...opts}) ], spacing:{after:100} });
}
function bullet(text){
  return new Paragraph({ text, bullet:{level:0}, spacing:{after:60} });
}

function cell(text, {header=false, width}={}){
  return new TableCell({
    width:{size:width, type:WidthType.DXA},
    shading: header? {type:ShadingType.CLEAR, fill:NAVY} : undefined,
    children:[ new Paragraph({ children:[ new TextRun({text, bold:header, color: header?"FFFFFF":"12202E", size:20}) ] }) ]
  });
}

function simpleTable(headers, rows, widths){
  return new Table({
    width:{size:9020, type:WidthType.DXA},
    columnWidths: widths,
    rows:[
      new TableRow({ children: headers.map((h,i)=>cell(h,{header:true,width:widths[i]})) }),
      ...rows.map(r=> new TableRow({ children: r.map((c,i)=>cell(String(c),{width:widths[i]})) }))
    ]
  });
}

const doc = new Document({
  sections: [{
    properties:{ page:{ size:{ width:12240, height:15840 } } },
    children: [
      new Paragraph({ children:[ new TextRun({text:"Sales, Customer & Inventory Analytics", bold:true, size:44, color:NAVY}) ], spacing:{after:80} }),
      new Paragraph({ children:[ new TextRun({text:"End-to-End Business Intelligence Project — ETL, Data Modeling, DAX & Dashboarding", size:24, color:GREY}) ], spacing:{after:300} }),

      h1("1. Project Summary"),
      body("This project builds a complete sales analytics solution from six raw source files (transactions, customers, products, inventory and monthly targets) through data cleaning, relational modeling, KPI/DAX measures, and a 4-page interactive dashboard. It follows a standard BI workflow — the same one used in production analytics teams — and is designed to be reproducible in Power BI as well as in the Python/HTML build included in this package."),

      h1("2. Objective"),
      bullet("Consolidate five disconnected source files into one clean, analysis-ready model."),
      bullet("Resolve real-world data quality issues: mixed date formats, duplicate rows, missing categories/regions, and returns (negative quantities)."),
      bullet("Build a star-schema data model with a dedicated date table for time intelligence."),
      bullet("Define reusable KPI measures (revenue, profit, margin, YoY growth, rank, contribution %)."),
      bullet("Deliver a decision-ready dashboard: Executive Summary, Sales Analysis, Customer Insights, and Inventory."),

      h1("3. Tools & Tech Stack"),
      simpleTable(
        ["Layer","Tool"],
        [
          ["ETL / Data cleaning","Power Query (Power BI) — reproduced in Python (pandas) in this package"],
          ["Data modeling","Power BI star schema (Sales as fact; Customers, Products, DateTable as dimensions)"],
          ["Metrics layer","DAX (see DAX_Measures.txt)"],
          ["Dashboard","Power BI report pages — reproduced as an interactive HTML/Chart.js dashboard for portability"],
        ],
        [3200,5820]
      ),

      h1("4. Source Data"),
      simpleTable(
        ["File","Rows","Grain"],
        [
          ["Sales_Transactions.csv","8 (7 after cleaning)","One row per order line"],
          ["Customers.xlsx","6","One row per customer"],
          ["Products.xlsx","6","One row per product"],
          ["Inventory.csv","6","One row per product stock snapshot"],
          ["Targets.xlsx","4","One row per month/category target"],
        ],
        [3600,2500,2920]
      ),

      h1("5. ETL / Data Cleaning Steps"),
      h2("5.1 Fix dates"),
      body("OrderDate in Sales mixed three formats (YYYY-MM-DD, YYYY/MM/DD, DD-MM-YYYY). Standardized to a single date type using locale-aware parsing."),
      h2("5.2 Remove duplicates"),
      body("Order 1007 was an exact duplicate of order 1006 (same customer, product, date, amounts) and was removed."),
      h2("5.3 Handle nulls"),
      bullet("Customers.Region: blank → \"Unknown\" (customer C003)."),
      bullet("Products.Category: blank → \"Others\" (product P106)."),
      h2("5.4 Handle returns"),
      body("Order 1008 has Quantity = -1, treated as a return and flagged (IsReturn) rather than dropped, so it correctly reduces revenue and profit instead of distorting counts."),
      h2("5.5 Calculated columns"),
      bullet("SalesAmount = Quantity × UnitPrice × (1 − Discount)"),
      bullet("Profit = (UnitPrice − Cost) × Quantity"),
      h2("5.6 Merge tables"),
      bullet("Sales + Products → Category, SubCategory"),
      bullet("Sales + Customers → Region"),

      h1("6. Data Model (Star Schema)"),
      body("Fact table: Sales. Dimension tables: Customers, Products, DateTable. Relationships: Sales→Customers (CustomerID), Sales→Products (ProductID), Sales→DateTable (OrderDate) — all single-direction, no many-to-many."),

      h1("7. DAX Measure Library"),
      body("Full formulas are in DAX_Measures.txt. Core measures:"),
      simpleTable(
        ["Measure","Purpose"],
        [
          ["Total Revenue","SUMX over Quantity × UnitPrice × (1-Discount)"],
          ["Total Profit","SUM of line Profit"],
          ["Profit Margin %","Total Profit ÷ Total Revenue"],
          ["Total Orders","DISTINCTCOUNT of OrderID"],
          ["Avg Order Value","Total Revenue ÷ Total Orders"],
          ["YTD Sales / YoY Growth %","Time-intelligence via DateTable"],
          ["Rank Product / Contribution %","RANKX and share-of-total for products"],
        ],
        [3200,5820]
      ),

      h1("8. Dashboard Pages"),
      bullet("Executive Dashboard — KPI cards (Revenue, Profit, Margin), sales trend line, category performance bar, revenue vs. target."),
      bullet("Sales Analysis — Category → Product matrix, top products by revenue."),
      bullet("Customer Insights — Revenue by region, customer-level order & revenue table."),
      bullet("Inventory — Stock vs. sold quantity, low-stock flags."),
      body("Slicers for Region and Category apply across the Executive, Sales and Customer pages. See dashboard/Dashboard.html for the working, filterable build."),

      h1("9. Key Insights"),
      bullet("Total revenue ₹4,740 across 7 orders, with 42.8% profit margin (₹2,030 profit)."),
      bullet("Electronics drives 80% of gross category revenue (₹3,790) vs. Furniture (₹1,310); the one return (Phone Y, \"Others\" category) reduced net revenue by ₹360."),
      bullet("Laptop A is the top product by revenue (₹1,400), followed by Phone X (₹1,200) and Laptop B (₹1,190)."),
      bullet("North region leads at ₹1,260, but \"Unknown\" region (missing customer data) already accounts for ₹950 — a data-quality gap worth closing at source."),
      bullet("All products are under-sold relative to stock (e.g. Laptop A: 3 sold vs. 50 in stock) — no stock-out risk in this sample, but Phone Y has zero recorded sales, which combined with its return is worth a closer look."),
      bullet("Electronics is tracking well below its Jan+Feb target (₹110,000 target vs. ₹3,790 actual in this sample dataset), reflecting the small sample size rather than a real shortfall."),

      h1("10. Package Contents"),
      simpleTable(
        ["Path","Description"],
        [
          ["data/raw/","Original uploaded files, untouched"],
          ["data/cleaned/Cleaned_Sales_Data.xlsx","Cleaned & merged tables, ready to load into Power BI"],
          ["scripts/etl.py","Reproducible Python ETL matching the Power Query steps"],
          ["dax/DAX_Measures.txt","All DAX formulas, ready to paste into Power BI"],
          ["dashboard/Dashboard.html","Working interactive dashboard (open directly in any browser)"],
          ["Project_Documentation.docx","This document"],
        ],
        [3600,5420]
      ),

      h1("11. How to Present This for Placement"),
      body("Suggested resume bullet:"),
      new Paragraph({
        children:[ new TextRun({text:"\"Built an end-to-end sales analytics solution (ETL → star-schema data model → DAX KPIs → 4-page interactive dashboard) from 5 raw sources, resolving data quality issues (mixed date formats, duplicates, nulls, returns) and surfacing category, regional and inventory insights.\"", italics:true, color:GREY}) ],
        spacing:{after:120}
      }),
      body("To extend for an interview: open Dashboard.html, walk through each of the 4 pages, and be ready to explain the star schema and any 2-3 DAX measures in depth — these are the most commonly asked follow-up questions for a BI/analyst placement round."),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/home/claude/project/Project_Documentation.docx", buf);
  console.log("written");
});
