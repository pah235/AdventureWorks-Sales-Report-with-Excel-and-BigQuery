# AdventureWorks-Sales-Report-with-Excel-BigQuery
This exercise demonstrates a Business Intelligence solution using Google BigQuery and Excel/Google Sheets to analyze AdventureWorks2022 sales data. It delivers a dynamic, refreshable dashboard that tracks monthly revenue, customer segmentation, product performance, and promotion impact.

# 📊 Sales Report with Excel & BigQuery

<img width="683" alt="image" src="https://github.com/user-attachments/assets/bb7ae424-f2d5-4f24-ace5-0d07ecadeb30" />

---

## 🅰️ A. Project Overview

### 1. Tools & Skills Used
- **Google BigQuery & SQL Server**  
  - Simple & advanced querying  
  - DDL (“CREATE”, “ALTER”) and data import/export  
  - Complex `JOIN` across 7+ tables  
- **Excel & Google Sheets**  
  - Direct BigQuery connector for live data  
  - Advanced lookup functions: `VLOOKUP`, `XLOOKUP`, `INDEX` & `MATCH`, `SUMIFS`  
  - PivotTables & PivotCharts for multi-source reporting  
  - Conditional Formatting & Slicers for interactivity  
  - Power Query & Power Pivot: large-scale ETL, Data Model, basic DAX  
  - In-depth DAX measures: `RANKX`, `SAMEPERIODLASTYEAR`, `DATEADD`, `TOTALYTD`

---

## 🅱️ B. Business Background

**AdventureWorks** is a U.S.-based bicycle manufacturer and retailer operating through two main channels:  
- **IN** – Direct online sales to individual customers.  
- **CS** – Wholesale and distributor relationships.  

They maintain:  
1. An **OLTP database** (`AdventureWorks`) on SQL Server, covering Purchasing raw materials, Production, Sales Order Processing, Product Management, Contact Management, and Human Resources.  
2. A **Data Warehouse** (`AdventureWorksDW`) for historical analytics and reporting.  

**Dataset Description:**  
- Tab **Data** with fields:  
  - `OrderDate`, `ShipDate`, `DueDate`, `OrderQty`  
  - `UnitPrice`, `LineTotal`  
  - `CustomerID`  
  - `SpecialOffer`, `Territory`  
  - `Group`, `SalesPerson`, `Product`, `ProductSubcategory`, `ProductCategory`
  - `ProductLine`, `Product`, `Customer`, `CustomerType`   

Analysts across departments use this data to clean, model, and build reports on sales performance, customer behavior, territory trends, and promotion effectiveness.

---

## 📊 C. Project Goals & Stakeholder Needs

1. **Role & Scenario**  
   You are a BI Analyst asked to pull sales data from BigQuery/SQL and present it in Excel or Google Sheets (see the **Data** tab).  
2. **Manager Requirements**  
   - A refreshable overview dashboard that updates when new data is added  
   - Visuals & tables to track daily/monthly revenue and compare to the prior month  
   - Breakdowns by customer type, territory, product category, and special offers  
3. **Data Scope**  
   All 2013 orders for **IN** (individual) and **CS** (sales-contact) customers, including product hierarchy, territory, special offers, and line-item revenue (`LineTotal`).
   Dataset: https://learn.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver16&tabs=ssms
5. **Deliverables**  
   - Month-over-prior-month revenue & growth rate  
   - Monthly trend analysis for the current year  
   - Customer-type share and growth rates  
   - Revenue by category/subcategory/product vs. prior month (with filters)  
   - Promotion impact on sales  
   - Any additional insights uncovered  

---

## 🛠️ D. Project Execution

### 1. Data Collection & Understanding
- Downloaded AdventureWorks sample from Microsoft Learn  
- Loaded tables into Google BigQuery  
- Explored schema and data relationships  

### 2. Data Preparation  
Queried & joined 10+ tables in BigQuery, then connected to Google Sheets/Excel:

```sql
SELECT 
    h.OrderDate, 
    h.ShipDate, 
    h.DueDate, 
    d.OrderQty, 
    d.UnitPrice, 
    d.LineTotal, 
    h.CustomerID, 
    so.Category AS SpecialOffer, 
    t.Name AS Territory, 
    t.`Group`, 
    CONCAT(pp.FirstName,' ',
           CASE WHEN pp.MiddleName = 'NULL' THEN '' ELSE pp.MiddleName END,' ',
           pp.LastName) AS SalesPerson,
    p.Name AS Product, 
    ps.Name AS ProductSubcategory, 
    pc.Name AS ProductCategory, 
    p.ProductLine,
    CONCAT(pp2.FirstName,' ',
           CASE WHEN pp2.MiddleName = 'NULL' THEN '' ELSE pp2.MiddleName END,' ',
           pp2.LastName) AS Customer, 
    pp2.PersonType AS CustomerType
FROM `adventure-works-analysis.AdventureWorks2022.Sales_SalesOrderDetail` AS d
JOIN `adventure-works-analysis.AdventureWorks2022.Sales_SalesOrderHeader_fixed` AS h 
    ON d.SalesOrderID = h.SalesOrderID
JOIN `adventure-works-analysis.AdventureWorks2022.Sales_SpecialOffer` AS so 
    ON d.SpecialOfferID = so.SpecialOfferID 
JOIN `adventure-works-analysis.AdventureWorks2022.Sales_SalesTerritory` AS t 
    ON h.TerritoryID = t.TerritoryID
LEFT JOIN `adventure-works-analysis.AdventureWorks2022.Sales_SalesPerson` AS sp 
    ON h.SalesPersonID = sp.BusinessEntityID
LEFT JOIN `adventure-works-analysis.AdventureWorks2022.Person_Person` AS pp 
    ON sp.BusinessEntityID = pp.BusinessEntityID
JOIN `adventure-works-analysis.AdventureWorks2022.Production_Product_fixed` AS p 
    ON d.ProductID = p.ProductID
JOIN `adventure-works-analysis.AdventureWorks2022.Production_ProductSubcategory` AS ps 
    ON p.ProductSubcategoryID = ps.ProductSubcategoryID
JOIN `adventure-works-analysis.AdventureWorks2022.Production_ProductCategory` AS pc 
    ON ps.ProductCategoryID = pc.ProductCategoryID
JOIN `adventure-works-analysis.AdventureWorks2022.Sales_Customer_fixed2` AS c 
    ON h.CustomerID = c.CustomerID
JOIN `adventure-works-analysis.AdventureWorks2022.Person_Person` AS pp2 
    ON c.PersonID = pp2.BusinessEntityID
WHERE pp2.PersonType IN ('IN', 'SC') 
  AND EXTRACT(YEAR FROM h.OrderDate) = 2013;
```
Connected the result set to Google Sheets / Excel for live or scheduled refresh.

Used VLOOKUP, INDEX & MATCH, SUMIFS to enrich and lookup supporting data.

Applied Power Query to shape, clean, and load into an Excel Data Model.

### 3. Data Visualization & Dashboard Design
<img width="583" alt="image" src="https://github.com/user-attachments/assets/973509c8-ae76-4f90-930e-9ebbf9f0bc25" />

Created multiple sheets/pages for:

Revenue Overview (monthly trend, variance)

Product Performance (category/subcategory/product)

Territory & Promotion (map, bar charts)

Utilized PivotTables, PivotCharts, column/bar/donut/time‐series charts

Added Slicers & Filters for Order Date, Customer Type, Territory, Special Offer
