# Sales-Management
analysing internet sales trends from 2019 to 2021 
<table>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/6e3f759b-d8f2-4e3c-9dd7-cd34c2c56fb7" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/96b2d6a7-88d0-49b6-9094-a0df69a626c3" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/adca6c4e-4121-4b03-b783-2037cfe28b98" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/68a4a0d1-d61d-46c5-96cc-625239dc34ce" width="400"/></td>
  </tr>
</table>


Business Request & User Stories

The business request for this data analyst project was an executive sales report for sales managers. Based on the request that was made from the business we following user stories were defined to fulfill delivery and ensure that acceptance criteria’s were maintained throughout the project.


# Sales Management Dashboard Requirements

| #  | As a (role)               | I want (request/demand)                                  | So that I (user value)                                 | Acceptance Criteria                                                                 |
|----|---------------------------|----------------------------------------------------------|-------------------------------------------------------|-------------------------------------------------------------------------------------|
| 1  | Sales Manager             | To get a dashboard overview of internet sales            | Can follow better which customers and products sell the best | A Power BI dashboard which updates data once a day                                 |
| 2  | Sales Representative      | A detailed overview of Internet Sales per Customers      | Can follow up my customers that buy the most and who we can sell more to | A Power BI dashboard which allows me to filter data for each customer               |
| 3  | Sales Representative      | A detailed overview of Internet Sales per Products       | Can follow up my products that sell the most           | A Power BI dashboard which allows me to filter data for each product                |
| 4  | Sales Manager             | A dashboard overview of internet sales                   | Follow sales over time against budget                  | A Power BI dashboard with graphs and KPIs comparing against budget                  |

## Data Cleansing & Transformation (SQL)

To create the necessary data model for doing analysis and fulfilling the business needs defined in the user stories, the following tables were extracted using SQL.

One data source (sales budgets) was provided in Excel format and was connected in the data model in a later step of the process.

Below are the SQL statements for cleansing and transforming the necessary data.

### DIM_Calendar:


```
-- Cleansed DIM_Date Table --
SELECT 
  [DateKey], 
  [FullDateAlternateKey] AS Date, 
  --[DayNumberOfWeek], 
  [EnglishDayNameOfWeek] AS Day, 
  --[SpanishDayNameOfWeek], 
  --[FrenchDayNameOfWeek], 
  --[DayNumberOfMonth], 
  --[DayNumberOfYear], 
  --[WeekNumberOfYear],
  [EnglishMonthName] AS Month, 
  Left([EnglishMonthName], 3) AS MonthShort,   -- Useful for front end date navigation and front end graphs.
  --[SpanishMonthName], 
  --[FrenchMonthName], 
  [MonthNumberOfYear] AS MonthNo, 
  [CalendarQuarter] AS Quarter, 
  [CalendarYear] AS Year --[CalendarSemester], 
  --[FiscalQuarter], 
  --[FiscalYear], 
  --[FiscalSemester] 
FROM 
 [AdventureWorksDW2019].[dbo].[DimDate]
WHERE 
  CalendarYear >= 2019

```
### DIM_Customer:



```
-- Cleansed DIM_Customers Table --
SELECT 
  c.customerkey AS CustomerKey, 
  --      ,[GeographyKey]
  --      ,[CustomerAlternateKey]
  --      ,[Title]
  c.firstname AS [First Name], 
  --      ,[MiddleName]
  c.lastname AS [Last Name], 
  c.firstname + ' ' + lastname AS [Full Name], 
  -- Combined First and Last Name
  --      ,[NameStyle]
  --      ,[BirthDate]
  --      ,[MaritalStatus]
  --      ,[Suffix]
  CASE c.gender WHEN 'M' THEN 'Male' WHEN 'F' THEN 'Female' END AS Gender,
  --      ,[EmailAddress]
  --      ,[YearlyIncome]
  --      ,[TotalChildren]
  --      ,[NumberChildrenAtHome]
  --      ,[EnglishEducation]
  --      ,[SpanishEducation]
  --      ,[FrenchEducation]
  --      ,[EnglishOccupation]
  --      ,[SpanishOccupation]
  --      ,[FrenchOccupation]
  --      ,[HouseOwnerFlag]
  --      ,[NumberCarsOwned]
  --      ,[AddressLine1]
  --      ,[AddressLine2]
  --      ,[Phone]
  c.datefirstpurchase AS DateFirstPurchase, 
  --      ,[CommuteDistance]
  g.city AS [Customer City] -- Joined in Customer City from Geography Table
FROM 
  [AdventureWorksDW2019].[dbo].[DimCustomer] as c
  LEFT JOIN dbo.dimgeography AS g ON g.geographykey = c.geographykey 
ORDER BY 
  CustomerKey ASC -- Ordered List by CustomerKey
```
### DIM_Products:
```
-- Cleansed DIM_Products Table --
SELECT 
  p.[ProductKey], 
  p.[ProductAlternateKey] AS ProductItemCode, 
  --      ,[ProductSubcategoryKey], 
  --      ,[WeightUnitMeasureCode]
  --      ,[SizeUnitMeasureCode] 
  p.[EnglishProductName] AS [Product Name], 
  ps.EnglishProductSubcategoryName AS [Sub Category], -- Joined in from Sub Category Table
  pc.EnglishProductCategoryName AS [Product Category], -- Joined in from Category Table
  --      ,[SpanishProductName]
  --      ,[FrenchProductName]
  --      ,[StandardCost]
  --      ,[FinishedGoodsFlag] 
  p.[Color] AS [Product Color], 
  --      ,[SafetyStockLevel]
  --      ,[ReorderPoint]
  --      ,[ListPrice] 
  p.[Size] AS [Product Size], 
  --      ,[SizeRange]
  --      ,[Weight]
  --      ,[DaysToManufacture]
  p.[ProductLine] AS [Product Line], 
  --     ,[DealerPrice]
  --      ,[Class]
  --      ,[Style] 
  p.[ModelName] AS [Product Model Name], 
  --      ,[LargePhoto]
  p.[EnglishDescription] AS [Product Description], 
  --      ,[FrenchDescription]
  --      ,[ChineseDescription]
  --      ,[ArabicDescription]
  --      ,[HebrewDescription]
  --      ,[ThaiDescription]
  --      ,[GermanDescription]
  --      ,[JapaneseDescription]
  --      ,[TurkishDescription]
  --      ,[StartDate], 
  --      ,[EndDate], 
  ISNULL (p.Status, 'Outdated') AS [Product Status] 
FROM 
  [AdventureWorksDW2019].[dbo].[DimProduct] as p
  LEFT JOIN dbo.DimProductSubcategory AS ps ON ps.ProductSubcategoryKey = p.ProductSubcategoryKey 
  LEFT JOIN dbo.DimProductCategory AS pc ON ps.ProductCategoryKey = pc.ProductCategoryKey 
order by 
  p.ProductKey asc
```

### FACT_InternetSales


```
-- Cleansed FACT_InternetSales Table --
SELECT 
  [ProductKey], 
  [OrderDateKey], 
  [DueDateKey], 
  [ShipDateKey], 
  [CustomerKey], 
  --  ,[PromotionKey]
  --  ,[CurrencyKey]
  --  ,[SalesTerritoryKey]
  [SalesOrderNumber], 
  --  [SalesOrderLineNumber], 
  --  ,[RevisionNumber]
  --  ,[OrderQuantity], 
  --  ,[UnitPrice], 
  --  ,[ExtendedAmount]
  --  ,[UnitPriceDiscountPct]
  --  ,[DiscountAmount] 
  --  ,[ProductStandardCost]
  --  ,[TotalProductCost] 
  [SalesAmount] --  ,[TaxAmt]
  --  ,[Freight]
  --  ,[CarrierTrackingNumber] 
  --  ,[CustomerPONumber] 
  --  ,[OrderDate] 
  --  ,[DueDate] 
  --  ,[ShipDate] 
FROM 
  [AdventureWorksDW2019].[dbo].[FactInternetSales]
WHERE 
  LEFT (OrderDateKey, 4) >= YEAR(GETDATE()) -2 -- Ensures we always only bring two years of date from extraction.
ORDER BY
  OrderDateKey ASC

```

# Data Model

Below is a screenshot of the data model after cleansed and prepared tables were read into Power BI.
This data model also shows how FACT_Budget hsa been connected to FACT_InternetSales and other necessary DIM tables.




