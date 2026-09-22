/* =========================================================
   MINI-PROJECT 2
   Advanced SQL Data Warehouse & Business Analytics
   STEP 1: Create Database
   ========================================================= */

CREATE DATABASE Central_Superstore;
GO

USE Central_Superstore;
GO

USE Central_Superstore;
GO

/* =========================================================
   STEP 2: Create Staging Table
   Purpose:
   Store the raw data imported from Central_Superstore.xlsx
   ========================================================= */

CREATE TABLE dbo.Staging_Central_Superstore
(
    [Row ID]        INT,
    [Order ID]      VARCHAR(20),
    [Order Date]    DATE,
    [Ship Date]     DATE,
    [Ship Mode]     VARCHAR(50),
    [Customer ID]   VARCHAR(20),
    [Customer Name] VARCHAR(100),
    [Segment]       VARCHAR(50),
    [Country]       VARCHAR(100),
    [City]          VARCHAR(100),
    [State]         VARCHAR(100),
    [Postal Code]   VARCHAR(20),
    [Region]        VARCHAR(50),
    [Product ID]    VARCHAR(30),
    [Category]      VARCHAR(50),
    [Sub-Category]  VARCHAR(50),
    [Product Name]  VARCHAR(255),
    [Sales]         DECIMAL(18,4),
    [Quantity]       INT,
    [Discount]       DECIMAL(10,4),
    [Profit]         DECIMAL(18,4)
);
GO
SELECT COUNT(*) 
FROM dbo.Staging_Central_Superstore

/* =========================================================
   STEP 3: Load Raw CSV Data into Staging
   ========================================================= */

USE Central_Superstore;
GO

BULK INSERT dbo.Staging_Central_Superstore
FROM 'C:\Users\DELL\Downloads\DEPI\Central_Superstore.csv'
WITH
(
    FORMAT = 'CSV',
    FIRSTROW = 2,
    FIELDQUOTE = '"',
    FIELDTERMINATOR = ',',
    ROWTERMINATOR = '0x0a',
    TABLOCK
);
GO


/* =========================================================
   STEP 4: Verify Staging Data
   ========================================================= */

SELECT COUNT(*) AS Staging_Row_Count
FROM dbo.Staging_Central_Superstore;

SELECT TOP 10 *
FROM dbo.Staging_Central_Superstore;
GO
SELECT TOP 10 * 
FROM dbo.Staging_Central_Superstore

USE Central_Superstore;
GO

CREATE TABLE dbo.Dim_Customer
(
    Customer_Key INT IDENTITY(1,1) PRIMARY KEY,
    Customer_ID VARCHAR(20) NOT NULL,
    Customer_Name VARCHAR(100),
    Segment VARCHAR(50),
    Country VARCHAR(100),
    City VARCHAR(100),
    State VARCHAR(100),
    Postal_Code VARCHAR(20),
    Region VARCHAR(50)
);
GO

USE Central_Superstore;
GO

INSERT INTO dbo.Dim_Customer
(
    Customer_ID,
    Customer_Name,
    Segment,
    Country,
    City,
    State,
    Postal_Code,
    Region
)
SELECT DISTINCT
    [Customer ID],
    [Customer Name],
    [Segment],
    [Country],
    [City],
    [State],
    [Postal Code],
    [Region]
FROM dbo.Staging_Central_Superstore;
GO

SELECT COUNT(*) AS Customer_Count 
FROM dbo.Dim_Customer
SELECT TOP 10 * 
FROM dbo.Dim_Customer

USE Central_Superstore;
GO

CREATE TABLE dbo.Dim_Product
(
    Product_Key INT IDENTITY(1,1) PRIMARY KEY,
    Product_ID VARCHAR(30) NOT NULL,
    Product_Name VARCHAR(255),
    Category VARCHAR(50),
    Sub_Category VARCHAR(50)
);
GO

INSERT INTO dbo.Dim_Product
(
    Product_ID,
    Product_Name,
    Category,
    Sub_Category
)
SELECT DISTINCT
    [Product ID],
    [Product Name],
    [Category],
    [Sub-Category]
FROM dbo.Staging_Central_Superstore;
GO

SELECT COUNT(*) AS Product_Count
FROM dbo.Dim_Product;

SELECT TOP 10 *
FROM dbo.Dim_Product;

USE Central_Superstore;
GO

CREATE TABLE dbo.Dim_Date
(
    Date_Key INT PRIMARY KEY,
    Full_Date DATE NOT NULL,
    Day_Number INT,
    Month_Number INT,
    Month_Name VARCHAR(20),
    Quarter_Number INT,
    Year_Number INT
);
GO

INSERT INTO dbo.Dim_Date
(
    Date_Key,
    Full_Date,
    Day_Number,
    Month_Number,
    Month_Name,
    Quarter_Number,
    Year_Number
)
SELECT DISTINCT
    CONVERT(INT, FORMAT([Order Date], 'yyyyMMdd')) AS Date_Key,
    [Order Date] AS Full_Date,
    DAY([Order Date]) AS Day_Number,
    MONTH([Order Date]) AS Month_Number,
    DATENAME(MONTH, [Order Date]) AS Month_Name,
    DATEPART(QUARTER, [Order Date]) AS Quarter_Number,
    YEAR([Order Date]) AS Year_Number
FROM dbo.Staging_Central_Superstore
WHERE [Order Date] IS NOT NULL;
GO
SELECT COUNT(*) AS Date_Count
FROM dbo.Dim_Date;

SELECT TOP 10 *
FROM dbo.Dim_Date
ORDER BY Full_Date;

USE Central_Superstore;
GO

CREATE TABLE dbo.Dim_Ship_Mode
(
    Ship_Mode_Key INT IDENTITY(1,1) PRIMARY KEY,
    Ship_Mode VARCHAR(50) NOT NULL
);
GO

INSERT INTO dbo.Dim_Ship_Mode
(
    Ship_Mode
)
SELECT DISTINCT
    [Ship Mode]
FROM dbo.Staging_Central_Superstore
WHERE [Ship Mode] IS NOT NULL;
GO
SELECT COUNT(*) AS Ship_Mode_Count
FROM dbo.Dim_Ship_Mode;

SELECT *
FROM dbo.Dim_Ship_Mode
ORDER BY Ship_Mode_Key;

USE Central_Superstore;
GO

CREATE TABLE dbo.Fact_Sales
(
    Sales_Key INT IDENTITY(1,1) PRIMARY KEY,
    Order_ID VARCHAR(20) NOT NULL,
    Customer_Key INT NOT NULL,
    Product_Key INT NOT NULL,
    Date_Key INT NOT NULL,
    Ship_Mode_Key INT NOT NULL,
    Sales DECIMAL(18,4),
    Quantity INT,
    Discount DECIMAL(10,4),
    Profit DECIMAL(18,4)
);
GO

USE Central_Superstore;
GO

INSERT INTO dbo.Fact_Sales
(
    Order_ID,
    Customer_Key,
    Product_Key,
    Date_Key,
    Ship_Mode_Key,
    Sales,
    Quantity,
    Discount,
    Profit
)
SELECT
    s.[Order ID],
    c.Customer_Key,
    p.Product_Key,
    d.Date_Key,
    sm.Ship_Mode_Key,
    s.[Sales],
    s.[Quantity],
    s.[Discount],
    s.[Profit]
FROM dbo.Staging_Central_Superstore s
INNER JOIN dbo.Dim_Customer c
    ON s.[Customer ID] = c.Customer_ID
INNER JOIN dbo.Dim_Product p
    ON s.[Product ID] = p.Product_ID
INNER JOIN dbo.Dim_Date d
    ON CONVERT(INT, FORMAT(s.[Order Date], 'yyyyMMdd')) = d.Date_Key
INNER JOIN dbo.Dim_Ship_Mode sm
    ON s.[Ship Mode] = sm.Ship_Mode;
GO

SELECT COUNT(*) AS Fact_Row_Count
FROM dbo.Fact_Sales

USE Central_Superstore;
GO

SELECT
    Customer_ID,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Customer
GROUP BY Customer_ID
HAVING COUNT(*) > 1;

SELECT
    Product_ID,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Product
GROUP BY Product_ID
HAVING COUNT(*) > 1;

SELECT
    Date_Key,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Date
GROUP BY Date_Key
HAVING COUNT(*) > 1;

SELECT
    Ship_Mode,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Ship_Mode
GROUP BY Ship_Mode
HAVING COUNT(*) > 1;

USE Central_Superstore
GO

TRUNCATE TABLE dbo.Fact_Sales
GO

TRUNCATE TABLE 
dbo.Dim_Customer
GO


WITH Customer_Dedup AS
(
    SELECT
        [Customer ID],
        [Customer Name],
        [Segment],
        [Country],
        [City],
        [State],
        [Postal Code],
        [Region],
        ROW_NUMBER() OVER
        (
            PARTITION BY [Customer ID]
            ORDER BY [Row ID]
        ) AS rn
    FROM dbo.Staging_Central_Superstore
)
INSERT INTO dbo.Dim_Customer
(
    Customer_ID,
    Customer_Name,
    Segment,
    Country,
    City,
    State,
    Postal_Code,
    Region
)
SELECT
    [Customer ID],
    [Customer Name],
    [Segment],
    [Country],
    [City],
    [State],
    [Postal Code],
    [Region]
FROM Customer_Dedup
WHERE rn = 1;
GO

USE Central_Superstore;
GO

SELECT
    COUNT(*) AS Customer_Count
FROM dbo.Dim_Customer;

SELECT
    Customer_ID,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Customer
GROUP BY Customer_ID
HAVING COUNT(*) > 1;

USE Central_Superstore;
GO

SELECT COUNT(*) AS Staging_Count
FROM dbo.Staging_Central_Superstore;

SELECT COUNT(*) AS After_Customer
FROM dbo.Staging_Central_Superstore s
INNER JOIN dbo.Dim_Customer c
    ON s.[Customer ID] = c.Customer_ID;

SELECT COUNT(*) AS After_Product
FROM dbo.Staging_Central_Superstore s
INNER JOIN dbo.Dim_Customer c
    ON s.[Customer ID] = c.Customer_ID
INNER JOIN dbo.Dim_Product p
    ON s.[Product ID] = p.Product_ID;

SELECT COUNT(*) AS After_Date
FROM dbo.Staging_Central_Superstore s
INNER JOIN dbo.Dim_Customer c
    ON s.[Customer ID] = c.Customer_ID
INNER JOIN dbo.Dim_Product p
    ON s.[Product ID] = p.Product_ID
INNER JOIN dbo.Dim_Date d
    ON CONVERT(INT, FORMAT(s.[Order Date], 'yyyyMMdd')) = d.Date_Key;

SELECT COUNT(*) AS After_Ship_Mode
FROM dbo.Staging_Central_Superstore s
INNER JOIN dbo.Dim_Customer c
    ON s.[Customer ID] = c.Customer_ID
INNER JOIN dbo.Dim_Product p
    ON s.[Product ID] = p.Product_ID
INNER JOIN dbo.Dim_Date d
    ON CONVERT(INT, FORMAT(s.[Order Date], 'yyyyMMdd')) = d.Date_Key
INNER JOIN dbo.Dim_Ship_Mode sm
    ON s.[Ship Mode] = sm.Ship_Mode;

USE Central_Superstore;
GO

SELECT
    Product_ID,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Product
GROUP BY Product_ID
HAVING COUNT(*) > 1;

USE Central_Superstore
GO
TRUNCATE TABLE dbo.Dim_Product
GO

USE Central_Superstore;
GO

WITH Product_Dedup AS
(
    SELECT
        [Product ID],
        [Product Name],
        [Category],
        [Sub-Category],
        ROW_NUMBER() OVER
        (
            PARTITION BY [Product ID]
            ORDER BY [Row ID]
        ) AS rn
    FROM dbo.Staging_Central_Superstore
)
INSERT INTO dbo.Dim_Product
(
    Product_ID,
    Product_Name,
    Category,
    Sub_Category
)
SELECT
    [Product ID],
    [Product Name],
    [Category],
    [Sub-Category]
FROM Product_Dedup
WHERE rn = 1;
GO

USE Central_Superstore;
GO

SELECT
    COUNT(*) AS Date_Count
FROM dbo.Dim_Date;

SELECT
    Date_Key,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Date
GROUP BY Date_Key
HAVING COUNT(*) > 1;

USE Central_Superstore;
GO

SELECT
    COUNT(*) AS Ship_Mode_Count
FROM dbo.Dim_Ship_Mode;

SELECT
    Ship_Mode,
    COUNT(*) AS Duplicate_Count
FROM dbo.Dim_Ship_Mode
GROUP BY Ship_Mode
HAVING COUNT(*) > 1;

USE Central_Superstore;
GO

INSERT INTO dbo.Fact_Sales
(
    Order_ID,
    Customer_Key,
    Product_Key,
    Date_Key,
    Ship_Mode_Key,
    Sales,
    Quantity,
    Discount,
    Profit
)
SELECT
    s.[Order ID],
    c.Customer_Key,
    p.Product_Key,
    d.Date_Key,
    sm.Ship_Mode_Key,
    s.[Sales],
    s.[Quantity],
    s.[Discount],
    s.[Profit]
FROM dbo.Staging_Central_Superstore s
INNER JOIN dbo.Dim_Customer c
    ON s.[Customer ID] = c.Customer_ID
INNER JOIN dbo.Dim_Product p
    ON s.[Product ID] = p.Product_ID
INNER JOIN dbo.Dim_Date d
    ON CONVERT(INT, FORMAT(s.[Order Date], 'yyyyMMdd')) = d.Date_Key
INNER JOIN dbo.Dim_Ship_Mode sm
    ON s.[Ship Mode] = sm.Ship_Mode;
GO

SELECT COUNT(*) AS Fact_Row_Count
FROM dbo.Fact_Sales


USE Central_Superstore;
GO

ALTER TABLE dbo.Fact_Sales
ADD CONSTRAINT FK_FactSales_Customer
FOREIGN KEY (Customer_Key)
REFERENCES dbo.Dim_Customer(Customer_Key);
GO

USE Central_Superstore;
GO

ALTER TABLE dbo.Fact_Sales
ADD CONSTRAINT FK_FactSales_Product
FOREIGN KEY (Product_Key)
REFERENCES dbo.Dim_Product(Product_Key);
GO

USE Central_Superstore;
GO

ALTER TABLE dbo.Fact_Sales
ADD CONSTRAINT FK_FactSales_Date
FOREIGN KEY (Date_Key)
REFERENCES dbo.Dim_Date(Date_Key);
GO

USE Central_Superstore;
GO

ALTER TABLE dbo.Fact_Sales
ADD CONSTRAINT FK_FactSales_ShipMode
FOREIGN KEY (Ship_Mode_Key)
REFERENCES dbo.Dim_Ship_Mode(Ship_Mode_Key);
GO


-- Query 1: Overall Sales, Profit and Quantity

USE Central_Superstore;
GO

SELECT
    SUM(Sales) AS Total_Sales,
    SUM(Profit) AS Total_Profit,
    SUM(Quantity) AS Total_Quantity
FROM dbo.Fact_Sales;

-- Query 2: Sales and Profit by Product Category

SELECT
    p.Category,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Product p
    ON f.Product_Key = p.Product_Key
GROUP BY
    p.Category
ORDER BY
    Total_Sales DESC;

-- Query 3: Sales and Profit by Sub-Category

SELECT
    p.Sub_Category,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Product p
    ON f.Product_Key = p.Product_Key
GROUP BY
    p.Sub_Category
ORDER BY
    Total_Sales DESC;

-- Query 4: Sales and Profit by Customer Segment

SELECT
    c.Segment,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Customer c
    ON f.Customer_Key = c.Customer_Key
GROUP BY
    c.Segment
ORDER BY
    Total_Sales DESC;

-- Query 5: Monthly Sales Trend

SELECT
    d.Year_Number,
    d.Month_Number,
    d.Month_Name,
    SUM(f.Sales) AS Total_Sales
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Date d
    ON f.Date_Key = d.Date_Key
GROUP BY
    d.Year_Number,
    d.Month_Number,
    d.Month_Name
ORDER BY
    d.Year_Number,
    d.Month_Number;

-- Query 6: Top 10 Customers by Sales

SELECT TOP 10
    c.Customer_ID,
    c.Customer_Name,
    SUM(f.Sales) AS Total_Sales
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Customer c
    ON f.Customer_Key = c.Customer_Key
GROUP BY
    c.Customer_ID,
    c.Customer_Name
ORDER BY
    Total_Sales DESC;

-- Query 7: Top 10 Products by Profit

SELECT TOP 10
    p.Product_ID,
    p.Product_Name,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Product p
    ON f.Product_Key = p.Product_Key
GROUP BY
    p.Product_ID,
    p.Product_Name
ORDER BY
    Total_Profit DESC;

-- Query 8: Sales by Region

SELECT
    c.Region,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Customer c
    ON f.Customer_Key = c.Customer_Key
GROUP BY
    c.Region
ORDER BY
    Total_Sales DESC;

-- Query 9: Profitability Classification using CASE

SELECT
    p.Category,
    SUM(f.Profit) AS Total_Profit,
    CASE
        WHEN SUM(f.Profit) > 0 THEN 'Profitable'
        WHEN SUM(f.Profit) < 0 THEN 'Loss'
        ELSE 'Break Even'
    END AS Profitability_Status
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Product p
    ON f.Product_Key = p.Product_Key
GROUP BY
    p.Category
ORDER BY
    Total_Profit DESC;

-- Query 10: Average Order Value

SELECT
    SUM(Sales) / COUNT(DISTINCT Order_ID) AS Average_Order_Value
FROM dbo.Fact_Sales;

-- Query 11: Sales by Ship Mode

SELECT
    sm.Ship_Mode,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit,
    COUNT(*) AS Number_of_Sales
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Ship_Mode sm
    ON f.Ship_Mode_Key = sm.Ship_Mode_Key
GROUP BY
    sm.Ship_Mode
ORDER BY
    Total_Sales DESC;

-- Query 12: Customers with Sales Above Average using Subquery

SELECT
    c.Customer_ID,
    c.Customer_Name,
    SUM(f.Sales) AS Total_Sales
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Customer c
    ON f.Customer_Key = c.Customer_Key
GROUP BY
    c.Customer_ID,
    c.Customer_Name
HAVING
    SUM(f.Sales) >
    (
        SELECT AVG(Customer_Sales)
        FROM
        (
            SELECT
                Customer_Key,
                SUM(Sales) AS Customer_Sales
            FROM dbo.Fact_Sales
            GROUP BY Customer_Key
        ) AS Customer_Summary
    )
ORDER BY
    Total_Sales DESC;

-- Query 13: Sales and Profit by Year using CTE

WITH Yearly_Sales AS
(
    SELECT
        d.Year_Number,
        SUM(f.Sales) AS Total_Sales,
        SUM(f.Profit) AS Total_Profit
    FROM dbo.Fact_Sales f
    INNER JOIN dbo.Dim_Date d
        ON f.Date_Key = d.Date_Key
    GROUP BY
        d.Year_Number
)
SELECT
    Year_Number,
    Total_Sales,
    Total_Profit
FROM Yearly_Sales
ORDER BY
    Year_Number;

-- Query 14: Customer Sales Ranking using CTE

WITH Customer_Sales AS
(
    SELECT
        c.Customer_ID,
        c.Customer_Name,
        SUM(f.Sales) AS Total_Sales
    FROM dbo.Fact_Sales f
    INNER JOIN dbo.Dim_Customer c
        ON f.Customer_Key = c.Customer_Key
    GROUP BY
        c.Customer_ID,
        c.Customer_Name
)
SELECT
    Customer_ID,
    Customer_Name,
    Total_Sales,
    RANK() OVER (ORDER BY Total_Sales DESC) AS Sales_Rank
FROM Customer_Sales
ORDER BY
    Sales_Rank;

-- Query 15: Loss-Making Products

SELECT
    p.Product_ID,
    p.Product_Name,
    p.Category,
    p.Sub_Category,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Product p
    ON f.Product_Key = p.Product_Key
GROUP BY
    p.Product_ID,
    p.Product_Name,
    p.Category,
    p.Sub_Category
HAVING
    SUM(f.Profit) < 0
ORDER BY
    Total_Profit ASC;

-- Query 16: Profit Margin by Category

SELECT
    p.Category,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit,
    CAST(
        SUM(f.Profit) * 100.0 / NULLIF(SUM(f.Sales), 0)
        AS DECIMAL(10,2)
    ) AS Profit_Margin_Percentage
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Product p
    ON f.Product_Key = p.Product_Key
GROUP BY
    p.Category
ORDER BY
    Profit_Margin_Percentage DESC;

-- Query 17: Regional Customer and Sales Analysis

SELECT
    c.Region,
    COUNT(DISTINCT c.Customer_ID) AS Number_of_Customers,
    COUNT(DISTINCT f.Order_ID) AS Number_of_Orders,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Customer c
    ON f.Customer_Key = c.Customer_Key
GROUP BY
    c.Region
ORDER BY
    Total_Sales DESC;


-- Create Sales Analysis View

USE Central_Superstore;
GO

CREATE VIEW dbo.vw_Sales_Analysis
AS
SELECT
    f.Sales_Key,
    f.Order_ID,
    d.Full_Date AS Order_Date,
    d.Year_Number,
    d.Month_Number,
    d.Month_Name,
    c.Customer_ID,
    c.Customer_Name,
    c.Segment,
    c.Region,
    p.Product_ID,
    p.Product_Name,
    p.Category,
    p.Sub_Category,
    sm.Ship_Mode,
    f.Sales,
    f.Quantity,
    f.Discount,
    f.Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Customer c
    ON f.Customer_Key = c.Customer_Key
INNER JOIN dbo.Dim_Product p
    ON f.Product_Key = p.Product_Key
INNER JOIN dbo.Dim_Date d
    ON f.Date_Key = d.Date_Key
INNER JOIN dbo.Dim_Ship_Mode sm
    ON f.Ship_Mode_Key = sm.Ship_Mode_Key;
GO

SELECT TOP 20 * 
FROM dbo.vw_Sales_Analysis


-- Create Sales KPI Stored Procedure

USE Central_Superstore;
GO

CREATE PROCEDURE dbo.usp_Sales_KPI_By_Year
    @Year INT
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        d.Year_Number,
        SUM(f.Sales) AS Total_Sales,
        SUM(f.Profit) AS Total_Profit,
        SUM(f.Quantity) AS Total_Quantity,
        COUNT(DISTINCT f.Order_ID) AS Total_Orders,
        CAST(
            SUM(f.Profit) * 100.0 / NULLIF(SUM(f.Sales), 0)
            AS DECIMAL(10,2)
        ) AS Profit_Margin_Percentage
    FROM dbo.Fact_Sales f
    INNER JOIN dbo.Dim_Date d
        ON f.Date_Key = d.Date_Key
    WHERE d.Year_Number = @Year
    GROUP BY
        d.Year_Number;
END;
GO

SELECT DISTINCT
    Year_Number
FROM dbo.Dim_Date
ORDER BY Year_Number

USE Central_Superstore
GO
EXEC dbo.usp_Sales_KPI_By_Year
    @Year = 2016


-- Create Performance Indexes

USE Central_Superstore;
GO

CREATE INDEX IX_Fact_Sales_Customer_Key
ON dbo.Fact_Sales(Customer_Key);
GO

CREATE INDEX IX_Fact_Sales_Product_Key
ON dbo.Fact_Sales(Product_Key);
GO

CREATE INDEX IX_Fact_Sales_Date_Key
ON dbo.Fact_Sales(Date_Key);
GO

CREATE INDEX IX_Fact_Sales_Ship_Mode_Key
ON dbo.Fact_Sales(Ship_Mode_Key);
GO

CREATE INDEX IX_Fact_Sales_Order_ID
ON dbo.Fact_Sales(Order_ID);
GO

-- Execution Plan Review

USE Central_Superstore;
GO

SELECT
    c.Segment,
    SUM(f.Sales) AS Total_Sales,
    SUM(f.Profit) AS Total_Profit
FROM dbo.Fact_Sales f
INNER JOIN dbo.Dim_Customer c
    ON f.Customer_Key = c.Customer_Key
GROUP BY
    c.Segment
ORDER BY
    Total_Sales DESC;

-- Final Business KPI Summary

USE Central_Superstore;
GO

SELECT
    SUM(Sales) AS Total_Sales,
    SUM(Profit) AS Total_Profit,
    SUM(Quantity) AS Total_Quantity,
    COUNT(DISTINCT Order_ID) AS Total_Orders,
    CAST(
        SUM(Profit) * 100.0 / NULLIF(SUM(Sales), 0)
        AS DECIMAL(10,2)
    ) AS Profit_Margin_Percentage
FROM dbo.Fact_Sales;