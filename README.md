Data exploration and Analysis
# Excel Data exploration and Analysis Assignment

## Project Overview
This project focuses on analyzing a product dataset using Microsoft Excel. I tried using power query editor in excel as well. It demonstrates the use of basic Excel functions, logical functions, conditional calculations, and text functions to summarize product information and extract meaningful insights.

## Objectives
- Calculate summary statistics for product prices.
- Identify minimum and maximum product prices.
- Categorize products based on their price.
- Perform conditional calculations using SUMIF and COUNTIF.
- Extract information from Product IDs using LEFT, RIGHT, and MID.

## Tools Used
- Microsoft Excel
- Excel formulas and functions
- Power Query Editor in excel

## Tasks Performed: separate sheets "Qns Answers of data exploration". created a table (Data_Exploration_EA1) using raw data set and performed the following functions.

### 1. Sum, Count, and Average
Used Excel functions to analyze product prices.

- Total Price: Calculated the sum of all product prices using SUM.
- Total Products: Counted the number of products using COUNTA.
- Average Price: Calculated the average product price using AVERAGE.

**Excel Formulas:**

    =sum(Data_Exploration_EA1[Price ($)])

    =COUNTA(Data_Exploration_EA1[Price ($)])

    =AVERAGE(Data_Exploration_EA1[Price ($)])

### 2. Minimum and Maximum Values
Identified the lowest and highest product prices in the dataset.

- Minimum Price: Used MIN to find the lowest price.
- Maximum Price: Used MAX to find the highest price.

**Excel Formulas:**

    =MIN(Data_Exploration_EA1[Price ($)])

    =MAX(Data_Exploration_EA1[Price ($)])

### 3. Logical Function - IF
Created a new column named Price Range to classify products based on their prices.

- If the price is greater than or equal to $500, the product is classified as High Price.
- If the price is less than $500, the product is classified as Standard Price.
- logical function: using power query editor created a column as price range.

**Excel Formula:**

    =IF([@[Price ($)]]>=500,"High Price","Standard Price")

**power Query Editor:**

 Add column-> conditional column -> enter the condition ">=500" -> ok

### 4. Conditional Functions - SUMIF and COUNTIF
Used conditional functions to analyze specific product groups.

- Total Electronics Price: Calculated the total price of products in the Electronics category using SUMIF.
- Products Below $100: Counted the products priced less than $100 using COUNTIF.

**Excel Formulas:**

    =SUMIF(Data_Exploration_EA1[Category],"Electronics",Data_Exploration_EA1[Price ($)])

    =COUNTIF(Data_Exploration_EA1[Price ($)],"<100")

### 5. Text Functions - LEFT, RIGHT, and MID
Created three new columns to extract information from the Product ID.

- Day: Extracted the first two characters of the Product ID using LEFT.
- Country Code: Extracted the last two characters of the Product ID using RIGHT.
- Month: Extracted the 4th to 6th characters of the Product ID using MID.

**Excel Formulas:**

    =LEFT(Data_Exploration_EA1[Product ID],2)

    =RIGHT(Data_Exploration_EA1[Product ID],2)

    =MID(Data_Exploration_EA1[Product ID],4,3)

## Project Structure

    Excel-Data-Analysis/
    |
    |-- README.md
    |-- EA1 Data exploration and analysis Assignment Document.xlsx
    |-- screenshots/
        |-- EA1 data exploration Table and Screen Shots document.png

## Key Learnings
- Using Excel functions for data summarization.
- Applying logical conditions to categorize data.
- Performing conditional calculations.
- Extracting specific characters from text fields.
- Organizing and documenting a data analysis project for GitHub.

## Conclusion
This assignment demonstrates how Excel functions can be used to analyze product data, summarize prices, categorize products, calculate category-specific metrics, and extract useful information from Product IDs.

The project provides a foundation for further learning in data cleaning, data analysis, and reporting using spreadsheet tools.

## Project Information
- Project Type: Excel Data exploration and Analysis
- Tool: Microsoft Excel with power Query editor
