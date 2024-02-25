# VBA-challenge
## Overview:
This repository contains a stock analysis script written in VBA. The script is designed to analyze stock price information, which is provided in the Multiple_year_stock_data workbook. The workbook includes data such as date, opening price, highest and lowest points of the day, closing market price, and total volume of stocks traded.

## Functionality:

The script iterates through all the information, tables, and worksheets within the Multiple_year_stock_data workbook.
For each ticker name, it automatically calculates the yearly change between the opening price at the beginning of the year and the closing price recorded at the end of that year, as well as the total stock volume.
Additionally, the script identifies which ticker had the highest percent change in increasing and decreasing price within the year, and also highlights the stock with the highest total stock volume.

## Usage:

- Open the Multiple_year_stock_data workbook.
- Run the VBA script provided in this repository.
- The script will analyze the stock data and generate insights automatically.
Note: Make sure to enable macros in Excel to run the VBA script successfully.

## Contents:

- VBAScript: Stock_Analysis visual basic script
- Results_year: Results Screenshots 
- README.md: Documentation providing an overview of the repository and instructions for usage.
- Multiple_year_stock_data.xlsm: Excel workbook containing stock price information from 2018 to 2020.
