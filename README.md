# VBA-challenge

# Stock Data Analysis 

This VBA macro is designed to calculate yearly stock data across multiple sheets within an Excel workbook. It automates the process of analyzing stock performance by calculating yearly change, percentage change, and total stock volume for each ticker. Furthermore, it identifies the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume, yearly.

# Features
- Iterates through each stock ticker within a sheet and calculates:
  - Yearly change from the opening price at the beginning of the year to the closing price at the end of the year.
  - The percentage change from the opening price at the beginning of the year to the closing price at the end of the year.
  - The total stock volume over the year.
- Identifies and records:
  - The stock with the greatest percentage increase, greatest percentage decrease, and greatest     total volume over the year.
- Processes data across all sheets in a workbook, catering to multiple years of stock data.

# Output
- The macro outputs the calculated data into the following columns on each sheet:
  - I: Ticker
  - J: Yearly Change
  - K: Percentage Change
  - L: Total Stock Volume
- Additionally, it outputs the summary of greatest increase, decrease, and volume in a designated area on each sheet.
