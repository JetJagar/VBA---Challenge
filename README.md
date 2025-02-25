# VBA Stock Analysis Challenge

## Overview
This project involves creating a VBA script that analyzes stock data for multiple quarters. The script will iterate through each stock ticker and compute key metrics, including price changes, percentage changes, and total volume. Additionally, it will identify the stocks with the greatest percentage increase, greatest percentage decrease, and highest total volume. The script should be capable of running across multiple worksheets representing different quarters.

## Requirements

### Core Functionality
- Loop through all stock tickers for each quarter.
- Calculate the following metrics:
  - **Ticker Symbol**
  - **Quarterly Change**: Difference between the opening price at the start of the quarter and the closing price at the end.
  - **Percentage Change**: Percentage difference relative to the opening price.
  - **Total Stock Volume**: Sum of the volume traded for each stock during the quarter.

### Additional Functionality
- Determine and display the following:
  - **Greatest % Increase**
  - **Greatest % Decrease**
  - **Greatest Total Volume**

### Advanced Functionality
- Modify the script to execute across all worksheets (i.e., all quarters) automatically.

## Implementation Steps
1. **Initialize Variables**:
   - Create variables for ticker symbols, quarterly changes, percentage changes, and total volume.
   - Use helper variables to track the opening price for each ticker.
   - Store maximum increase, decrease, and volume values.

2. **Loop Through Stock Data**:
   - Identify where each ticker starts and ends within the dataset.
   - Calculate the required metrics at the end of each ticker's data.
   - Format the output for clarity.

3. **Determine Greatest Values**:
   - Compare all percentage changes and total volumes to track the greatest values.
   - Store the corresponding ticker symbols for each.

4. **Optimize for Multiple Worksheets**:
   - Implement a loop to process each worksheet dynamically.
   - Ensure results are recorded separately for each quarter.

## Expected Output
- A structured table displaying:
  - **Ticker Symbol**
  - **Quarterly Change**
  - **Percentage Change**
  - **Total Volume**
- Additional summary statistics for highest gainers, losers, and most traded stocks.

## Notes
- Ensure proper formatting for percentage changes.
- Utilize conditional formatting to highlight gains and losses.
- Use efficient looping techniques to optimize performance for large datasets.
- Test across multiple worksheets to validate accuracy.

By following these steps, the script will provide a comprehensive stock analysis for each quarter, facilitating easy comparison and insights into market trends.




