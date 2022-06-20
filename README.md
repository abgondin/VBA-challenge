# 02 VBA Homework - The VBA of Wall Street

## Background
Here, I use VBA scripting to analyze real stock market data. 
The goal is to create a script that runs across different worksheets and that will loop through the daily stock trading information for a given year and output:
1. The ticker symbol.
2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
4. The total stock volume of the stock.
5. The stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".


## Skills
VBA scripting | loops | conditional formatting

## Conclusion

The alphabetical_testing.xlsx file was used for testing while generating the script. This data set is smaller and allowed me to test faster.

I then applied the script to Multiple_year_stock_data.xlsx to analyse all stocks traded each year from 2014-2016. Multiple_year_stock_data.xlsx contains a worksheet for each year and ~800K rows each worksheet.

In 2014, the stock that returned the most value was DM (5581% increase) and the most traded stock was BAC.
<img width="1136" alt="Screen Shot 2021-02-15 at 3 45 21 pm" src="https://user-images.githubusercontent.com/77761497/174531654-ff231a9e-809e-4a6d-8b93-a11c61c8c15c.png">

In 2015, the stock that returned the most value was ARR (491% increase) and the most traded stock continued being BAC.

<img width="1142" alt="Screen Shot 2021-02-15 at 3 45 07 pm" src="https://user-images.githubusercontent.com/77761497/174532062-6b612acb-b39c-4c66-b6ec-ec2ea05c605d.png">

In 2016, the stock that returned the most value was SD (11675% increase) and the most traded continued being BAC.

<img width="1235" alt="Screen Shot 2021-02-15 at 3 44 46 pm" src="https://user-images.githubusercontent.com/77761497/174532089-e1638397-827a-4808-9b20-b389ff21a157.png">

