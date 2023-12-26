# VBA Challenge
Week 2 - Data Analytics Boot Camp - University of Oregon

![VBA Challenge](/images/banner.jpg)

## Instructions
Create a script that loops through all the stocks for one year and outputs the following information:
- The ticker symbol
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.
- Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
- Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

## Approach
For this problem, I chose to create a Ticker Summary Chart for each ticker symbol. As the script moved through each row of data, if the current row matched the previous ticker symbol, it would dynamically update the values on the chart.  If they did not match, the script would create a new row where it would start a new record.

Once it has moved through all the rows in the worksheet and completed the Ticker Summary Chart, it created the Global Summary Report.  I used Excel functions to identify the *Min/Max* value.  Then I used the *Match* function to find the location of the *Min/Max* value and return the corresponding row.  From there, I was able to then populate the Global Summary chart with the values from that row.

Once both charts are complete, the script moves onto the next worksheet in the workbook, summarizing that data.

## Results

![Results 2018 Tab](/images/results_2018.JPG)
*Results 2018 Tab*

![Results 2019 Tab](/images/results_2019.JPG)
*Results 2019 Tab*

![Results 2020 Tab](/images/results_2020.JPG)
*Results 2020 Tab*
