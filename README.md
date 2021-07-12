# VBA Challenge

## Project Overview
The purpose of this analysis is to help Steve to analyze the historical stock performance in 2017 and 2018. At the meantime, to explore the automation process using VBA code in Excel and to understand the importance and approach to improve the efficiency of the code.

## Dataset
The original dataset consists of thousands of rows of data, each includes the ticker name, trading date, closing price and day trading volume. The macro is created to quicky calculate the total trading volume of the stock over the year and the yearly return by finding the starting price and ending price.
## Comparison

Utilizing the for-loop function and the if condition the macro can automate the process of analysis. The original code could run fast, around 0.8s to complete the analysis of one year for 12 different tickers. The refactored can run even faster, around 0.1s for both the year of 2017 and 2018 to complete the same analysis. ![Alt text](resources/ VBA_Challenge_2017.png?raw=true ) ![Alt text](resources/ VBA_Challenge_2018.png?raw=true ) The main difference in the methodology is that in the original takes a ticker name and run through the whole list before taking another. Instead, the refactored code switch to the next ticker when the previous row was determined to be the last row for last ticker. Thus, the refactored code goes through the list only once. Thus, the refactored code requires less execution time. It may not sound a lot in difference, but in cases that if we are going to analyze thousands more tickers, and the dataset is thousands of times larger. The difference would be minutes comparing to hours.

## Limitation

Although the refactored method is faster in this project, there is also limitation to this approach. One of the contributors to this case is that the original dataset is compiled in order of the ticker name and the tickers array is also hard coded according to the appearance order of the ticker in the original dataset. if the data points are randomly placed in order the refactored code would fail, yet the slower code would run at the same execution time.
