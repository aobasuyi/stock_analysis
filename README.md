# VBA stock_analysis
Performing analysis on green energy stock data.

## Overview of Project
This project used Visual Basic for Applications (VBA) programming language to create flexible and interactive macros to run analyses on multiple stocks. The results of the analyses provide insights on the trading volume and the performance of a green energy stock, DAQO New Energy Corp (DQ) and will guide decisions on how to diversify the green energy stock portfolio. The analyses will also provide information on the cost of running the VBA automated scripts. The analyses were performed using the [Stock_analysis](https://github.com/aobasuyi/stock_analysis/blob/main/VBA_Challenge.xlsm) dataset.

### Purpose
To explore green energy stock performance by analyzing financial data using VBA and to refactor codes to make the VBA scripts run faster.

## Results

**Comparison of stock performance for 2017 and 2018:**<br />
**2017:** *DQ* had the highest return of all stocks in 2017 at about 200%. *TERP* had the least return and dropped about 7%. *FSLR* was the most actively traded stock and *DQ* was the least actively traded stock.<br /><br /> *![VBA_Challenge 2017](Module%202_Resources/VBA_Challenge_All%20Stocks_2017.png)*<br />

**2018:** *RUN* had the highest return at 84% followed by the *ENPH* at 82%. *ENPH* was the most actively traded stock of the year. *DQ* dropped over 63% and had the least performance of the year. <br /><br />![image](Module%202_Resources/VBA%20Challenge_All%20Stocks_2018.png)
<br />

**Comparison of codes and execution time for 2017 and 2018:** <br />
To calculate the yearly return and how actively green energy stocks were traded for each year, nested ***For*** Loop was used in the original script analysis, while ***For*** Loops only, were used in the refactored script as shown below.<br /><br />
**Original all stocks analysis code:**
```
#Loop through the tickers.
     For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
  Worksheets(yearValue).Activate
        For j = 2 To RowCount
    If Cells(j, 1).Value = ticker Then
    	totalVolume = totalVolume + Cells(j, 8).Value
        End If
    If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
        'set starting price
            startingPrice = Cells(j, 6).Value
        End If
    If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
       'set ending price
            endingPrice = Cells(j, 6).Value
        End If
    Next j
  #Output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
Next i
```
**Refactored all stocks analysis code:**
```
#Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    For i = 2 To RowCount
              tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           End If
       If the next row???s ticker doesn???t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        #Increase the tickerIndex.
             tickerIndex = tickerIndex + 1
             End If
    Next i
```
**Execution time for 2017 and 2018:** <br /><br />
**2017:** The original script execution time was **0.8085938** seconds while refactoring the codes resulted in a lower execution time of **0.09375** seconds respectively as shown below. <br />
| Original script time  | Refactored script time |
| ------------- | ------------- |
| ![Original_2017](Module%202_Resources/VBA_Original_2017.png)  | ![Refactored_2017](Module%202_Resources/VBA_Challenge_2017.png)  |

**2018:** The original script execution time was **0.796875** seconds while refactoring the codes also resulted in a lower execution time of **0.0859375** seconds respectively as shown below .<br />

| Original script time  | Refactored script time |
| ------------- | ------------- |
| ![Original_2018](Module%202_Resources/VBA_Original_2018.png)  | *![Refactored_2018](Module%202_Resources/VBA_Challenge_2018%20.png)*  |

## Summary
**Pros and Cons of refactoring code in general:** <br />
Refactoring is a key part of the coding process to make a code more efficient. It does not add new features or functionalities. Refactoring may involve taking fewer steps or it may involve improving the logic of the code to make it easier for users to read. Some advantages include that codes are easier to read and understand. Codes are also less complex or may take fewer steps thereby using less memory. However, refactoring is time consuming and can introduce bugs. The cost of refactoring therefore, can be higher than rewriting the code from scratch.<br /><br />
**Pros and Cons of the original and refactored VBA script:**<br />
The refactored VBA script for this project was less complex and easier to read because it did not use nested ***For*** loops. The execution time for the refactored script reduced significantly which could lead to cost savings in real life situations. However, refactoring the original VBA script introduced bugs to the codes. It was time consuming to debug the code using ***Toggle Breakpoints*** to track each line of code to fix the bugs.
