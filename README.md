# VBA stock_analysis
Performing analysis on green energy stock data.

## Overview of Project
This project used Visual Basic for Applications (VBA) programming language to create flexible and interactive macros to run analyses on multiple stocks. The results of the analyses provide insights on the trading volume and the performance of a green energy stock, DAQO New Energy Corp (DQ) and will guide decisions on how to diversify the green energy stock portfolio. The analyses will also provide information on the cost of running the VBA automated scripts. The analysis was performed using the [Stock_analysis](https://github.com/aobasuyi/stock_analysis/blob/main/VBA_Challenge.xlsm) dataset.

### Purpose
To explore green energy stock performance by analyzing financial data using Visual Basic for Applications (VBA) and to refactor codes to make the VBA scripts run faster.

## Results

**Comparison of stock performance between 2017 and 2018:**<br />
- **2017:** *DQ* had the highest return of all stocks in 2017 at about 200%. *TERP* had the least return and dropped about 7%. *FSLR* was the most actively traded stock and *DQ* was the least actively traded stock.<br /><br /> *![VBA_Challenge 2017](Module%202_Resources/VBA_Challenge_All%20Stocks_2017.png)*<br />

- **2018:** *RUN* had the highest return at 84% followed by the *ENPH* at 82%. ENPH was the most actively traded stock of the year. *DQ* dropped over 63% and had the least performance of the year. <br /><br />![image](Module%202_Resources/VBA%20Challenge_All%20Stocks_2018.png)
<br />

**Comparison of VBA scripts execution time between 2017 and 2018** <br /><br />
- **2017:** The original VBA script execution time was **0.8085938** seconds. <br />
*![Original_2017](Module%202_Resources/VBA_Original_2017.png)*. <br /> 

The refactored VBA script execution time was **0.09375** seconds respectively. <br />
*![Refactored_2017](Module%202_Resources/VBA_Challenge_2017.png)* <br />

- **2018:** The executive time of the original VBA script was **0.796875** seconds.<br />
*![Original_2018](Module%202_Resources/VBA_Original_2018.png)*<br />
While the refactored VBA script execution time was **0.0859375** seconds respectively.
<br />, *![Refactored_2018](Module%202_Resources/VBA_Challenge_2018%20.png)*
 
## Summary
**Advantages and disadvantages of refactoring code in general:** <br />
Refactoring is a key part of the coding process to make a code more efficient. It does not add new features or functionalities. Refactoring may involve taking fewer steps or improving the logic of the code to make it easier for future users to read. Some advantages include refactored codes are less complex and are easier to understand or read.  Refactoring takes make codes more by taking fewer steps thereby using less memory. However, refactoring is time consuming and can introduce bugs. The cost of refactoring therefore can be higher than rewriting the code from scratch.<br /><br />
**Advantages and disadvantages of the original and refactored VBA script:**<br />
After refactoring, the VBA codes were less complex and easier to read because it did not used nested “For” loops. The execution time for the refactored codes reduced significantly which could lead to cost savings in real life situations. However, refactoring the original VBA scripts introduced bugs to the codes. It was time consuming to debug the code using ***Toggle Breakpoints*** to track each line of code to fix the bugs.
