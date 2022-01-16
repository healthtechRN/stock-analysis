# stock-analysis
## Overview of Project
Using VBA to analyze 2017 and 2018 stock data.
### Purpose
In this project, 2017 and 2018 Stock Market Data was analyzed to create a VBA macro that can trigger pop-ups and inputs, read and change cell values, and format cells. Loops, nested for loops and conditionals were used to direct flow. Finally, the code was made more efficient by taking fewer steps, using less memory, and improving the logic of the code to find the total daily volume and yearly return for each stock in our dataset for 2017 and 2018. Coding skills were applied such as syntax recollection, pattern recognition, problem decomposition, and debugging.
##Results: 
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
![VBA Challenge 2017](Resources/VBA_Challenge_2017.png)
![VBA Challenge 2018](Resources/VBA_Challenge_2018.png)
## Summary: In a summary statement, address the following questions.
- What are the advantages or disadvantages of refactoring code?
- How do these pros and cons apply to refactoring the original VBA script?

![Results 2017](Resources/2017_Results.png)
![Results 2018](Resources/2018_Results.png)
```VBScript
'1a) Create a ticker Index
    For i = 0 To 11
        tickerIndex = tickers(i)
```
