# stock-analysis
## Overview of Project
This challenge was focused on VBA and creating macros that would run through stock data then show returns for each stock.

## Results
Once the macro calculated the returns, the first thing noticed is that only ENPH and RUN had growth in both years. Also, stock focused on the module exercises, DQ, returned the most significant loss between 2017 and 2018. Lastly, comparing the module script to the challenge, the refactored script ran significantly faster than the original. The results of each time are below.

2017 Module Timed Results
![VBA_Module_2017](https://github.com/racruz25/stock-analysis/blob/main/Resources/VBA_Module_2017.png)

2017 Challenge Timed Results
![VBA_Challenge_2017](https://github.com/racruz25/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

2018 Module Timed Results
![VBA_Module_2018](https://github.com/racruz25/stock-analysis/blob/main/Resources/VBA_Module_2018.png)

2018 Challenge Timed Results
![VBA_Challenge_2018](https://github.com/racruz25/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary
The refactored code ran faster. However, the original script contained the module's "DQ Analysis" macro, which may have affected the time. The main difference in the actual code between the two was the `tickerIndex`. The refactored script allowed for the use of `tickerIndex` rather than an `If` statement to find the `tickerVolumes`. 
