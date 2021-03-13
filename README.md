# stock-analysis
Module 2: VBA Of Wall Street

# VBA Stock Analysis Challenge

## Overview
### General Intent
The purpose of this VBA code was to create a way to analyze a set of data related to stocks over multiple years. The key factors were: 
1. Measuring if each stock's prices increased or decreased from the beginning of the year to the end of the year thus determining return on investment. 
2. To determine how much each stock was traded.

This code was designed to allow us to not only assess the two years worth of data that we had, but also to be able to apply it to other years as well. 

### Coding Intnet
This project was used because it allowed us to utilize For Loops, If/Then statements, nested Loops, varibales, arrays, formatting and conditional formatting in a VBA enviroment. This project also required debugging and refactoring to ensure that the code worked as intended and worked more efficiently than the code generated in the module. 

## Results
### 2017 Stock Analysis
The outcomes of the analysis for 2017 showed that all stocks save for TERP had a positive return on investment. Of these profitable stocks, DQ showed the highest return at 199.4% followed by SEDG with 184.5% positive return. TERP showed a -7.2% return. 

![2017 Stocks Analysis](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_2017%20output.png)

### 2018 Stock Anaylysis
The outcomes of the analysis for 2018 showed that all stocks had a positive return on investment wth ENPH coming in the highest with 285.6%, JKS with a 245.2% return, and DQ with a 226.2% return. 

![2018 Stocks Analysis](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_2018%20output.png)

### Coding Results
Refactoring the code from the original version to the new version showed varying improvement in processing times for each year, however there was an improvement both times. If this code is to be edited to take in thousands of stocks, the small improvements in processing time could add up to significant changes. For 2017, the old code originally took 1.300781 seconds where the new code only too 0.2089844 seconds. This is a difference of 1.0917966 seconds for only 12 stocks. If this formula is expanded to thousands of stocks, the time savings of the refactored code would be immense. For 2018, the old code took 0.2148438 seconds where the refactored code took 0.1992188 sceonds for an improvement of 0.015625 seconds. This is a much smaller difference than the 2017 sheet, however when expanded from 12 stocks to thousands, the difference would add up significantly.

#### 2017 Original Code Run Time
![2017 Old Code Run Time](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_OldCode_2017.png)

#### 2017 Refactored Code Run Time
![2017 Refactored Code](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_2017.png)

#### 2018 Original Code Run Time
![2018 Old Code Run Time](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_OldCode_2018.png)

#### 2018 Refactored Code Run Time
![2018 Refactored Code Run Time](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_2018.png)

The reason for this improvement in time is related to the methodology of how the code processes the data. In the original code, we used a double For Loop which required the code to process the entirety of the dataset for each of the stock tickers. This process takes a lot of time. The refactored code used only one For Loop and utilized If/Then statements with each pass to populate arrays with the desired data so only one pass through the stock data was required.

#### Original Code Double For Loops
![Original Code Double For Loops](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_OldCode_Double_Loops.png)

#### Refactored Code For Loop With If/Then Statements
![Refactored Code For Loop With If/Then Statements](https://github.com/ghynox/stock-analysis/blob/main/VBA_Challenge_RefactoredCode_Single_Loop.png)

## Summary
### Code Refactoring
Refactoring of code can have both advantages and disadvantages, and the ratio of those will determine if it shoudld be done. Advantages of refactoring are fairly straight forward. By reworking code, it can be made to be more efficient time wise, space wise, and could make it easier for subsequent coders to understand the code and update or edit it later. In the case of this project, we imprved the efficiency of the code and saved potential run times in the future. The disadvantages of refactoring are more related to the time spent having to rethink the structuring of the code and then debugging any issues that come up along the way. This could require quite a bit of work and creative thinking. The disadvantages are more front loaded with work time, however the advantages are realized in the long run. In this case, the benefits outweighed the costs and it was worth it to refactor and produce something more efficient and more easily adapted to larger scale operations. 
