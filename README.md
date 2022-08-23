## Module 2 Challenge: Stock Analysis

### Project Overview

* Using previously written Excel VBA code for analyzing the total stock volume, and price percent change, for twelve stocks across two years (2017 and 2018), demonstrate the advantage and disadvantages of refactoring code. 

### Refactoring Results and Walkthrough

* The code used in the original Green Stocks analysis utilzed two "For..To" loops to gather and store the data for each variable (stock volume, starting price and ending price) for the pre-defined array and indices. Essentially, each time the macro ran, Excel would loop through the entire data set 12 times (the pre-defined number of indices in the array) for each row (up to 3,013 rows) and, for every ticker symbol, record the total volume (as added across the year) and document the starting/ending prices to complete the percent change calculation. See the original code below:

  **For i = 0 To 11**
  
    **ticker = tickers(i)**
    **totalVolume = 0**
    
    **Sheets(yearValue).Activate**
        
    **For j = 2 To RowCount**
        
    **If Cells(j, 1).Value = ticker Then**
        
    **totalVolume = totalVolume + Cells(j, 8).Value**
        
    **End If**
        
    **If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then**
        
    **startingPrice = Cells(j, 6).Value**
            
    **End If**
        
    **If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then**
        
    **endingPrice = Cells(j, 6).Value**
            
    **End If**   
        
    **Next j** 
    
  **Worksheets("All Stocks Analysis").Activate**

  **Cells(4 + i, 1).Value = ticker**
  **Cells(4 + i, 2).Value = totalVolume**
  **Cells(4 + i, 3).Value = endingPrice / startingPrice - 1**
  
  **Next i**
  
  
 While this version of code works for a smaller array, if the array had hundreds of stock ticker symbols, looping through the entire data set thousands of times would make the code inefficient and result in a longer execution time. 
 
 By refactoring the code slightly and keeping key details from the original code, such as the array definitions, the yearValue definition, the calculations to find ticker volumes, ticker starting price and ticker ending price, we can make the code more efficient. 
 
 As seen below, unlike the original code, we opted to define three output arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices), each with an index count equal to our original "tickers(11)" array indices: (12). We can have the code loop through the data and store each value once in these output arrays for each index. 
 
 ###### **Refactored Code with New Output Arrays**
 
 ![New_Code_Screen_Shot](https://github.com/dpTuttle/stock-analysis/blob/main/Code_Screenshot_1_VBA_Challenge.png)
 
 
 To do so, we removed the innner 'For...To' loop that required the code to loop through the data for each ticker and instead have the code loop through the entire data set once for each index (adding to the index count each time it came across a new ticker symbol) and recording the total volume, starting price and ending price for each ticker symbol in the output array. This eliminates the need to check for the ticker symbol against the ticker array index row-by-row to gather the same data. See the above refactored code with seperate 'For..To' loops instead of nested 'For...To' loops. 

 
 The result of refactoring this code is more efficiency and faster execution times as detailed below. For the 2017 analysis, the original code ran in .34 seconds:
 
 ![2017_Original_Analysis](https://github.com/dpTuttle/stock-analysis/blob/main/Green_Stocks_Code_2017.png)
 
 The refactored code for the 2017 stock data ran in .09 seconds:
 
 ![2017_New_Code_Data](https://github.com/dpTuttle/stock-analysis/blob/main/VBA_Challenge_2017.png)
 
 Likewise, for the 2018 data set, the original code ran in .33 seconds:
 
 ![2018_Orginal_Analysis](https://github.com/dpTuttle/stock-analysis/blob/main/Green_Stocks_Code_2018.png)
 
 The refactored code for the 2018 data set ran in .10 seconds:
 
 ![2018_New_Code_Data](https://github.com/dpTuttle/stock-analysis/blob/main/VBA_Challenge_2018.png)
 
 I also took the opportunity to add a formatting subroutine into the refactored code to highlight the headers, bold the font, and center justify the text and adjust the column width. 

### Summary

###### **Pros / Cons of Refactoring in General**

* It is clear that using refactored code saves time in coding the data and ensures continuity between code sets --making code analysis easier to follow and amend in the future. Conversely, refactoring some code may not completlely do the job we want and also limits us to making adjustments only for the orgininal data set provided. 

###### **Pros / Cons of Refactoring this Code**

* Refactoring this code provided a quick and easier method for efficiently analyzing the predefined list of stocks the client provided. As such, we cab turn over a ready to use analysis that quickly analyzing the data for this portfolio across multiple years. 

* Unfortunately, this code only works for the twelve stocks defined in the 'tickers' array. If this data set contained 200 stock ticker symbols, the code would not work without defining those same stocks in the array. In this case, we may write different code that did the array analysis for us prior to looping through the data. Additionally, the code analyzes one year at a time. If the client would like to see the output for these same stocks across multiple years at once, this would require a rewrite of the code itself. 

