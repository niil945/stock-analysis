# stock-analysis

## Overview

This weeks challenge focused on refactoring code to improve execution time. In our original iteration we looped through the data once for each of the 12 tickers in the dataset. The objective of the refactoring was to see the result of minimizing looping through the dataset to a single loop, tabulating values based on indexes, and ensuring the final values received matched the original data.

## Results

In the initial code we hardcoded the ticker values, looped through each of the 3011 rows of data once for each of those 12 values via nested for loops, resulting in parsing 36,132 rows of data as seen here:

```
    'Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        'Loop through rows in the data
        For j = 2 To RowCount
            'Activate data worksheet
            Worksheets(yearValue).Activate
            'increase totalVolume if ticker is "DQ"
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
        'Output results
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
```

Tabulating the data this way resulted in the output requiring ~12-13 seconds to run.

[Pre-Refactor 2017 Timer](Resources/Pre-Refactor_Timer_2017.png)

[Pre-Refactor 2018 Timer](Resources/Pre-Refactor_Timer_2018.png)

Instead of looping through the data repeatedly for each ticker, we utilized an index and stored the values within the indexed arrays for each of the tickers we wanted to capture data for. Resulting code was marginally more complex but comparable length. The most important factor is that instead of parsing 36,132 rows of data the output is calculated parsing each line only once for a total of 3011 rows of data being parsed.

```
    For j = 2 To RowCount
    
        'store cumulative daily volume by ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        ticker(tickerIndex) = Cells(j, 1).Value
        
        'set ticker starting price if first instance of ticker
        If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then
        
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
        End If
        
        'set ending price and advance index if last instance of ticker
        If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
        
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                tickerIndex = tickerIndex + 1
                
        End If
       
    Next j
```

Tabulating the data within a single loop resulted in the output generating in ~.1 second. 

[Post-refactor 2017 Timer](Resources/VBA_Challenge_2017.png)

[Post-refactor 2018 Timer](Resources/VBA_Challenge_2018.png)

## Summary

There are a variety of advantages and disadvantages to consider when refactoring code. In this instance we were able to clean up the code and vastly improve the rate at which it was able to output the necessary data by a factor of ~100. This result was better than expected in that we cut the number of rows of data being parsed by 1/12th th total. If that was the primary factor in time inefficiency we should see a difference of a factor of ~10. Reviewing the initial pre-refactored code also lead to the insight that each loop was also activating a specific sheet. Utilizing the refactored code I tested by inserting a duplicate activation inside the for loop for the already active sheet. Doing so raised the run time to ~1.1 seconds. The sheet activation seems to have a significantly larger impact on run time than the repeated loops and demonstrates that things that seem like it would have a trivial impact may in fact be the largest source of delay to how long it takes code to process. When refactoring code everything should be scrutinized to assess performance gains and code simplification. Example code demonstrating the activation of an already active sheet and the timer result as follows:

```
    Worksheets(yearValue).Activate
    'Find number of rows
    rowStart = 2
    'RowCount code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    
    For j = 2 To RowCount
        'testing activating worksheet impact on timer
        Worksheets(yearValue).Activate
        'store cumulative daily volume by ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        ticker(tickerIndex) = Cells(j, 1).Value
```

[Post-refactor 2017 Activate Sheet Test Timer](Resources/Post-Refactor_Sheet_Activation_Test_Timer_2017.png)

However refactoring code requires care as making changes without understanding how everything works can result in errors and non-functional code. For this reason it is important to always comment providing insight for anyone that may refactor the code in the future and to ensure that no 'magic numbers' are being utilized that aren't comprehensible to another developer who could work on the code in the future. It requires verification and re-verification to ensure that the data output or function of the code is the same as the previous iteration, that it works with all other dependent code, and is easy to analyze and refactor again in the future.

For this specific piece of code refactoring it had other benefits beyond decreasing the time the code took to run. The original code had hardcoded values for the tickers which may work with the initial set of data we have but any future changes to the market may result in additions to or removals from that list of energy companies. While it would require additional refactoring, these changes result in the code being more easily modifiable to handle variations in the data for potential future years that could require analysis. Doing so would require counting unique instances of tickers at the start of the class and utilizing that value as a variable to create the for loops. I can see no disadvantage in refactoring the code as every aspect of the changes resulted in a positive improvement to the output.
