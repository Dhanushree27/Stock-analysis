# Automating excel for Stock Market Analysis

## Overview of the project
This project was undertaken to assist Steve with analysis of stock market data, to provide suggestions for his parents who had enlisted his help. To optimize this task, macros were created using Microsoft's Visual Basic for Applications tool which is in-built inside excel. Macros were constructed to calculate the 'Total volume' and 'Yearly return' for each stock for years 2017 and 2018. The macro built initially using nested for loops held up well against the existing data, but would have had performance issues with larger datasets. Therefore, it was refactored using arrays to improve its performance

## Results

### Analysis of Data
From the analysis, it was observed that the 'DQ' stock that Steve's parents had invested in had performed poorly in 2018. Therefore, the performance of the other stocks was also analysed to identify which stock to invest in. From the analysis it was observed that while most stocks had performed well in 2017, their performance had dropped drastically in 2018. The 'DQ' stock which had a return of **199.45%** in 2017 dropped to **-62.6%** in 2018. The 'ENPH' and 'RUN' stocks were the only stocks with positive returns in 2018, out of which 'RUN' was the only stock which demonstrated a increase in return from **5.55%** to **83.5%**, with almost a 100% increase in total volume. On the other hand, the volume of 'ENPH' stocks increased from **221,772,100** in 2017 to **607,473,500** in 2018.

### Analysis of Performance
To improve the performance of the code for larger datasets, the code was refactored to use arrays instead of nested for loops. Below is a comparison of the initial code and the refactored code.

#### Initial code
The code was initially constructed using nested for loops, in which the code analysed all the rows for each stock. It also compared each row with the stock name for each measure. This resulted in an average run time of **1.5 seconds** for each year.

##### Code

'Loop for each ticker' _<- Loop 1_
 For i = 0 To 11
    TotalVolume = 0
    Sheets(yearValue).Activate _<-Reactivation of sheet for each loop_

    'Loops over all rows
    For j = RowStart To Rowend _<-Loop 2_
        'Calculate Total Volume
        If Cells(j, 1).Value = tickers(i) Then _<-Repeated checks for stock name_
            TotalVolume = TotalVolume + Cells(j, 8).Value
        End If
        'Calculates starting price for yearly return'
        If Cells(j, 1).Value = tickers(i) And Cells(j - 1, 1).Value <> tickers(i) Then _<-Repeated checks for stock name_
            StartingPrice = Cells(j, 6).Value
        End If
        'Calculates closing price for yearly return'
        If Cells(j, 1).Value = tickers(i) And Cells(j + 1, 1).Value <> tickers(i) Then _<-Repeated checks for stock name_
            EndingPrice = Cells(j, 6).Value
        End If
    Next j
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = TotalVolume
    Cells(4 + i, 3).Value = ((EndingPrice / StartingPrice) - 1)
Next i

![Initial Code_2017]https://github.com/Dhanushree27/Stock-analysis/blob/main/Resources/InitialCode_2017.PNG) 

![Initial Code_2018]https://github.com/Dhanushree27/Stock-analysis/blob/main/Resources/InitialCode_2018.PNG)

#### Refactored code
Since the data was in sequence and the sequence was defined in the tickers array, analysing each row for each stock was redundant. Therefore, the code was refactored to use arrays for the calculated measures as well. This reduced the run time considerably to about **.25 seconds**, which would vastly improve the performance of larger datasets.

##### Code

'Declaring ticker index variables to store results *<-Creation of arrays*
tickerIndex = 0
Dim tickerVolumes(0 To 11) As Long
Dim tickerStartingPrice(0 To 11) As Single
Dim tickerEndingPrice(0 To 11) As Single
                    
'Initializing ticker volume to zero *<-Initializing the arrays*
For i = 0 To 11
    tickerVolumes(i) = 0
Next i
'Loop through all rows *<-One loop for calculation with reduced repitions*
For i = 2 To RowCount

    'Increasing volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Cells

    'Identifying starting price for current ticker
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
    End If

    'Identifying ending price for current ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
        'Increasing ticker index
        tickerIndex = tickerIndex + 1
    End If
Next i       
'Populating results from arrays into cells _<-Separate loop for results reducing transition time_
For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = ((tickerEndingPrice(i) / tickerStartingPrice(i)) - 1)
Next i

![Refactored Code_2017]https://github.com/Dhanushree27/Stock-analysis/blob/main/Resources/Refactored_2017.PNG) 

![Refactored Code_2018]https://github.com/Dhanushree27/Stock-analysis/blob/main/Resources/Refactored_2018.PNG)

## Summary
Generally, the initial code may be written with the purpose of arriving at results. This might not always be the optimal or better performing code. Therefore, reviewing the code again might provide insights into:
a. Reducing redundancy, repetitions
b. Fine tuning for details
c. Correcting any missed errors
d. Formatting/ restructuring for better understanding
While these are some of the advantages, it is also possible to:
a. Break some existing functionality
b. Break links to other parts of the code
c. Lose track of minor details of the original requirement over time, if code is refactored multiple times

In this case, refactoring had improved the performance of the code by reducing redundancy, repetitive tasks and navigation time between sheets, but it also narrowed the usage of the code. The code will only be able to handle sequential data and it is necessary that the data in ticker array is also in sequence. On the other hand, the original code would have been able to handle disordered data as well.
