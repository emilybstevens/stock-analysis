# VBA of Wall Street
Utilizing VBA to analyze stock market data for trends

## Overview of Project
VBA is a programming language used for Excel and other Microsoft programs. It provides a way to automate Excel spreadsheets to analyze large data sets. 
### Purpose
The purpose of this project is to provide the client with a convenient way to analyze stock market data to assist in data-driven market decisions. 

## Results
### Results of Stock Analysis
Overall, stocks generally did better in 2017 than they did in 2018, with a larger percentage of stocks resulting in a positive return. </br></br>
![Analysis of 2017 Stock Returns](resources/2017_refactored.png)
![Analysis of 2018 Stock Returns](resources/2018_refactored.png)</br></br>
As the client's parents are invested in the stock DAQO ("DQ" for short), it should be noted that DQ had a very poor return in 2018, even when compared against all of the poor stock returns for the same year. While the majority of stocks analyzed had negative returns in 2018, only two of them - including DQ - had a negative return larger than -60%. In 2017, on the other hand, DQ had a very successful year when compared against other stock returns in the same time frame. With a return of +199.4% in 2017 and a return of -62.6% in 2018, DQ has proven to be a particularly volatile stock. </br></br>
If the client's parents prefer to work with a volatile stock, DQ would be a good option. If they would prefer something more stable, DQ would not be a good option. Instead, ENPH might prove more to their taste, with a return of +129.5% (2017) and +81.9% (2018).  
### Results of Refactoring Code
Refactoring the code produced the same information at greater speeds. The original code returned results for the year 2017 in 0.78125 seconds, while the refactored code returned results in 0.1523438 seconds. Similarly, the original code for the year 2018 returned results in 0.796875 seconds, while the refactored code returned results in 0.1289063 seconds. While that difference might not seem very large, it should be noted that this code was ran on a relatively small data set. On a much larger set of data, the difference between run times could potentially prove to be more impressive.</br> 
#### Refactored Code Run Times
![Return Speed 2017 Refactored](resources/VBA_Challenge_2017.png)
![Return Speed 2018 Refactored](resources/VBA_Challenge_2018.png)</br>
#### Refactored vs. Original Code
The refactored code was written to provide the client with a program that could efficiently analyze larger data sets with ease. The original code required three separate passes through the data in order to obtain all values needed, while the refactored code only required one pass. This saves time and processing power. </br></br>
The loop in the original code was written as such:</br>
```
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
            Worksheets(yearValue).Activate
            For j = 2 To RowCount
                
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
                
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
                
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
            
            Next j

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
```
While the loop in the refactored code was written: 
```
    For i = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
            
    Next i

    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
    
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1

    Next i
```
By refactoring the code in this manner, the process of looping through the data was simplified, and as a result, the total run time was shortened. This will allow the code to run on larger data sets without causing too much strain on the computer or on the client's time. 
## Summary
By refactoring the code, the process of running the code can be simplified and the run time reduced. Unfortunately, the process of refactoring the code can be difficult and confusing if the coder does not fully understand the parts of the code they've written. In general, refactoring code would provide a more versatile tool for clients who may not have coding skills or the necessary knowledge to rework the non-refactored code themselves for other data sets. When a code is refactored, it can become more efficient and can be scaled to much larger projects without strain. This means that in the future, a coder will have significantly less work on their hands if they need to work on different data than originally provided. </br></br>
On the other hand, the original script without the refactoring could be considered "easier" to write, since there's less confusion over referencing indexes. Still, the original code is slow and puts more strain on the computer, since it has to run through the loop a few times to retrieve all of the necessary data. This means that on larger data sets, the non-refactored code would not be the most convenient option for a client to use. </br></br>

