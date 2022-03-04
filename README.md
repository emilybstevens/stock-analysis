# VBA of Wall Street
Utilizing VBA to analyze stock market data for trends

## Overview of Project
VBA is a programming language used for Excel and other Microsoft programs. 
### Purpose
The purpose of this project is to provide the client with a convenient way to analyze stock market data to assist in data-driven market decisions. 

## Results
### Results of Stock Analysis
Overall, stocks generally did better in 2017 than they did in 2018, with a larger percentage of stocks resulting in a positive return. </br></br>
![Analysis of 2017 Stock Returns](resources/2017_refactored.png)
![Analysis of 2018 Stock Returns](resources/2018_refactored.png)</br></br>
As the client's parents are invested in the stock DAQO ("DQ" for short), it should be noted that DQ had a very poor return in 2018, even when compared against all the poor stock returns for the same year. While the majority of stocks analyzed had negative returns, only two of them had a negative return larger than -60%. In 2017, on the other hand, DQ had a very successful year when compared against other stock returns. With a return of +199.4% in 2017 and a return of -62.6% in 2018, DQ has proven to be a particularly volatile stock. </br></br>
If the client's parents prefer to work with a volatile stock, DQ would be a good option. If they would prefer something more stable, DQ would not be a good option. Instead, ENPH might prove more to their taste, with a return of +129.5% (2017) and +81.9% (2018).  
### Results of Refactoring Code
Refactoring the code produced the same information at greater speeds. Listed speeds for both original and refactored code are as follows:
* speed of original code (2017): 0.78125 s
* speed of original code (2018): 0.796875 s
* speed of refactored code (2017): 0.1171875 s
* speed of refactored code (2018): 0.1210938 s </br></br>
The refactored code was written to provide the client with a program that can efficiently analyze larger data sets with ease. The original code required three separate passes through the data in order to obtain all values needed, while the refactored code only required one pass. This saves time and processing power. </br></br>
The loop in the original code was written as such:</br></br>
`    For i = 0 To 11
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
    Next i`

## Summary
* What are the advantages of refactoring code? 
* What are the disadvantages of refacotring code? 
* What are the advantages of the original VBA script? 
* What are the disadvantages of the original VBA script? 
* What are the advantages of the refactored script?
* What are the disadvantages of the refacotred script?  