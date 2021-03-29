# Stock-Analysis

# Overview of Project

## Purpose

The purpose of this project was to refactor Microsoft Excel VBA code to collect total daily volume and return for 12 green energy stock prices for the years 2017 and 2018 and determine if these stocks are worthwhile investing in. The starting point was to reuse code to set up a similar structure. The goal of the challenge was to reuse and refactor the original code to perform the analysis for 2017 and 2018 stocks and improve the overall efficiency of the code.

## Data

The data used in this analysis were two separate worksheets for 2017 and 2018 stock information. The data contained therein included the ticker index, the stock issue date, the opening price, the high and low price, the adjusted closing price, and the volume for each stock. The objective of the analysis was to retrieve the ticker, the total daily volume, and the return on each stock.

The challenge instructions were to add code to determine the total daily volume and return for each stock for each year and refactor the code structure.

## Results

Before refactoring the code, I reviewed the code that I could reuse and copied the code to create the table headers,the input box for the year that I want to run the analysis, the ticker array, and to format the header rows and conditional color formatting. I also made sure that I activated the correct worksheet to output the data.

The instruction and code as written in the file are below.


    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
     For i = 0 To 11
            tickerVolumes(i) = 0
            
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
            
    
        '3a) Increase volume for current ticker
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
  
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value 
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
                    
                        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                        
                        '3d Increase the tickerIndex
                        
                        tickerIndex = tickerIndex + 
                End If
                
Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
      Next i
            

# Summary

## Advantages and Disadvantages of Refactoring Code

Refactoring code has several advantages. It can help to reduce complexity and duplication in the code, especially for legacy code. It also improves the organization and efficiency of the code to achieve better performance and increase its overall readability and maintability over time. Another key benefit is that it makes the code more extensible and flexible. Moreover, it can help other developers understand the code better. It can be beneficial in situations where there are multiple developers collaborating on a project but on a different aspects of the code; refactoring helps developers to understand the end to end logic and flow of the code if the code is well organized and clean.

There are some disadvantages with refactoring code. Refactoring can introduce defects but this can be mitigated with proper peer review and testing. It can be time consuming, especially for a large or legacy code base. Another disadvantage is that refactoring does not address any underlying architectural issues. Lack of or insufficient test cases can also add risks to refactoring.

## Advantages and Disadvantages of the original and refactored VBA script

As a result of refactoring the original VBA script, the code run time performance improved from minutes to well under one second. Another major benefit is that the code organization and readability improved resulting in performance efficiency.

Attached below are the screenshots to demonstrate the improved run time for the refactored code and the summary table of the stocks and their return.

<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/80140082/112773639-8d8f5700-8feb-11eb-92c4-c565e7ababf8.png">

<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/80140082/112773592-6cc70180-8feb-11eb-98ec-305558a02c67.png">

![image](https://user-images.githubusercontent.com/80140082/112773684-ba436e80-8feb-11eb-96e6-743bcfb253d8.png)

![image](https://user-images.githubusercontent.com/80140082/112773715-e5c65900-8feb-11eb-870c-5e26c3116a37.png)


