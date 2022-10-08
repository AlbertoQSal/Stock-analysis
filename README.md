# Stock-analysis
## Second Challenge

### Overview
The main objective of the proyect was to help Steve to show his parents the behavior of different stocks so they can be able to invest their money in the stocks that gives them the best profit.

In this proyect I refactored the code to loop through all the data in the VBA_Challenge.xslm file, with the purpose to make the code more efficient. The code created had the objective of presenting the Total Daily Volume and the Return (with conditional formatting) for each stock. This code was modified from the code belonging to the green_stocks.xslm file.

### Results
The Results of the analysis were a table that presented the Total Daily Volume and the Return for each stock in the year 2017 or 2018. The year was provided by the user and the analysis ran according to this iformation. 
  
The code was used to navigate through the spreadseets with all the stocks data, and to calculate the Total Daily Volume and the Return of each stock. In order to present the information in the table, four arrays were used, this arrays were: tickers, tickersVolumnes, tickerEndingPrices and tickerStartingPrices.

Te results obtained were: 

For the year 2017:

![Results_2017](https://user-images.githubusercontent.com/93279134/194717928-a72379b9-133c-42b3-993e-424cc318f6ea.png)

For the year 2018: 

![Results_2018](https://user-images.githubusercontent.com/93279134/194717933-d138f5fb-678a-49c9-a03e-55dd966ac5ae.png)

### Summary
The main advantage of the refactoring of the code is that it makes it better and faster. The use of differents arrays and avoiding the use of nested for loops allows us to have a cleaner code and to avoid mistakes in the navigation through the arrays and the worksheets. Te disadvantage of this is the use of different indexes to navigate trough the different arrays, so if you're mistaken in one or more indexes it may take you a while to find the mistake.

The general structure of the original and the refactored code are shown below:

Original code:

'Nested loops
 For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
    '5) Loop through rows in the data.

    Worksheets(yearValue).Activate
    For j = 2 To rowEnd
        '5.1) Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then
    
            totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
        '5.2) Find the starting price for the current ticker.
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
            'Set starting price
            startingPrice = Cells(j, 6).Value
            
        End If
        '5.3) Find the ending price for the current ticker.
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
            'Set ending price
            endingPrice = Cells(j, 6).Value
    
        End If
    
    Next j
    
    'Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
    
Next i

Refactored code:

Create a for loop to initialize the tickerVolumes array to zero.
    For i = 0 To 11
        tickerVolumes(tickerIndex) = 0
    Next i
    
    'For loop to travel across all the rows in the spreadsheet.
        For i = 2 To RowCount
        ' Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        ' Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        ' Check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            ' Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    
        Next i
    
    ' Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

In the original code I used two nested for loops that allowed to obtain the information form each worksheet, assigned it to a variable and to make the result an output. The main disadvantage of this process is that once calculated the result, the variable was set to zero, so the information was lost and if you wanted to work with it later, you need to recalculate all the information. Also, the use of nested for loops add a certain level of difficultie because you need to understand what is going on at each loop.

On the other side, the refactored code used single loops, this helps to isolate the procceses and to make the code run faster. Also, the use of arrays allowed to save each result, so if any other procces were needed you could only call the array and access to the specific information that you need.

The refactored code had a difference in running time betwen 0.5 and 1 second for each year. And the results of the performing of the refactored code are chown below.

For year 2017:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/93279134/194719096-6e4f0722-100b-46da-9fa3-8ad28d08de3d.png)

For year 2018:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/93279134/194719118-7e953f1d-5180-4b53-87a9-cc651a81d267.png)


