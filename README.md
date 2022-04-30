# stock-analysis
# Overview of Project

## Background

The main objective of this project is to refactor or edit, the Stock-Market Dataset with VBA solution code to loop through all the data one time in order to collect the same information that we did in this module.Then, we will determine wheather refactoring improve its efficiency and clearity or not.Refactoring is a key part of the coding process. When refactoring code, we are adding new functionality; and making the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Purpose

The main purpose of this analysis is to provide Steve a complete overview of green-energy stock market for his parents.Before investing their money in DAQO stocks,Steve wants to analyize DQ stock performance comparison to other green-energy stocks.The main purpose of this challenge is to refactor the original VBA code to loop only once, and determine whether refactoring VBA code successfully made the VBA script run faster.

## Results

### Analysis

In 2017,the green energy stock categories from this year tended to have a successful yearly return.Among 12 of them,only one stock "TERP" had a negative yearly return.The DQ stock had almost 200% positive yearly return.From this data, we can conclude that people probably made money off their stocks in 2017.

![](https://github.com/akthersr/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

In 2018,the majority of the stocks had a negative returns.The DQ stock had almost negative 63% yearly return. In this year there was a much higher chance for their stock to have gone down in value.These results indicate us that the stock trend is not stable for DQ stocks and the investment might be risky.

![](https://github.com/akthersr/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

### Refactor VBA Code

In this analysis both scripts had the same output.Only difference between them was in the execution time.At first,I have created tickerIndex and set it equal to zero before looping over the rows.Next,arrays are created for tickers,tickerVolumes,tickerStartingPrices and tickerEndingPrices.

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row's ticker doesn't match, increase the tickerIndex.
For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
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

Then,a for loop created to initialize the tickerVolumes to zero.If the next row's ticker doesn't match, increase the tickerIndex.We created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes and adds the ticker volume for the current stock ticker.

For loops are responsible for executing the code in a monotonous manner until the condition is fullfiled.Here,tickerIndex = tickerIndex + 1 to move to the next ticker and initialize the arrays by 
tickerVolumes(tickerIndex)= 0 before entering the loop again.At last,in order to make the table more visulize and pleasent the code cotain the fromatting syntex too.

    Range("A3:C3").Font.FontStyle = "Bold"
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            End If

!execution times of the original script()
![](https://github.com/akthersr/stock-analysis/blob/main/Resources/execution%20time%20of%202017.png)

!execution times of the refactored script()
![](https://github.com/akthersr/stock-analysis/blob/main/Resources/execution%20time%20of%202017%20refactor.png)


## Summary

Refactoring code was a extremely beneficial procees to explore finding alternative methods to a previously successful one.By refactoring we are making the code more efficient without adding no new functionality.

### Advantages

The main goal with refactoring was to improve readability, reduce complexity , increase speed, making it easier to maintain or extend.

It also takes fewer steps, uses less memory and  improves the logic of the code to make it easier for future users to read.

We can also debug different types of coding isues(duplicate code).

### Disadvantages

Refactoring could be time or money consuming and frastrating.If the code breaks while refactoring the dataset,it could be hard to find the root issue or debugging the errors.

 Having worked with both scripts, I do like the speed and efficiency of the refactored code,but requires more attention and effort.It decresed logarithimic response and gives quicker results.By removing the nested loops,the codes look more readable and simple.

 On the other hand the original script was also easier to follow the logic of the code,was functioning properly and execute the desired output in a decent amount of time. The new code ran much faster,and because of this, we could reuse the new code on a much larger data set. 
  














