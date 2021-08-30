# VBA Challenge Attempt

## Overview of Project

### Purpose and Background
The purpose of this project was to evaluate the performance of certain stocks of interested in 2017 versus 2018 and how refactor VBA code would improve the performance to analysis the data. An initial code was written to perform the analysis; the refactor was utilized to compare and contrast speed of analysis.

## Results and Code

### Results
For the original code, the time to complete the 2017 and 2018 analysis were as follows:

<img width="156" alt="Original Code Time for 2017" src="https://user-images.githubusercontent.com/88955412/131284356-75de46e1-a6d9-476f-83e8-5641fff00dfa.png">
<img width="156" alt="Original Code Time for 2018" src="https://user-images.githubusercontent.com/88955412/131284361-23ae67ac-1ff8-4a18-9726-2a0ae435c799.png">

The code was as follows:

Sub AllStocksAnalysis()

Dim startTime As Single
Dim endTime As Single

yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

'Format the output sheet on the "All Stocks Analysis" worksheet

Worksheets("All Stocks Analysis").Activate

Range("A1").Value = "All Stocks (" + yearValue + ")"

'Create a header row
Cells(3, 1) = "Ticker"
Cells(3, 2) = "Total Daily Volume"
Cells(3, 3) = "Return"

'Initialize an array of all tickers

Dim tickers(12) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

'Initialize variables for the starting price and ending price
Dim startingPrice As Double
Dim endingPrice As Double
    
'Activate the data worksheet

Sheets(yearValue).Activate

'Find the number of rows to loop over

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Loop through the tickers

For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
'Loop through rows in the data
Worksheets(yearValue).Activate
    For j = 2 To RowCount
        'Find total volume for the currect ticker
        If Cells(j, 1).Value = ticker Then
        
        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
        startingPrice = Cells(j, 6).Value
        
        End If
        
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
        endingPrice = Cells(j, 6).Value
        
        End If
        
        
        'Find starting price for the current ticker
        'Find ending price for the current ticker
    
    Next j
    'Output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
Worksheets("All Stocks Analysis").Activate
Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A3:C3").Font.Italic = True
Range("A3:C3").Font.Color = vbBlue
Range("A3:C3").Font.Size = 24

Range("B4:B15").NumberFormat = "$#,##0.00"
Range("C4:C15").NumberFormat = "0.00%"

Columns("A").AutoFit
Columns("B").AutoFit
Columns("C").AutoFit

dataRowStart = 4
dataRowEnd = 15

For i = dataRowStart To dataRowEnd

If Cells(i, 3) > 0 Then
    'Color the cell green
    Cells(i, 3).Interior.Color = vbGreen

ElseIf Cells(i, 3) < 0 Then
    'Color the cell red
    Cells(i, 3).Interior.Color = vbRed
    
Else
    'Clear the cell color
    Cells(i, 3).Interior.Color = xlNone
    
End If

Next i

End Sub

Unfortunately, when attempting to code the refactor code, I was unable to debug the following error

<img width="124" alt="Error Message" src="https://user-images.githubusercontent.com/88955412/131282160-2786a88c-be32-4403-b41d-fe0569eade02.png">

I attempted to better define the variables, concerned that *tickerIndex* was not properly assigned. I looked up ways to understand looping with a variable as an index. I believe the fact I was unable to successfully increase the *tickerVolumes* may have contributed to the output strictly showing the first ticker **AY** 

The code I attempted to use is as follows (and I know once the error is brought to my attention, I will be kicking myself but as I learn this new language, mistakes only help me learn...hopefully. And I suppose one upside was that I definitely Googled a lot and got a chance to explore coders and the language and syntax they use. 

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Sheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Variant
    
    'tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
       
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerIndex = tickers(i)
    tickerVolumes = 0
    
        ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
            If Cells(j, 1).Value = tickerIndex Then
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            
            End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            tickerStartingPrices = Cells(j, 6).Value
        
            End If
     
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
            If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            tickerEndingPrices = Cells(j, 6).Value
        
        End If
        
       

            '3d Increase the tickerIndex.
            
        Next j
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers
        Cells(4 + i, 2).Value = tickerVolumes
        Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
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
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

## Advantages and Disadvantages

### In General 
As definied in the challenge, refactoring is all about making code more efficient. At the heart of it, that is a great advantage because clean code makes it easier to read, especially if you have to go back to it after some time. However, the downside is that your original code may have worked for reasons you're not sure and then when you refactor, you may run into errors because you aren't changing the functionality, just cleaning it up and so if there was a seredipitous code, upon cleaning, you may not have quite some debugging to do (as I learned...)

### For VBA script 
I know VBA script is not quite the most robust coding language but it is a great stepping stone. Despite the efficiencies that refactoring VBA script provides, VBA is limited in it's capabilities and requires far more definitions than other coding languages which I can imagine would add a level of unncessary work while trying to clean up code.

Although there is not much to share, additional information can be found in the Resources folder
https://github.com/jaimiesj/stock-analysis/commit/33e92ac8feb3fa9d20a6b2095a9e76ffa3dc9987


