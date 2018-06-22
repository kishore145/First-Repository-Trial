'###########################################################################################################
'
'                   Unit 2 | Assignment - The VBA of Wall Street
'                       Author: Kishore Ramakrishnan
'                      Assignment Level : Hard/Challenger
'
' Data constraints:
'------------------
' Data must be sorted without blank lines between on ticker, then date
' Data should be arranged in the order from column A with header in row 1
' <ticker>  <date>  <open>  <high>  <low>   <close> <vol>
'
'
' Running instructions:
'----------------------
' Execute the subroutine worksheetLoop() to calculate stock information for all worksheets
'
'###########################################################################################################

' Setting header for consolidated stock calculation and percentage format.
Sub setColumnHeader()
    'Set Header value for moderate difficulty assignment
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Range("K:K").NumberFormat = "0.00%"             ' Set number format for percentage
    
    'Set Header value for challenger difficulty assignment
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest total volume"
    Range("Q2:Q3").NumberFormat = "0.00%"           ' Set number format for percentage
End Sub

'Main subroutine
Sub tickerCalculatorModerate()
 
    ' rowCounter stores the current row in being accessed
    ' tickerRowNum stores the row number of consolidated stock data
    Dim rowCounter, tickerRowNum As Long
    tickerRowNum = 2: rowCounter = 2
    
    ' volume stores the consolidated volume of current ticker, openingPrice stores the opening price of current
    ' ticker, closingPrice stores the closing price of current ticker, percentIncrease stores the percentage change
    ' between opening and closing price of current ticker
    Dim volume, openingPrice, closingPrice, percentIncrease, maxPercent, minPercent, maxVolume As Double
    volume = 0: openingPrice = 0: closingPrice = 0: yearlyChange = 0
    maxPercent = 0: minPercent = 0: maxVolume = 0

    ' Loop to go through all rows in the list till a blank like is obtained.
    Do While (Cells(rowCounter, 1) <> "")
        
        ' Find the opening entry of each ticker and assign it to variable openingPrice
        If (Cells(rowCounter, 1).Value <> Cells(rowCounter - 1, 1).Value) Then      ' Find opening value by comparing ticker value with prior row to find first entry of ticker
            openingPrice = Cells(rowCounter, 3)
        End If
        
        ' Find the last entry of each ticker and complete all calculations like
        ' closingPrice, volume, yearlyChange, percentIncrease
        If (Cells(rowCounter + 1, 1).Value <> Cells(rowCounter, 1).Value) Then      ' Find closing value by comparing ticker value with next row to find first entry of ticker
            
            Cells(tickerRowNum, 9).Value = Cells(rowCounter, 1)             ' Set current ticker value to consolidated table
            volume = volume + Cells(rowCounter, 7)                          ' calculating final volume
            Cells(tickerRowNum, 12).Value = volume                          ' Assigning final volume to consolidated table
            closingPrice = Cells(rowCounter, 6)                             ' calculating closing Price and assigning it to variable
            yearlyChange = closingPrice - openingPrice                      ' calculating yearly Change by subtracting opening price from closing price
            Cells(tickerRowNum, 10).Value = yearlyChange                    ' Assiging yearly Change to consolidated table
                        
            ' Calculating percentage increase with exception handling when openingPrice is 0
            If (openingPrice = 0) Then
                percentIncrease = 0
                Cells(tickerRowNum, 11).Value = "Cannot calculate as opening price is 0"
            Else
                percentIncrease = yearlyChange / openingPrice               ' calculating perccent Increase
                Cells(tickerRowNum, 11).Value = percentIncrease             ' Assigning percentIncrease to consolidated table
            End If
           
            
            'Setting conditional formatting for yearlyChange
            If (yearlyChange < 0) Then
                Cells(tickerRowNum, 10).Interior.ColorIndex = 3             ' Change color of cell to red if yearly change is negative
            Else
                Cells(tickerRowNum, 10).Interior.ColorIndex = 4             ' Change color of cell to green if yearly change is greater than or equal to zero
            End If
                     
            'Calculating max volume by comparing current volume to maxvolume and changing if current volume is greater than max volume
            If (volume >= maxVolume) Then
                maxVolume = volume
                Cells(4, 16).Value = Cells(rowCounter, 1).Value             ' Storing Ticker to Consolidated Table
                Cells(4, 17).Value = maxVolume                              ' Storing maxVolume to consolidated table
            End If
            'Calculating min Percent
            If (percentIncrease <= minPercent) Then
                minPercent = percentIncrease
                Cells(3, 16).Value = Cells(rowCounter, 1).Value             ' Storing max percentage ticker to consolidated table
                Cells(3, 17).Value = minPercent                             ' storing min percentage to consolidated table
            End If
            'Calculating max Percentt
            If (percentIncrease >= maxPercent) Then
                maxPercent = percentIncrease
                Cells(2, 16).Value = Cells(rowCounter, 1).Value            ' Storing min percentage ticker to consolidated table
                Cells(2, 17).Value = maxPercent                             ' storing max percentage to consolidated table
            End If
            
            volume = 0: openingPrice = 0: closingPrice = 0                  ' Resetting volume, openingPrice & closingPrice to zero for next iteration.
            tickerRowNum = tickerRowNum + 1                                 ' Increasing the row count for consolidated table
        
        ' Continue looping till ticker changes and add to volume
        Else
            volume = volume + Cells(rowCounter, 7)                          ' Add volume information
        End If
        
        rowCounter = rowCounter + 1                                         ' Increase rowCounter by 1
        Loop
        
            
                          
End Sub

' Challenger scenario - Create a loop to go through all worksheets in the current workbook using worksheet object type.
Sub worksheetLoop()

    ' Declare Current as a worksheet object variable.
        Dim Current As Worksheet
        ' Loop through all of the worksheets in the active workbook.
        For Each Current In Worksheets
            Current.Activate                ' Activating the worksheet
            setColumnHeader                 ' Calling subroutine to set static text
            tickerCalculatorModerate        ' Calling subroutine to calculate and set consolidated ticker values.
        Next

End Sub

