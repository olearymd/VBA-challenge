Sub stocks()

    ' Create a script that loops through all the stocks for one year and outputs the following information:
    ' The ticker symbol
    ' Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' The total stock volume of the stock.
    ' Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    ' Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

    ' Initilize variables
    Dim lastRow As Double
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    Dim volume As Double
    Dim totalVolume As Double
    Dim row As Long
    Dim outputRow As Integer
    Dim greatestPercentageIncreaseTicker As String
    Dim greatestPercentageDecreaseTicker As String
    Dim greatestPercentageTotalVolumeTicker As String
    Dim greatestPercentageIncrease As Double
    Dim greatestPercentageDecrease As Double
    Dim greatestTotalVolume As Double
           
   For Each ws In Worksheets
           
        ' Determine last row in worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Provide headers for output colums
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize numeric variables to 0, since we will be doing math with them
        volume = 0
        totalVolume = 0
        greatestTotalVolume = 0
        greatestPercentageIncrease = 0
        greatestPercentageDecrease = 0
            
        ' Initilize output row
        outputRow = 2
        
        ' Get ticker and opening price before entering loop
         ticker = ws.Cells(2, 1).Value
         openPrice = ws.Cells(2, 3).Value
        
        ' Loop through each row in worksheet, skipping Header
        For row = 2 To lastRow
                                    
            ' Add volume for each row
            volume = ws.Cells(row, 7).Value
            totalVolume = totalVolume + volume
                    
            ' If the next row is not equivalent to the current row, then calculate closing values for current stock
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
                ' Get the closing price from current row
                closePrice = ws.Cells(row, 6).Value
                
                ' Print output values back to spreadsheet
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = closePrice - openPrice
                
                ' If close is lower than open, set cell color to red:
                If (closePrice - openPrice > 0) Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                ' If close is higher than open, set cell color to green:
                ElseIf (closingPrice - openPrice < 0) Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                End If
                
                ' Calculate percent change and output it to spreadsheet
                percentChange = (closePrice - openPrice) / openPrice
                ws.Cells(outputRow, 11).Value = percentChange
                ' Change output format to percent
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
                ' Test whether percentChange is greatest so far
                If (percentChange > greatestPercentageIncrease) Then
                    greatestPercentageIncrease = percentChange
                    greatestPercentageIncreaseTicker = ticker
                End If
                
                ' Test whether percentChange is least so far
                If (percentChange < greatestPercentageDecrease) Then
                    greatestPercentageDecrease = percentChange
                    greatestPercentageIDecreaseTicker = ticker
                End If
                
                ' Output total stock volume to spreadsheet
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Test whether volume is the highest so far
                If (totalVolume > greatestTotalVolume) Then
                    greatestTotalVolume = totalVolume
                    greatestTotalVolumeTicker = ticker
                End If
                
                ' Reset volume to 0 and totalVolume before proceeding to next stock re-entering loop
                volume = 0
                totalVolume = 0
            
                ' Get next stock ticker before re-entering loop
                ticker = ws.Cells(row + 1, 1).Value
            
                ' Get opening price for next stock before re-entering loop
                openPrice = ws.Cells(row + 1, 3).Value
                            
                outputRow = outputRow + 1
            End If
        Next row
        
        ' Provide headers for new output columns
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = greatestPercentageIncreaseTicker
        ws.Cells(3, 16).Value = greatestPercentageIDecreaseTicker
        ws.Cells(4, 16).Value = greatestTotalVolumeTicker
        ws.Cells(2, 17).Value = greatestPercentageIncrease
        ws.Cells(3, 17).Value = greatestPercentageDecrease
        ws.Cells(4, 17).Value = greatestTotalVolume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
    Next ws
    
End Sub



