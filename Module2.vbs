Attribute VB_Name = "Module1"
' this subroutine will loop through all the stocks for one year and output
' the ticker symbol
' yearly change of the opening price of a stock in a given year to the closing price at the end of that year
' the percentage change of the opening price in a given year to the closing price at the end of the year
' the total volume of the stock

Sub Module2Challenge():
    For Each ws In Worksheets
        ' I1 should be "Ticker
        ws.Cells(1, 9).Value = "Ticker"
        ' J1 should be "Yearly Change"
        ws.Cells(1, 10).Value = "Yearly Change"
        ' K1 should be "Percent Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ' L1 should be "Total Stock Volume"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ' let's focus on the tickers first
        ' make a var called col, set to 1, so we focus on the column with the tickers
        Column = 1
    
        ' make a var for the row number in the tick column, which will increment to list the ticker values
        tick = 2
    
        ' make a var to hold the value of the last cell with data in this sheet - refer to class vb scripts
        ' specifically star_counter
        lastRow = ws.Cells(Rows.Count, Column).End(xlUp).Row
    
        ' make an openVal and set it to the first opening value
        openVal = Cells(2, 3).Value
    
        ' make a perChange as a double
        Dim perChange As Double
    
        ' start a for loop to go from the 2nd row (start of data) to last row
        For i = 2 To lastRow
            ' add all the values in row 7 to a variable that resets to 0 when the ticker changes
            totalVol = ws.Cells(i, 7).Value + totalVol
        
            ' check when the value of the next cell down is different from the value of the current cell
            ' next_cells vba script should have it
            If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then
                ' the value of the cell with column "ticker" & row "tick" should be the value of column 1, row i
                ws.Cells(tick, 9).Value = ws.Cells(i, Column).Value
            
                ' make a closeVal and set it to the closing value when the next row's ticker is not equal to this row's ticker
                closeVal = ws.Cells(i, 6).Value
                ' yearChange represents the Yearly Change, or the difference between a ticker's first opening value and that same ticker's last closing value in a given year
                yearChange = closeVal - openVal
                ' put this yearChange value into the column to correspond with the row with the ticker with that value
                ws.Cells(tick, 10).Value = yearChange
            
                ' do an if statement (remember to EndIf) - if the value is negative, then turn the cell red, and if it's positive, then turn the cell green
                If yearChange > 0 Then
                    ws.Cells(tick, 10).Interior.ColorIndex = 4
                ElseIf yearChange < 0 Then
                    ws.Cells(tick, 10).Interior.ColorIndex = 3
                End If
            
                ' perChange is equal to the difference between the last close and the first open, all divided by the first open val
                If openVal = 0 Then
                    perChange = 100 * yearChange
                Else
                    perChange = 100 * yearChange / openVal
                End If
                ' round perChange to 2 decimal places
                perChange = Round(perChange, 2)
                ' put this perChange value into the corresponding column
                ws.Cells(tick, 11).Value = perChange
            
                ' drop the totalVol value into column L (the 12th column) for each ticker
                ws.Cells(tick, 12).Value = totalVol
                ' reset totalVol to 0 so that it can now hold the sums of the Volumes for the next ticker
                totalVol = 0
            
                ' reset openVal so that it's now equal to the opening value of the ticker in the next cell
                openVal = Cells(i + 1, 3).Value
            
                ' increment tick variable
                tick = tick + 1
            End If
        Next i
    Next ws
    
End Sub
