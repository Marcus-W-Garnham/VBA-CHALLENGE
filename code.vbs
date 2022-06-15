
Sub Advanced()

Dim Ticker As String
Dim TotalRows As Long
Dim RowCount As Long
Dim Line As Integer
Dim First As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim PercentYearly As Double
Dim YearlyChange As Double
Dim StockTotal As Double
Dim GreatestInc As Double
Dim GreatestDec As Double
Dim GreatestVol As Double
Dim TotalRowsOutput As Long

' find number of rows
TotalRows = Cells(Rows.count, 1).End(xlUp).Row

' Setup headers for the outputs
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percentage Change"
Cells(1, 12) = "Total Stock Volume"

' Table for targeted or most intesting data
Cells(2, 14) = "Greatest Increase %"
Cells(3, 14) = "Greastest Decrease %"
Cells(4, 14) = "Greatest Total Volume"
Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"

'Make sure values empyty for Greatest
GreatestInc = 0
GreatestDec = 0
GreatestVol = 0

' Line/Row we will output are data on
Line = 2

' Setup row values we will count through for the data
RowCount = 2

' Gets first ticker we will be using
Ticker = Cells(RowCount, 1)

'Set First to 1 so we can use it as a switch
First = 1

    ' While will count through all the rows
    While RowCount < TotalRows
        
        ' Checks if the ticker is the same and is the first of its ticker
        If Ticker = Cells(RowCount, 1) And First = 1 Then
            
            ' Takes the open price and since data is already sorted by date then it will be first for that year
            OpenPrice = Cells(RowCount, 3)
            
            ' Fetches first closing price incase there is only one entry
            ClosePrice = Cells(RowCount, 6)
            
            ' Update the stock total volume
            StockTotal = StockTotal + Cells(RowCount, 7).Value
            
            ' Moves on to the next row
            RowCount = RowCount + 1
            
            'So it won't enter this if again until next ticker
            First = 0
            
        ' Just checks if the ticker is the same
        ElseIf Ticker = Cells(RowCount, 1) Then
            
            ' Constanly fetches the last closing price
            ClosePrice = Cells(RowCount, 6)
            
            ' Update the stock total volume
            StockTotal = StockTotal + Cells(RowCount, 7).Value
            
            ' Moves on to the next row
            RowCount = RowCount + 1
        
        
        ' When ticker is not the same
        Else
            ' Prints out the total stock for ticker we were looking at
            Cells(Line, 12).Value = StockTotal
            
            ' Prints the ticker out on same line as its stock value
            Cells(Line, 9).Value = Ticker
            
            ' Calculates the early change in the stock price
            YearlyChange = ClosePrice - OpenPrice
            
            ' Puts the yearly change into its cell for that ticker
            Cells(Line, 10).Value = YearlyChange
            
            ' Calculates yearly percent change and prints it out with percent sign which atomatically makes data go to 2 decimal places
            PercentYearly = ((YearlyChange / OpenPrice) * 100)
            Cells(Line, 11).Value = PercentYearly & "%"
                
              ' Checks if yearly change is appove zero/positive
                If YearlyChange > 0 Then
                ' Colours the cell green to show its doing good
                    Cells(Line, 10).Interior.ColorIndex = 4
                    
                ' checks if values are under zero/negative
                ElseIf YearlyChange < 0 Then
                ' colours the cell red to show its doing bad
                    Cells(Line, 10).Interior.ColorIndex = 3
                    
                ' Checks unlikely case that price has not changed
                ElseIf YearlyChange = 0 Then
                ' Colours the cell blue
                    Cells(Line, 10).Interior.ColorIndex = 8
                Else
                    ' empty else that probably unneeded but just incase of blank
                End If
                    
        
                   
        
            ' Move output to next line for the next ticker
            Line = Line + 1
        
            ' Set Ticker to the new Ticker so we can use the if above for it
            Ticker = Cells(RowCount, 1)
            
            ' Set the stock total to zero for the new ticker
            StockTotal = 0
            
            First = 1
            
        ' Is the end of the if and else statements
        End If
        
    ' Wend is the end of the while statement
    Wend
    
'Need length of new table
TotalRowsOutput = Cells(Rows.count, 9).End(xlUp).Row

' Line/Row we will output are data on
Line = 2

    ' Check values in range of new
    For RowCount = 2 To TotalRowsOutput
        
        ' if there is bigger value then replace and print
        If Cells(RowCount, 10) > GreastInc Then
        
            Cells(2, 16) = Cells(RowCount, 10) & "%"
            GreastInc = Cells(RowCount, 10)
            Cells(2, 15) = Cells(RowCount, 9)
        
        ' if there is smaller value then replace and print
        ElseIf Cells(RowCount, 10) < GreastDec Then
        
            Cells(3, 16) = Cells(RowCount, 10) & "%"
            GreastDec = Cells(RowCount, 10)
            Cells(3, 15) = Cells(RowCount, 9)
        
        ' if there is greater volume replace and print
        ElseIf Cells(RowCount, 12) > GreastVol Then
            Cells(4, 16) = Cells(RowCount, 12)
            GreastVol = Cells(RowCount, 12)
            Cells(4, 15) = Cells(RowCount, 9)
        
        Else
        'empty for everything else
        
        End If
    
    'next row
    Next RowCount