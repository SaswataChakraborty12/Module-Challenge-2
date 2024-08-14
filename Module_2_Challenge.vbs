Sub Multiple_Year_Stock_Data()

' Define variables
    Dim tickerSymbol As String
    Dim stockOpen As Double
    Dim stockClose As Double
    Dim stockVolume As Double
    Dim lastRow As Long
    Dim SummaryRow As Long
    Dim GreatestInc As Double
    Dim GreatestDec As Double
    Dim GreatestVol As Double
    Dim GreatestIncTckr As String
    Dim GreatestDecTckr As String
    Dim GreatestVolTckr As String
        
' Loop through each row
    For Each ws In Worksheets
    
' Find the last row
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

' Start variables at initial values
        SummaryRow = 2
        GreatestInc = 0
        GreatestDec = 0
        GreatestVol = 0
        
          
' Generate the summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
' Define opening price of the first ticker symbol
        stockOpen = ws.Cells(2, 3).Value
    
' Loop through the data and set sum of stock volume for each ticker
        For i = 2 To lastRow
            tickerSymbol = ws.Cells(i, 1).Value
            stockVolume = stockVolume + ws.Cells(i, 7).Value
        
' Look for same value immediately next to the first value, and if not
            If ws.Cells(i + 1, 1) <> tickerSymbol Then
                stockClose = ws.Cells(i, 6).Value
                ws.Cells(SummaryRow, 9).Value = tickerSymbol
                ws.Cells(SummaryRow, 10).Value = stockClose - stockOpen
                
' When the opening price is 0
                If stockOpen = 0 Then

' And If the closing price is also 0, set 0% to the total percent change
                    If stockClose = 0 Then
                        ws.Cells(SummaryRow, 11).Value = 0

' If the closing price is greater to 0,
                    Else
                        ws.Cells(SummaryRow, 11).Value = stockClose / stockClose
                    End If
' If the total opening price is greater than 0, calculate percent change
                Else
                    ws.Cells(SummaryRow, 11).Value = stockClose / stockOpen - 1
                End If
                
' Format the percent change
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
' Get stock volume value in the summary table
                ws.Cells(SummaryRow, 12).Value = stockVolume
            
                
' Determine greatest volume  value in the variables
                If ws.Cells(SummaryRow, 12).Value > GreatestVol Then
                    GreatestVol = ws.Cells(SummaryRow, 12).Value
                    GreatestVolTckr = tickerSymbol
                End If
                
' Determine greatest percent increase value in the variables
                If ws.Cells(SummaryRow, 11).Value > GreatestInc Then
                    GreatestInc = ws.Cells(SummaryRow, 11).Value
                    GreatestIncTckr = tickerSymbol
                    
' Determine greatest percent increase value in the variables
                ElseIf ws.Cells(SummaryRow, 11).Value < GreatestDec Then
                    GreatestDec = ws.Cells(SummaryRow, 11).Value
                    GreatestDecTckr = tickerSymbol
                End If
            
' Conditional formatting of quarterly change
                               
                If ws.Cells(SummaryRow, 10) < 0 Then
                   ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(SummaryRow, 10) > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                End If
                
' If the next symbol is different, add one to the summary row and reset variables to 0.
                SummaryRow = SummaryRow + 1
                stockOpen = ws.Cells(i + 1, 3).Value
                stockClose = 0
                stockVolume = 0
        
            End If
        
        Next i
    
' Header for greatest volume, greatest increase and greatest decrease
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
    
    
' Print the greatest volume, greatest increase and greatest decrease in the short summary table
        ws.Cells(2, 16).Value = GreatestIncTckr
        ws.Cells(2, 17).Value = GreatestInc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = GreatestDecTckr
        ws.Cells(3, 17).Value = GreatestDec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = GreatestVolTckr
        ws.Cells(4, 17).Value = GreatestVol
        
                
    Next ws
 
End Sub

