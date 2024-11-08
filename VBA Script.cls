VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 Sub stockloop()
    
    For Each ws In Worksheets
    
            'Varibles
            Dim stockname As String 'holds stock name
            Dim stockvol As Double 'holds stock total volume'
            stockvol = 0
            
            Dim stockrow As Integer 'tracks stock in the table
            stockrow = 2
            
            rowsincolumn = ws.Cells(Rows.Count, 1).End(xlUp).Row 'count the rows in the stock name column
            
            Dim openprice As Double
            Dim quarterlychange As Double
            Dim percentchange As Double
            openprice = ws.Cells(2, 3).Value ' starting opening price
            
            '1st table headers
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            '2nd table headers
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            
            '2nd table columns
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            
            'Quarterly change (Close price-open price)
            'Percent Change((Close Price-Open Price)/Open Price * 100)
            
            For i = 2 To rowsincolumn
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    stockname = ws.Cells(i, 1).Value ' stock name
                    stockvol = stockvol + ws.Cells(i, 7).Value 'stock volume
                    closeprice = ws.Cells(i, 6).Value
                    quarterlychange = closeprice - openprice
                    ws.Range("I" & stockrow).Value = stockname 'places stock name in 1st table
                    ws.Range("L" & stockrow).Value = stockvol 'places stock volume in 1st table
                    ws.Range("J" & stockrow).Value = quarterlychange 'places stock quarterlychange in 1st table
                    
                    If openprice = 0 Then
                        percentchange = 0
                    Else
                        percentchange = quarterlychange / openprice
                        
                    End If
                    
                    ws.Range("K" & stockrow).Value = percentchange
                    ws.Range("K" & stockrow).NumberFormat = "0.00%"
                    
                    
          'reset stock row,volume, and open price
                    stockrow = stockrow + 1
                    stockvol = 0
                    openprice = ws.Cells(i + 1, 3)
                Else
                    stockvol = stockvol + ws.Cells(i, 7).Value
                End If
            Next i
            
            'conditional formatting
            stockrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            For i = 2 To stockrow
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
            Next i
            
           'places data in 2nd table
            For i = 2 To rowsincolumn
                'max percent change
                If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & rowsincolumn)) Then
                        ws.Cells(2, 16).Value = Cells(i, 9).Value
                        ws.Cells(2, 17).Value = Cells(i, 11).Value
                        ws.Cells(2, 17).NumberFormat = "0.00%"
                        
                   'min percent change
                ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & rowsincolumn)) Then
                        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                        ws.Cells(3, 17).NumberFormat = "0.00%"
                     
                    'min volume change
                ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & rowsincolumn)) Then
                        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                        
                End If
                
            Next i
            
         Next
         
        End Sub
        
