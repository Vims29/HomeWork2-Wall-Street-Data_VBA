Sub Stockdata_hard()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    
    WS.Activate
      'Find the Last Row
      LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
      
      'Headings
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim TickerName As String
    Dim PercentChange As Double
    Dim StockVolume As Double
    
    Volume = 0
    Dim r As Double
    r = 2
    Dim c As Integer
    c = 1
    Dim i As Long
    
    'initial price
    
    OpenPrice = Cells(2, c + 2).Value
    ' Loop
    
    For i = 2 To LastRow
    
        If Cells(i + 1, c).Value <> Cells(i, c).Value Then
        
        TickerName = Cells(i, c).Value
        Cells(r, c + 8).Value = TickerName
        
        'Close Price
        
        ClosePrice = Cells(i, c + 5).Value
        YearlyChange = ClosePrice - OpenPrice
        Cells(r, c + 9).Value = YearlyChange
        
        'Find percent Change
        
        If (OpenPrice = 0 And ClosePrice = 0) Then
            PercentChange = 0
        
        ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
            PercentChange = 1
            
        Else
            PercentChange = YearlyChange / OpenPrice
            Cells(r, c + 10).Value = PercentChange
            Cells(r, c + 10).NumberFormat = "0.00%"
        
        End If
        
        Volume = Volume + Cells(i, c + 6).Value
        Cells(r, c + 11).Value = Volume
        
        r = r + 1
        
        'reset open price
        
        OpenPrice = Cells(i + 1, c + 2)
        
        Volume = 0
        
        Else
        
        Volume = Volume + Cells(i, c + 6).Value
        
        End If
        
    Next i
        
        'Locate the last row of Yearly change
        
        YChangeLastRow = WS.Cells(Rows.Count, c + 8).End(xlUp).Row
        
        'Cell Colors
        
    For j = 2 To YChangeLastRow
            If (Cells(j, c + 9).Value > 0 Or Cells(j, c + 9).Value = 0) Then
                Cells(j, c + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, c + 9).Value < 0 Then
                Cells(j, c + 9).Interior.ColorIndex = 3
            End If
       
            
    Next j
    
    
    'Determine Greatest % increase,,% decrease and Total volume
    
   Cells(2, c + 14).Value = "Greatest % increase"
   Cells(3, c + 14).Value = "Greatest % decrease"
   Cells(4, c + 14).Value = "Greatest Total Volume"
   Cells(1, c + 15).Value = "Ticker"
   Cells(1, c + 16).Value = "Value"
   
   'Search each row to find greatest value and its ticker
   
   For Z = 2 To YChangeLastRow
   
        If Cells(Z, c + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YChangeLastRow)) Then
            Cells(2, c + 15).Value = Cells(Z, c + 8).Value
            Cells(2, c + 16).Value = Cells(Z, c + 10).Value
            Cells(2, c + 16).NumberFormat = "0.00%"
        ElseIf Cells(Z, c + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YChangeLastRow)) Then
            Cells(3, c + 15).Value = Cells(Z, c + 8).Value
            Cells(3, c + 16).Value = Cells(Z, c + 10).Value
            Cells(3, c + 16).NumberFormat = "0.00%"
        ElseIf Cells(Z, c + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YChangeLastRow)) Then
            Cells(4, c + 15).Value = Cells(Z, c + 8).Value
            Cells(4, c + 16).Value = Cells(Z, c + 11).Value
        End If
        
    Next Z
           
    
 Next WS
        
            
                
End Sub

