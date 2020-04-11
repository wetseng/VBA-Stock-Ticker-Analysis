Attribute VB_Name = "Module1"
Sub stockMarket():
    
    Dim ticker As String
    Dim volumn As Double
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    Dim gPincreaseT As String
    Dim gPdecreaseT As String
    Dim gPincrease As Double
    Dim gPdecrease As Double
    Dim gTvolumnT As String
    Dim gTvolumn As Double
      
    Dim lRow As Double
    Dim olRow As Double
    
    
    For Each ws In Worksheets
        
        '''''''''''''''''''''''''''''''''''' Loop Through ''''''''''''''''''''''''''''''''''''
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volumn"
        ws.Range("M1").Value = "Open Price"
        ws.Range("N1").Value = "Close Price"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volumn"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Initial var
        volumn = 0
        openPrice = 0
        closePrice = 0
        yearlyChange = 0
        percentChange = 0
        
        gPincrease = 0
        gPdecrease = 0
        
        lRow = 0
        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        olRow = 1
        
        For i = 2 To lRow
        
            ticker = ws.Cells(i, 1).Value
            volumn = volumn + ws.Cells(i, 7).Value
            
            If ticker <> ws.Cells(i - 1, 1).Value Then
                openPrice = ws.Cells(i, 3).Value
            End If
            
            ' Check if the next ticker still the same, if not output the result
            If ticker <> ws.Cells(i + 1, 1).Value Then
            
                'olRow = Cells(Rows.Count, 9).End(xlUp).Row
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                
                If openPrice <> 0 Then
                    percentChange = (closePrice - openPrice) / openPrice
                Else
                    percentChange = 0
                End If
                
                ''''''''''''''''''''''''''' CHALLENGES '''''''''''''''''''''''''''
                If (gPincrease < percentChange) Then
                    gPincrease = percentChange
                    gPincreaseT = ticker
                ElseIf (gPdecrease > percentChange) Then
                    gPdecrease = percentChange
                    gPdecreaseT = ticker
                End If
                
                If (gTvolumn <= volumn) Then
                    gTvolumn = volumn
                    gTvolumnT = ticker
                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                If (yearlyChange >= 0) Then
                    ws.Cells(olRow + 1, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(olRow + 1, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(olRow + 1, 10).Value = yearlyChange
                ws.Cells(olRow + 1, 11).Value = percentChange
                ws.Cells(olRow + 1, 11).Style = "Percent"
                ws.Cells(olRow + 1, 11).NumberFormat = "0.00%"
                
                ws.Cells(olRow + 1, 13).Value = openPrice
                ws.Cells(olRow + 1, 14).Value = closePrice
                
                ws.Cells(olRow + 1, 9).Value = ticker
                ws.Cells(olRow + 1, 12).Value = volumn
                
                olRow = olRow + 1 ' Next output row
                
                volumn = 0 ' reset
                yearlyChange = 0
                percentChange = 0
                
            End If
            
        Next i
        
    
        ws.Range("P2").Value = gPincreaseT
        ws.Range("Q2").Value = gPincrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
    
        ws.Range("P3").Value = gPdecreaseT
        ws.Range("Q3").Value = gPdecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("P4").Value = gTvolumnT
        ws.Range("Q4").Value = gTvolumn
        
        MsgBox ("Completed " & ws.Name)
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Next ws
    
End Sub


