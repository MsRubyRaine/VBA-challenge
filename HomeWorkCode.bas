Sub StockMarket()
    
    Dim wsname As String
    Dim TickerName As String
    Dim OpenCloseChange As Double
    Dim PercentChange As Long
    Dim TotalStockVolume As Double
    Dim FirstRow As Double
    Dim LastRow As Double
    Dim VolumeTotal As Double
    

    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        OpenCloseChange = 0
        VolumeTotal = 0
        YearlyChange = 2
        FirstRow = 2
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
        
            For i = 2 To LastRow
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    TickerName = Cells(i, 1).Value
                    
                    
                    Range("I" & YearlyChange).Value = TickerName
                    Range("J" & YearlyChange).Value = OpenCloseChange
                    
                    If Cells(FirstRow, 3).Value = 0 Then
                            Cells(YearlyChange, 11) = 0
                            Cells(YearlyChange, 12) = 0
                            
                    YearlyChange = YearlyChange + 1
                    OpenCloseChange = 0
                    FirstRow = i + 1
                    VolumeTotal = 0
                            
                    Else
                        
                    
                    Cells(YearlyChange, 10) = Cells(i, 6) - Cells(FirstRow, 3)
                        If Cells(YearlyChange, 10) > 0 Then
                            Cells(YearlyChange, 10).Interior.Color = vbGreen
                            
                        Else: Cells(YearlyChange, 10).Interior.Color = vbRed
                        
                            End If
                        
                            
                    Cells(YearlyChange, 11) = (Cells(i, 6) - Cells(FirstRow, 3)) / Cells(FirstRow, 3)
                    Cells(YearlyChange, 11).Select
                    Selection.Style = "Percent"
                    
                    Cells(YearlyChange, 12) = (VolumeTotal + Cells(i, 7).Value)
                    
                        
                    YearlyChange = YearlyChange + 1
                    OpenCloseChange = 0
                    FirstRow = i + 1
                    VolumeTotal = 0
                        
                     End If
                
                Else
                    VolumeTotal = VolumeTotal + Cells(i, 7)
                    
                    
                    
               
                    
                    
            End If
             
            
           Next i
    

End Sub


