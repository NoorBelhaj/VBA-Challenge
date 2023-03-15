Attribute VB_Name = "Module1"
Sub Ticker()

' First Sub for the Tiker 

    Dim Row, NextRow, TotalRecords As Integer
   
    
    NextRow = 1
    
     
   TotalRecords = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox (TotalRecords)
    
        
    For Row = 1 To TotalRecords
                
        If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
            Cells(NextRow, 9).Value = Cells(Row, 1).Value
              ' MsgBox (Cells(Row, 1).Value)
              NextRow = NextRow + 1
                
        End If
                   
    
   Next Row
   
End Sub

Sub YearlyPricingChange()


    Dim Opening, Closing, Variation, TotalStockVolume, Row, RowCount, TotalRecords As Long
    
    
    RowCount = 2
     
   TotalRecords = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (TotalRecords)
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "YearlyChange"
    Cells(1, 11).Value = "% Change"
    Cells(1, 12).Value = "TotalStockVolume"
    Opening = Cells(2, 3).Value
    TotalStockVolume = 0
     
    For Row = 2 To TotalRecords
    
            TotalStockVolume = TotalStockVolume + Cells(Row, 7).Value
            
            If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
              Cells(RowCount, 9).Value = Cells(Row, 1).Value
              Cells(RowCount, 10).Value = Cells(Row, 6).Value - Opening
              Cells(RowCount, 10).Interior.Color = RGB(0, 256, 0)
              If Cells(RowCount, 10).Value < 0 Then
                    Cells(RowCount, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                    Cells(RowCount, 10).Interior.Color = RGB(0, 255, 0)
                End If
              Cells(RowCount, 11).Value = (Cells(Row, 6).Value) / Opening - 1
              Cells(RowCount, 11).NumberFormat = "0.00%"
                            
              Cells(RowCount, 12).Value = TotalStockVolume
              Cells(RowCount, 12).NumberFormat = "0"
              RowCount = RowCount + 1
              Opening = Cells(Row + 1, 3).Value
              
            End If
            
            Next Row
  
End Sub

Sub YearlyPricingChangewithMaxMinVar()


    Dim Opening, Closing, Variation, Greatest_Increase, Greatest_Decrease, VolumeVar As Long
    Dim TotalStockVolume, Greatest_Tot_Vol As LongLong
    Dim Row, RowCount, TotalRecords As Double
    Dim MaxTicker, MinTicker, VolTicker As String
        
    RowCount = 2
     
   TotalRecords = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (TotalRecords)
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "YearlyChange"
    Cells(1, 11).Value = "% Change"
    Cells(1, 12).Value = "TotalStockVolume"
    Opening = Cells(2, 3).Value
    TotalStockVolume = 0
    Greatest_Tot_Vol = 0
    
    
    VolumeVar = 0
    
    
     
    For Row = 2 To TotalRecords
    
            TotalStockVolume = TotalStockVolume + Cells(Row, 7).Value
            
            If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
              Cells(RowCount, 9).Value = Cells(Row, 1).Value
              Cells(RowCount, 10).Value = Cells(Row, 6).Value - Opening
              Cells(RowCount, 10).Interior.Color = RGB(0, 256, 0)
              If Cells(RowCount, 10).Value < 0 Then
                    Cells(RowCount, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                    Cells(RowCount, 10).Interior.Color = RGB(0, 255, 0)
                End If
              Cells(RowCount, 11).Value = (Cells(Row, 6).Value) / Opening - 1
              Variation = Cells(RowCount, 11).Value
              
              Cells(RowCount, 11).NumberFormat = "0.00%"
                            
              Cells(RowCount, 12).Value = TotalStockVolume
              Cells(RowCount, 12).NumberFormat = "0"
              RowCount = RowCount + 1
              Opening = Cells(Row + 1, 3).Value
              TotalStockVolume = 0
              
            End If
            
                
            Next Row
            
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            Cells(3, 15).Value = "Greatest%Increase"
            Cells(3, 17).Value = Greatest_Increase
            Cells(4, 15).Value = "Greatest%Decrease"
            Cells(4, 17).Value = Greatest_Decrease
            Cells(5, 15).Value = "Greatest Tot Volume"
           
            Greatest_Increase = 0
            Greatest_Decrease = 0
            Greatest_Tot_Vol = 0
            
            
            For Row = 2 To RowCount - 1
            
                If Cells(Row, 12).Value > Greatest_Tot_Vol Then
                    VolTicker = Cells(Row, 9).Value
                    Greatest_Tot_Vol = Cells(Row, 12).Value
                    End If
                    
                If Cells(Row, 11).Value > Greatest_Increase Then
                    MaxTicker = Cells(Row, 9).Value
                    Greatest_Increase = Cells(Row, 11).Value
                    ElseIf Cells(Row, 11).Value < Greatest_Decrease Then
                    MinTicker = Cells(Row, 9).Value
                    Greatest_Decrease = Cells(Row, 11).Value
                            
                End If
             Next Row
           
                      
            Cells(3, 16).Value = MaxTicker
            Cells(3, 17).Value = MaxTicker ' Application.WorksheetFunction.Max(Range("K:K")) ' Greatest_Increase
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(4, 16).Value = MinTicker
            Cells(4, 17).Value = Greatest_Decrease ' Application.WorksheetFunction.Min(Range("K:K"))
            Cells(4, 17).NumberFormat = "0.00%"
            Cells(5, 16).Value = VolTicker
            Cells(5, 17).Value = Greatest_Tot_Vol ' Application.WorksheetFunction.Max(Range("L:L"))
End Sub


Sub YearlyPricingChangewithMaxMinVarMultiYearVersion()


For Each ws In Worksheets
Dim WorksheetName As String

    Dim Opening, Closing, Variation, Greatest_Increase, Greatest_Decrease, VolumeVar As Long
    Dim TotalStockVolume, Greatest_Tot_Vol As LongLong
    Dim Row, RowCount, TotalRecords As Double
    Dim MaxTicker, MinTicker, VolTicker As String
        
    RowCount = 2
     
   TotalRecords = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "YearlyChange"
    ws.Cells(1, 11).Value = "% Change"
    ws.Cells(1, 12).Value = "TotalStockVolume"
    Opening = ws.Cells(2, 3).Value
    TotalStockVolume = 0
    Greatest_Tot_Vol = 0
    
    
    VolumeVar = 0
    
    
     
    For Row = 2 To TotalRecords
    
            TotalStockVolume = TotalStockVolume + ws.Cells(Row, 7).Value
            
            If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
              ws.Cells(RowCount, 9).Value = ws.Cells(Row, 1).Value
              ws.Cells(RowCount, 10).Value = ws.Cells(Row, 6).Value - Opening
              ws.Cells(RowCount, 10).Interior.Color = RGB(0, 256, 0)
              If ws.Cells(RowCount, 10).Value < 0 Then
                    ws.Cells(RowCount, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                    ws.Cells(RowCount, 10).Interior.Color = RGB(0, 255, 0)
                End If
              ws.Cells(RowCount, 11).Value = (ws.Cells(Row, 6).Value) / Opening - 1
              Variation = ws.Cells(RowCount, 11).Value
              
              ws.Cells(RowCount, 11).NumberFormat = "0.00%"
                            
              ws.Cells(RowCount, 12).Value = TotalStockVolume
              ws.Cells(RowCount, 12).NumberFormat = "0"
              RowCount = RowCount + 1
              Opening = ws.Cells(Row + 1, 3).Value
              TotalStockVolume = 0
              
            End If
            
                
            Next Row
            
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(3, 15).Value = "Greatest%Increase"
            ws.Cells(3, 17).Value = Greatest_Increase
            ws.Cells(4, 15).Value = "Greatest%Decrease"
            ws.Cells(4, 17).Value = Greatest_Decrease
            ws.Cells(5, 15).Value = "Greatest Tot Volume"
           
            Greatest_Increase = 0
            Greatest_Decrease = 0
            Greatest_Tot_Vol = 0
            
            
            For Row = 2 To RowCount - 1
            
                If ws.Cells(Row, 12).Value > Greatest_Tot_Vol Then
                    VolTicker = ws.Cells(Row, 9).Value
                    Greatest_Tot_Vol = ws.Cells(Row, 12).Value
                    End If
                    
                If ws.Cells(Row, 11).Value > Greatest_Increase Then
                    MaxTicker = ws.Cells(Row, 9).Value
                    Greatest_Increase = ws.Cells(Row, 11).Value
                    ElseIf ws.Cells(Row, 11).Value < Greatest_Decrease Then
                    MinTicker = ws.Cells(Row, 9).Value
                    Greatest_Decrease = ws.Cells(Row, 11).Value
                            
                End If
             Next Row
           
                      
            ws.Cells(3, 16).Value = MaxTicker
            ws.Cells(3, 17).Value = Greatest_Increase ' Application.WorksheetFunction.Max(Range("K:K"))
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(4, 16).Value = MinTicker
            ws.Cells(4, 17).Value = Greatest_Decrease ' Application.WorksheetFunction.Min(Range("K:K"))
            ws.Cells(4, 17).NumberFormat = "0.00%"
            ws.Cells(5, 16).Value = VolTicker
            ws.Cells(5, 17).Value = Greatest_Tot_Vol ' Application.WorksheetFunction.Max(Range("L:L"))

Next

End Sub


