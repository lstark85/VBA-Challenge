Attribute VB_Name = "Module1"
Sub Multi_Year_Data()

'Specify all objects
    Dim Ticker As String
    Dim Total_Stock_Vol As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Greatest_Perc_Increase As Double
    Dim Greatest_Perc_Decrease As Double
    Dim Greatest_Total_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim opening As Double
    Dim closing As Double

'Start the loop process
    For Each ws In Worksheets
    
    Summary_Table_Row = 2
    
    Total_Stock_Vol = 0

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    opening = ws.Cells(2, 3)
    
        For i = 2 To LastRow
        
            'Start counting ticker and stock volume
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            Ticker = ws.Cells(i, 1)
            
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7)
        
            closing = ws.Cells(i, 6)
           
           'Calculate yearly change
            Yearly_Change = closing - opening
            
            If opening = 0 Then
                Percent_Change = 0
                
            Else
                Percent_Change = Yearly_Change / opening * 100
                
            End If
            
            'create range for new data
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = (Percent_Change & "%")
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Vol
            
            'Calculate total stock volume
            Total_Stock_Vol = 0
       
            opening = ws.Cells(i + 1, 3)
            
             ' Conditional formatting-positive change in green and negative change in red
            If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
            
                Summary_Table_Row = Summary_Table_Row + 1
                
        Else
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7)
        
        End If
        
    Next i
    
    ' Create names for data
    
    ws.Range("I1") = "Ticker"
    ws.Range("I1").Columns.AutoFit
    ws.Range("P1") = "Ticker "
    ws.Range("P1").Columns.AutoFit
    ws.Range("J1") = "Yearly Change"
    ws.Range("J1").Columns.AutoFit
    ws.Range("K1") = "Percent Change"
    ws.Range("K1").Columns.AutoFit
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("L1").Columns.AutoFit
    ws.Range("Q1") = "Value"
    ws.Range("Q1").Columns.AutoFit
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("O4").Columns.AutoFit
    
Next ws


End Sub

Sub Greatest():

'Specify all objects
    Dim max1, gtv, min1 As Double
    Dim rng1 As Range
    Dim rng As Range
    Dim FndRng As Range
    Dim i, LastRow As Integer
    Dim ticker_min, ticker_max, ticker_total As String


    For Each ws In Worksheets
    
    'Set ranges for max, min, and total
        Set rng = ws.Columns(11)
        max1 = ws.Application.Max(rng)
        ws.Range("Q2") = max1
        
        min1 = ws.Application.Min(rng)
        ws.Range("Q3") = min1
        ws.Range("Q3").Columns.AutoFit
        
        Set rng1 = ws.Columns(12)
        gtv = ws.Application.Max(rng1)
        ws.Range("Q4") = gtv
    
    
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    For i = 2 To LastRow
        If ws.Cells(i, 11) = min1 Then
            ticker_min = ws.Cells(i, 9)
        End If
        
    
        If ws.Cells(i, 11) = max1 Then
            ticker_max = ws.Cells(i, 9)
            
        End If
            
        If ws.Cells(i, 12) = gtv Then
            ticker_total = ws.Cells(i, 9)
            
        End If
            
    Next i
    
    ws.Range("P3") = ticker_min
    
    ws.Range("P2") = ticker_max
    
    ws.Range("P4") = ticker_total
    
Next ws

End Sub

Sub Clear_data():

For Each ws In Worksheets

ws.Range("I:Q").Clear

Next ws


End Sub
