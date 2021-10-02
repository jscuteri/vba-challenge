Option Explicit

Sub Stocks()

MsgBox ("This is the Way")

'Dim ws As Worksheet
'For Each ws In Worksheets

'Insert header names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
'Designate types
    Dim Ticker As String
    
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    Dim Percent_Change As Double
    Percent_Change = 0
    
    Dim Open_Price As Double
    Open_Price = Cells(2, 3).Value
    
    Dim Total_Stock_Volumes As Variant
    Total_Stock_Volumes = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    Dim i As Long
    Dim x As Long
    
    Dim Worksheet_Count As Integer
    Worksheet_Count = ActiveWorkbook.Worksheets.Count
    
    Dim Final_Row As Long
    Final_Row = Cells(Rows.Count, 1).End(xlUp).Row
    
  '  ws.Activate
    
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Ticker = Cells(i, 1).Value
                Range("I" & Summary_Table_Row).Value = Ticker
             
                Yearly_Change = Cells(i, 6).Value - Open_Price
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                     If Open_Price = 0 Then
                        Percent_Change = 0#
                        Range("K" & Summary_Table_Row).Value = Percent_Change
                        
                     Else
                        Percent_Change = (Cells(i, 6).Value - Open_Price) / Open_Price
                        Range("K" & Summary_Table_Row).Value = Percent_Change

                     End If
            
                Total_Stock_Volumes = Total_Stock_Volumes + Cells(i, 7).Value
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volumes
                
                Total_Stock_Volumes = 0
            
                Summary_Table_Row = Summary_Table_Row + 1
                
                Open_Price = Cells(i + 1, 3).Value
            
            Else
            
                Total_Stock_Volumes = Total_Stock_Volumes + Cells(i, 7).Value
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volumes
            
            End If
        
        Next i
    
        For x = 2 To Cells(Rows.Count, 1).End(xlUp).Row
                If Cells(x, 10).Value < 0 Then
                Cells(x, 10).Interior.ColorIndex = 3
            
            Else
                Cells(x, 10).Interior.ColorIndex = 4
            
            End If
    
        Next x
    
    Range("J:J").NumberFormat = "0.00"
    Range("K:K").NumberFormat = "0.00%"
    Range("L:L").NumberFormat = "#,##0"
    
    Range("P2") = WorksheetFunction.Max(Range("K2:K" & Final_Row))
    Range("P3") = WorksheetFunction.Min(Range("K2:K" & Final_Row))
    Range("P4") = WorksheetFunction.Max(Range("L2:L" & Final_Row))

    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
    Range("P4").NumberFormat = "#,##0"

    ' Next
    
    End Sub

