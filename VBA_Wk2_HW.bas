Attribute VB_Name = "Module1"
Sub StockMarket()

'Loop through all the worksheet
For Each ws In Worksheets

'Insert column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volumn"

'Create initial variables
    Dim Ticker_Symbol As String
    Dim Total_Volumn As Variant
        Total_Volumn = 0
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    

'Loop through all stocks
    For Row = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        

'Create the conditionals
        If ws.Cells(Row, 1).Value <> ws.Cells(Row - 1, 1).Value Then
            Ticker_Symbol = ws.Cells(Row, 1).Value
            OpenPrice = ws.Cells(Row, 3).Value
            Total_Volumn = Total_Volumn + ws.Cells(Row, 7).Value
         End If
         If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
            ClosePrice = ws.Cells(Row, 6).Value
            Yearly_Change = ClosePrice - OpenPrice
            If OpenPrice <> 0 Then
                Percent_Change = Yearly_Change / OpenPrice
            Else
                Percent_Change = 0
            End If
            
        
'Print outcomes to the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            If Yearly_Change < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            ElseIf Yearly_Change > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            End If
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("L" & Summary_Table_Row).Value = Total_Volumn

'RESET
        Summary_Table_Row = Summary_Table_Row + 1
        Total_Volumn = 0
    
'if the cell immediatelly following a row is the same ticker...
        Else
            Total_Volumn = Total_Volumn + ws.Cells(Row, 7).Value

        End If
    Next Row
    
    
Next ws
End Sub
 


