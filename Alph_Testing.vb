Option Explicit

Sub Ticker_Count()

'Variable Declaration
Dim Ticker As String
Dim Volume As Double
    Volume = 0
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim Percent_change As Double
Dim Total_Stock_Volume As Integer

Dim RowCount As Long
Dim i As Long

Dim increase_number As Double
Dim decrease_number As Double
Dim volume_number As Double

Dim Ws As Worksheet


'Set header
    For Each Ws In Worksheets
    Ws.Cells(1, 9).Value = "Ticker"
    Ws.Cells(1, 10).Value = "Yearly Change"
    Ws.Cells(1, 11).Value = "Percent Change"
    Ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Total_Stock_Volume = 2
    
    RowCount = Ws.Cells(Rows.Count, "A").End(xlUp).Row
    
'start loop
    
    
    For i = 2 To RowCount

      If year_open = 0 Then
          year_open = Ws.Cells(i, 3).Value
      End If

      If Ws.Cells(i - 1, 1) = Ws.Cells(i, 1) And Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
          year_close = Ws.Cells(i, 6).Value
          yearly_change = year_close - year_open
         
         Percent_change = yearly_change / year_open
          
          Ticker = Ws.Cells(i, 1).Value
          Volume = Volume + Cells(i, 7).Value
  
  'Find the values
  
          Ws.Range("I" & Total_Stock_Volume).Value = Ticker
          Ws.Range("j" & Total_Stock_Volume).Value = yearly_change
          Ws.Range("K" & Total_Stock_Volume).Value = Percent_change
          Ws.Range("L" & Total_Stock_Volume).Value = Volume

          Total_Stock_Volume = Total_Stock_Volume + 1

          Volume = 0


      Else

          Volume = Volume + Ws.Cells(i, 7).Value
          
   
          End If
Next i


'color part

Dim rg As Range
Dim j As Long
Dim column As Long
Dim color_cell As Range

    
    Set rg = Ws.Range("J2", Ws.Range("J2").End(xlDown))
    column = rg.Cells.Count
    Ws.Range("K2:K" & column).Style = "Percent"
 'loop
    
    For j = 2 To column
    'Set color_cell = rg(j)
    'Select Case color_cell
       If Ws.Cells(j, 10).Value >= 0 Then
            Ws.Cells(j, 10).Interior.Color = vbGreen
            'End With
       Else 'Ws.Cells(j, 11).Value < 0 Then
            Ws.Cells(j, 10).Interior.Color = vbRed
            'End With
       'End Select
       End If
       
    Next j

   
'Challenges



'take the max and min and place them in a separate part in the worksheet
Ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100
Ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & RowCount)) * 100
Ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & RowCount))

' returns one less because header row not a factor
increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Ws.Range("K2:K" & RowCount)), Ws.Range("K2:K" & RowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Ws.Range("K2:K" & RowCount)), Ws.Range("K2:K" & RowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Ws.Range("L2:L" & RowCount)), Ws.Range("L2:L" & RowCount), 0)

'final ticker symbol for  total, greatest % of increase and decrease, and average
Ws.Range("P2") = Ws.Cells(increase_number + 1, 9)
Ws.Range("P3") = Ws.Cells(decrease_number + 1, 9)
Ws.Range("P4") = Ws.Cells(volume_number + 1, 9)


'set header

Ws.Cells(1, 16).Value = "Ticker"
Ws.Cells(1, 17).Value = "Value"
Ws.Cells(2, 15).Value = "Greatest & Increase"
Ws.Cells(3, 15).Value = "Greatest & Decrease"
Ws.Cells(4, 15).Value = "Greatest Total Volume"

Next Ws
End Sub