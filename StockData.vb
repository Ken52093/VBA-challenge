Sub loops()

'Set headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' Set initial variable for holding the ticker name
Dim Ticker_Name As String
Ticker_Name = " "

' Set an initial variable for holding the total per ticker name
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

' Set variables
Dim Open_Price As Double
Open_Price = 0
Dim Close_Price As Double
Close_Price = 0
Dim Delta_Price As Double
Delta_Price = 0
Dim Delta_Percent As Double
Delta_Percent = 0
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
Dim Lastrow As Long
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Long


Open_Price = Cells(2, 3).Value
For i = 2 To Lastrow
'The ticker symbol

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker_Name = Cells(i, 1).Value

'Yearly Change & 'Percernt Change

Close_Price = Cells(i, 6).Value
Delta_Price = Close_Price - Open_Price

If Open_Price <> 0 Then
   Delta_Percent = (Delta_Price / Open_Price) * 100

End If

'Total Stock Volume

Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value


'Print
Range("I" & Summary_Table_Row).Value = Ticker_Name
Range("J" & Summary_Table_Row).Value = Delta_Price
Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

'Print Color
If (Delta_Price > 0) Then
Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

ElseIf (Delta_Price < 0) Then
Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

End If

' Add 1 to the summary table row count
Summary_Table_Row = Summary_Table_Row + 1
Delta_Price = 0
Close_Price = 0
Open_Price = Cells(i + 1, 3).Value
Total_Stock_Volume = 0

Else
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value


End If

Next i


End Sub



