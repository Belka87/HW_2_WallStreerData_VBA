Attribute VB_Name = "Module1"
Sub stock_data()
 'loop through all sheets
 ''''''''''''''''''''''''''
 
 
Dim ws As Worksheet

 For Each ws In Worksheets

   'add heading
   Cells(1, "J") = "Ticker"
   Cells(1, "K") = "Total Stock Volume"
   
 
   'set variables to hold value
   Dim Ticker_Symbol As String
   Dim Volume As Double
   Volume = 0
   Dim i As Long
   Dim Summary_Table_Row As Double
   Summary_Table_Row = 2
  
   'find the last row
   Dim Last_Row As Double
   Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 
   'loop through all ticker symbol
 
   For i = 2 To Last_Row
 
    'check if we are still within the same ticker symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
    'set ticker symbol
    Ticker_Symbol = Cells(i, 1).Value
    Cells(Summary_Table_Row, "J").Value = Ticker_Symbol
 
 
    'add volume
    Volume = Volume + Cells(i, 7).Value
    Cells(Summary_Table_Row, "K") = Volume
 
    'add one to the Summary Table Row
    Summary_Table_Row = Summary_Table_Row + 1
 
    'reset the volume column
    Volume = 0
 

   'if the next cell has the same ticker symbol
   Else

    Volume = Volume + Cells(i, 7).Value
 
 
    End If
 
   Next i
   
 Next ws
 

End Sub
