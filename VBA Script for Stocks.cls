VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks()


For Each ws In ActiveWorkbook.Worksheets


Dim I As Long

Dim Ticker_Symbol As String

Dim Summary_Table_Row As Integer

Dim LastRow As Long

Dim Year_open As Double
Dim Year_close As Double
Dim Yearly_change As Double

Dim Year_start As Long

Dim Year_end As Long





Summary_Table_Row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Year_start = ws.Cells(2, 2).Value

Year_end = ws.Cells(LastRow, 2).Value





For I = 2 To LastRow



    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

        Ticker_Symbol = ws.Cells(I, 1).Value

        ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
    
        Summary_Table_Row = Summary_Table_Row
        
        End If
        
    
    
    
   If ws.Cells(I, 2).Value = Year_start Then
     
     Year_open = ws.Cells(I, 3).Value
     
     End If

    
    If ws.Cells(I, 2).Value = Year_end Then
     
     Year_close = ws.Cells(I, 6).Value
     
     Yearly_change = Year_close - Year_open
     
     ws.Range("K" & Summary_Table_Row).Value = Yearly_change
     

     ws.Range("L" & Summary_Table_Row).Value = Yearly_change / Year_open
     
     ws.Range("L:L").NumberFormat = "0.00%"
     
     End If
     
    
     If ws.Cells(I, 11).Value > 0 Then
     ws.Cells(I, 11).Interior.ColorIndex = 4
     
     Else
     
     If ws.Cells(I, 11).Value < 0 Then
     
     ws.Cells(I, 11).Interior.ColorIndex = 3
     
     End If
     
     
    End If
     
    
     
     If ws.Cells(I + 1, 1).Value = ws.Cells(I, 1).Value Then
     
     Running_total = Running_total + ws.Cells(I, 7).Value
     
     ws.Range("M" & Summary_Table_Row).Value = Running_total
     
     Else
     
     If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
     
     'Running_total = Running_total + ws.Cells(I, 7).Value
     
     'ws.Range("M" & Summary_Table_Row).Value = Running_total
      
    Summary_Table_Row = Summary_Table_Row + 1
    
     
     End If
     
     Running_total = 0
     
     End If
    
     
      If ws.Cells(I, 12).Value > Max Then
      
           Max = ws.Cells(I, 12).Value
           
           Ticker_Name = ws.Cells(I, 10).Value
           
           ws.Cells(2, 16).Value = Max
           
           ws.Cells(2, 15).Value = Ticker_Name
           
           ws.Cells(2, 16).NumberFormat = "0.00%"
           
           ElseIf ws.Cells(I, 12).Value < Min Then
           
           Min = ws.Cells(I, 12).Value
           
           Ticker_Name = ws.Cells(I, 10).Value
           
           ws.Cells(3, 16).Value = Min
           
           ws.Cells(3, 15).Value = Ticker_Name
           
           ws.Cells(3, 16).NumberFormat = "0.00%"
           
           End If
           
           If ws.Cells(I, 13) > MaxVol Then
           
           MaxVol = ws.Cells(I, 13).Value
           
           Ticker_Name = ws.Cells(I, 10).Value
           
           ws.Cells(4, 16).Value = MaxVol
           
           ws.Cells(4, 15).Value = Ticker_Name
           
           End If
           
           

Next I

Max = 0
Min = 0
MaxVol = 0

Next ws

End Sub


