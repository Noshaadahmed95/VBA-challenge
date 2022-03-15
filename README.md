# VBA-challenge
VBA Homework

Sub Ticker()

    Dim ticker_name As String
    Dim Name As String
    Dim i As Integer
    Dim Summary_Table_Row As Integer
    Dim ticker_total As Double
    Dim ticker_Volume As Double
    ticker_total = 0
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'MsgBox (Cells(Rows.Count, 1).End(xlUp).Row)

        
    Summary_Table_Row = 2
    

    For i = 2 To LastRow
    
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker_name = Cells(i, 1).Value
            Range("I" & Summary_Table_Row).Value = ticker_name
            Range("J" & Summary_Table_Row).Value = ticker_total
            Range("L" & Summary_Table_Row).Value = ticker_Volume
        
            Summary_Table_Row = Summary_Table_Row + 1
            
            ticker_total = 0
            ticker_Volume = 0
            
        Else 'Cells(i, 1).Value <> Cells(i + 1, 1).Value
        
        ticker_total = ticker_total + Cells(i, 3).Value
        ticker_Volume = ticker_Volume + Cells(i, 7).Value
        
        
        
    If Cells(2, 10).Value <= 0 Then
    	Cells(2, 10).Interior.ColorIndex = 3
    	ElseIf Cells(2, 10).Value > 0 Then
    	Cells(2, 10).Interior.ColorIndex = 4
    End If
        
        
        End If
 
     Next i
     
        

End Sub
