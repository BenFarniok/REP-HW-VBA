Sub EasyHW()
     
    Dim ticker As String
    Dim Volume_Total As Double
    Dim Summary_Table_Row As Integer
    
    'Set Summary table row and the total volume
    Volume_Total = 0
    Summary_Table_Row = 2
    'Loop through entire data set
    
    For i = 2 To 70926
        
        'Check if we are still in the same ticker if not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'set the ticker name
            ticker = Cells(i, 1)
           'Add to the ticker's total volume
            Volume_Total = Volume_Total + Cells(i, 7).Value
            ' Print the ticker symbol in the Summary Table
             Range("I" & Summary_Table_Row).Value = ticker
             'Print the total volume in the summary table
             Range("J" & Summary_Table_Row).Value = Volume_Total
             'Set to go to next summary table row
             Summary_Table_Row = Summary_Table_Row + 1
             'Reset the volume total
             Volume_Total = 0
        'If we are still in the same ticker
        Else
        
            Volume_Total = Volume_Total + Cells(i, 7)
        End If
        Next i
             
End Sub
