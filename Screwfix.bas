Attribute VB_Name = "Screwfix"
Sub UKDateFormatterByCol()


    Dim i As Long
    Dim Rng As Range
    Dim r1 As Long, r2 As Long, c1 As Long, colNo As Long
    Dim ColLetter As String
    
    r1 = 1
    
    c1 = 5 'convert letter E to 5 col index
    
    'We need to find the last row
    
    r2 = Cells(Rows.Count, 1).End(xlUp).Row 'rows.count will return all Excel rows!
    
    
    Set Rng = Range(Cells(r1, c1), Cells(r2, c1))
    
    
    With Rng
        .NumberFormat = "dd/mm/yyyy"
        
        For i = 1 To .Areas.Count
            .Areas(i).Formula = .Areas(i).Formula
        Next i
        
    End With
    
End Sub
    

