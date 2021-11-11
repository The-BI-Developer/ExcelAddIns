Dim r As String
Dim Rng As Range
Dim Book1 As Workbook, Book2 As Workbook
Dim Book1Rows As Long, Book2Rows As Long 'uses number of orders as rowNumber
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub OkButton_Click()

    Dim colN As Long
    Dim rowN As Long
    Dim fnd As Variant, rplc As Variant
    Dim NowDate As Date
    
    Set Book1 = ActiveWorkbook 'CHANGE THIS
    Book1Rows = Book1.Sheets("Sales Tracker").Cells(Rows.Count, 1).End(xlUp).Row
    
    Debug.Print Book1Rows
    
    NowDate = Format(Now(), "dd/mm/yyyy") 'Work
    
    Debug.Print Book1Rows
  
    Set Rng = Book1.Sheets("Sales Tracker").Range(Cells(1, 16), Cells(Book1Rows, 18)) 'Change the argument only
    
    
    Range("A3").End(xlToRight).Activate
    
    colN = ActiveCell.Column
    rowN = ActiveCell.Row
    
    Rng.Copy (Cells(rowN - 2, colN + 1))
    
    With ActiveCell
        .Offset(-2, 1).Value = NowDate - (WorksheetFunction.Weekday(NowDate, 2))
        .Offset(1, 1).Activate
    End With
    
    'BOOK2
    r = ActiveCell.Address
    
     With Application.FileDialog(msoFileDialogFilePicker)
        .ButtonName = "Select One File"
        If .Show = -1 Then
            f = .SelectedItems(1)
        Else
            MsgBox "You didn't select any file. Exiting macro.", vbExclamation
            Exit Sub
        End If
    End With
    
    Set Book2 = Workbooks.Open(f)
    'Starting another sub routine
    
    Extract_Data
    
    'Replace any _ for 0
    
    fnd = "—"
    rplc = 0
    
    Book1.Sheets("Sales Tracker").Cells.Replace what:=fnd, replacement:=rplc, _
    searchformat:=False, ReplaceFormat:=False
    
    'Dynamic autofill
    AutoFillCalc
    
    Unload Me 'close userform
        
End Sub
    
Sub Extract_Data()
    
    Dim lookupVal As Range, srcArr As Range, retArr(1 To 2) As Range
    Dim r1 As Long
    
    'Count rows on both workbooks
    Book2Rows = Book2.ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set range to set for lookup range (note this is only for column G now!)
    
    With Book2.ActiveSheet
        Set srcArr = .Range(Cells(3, 1), Cells(Book2Rows, 1))
    
        Set retArr(1) = .Range(Cells(3, 7), Cells(Book2Rows, 7))
        Set retArr(2) = .Range(Cells(3, 12), Cells(Book2Rows, 12))
    End With
    
    For r1 = 1 To Book1Rows
    
        Set lookupVal = Book1.Sheets("Sales Tracker").Cells(r1 + 3, 1)
        
        With Book1.Sheets("Sales Tracker").Range(r)
        
            .Offset(r1 - 1, 0).Value _
            = Application.WorksheetFunction.XLookup _
            (lookupVal, srcArr, retArr(1), 0)
            
            .Offset(r1 - 1, 1).Value _
            = Application.WorksheetFunction.XLookup _
            (lookupVal, srcArr, retArr(2), 0)
        End With
        
    Next r1
    
    Book2.Close savechanges:=False


End Sub

Sub AutoFillCalc()
    Dim Adrs As String
    Dim ActR As Long, ActC As Long
    
    ActiveWorkbook.Sheets("Sales Tracker").Range("a3"). _
    End(xlToRight).Offset(1, 0).Activate
    
    ActR = ActiveCell.Row
    ActC = ActiveCell.Column
    
    Debug.Print "-Row: " & ActR & "-Col: " & ActC
    Set sourceRng = ActiveWorkbook.Worksheets("Sales Tracker").Range(Cells(ActR, ActC), Cells(ActR + 10, ActC))
    Set fillRng = ActiveWorkbook.Worksheets("Sales Tracker").Range(Cells(ActR, ActC), Cells(Book1Rows, ActC))
    
    sourceRng.AutoFill Destination:=fillRng
    
End Sub

