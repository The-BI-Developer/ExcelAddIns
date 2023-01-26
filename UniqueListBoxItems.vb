Private Sub UserForm_Initialize()

'Dimensioning variables
Dim unique As New Collection
Dim item As Range
Dim cell As Range

Set rng = ThisWorkbook.Worksheets("Data Output").Range("data_out[[Date Source]]")


For Each cell In rng

    On Error Resume Next
    unique.Add cell, CStr(cell)
    
Next cell

On Error GoTo 0

For Each item In unique
   lst.AddItem Format(item, "dd/mm/yy")
Next item




End Sub

Private Sub deleteRng_Click()

Dim DateToDelete, ws As Worksheet

Application.DisplayAlerts = False

Set ws = ThisWorkbook.Worksheets("Data Output")

On Error Resume Next 'if there is autofilter applied
    ws.ShowAllData
On Error GoTo 0

DateToDelete = lst.Text

With ws.ListObjects("data_out").DataBodyRange

    .AutoFilter 23, DateToDelete
     .SpecialCells(xlCellTypeVisible).Delete
    
End With

If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then ws.ShowAllData

Application.DisplayAlerts = True

End Sub
