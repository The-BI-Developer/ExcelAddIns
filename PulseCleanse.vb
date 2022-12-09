Sub PulseCleansing()

'append the data first


ImportPulseData

MsgBox "Importing pulse data operation has been completed", vbInformation

Dim ws As Worksheet

On Error Resume Next 'if there is autofilter applied
    ws.ShowAllData
On Error GoTo 0

Set ws = Workbooks(1).Worksheets("Data Output")
col = 33 'this is AG - extract column
col1 = 32 'date column (AF)



Application.DisplayAlerts = False

 
With ws.ListObjects("data_out").DataBodyRange

    .AutoFilter col, "Previous" 'filter to previous then change it to PC if in first week
    
    firstWeek = MsgBox("Filtered by 'Previous'" & vbCrLf & vbCrLf & "Rename 'Previous' to 'PC Month' instead of deleting 'Previous'?", vbQuestion + vbYesNo, "First week")
    
    If firstWeek <> vbYes Then
        .SpecialCells(xlCellTypeVisible).Delete 'deleting previous filtered visible rows
        
    Else
        .Columns(col).SpecialCells(xlCellTypeVisible).Value2 = Format(Now(), "mmmm yyyy") & " PC" 'need to revise this
    End If
    
    ws.ShowAllData

    'replace current for previous
    'replace any empty (new data) values for Current IN COL 33
    'REPLACE any empty values in col 32 for todays date
    'change column number
    .Columns(col).Replace what:="Current", replacement:="Previous", lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False, searchformat:=False, ReplaceFormat:=False
    .Columns(col).Replace what:=Empty, replacement:="Current", lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False, searchformat:=False, ReplaceFormat:=False
    
    'date column
    .Columns(col1).Replace what:=Format(Now(), "dd/mm/yyyy"), replacement:="Current", lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False, searchformat:=False, ReplaceFormat:=False
    


    
End With

Application.DisplayAlerts = True

If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then ws.ShowAllData 'avoidingerrors

MsgBox "Operation is complete", vbOKOnly + vbInformation

End Sub
Sub ImportPulseData()

Dim pulseFile As Workbook, myWb As Workbook
Dim pulseRng As Range, dataRng As Range
Dim i As Long, j As Long
Dim f As FileDialog

Set f = Application.FileDialog(msoFileDialogFilePicker)

With f
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx?", 1
    .AllowMultiSelect = False
    .InitialFileName = "C:\Users\AWasay\Downloads"
    
    If .Show = vbFalse Then Exit Sub  '.show is actually a button -1 for vbtrue and 0 for vbfalse
    pth = .SelectedItems(1)
    Debug.Print pth

    
End With

Set myWb = Workbooks(1)
Set pulseFile = Workbooks.Open(pth)

Application.ScreenUpdating = False

i = getRowno(pulseFile.Sheets(1).Range("A4")) 'keyedin databody range starts at a4
Set pulseRng = pulseFile.Sheets(1).Range("a4:v" & i)

'get rows and set data range for pulse 2.1
myWb.Activate 'will this not result in error then?
j = getRowno(myWb.Sheets("Data Output").Range("A2"))
Set dataRng = myWb.Sheets("Data Output").Range("a" & j)

Debug.Print "i: " & i; "--" & "j: " & j

'define fully
pulseRng.Copy dataRng.Offset(1, 0)


pulseFile.Close
myWb.Activate

Application.ScreenUpdating = True



End Sub
Function getRowno(r As Range) As Long

getRowno = r.End(xlDown).Row

Debug.Print "Function returns - " & getRowno

End Function


