VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()
    Dim MName As String, DNum As String
    Dim SelectCells As Range
    Dim xCell As Object
    
    MName = Format(Now(), "mmm")
    
    Sheets(MName).Activate
    
    DNum = (Format(Now(), "d"))
    
    Debug.Print DNum
    
For Each xCell In ActiveSheet.UsedRange.Cells
    If xCell.Value = DNum Then
        If SelectCells Is Nothing Then
            Set SelectCells = Range(xCell.Address)
        Else
            Set SelectCells = Union(SelectCells, Range(xCell.Address))
        End If
    End If

Next xCell

'select the cells with specified value

SelectCells.Select

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim sh As Worksheet

For Each sh In ActiveWorkbook.Sheets
    
    If sh.Name Like "*Overall*" Then
    
        sh.Activate
        Exit For
        
    Else
    
        MsgBox "There is no sheet called 'SOP'" _
        & vbCrLf & vbCrLf & "Was this sheet possibly renamed?", vbExclamation
    
        Exit Sub
        
    End If

Next sh


End Sub

Private Sub Workbook_Open()

    MsgBox "This workbook will automatically select the correct month, and highlights the date." _
    & vbCrLf & vbCrLf & "All worksheet names must be in 'mmm' format!", vbInformation

End Sub


