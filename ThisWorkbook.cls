VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Dim d As Date

'd = Format(Now(), "dd / mm / yyyy")

d = "30/01/2021"

Debug.Print d


If d = WorksheetFunction.EoMonth(d, 0) Then
    MsgBox "You have to update PCM and PYM sheets"
Else
    Exit Sub
End If



End Sub
