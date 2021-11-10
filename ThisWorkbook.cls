VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------

Private Sub Workbook_Open()

Dim TodaysDay As Date
Dim EndOfMonth As Date

TodaysDay = Format(Now(), "dd/mm/yyyy")

EndOfMonth = WorksheetFunction.WorkDay(WorksheetFunction.EoMonth(TodaysDay, 0), -1)

Debug.Print "NOW(): " & TodaysDay & " " & "ENDOFMONTH " & EndOfMonth

If TodaysDay = EndOfMonth Then
    MsgBox "You have to update PCM and PYM sheets"
Else
    Exit Sub
End If

End Sub

