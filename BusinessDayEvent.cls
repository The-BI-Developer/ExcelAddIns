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

Dim lngEoM As Long

On Error GoTo handler

lngEoM = Date - Day(Date)
Debug.Print Day(Date)
 
If Date = lngEoM + Choose(Weekday(lngEoM, vbMonday), 0, 0, 0, 0, 0, 2, 1) Then MsgBox "Update the PCM and PYM sheets. I will take you to PCM sheet.", vbInformation

Sheets("U - PCM").Activate

Exit Sub 'it is just a placeholder, and if-else. This should rectify.

handler:
    MsgBox "There is no sheet named U - PCM! Was it renamed? E.g. Check whitespace characters.", vbExclamation
End Sub

