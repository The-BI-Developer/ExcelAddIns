VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveAsWithTime 
   Caption         =   "File name with current date"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "SaveAsWithTime.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveAsWithTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()
    Unload Me 'me = userform
End Sub



Private Sub GetFilePath_Click()
   
Dim sPath As String
Dim GetTheName As String
Dim wb As Workbook
Dim TodaysDate As String

    GetTheName = FileNameTextBox.Value
    TodaysDate = " " & Format(Now(), "dd.mm.yyyy")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .ButtonName = "Save in this folder"
        If .Show = -1 Then 'if ok is pressed
            sPath = .SelectedItems(1) & "\"
        End If
        
    End With
    

Set wb = ActiveWorkbook

    If sPath <> "" Then
    
        If NS = True Then
            wb.SaveAs Filename:=sPath & GetTheName & TodaysDate & ".xlsx", FileFormat:=xlWorkbookDefault
        ElseIf SS = True Then
            wb.SaveAs Filename:=sPath & GetTheName & TodaysDate & ".xlsb", FileFormat:=xlExcel12
        End If
        
    End If
        
    Unload Me
    
End Sub


Private Sub NS_Click()
 
End Sub

Private Sub SS_Click()

End Sub

