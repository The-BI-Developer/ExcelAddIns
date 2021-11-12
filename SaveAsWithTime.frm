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
    
    GetTheName = FileNameTextBox.Value
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .ButtonName = "Save in this folder"
        If .Show = -1 Then 'if ok is pressed
            sPath = .SelectedItems(1)
        End If
        
    End With

todaysDay = Format(Now(), "dd.mm.yyyy")

    If sPath <> "" Then ActiveWorkbook.SaveAs Filename:=sPath & "\" & GetTheName & " " & Format(Now(), "dd.mm.yyyy") & ".xlsb", FileFormat:=50
    
    Unload Me
    
End Sub
