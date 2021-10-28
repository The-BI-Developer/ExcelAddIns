VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileNameFormatter 
   Caption         =   "File name with current date"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "FileNameFormatter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileNameFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()
    Unload Me 'me = userform
End Sub

Private Sub FileNameTextBox_Change()


End Sub

Private Sub GetFilePath_Click()
    Dim sFolder As String
    Dim GetTheName As String
    
    StatusBox.Value = "Awaiting input from user"
    
    GetTheName = FileNameTextBox.Value
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .ButtonName = "Get Folder Path"
        If .Show = -1 Then 'if ok is pressed
            sFolder = .SelectedItems(1)
        End If
        
    End With
    
    If sFolder <> "" Then
        ActiveWorkbook.SaveAs FileName:=sFolder & "\" & GetTheName _
        & " " & Format(Now(), "dd.mm.yyyy") & ".xlsx"
    End If
    
    StatusBox.Value = "Operation completed. You may close now."
    
    Unload Me
    
    
End Sub

Private Sub StatusBox_Change()

End Sub

Private Sub UserForm_Click()

End Sub
