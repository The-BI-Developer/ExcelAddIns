
Private Sub CancelButton_Click()
    Unload Me 'me = userform
End Sub



Private Sub GetFilePath_Click()
    Dim sPath As String
    Dim GetTheName As String
    
    StatusBox.Value = "Awaiting input from user"
    
    GetTheName = FileNameTextBox.Value
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .ButtonName = "Save in this folder"
        If .Show = -1 Then 'if ok is pressed
            sPath = .SelectedItems(1)
        End If
        
    End With

todaysDay = Format(Now(), "dd.mm.yyyy")

    If sPath <> "" Then ActiveWorkbook.SaveAs Filename:=sPath & "\" & GetTheName & " " & Format(Now(), "dd.mm.yyyy") & ".xlsx", FileFormat:=51
    
    Unload Me
    
End Sub

