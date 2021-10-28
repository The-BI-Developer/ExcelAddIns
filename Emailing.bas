Attribute VB_Name = "Emailing"
'Is there a way to prepopulate a list?

Sub EmailCurrentBook()
Dim objEmail As Object
Dim objOutlook As Object
Dim x As String


Set objOutlook = CreateObject("Outlook.Application")


x = Application.ActiveWorkbook.FullName


'create email object

Set objEmail = objOutlook.CreateItem(olmailitem)

On Error GoTo NotSaved

With objEmail '(Change this)
    .display 'display message in outlook
    .attachments.Add (x)
End With

If x = "" Then
    
NotSaved:         MsgBox "You have not saved the file! Try again.", vbExclamation

Exit Sub

End If

'clear
Set objEmail = Nothing: Set objOutlook = Nothing

End Sub
