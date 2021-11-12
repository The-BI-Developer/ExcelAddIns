Attribute VB_Name = "Email"

Public Sub MailWorkbookWithSig()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String, HeaderInsr As String
    Dim ItemSelect As Object
    
      'This is important as we need to create object model first
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
   
    HeaderInsr = Application.InputBox("What is the report's name?", "Date with Report Title")
    
    On Error GoTo attachmentissue
    
    With OutMail
        .Display
        .Subject = HeaderInsr & " - " & Format(Now(), "dd mmmm yyyy")
        .HTMLBody = "Hi,<br><br>Many thanks, have a good day!" & .HTMLBody 'This ".htmlbody" seems to do the signature trick
        .Attachments.Add ActiveWorkbook.FullName
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Exit Sub 'End procedure here to begin exception handling

attachmentissue:     MsgBox "The file is too big to be attached in the mail", vbExclamation
 
End Sub

