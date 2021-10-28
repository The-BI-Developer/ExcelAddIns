Attribute VB_Name = "OOR"

Public Sub OOR()

    Dim CarryOn As Integer, OpenOutlook 'for messagebox
    Dim NameWithCurrentDate As String, InYear As String
    Dim Rng As Range, my_FileName As Variant

    CarryOn = MsgBox("This macro is applicable to IFS data. For further information read SOP on OOR report.", vbYesNo, "IFS data transformation")

    
    If CarryOn = vbYes Then
       OOR_TransformAndSelect
    Else
        Exit Sub
    End If
    
    NameWithCurrentDate = "OOR " & Format(Now(), "dd.mm.yyyy")
      
    InYear = Application.InputBox("Provide a folder with year e.g. 1996", "Year Input", Type:=2)
    
    Path2File = "G:\Business Support\.Data Services\Reporting\Retail\Outstanding Order Report\" & InYear & "\" & NameWithCurrentDate & ".xlsx"
    
    ActiveWorkbook.SaveAs (TodayDate & Path2File)
    

            
End Sub
Public Sub OOR_TransformAndSelect()

    Dim Rng As Range, FilterAnswer As Integer
    Dim sh As Worksheet
    
 
    '------------------------------
        
    'Find the worksheet containing IFS data sheet
    
     For Each sh In ActiveWorkbook.Sheets
    
         If sh.Name Like "*IFS*" Then
           sh.Activate
           Exit For
         End If
        
    Next sh
    
    IFSDataTab
    '--------------------------------
    
    On Error Resume Next 'turn off error for input box
   
    For Each sh In ActiveWorkbook.Sheets
    
       If sh.Name Like "*OOS*" Then
         sh.Activate
         Exit For
       End If
      
    Next sh

End Sub


'This is hard coded with worksheet names!
Public Sub IFSDataTab()

Dim Rng As Range, FilterAnswer As Integer

'-0-----filteration

    With ActiveSheet.Range("$A$1:$T$50000")
    'Order type
        .AutoFilter Field:=2, Criteria1:="*SFW*"
    
    'Payer
        .AutoFilter Field:=17, Criteria1:="*Screwfix Direct*"
    
    'xlAnd is the key [update: works without it]. Note won't show tick in AUTOFILTER!
        .AutoFilter Field:=20, Criteria1:="<>*@*", Operator:=xlAnd
    
    'other than released
    
        .AutoFilter Field:=7, Criteria1:="<>Released", Operator:=xlAnd
    End With
    
     'Start applying Reserved status
'-o-----set range to released and copy that range to screwfix tab
    ActiveSheet.Range("A1:T50000").Copy (Sheets("Screwfix").Range("A1"))
    Set Rng = Application.InputBox("Select the range to change to Released", Type:=8)
    
    For Each cell In Rng
        If cell.Value <> "Released" Then
         Rng.Value = "Released"
        End If
    Next cell
    
 
    
    'Final prompt
    
    FilterAnswer = MsgBox("Check if months of column T match with column M." _
    & vbCrLf & vbCrLf & _
    "This macro is finished." _
    & vbCrLf & vbCrLf & _
    "Would you like to reset the filters?", vbQuestion + vbYesNo + vbDefaultButton2)
         
    If FilterAnswer = vbYes Then
       ActiveSheet.ShowAllData
       
        MsgBox "Filters have been reset, and this step of SOP is complete. You can press OK now", vbInformation
       
    Else
        MsgBox "Change months, and reset filter manually!" & vbCrLf & vbCrLf & "This macro is done.", vbCritical
        Exit Sub
    End If
    
    


End Sub


