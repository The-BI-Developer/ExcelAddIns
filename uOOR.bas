Attribute VB_Name = "uOOR"
Public Sub Load_OOR()

OutstandingOrderReport.Show

SaveAsWithTime.Show
    
End Sub

Public Sub OOR_TransformAndSelect()

    Dim Rng As Range, FilterAnswer As Integer
    Dim sh As Worksheet
        
    'Find the worksheet containing IFS data sheet
    For Each sh In ActiveWorkbook.Worksheets
    
    If sh.Name Like "*IFS*" Then
        sh.Activate
        Exit For
    End If
        
    Next sh
    
    IFSDataTab
    
   
End Sub


Public Sub IFSDataTab()

Dim Rng As Range, FilterAnswer As Integer

On Error GoTo nocells

'filteration

    With ActiveSheet.Range("A1:T1")
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
ActiveSheet.Range("$A:$T").Copy (Sheets("Screwfix").Range("A1"))



With ActiveSheet
    .Range(Cells(2, 7), Cells(Rows.Count, 7).End(xlUp)).SpecialCells(xlCellTypeVisible).Value = "Released"
End With

If Cells(1, 7).Value <> "Order Status" Then Cells(1, 7).Value = "Order Status"

Exit Sub

nocells:
    MsgBox "There were no orders that didn't have a released status!", vbExclamation
    ActiveSheet.ShowAllData

End Sub


