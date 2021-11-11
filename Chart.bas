
Private Sub AllCharts_Click()

    Dim PowerPointApp As Object
    Dim myPresentation As Object
    Dim mySlide As Object
    Dim myShape As Object
    
    'Dim for loops
    
    Dim i As Integer 'for chartobjects

    If PowerPointApp Is Nothing Then _
    Set PowerPointApp = CreateObject(class:="PowerPoint.Application")
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    Set myPresentation = PowerPointApp.Presentations.Add
    
'Count number of worksheets and loop that many times
    For i = 1 To ActiveSheet.ChartObjects.Count
    
        
        Set mySlide = myPresentation.Slides.Add(1, 11) '11 = ppLayoutTitleOnly
        
        ActiveSheet.ChartObjects(i).Copy 'WITHOUT ACTIVESHEET
        mySlide.Shapes.Paste
        Set myShape = mySlide.Shapes(mySlide.Shapes.Count)
        
        myShape.Left = 200
        myShape.Top = 200
        
        PowerPointApp.Visible = True
        PowerPointApp.Activate
        
        Application.CutCopyMode = False
 

    Next i


End Sub

Private Sub CancelButton_Click()

    Unload ChartConverter

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TablesCharts_Click()

    Dim PowerPointApp As Object
    Dim myPresentation As Object
    Dim mySlide As Object
    Dim myShape As Object
    Dim oChrt As ChartObject
    Dim n As String
    
    
    Dim i As Integer 'counters chart
    
    'check for name of worksheet
    
    ActiveWindow.DisplayGridlines = False
    
    n = "wholesale metrics.xlsx" 'Name of the file
    
    If (ActiveWorkbook.Name Like n) Then
        If PowerPointApp Is Nothing Then Set PowerPointApp = CreateObject(class:="PowerPoint.Application")
        
        On Error Resume Next
        
        Application.ScreenUpdating = False
        
        Set myPresentation = PowerPointApp.Presentations.Add
        
        'Count number of worksheets and loop that many times
        
        For i = 1 To ActiveSheet.ChartObjects.Count
        'copy chart objects from worksheets
            
            Set oChrt = ActiveSheet.ChartObjects(i)
            Set mySlide = myPresentation.Slides.Add(1, 11) '11 = ppLayoutTitleOnly
            
            With rngochrt
                Set rngochrt = Range(oChrt.TopLeftCell, oChrt.BottomRightCell)
                'Range of cells that includes the chart and data (to copy)
                Set rngData = .Cells.Resize(.Rows.Count, .Columns.Count + 3)
                'Set rngData = .Offset(4, .Columns.Count).Cells(1).Resize(10, 3)
                rngData.Copy
            End With
            mySlide.Shapes.PasteSpecial DataType:=2
    
            Set myShape = mySlide.Shapes(mySlide.Shapes.Count)
            
            myShape.Left = 66
            myShape.Top = 100
            
            PowerPointApp.Visible = True
            PowerPointApp.Activate
            
            Application.CutCopyMode = False
    
        Next i
        
    Else
        MsgBox "This option is reserved for Wholesale Metrics.xlsm. Try again, or check filename to contain *wholesale*", vbInformation
        Unload ChartConverter 'close the userform
    End If

     ActiveWindow.DisplayGridlines = True
    

End Sub


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

Dim a As Long
    Me.WsNames.Clear
    For a = 1 To Sheets.Count
        Me.WsNames.AddItem Sheets(a).Name
    Next
    Me.WsNames.Value = ActiveSheet.Name
End Sub

Private Sub WsNames_Change()
    Dim actWs As String
    actWs = WsNames.Text
    Worksheets(actWs).Select

End Sub
