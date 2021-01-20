Attribute VB_Name = "Module1"
Sub selectSoftware()
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    
    Set Workbook = ActiveWorkbook
    Set compareSheet = Workbook.Worksheets(4)
    Set inputSheet = Workbook.Worksheets(2)
    
    Dim iCol As Integer

    Dim LastCol As Integer
    Dim inputStartRow As Integer
    inputStartRow = 10
    Dim startRow As Integer
    startRow = compareSheet.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    For i = 6 To 12
        Cells(currentRow, i).Interior.ColorIndex = 4
    Next i
    
    For iCol = 6 To 11
                
        compareSheet.Cells(startRow + 1, iCol - 5) = inputSheet.Cells(currentRow, iCol)
        compareSheet.Cells.HorizontalAlignment = xlCenter
        compareSheet.Cells.VerticalAlignment = xlCenter
        compareSheet.Cells.WrapText = True
    Next iCol
        

End Sub
