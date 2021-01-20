VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Team 11 Term Project
'Oct 28, 2020
'Description: Algorithm for sorting sheet of softwares based on a search criteria
'Input: Search Criteria
'Output: List of softwares matching criteria

Private Sub btn_input_Click()
    frm_input.Show
End Sub

Private Sub btnClearInput_Click()
    Range("F3:M1000").Delete
End Sub

Private Sub btnCompareSheet_Click()
    Worksheets("COMPARE").Activate
End Sub

Private Sub btnSearch_Click()

 Set actsheet = ActiveSheet 'set INPUT sheet to variable actsheet
 Set Workbook = ActiveWorkbook

'   Read the search criteria
'   Read all data from INPUT sheet
'   Filter all the data using project type, then budget, then team size
'   Output the data to the cells below the button

'setting variables so that we read the data from the input sheet
'these variables store the criteria of the search
Dim projectType As String
Dim budget As Integer
Dim teamSize As String
Dim lastRow1 As Integer
lastRow1 = actsheet.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
projectType = actsheet.Cells(2, "A")
budget = actsheet.Cells(2, "B")
teamSize = actsheet.Cells(2, "C")


'we declare outputSheet as the second sheet (OUTPUT) in the entire workbook
Set outputSheet = Workbook.Worksheets(3)


'declare lastRow as a function to find the last row
'required since the amount of rows is not constant and will always change based on the data
Dim lastRow As Long: lastRow = outputSheet.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
Dim LastCol As Integer: LastCol = 6 'declare LastCol to 6 since on the output sheet, we are only using 6 coloumns and this value remains constant
Dim iRow As Integer, jCol As Integer
Dim iRow2 As Integer, jCol2 As Integer
Dim outputStartRow As Integer: outputStartRow = 3 'declare row 10 as the first row for the results table after the program has finished sorting based on criteria

Dim buttonControl As Object
Dim selectButtonCell As String

Range("F3:M100").Delete
   
For iRow2 = 3 To lastRow1
        selectButtonCell = "L" & iRow2
        Set buttonControl = ActiveSheet.Buttons.Add(Range(selectButtonCell).Left, Range(selectButtonCell).Top, Range(selectButtonCell).Width, Range(selectButtonCell).Height)
        With buttonControl
            .OnAction = "'Term 11 Term Project FINAL.xlsm'!selectSoftware"
            .Caption = "Compare"
        End With
Next iRow2

'iRow is set to 2 since in our output sheet, the data starts from row 2
For iRow = 2 To lastRow 'loop to scan output sheet
    If outputSheet.Cells(iRow, 3) = projectType Then 'if starting from row 2 on column 3 of output sheet, the value is equal to the value of the project type criteria then check next condition
        If outputSheet.Cells(iRow, 4) <= budget Then 'if starting from row 2 on column 4 of output sheet, the value is less than or equal to the value of budget criteria, then check next condition
            If outputSheet.Cells(iRow, 5) = teamSize Then 'if starting from row 2 on column 5 of output sheet, the value is equal to the value of the team size criteria, then...
                
                For jCol = 1 To LastCol
                
                    actsheet.Cells(outputStartRow, jCol + 5) = outputSheet.Cells(iRow, jCol) 'if the conditions above are satisfied, copy the entire row from column 1-6 over the results table in the active sheet
              
                Next jCol 'continues copying each column until entire row is copied into results table
        
                outputStartRow = outputStartRow + 1 'for next loop, if another software satifies the conditons, then copy the row to the next row (row + 1) in the results table
            End If
        End If
    End If
Next iRow 'check next row of output sheet

Range("F2:L100").HorizontalAlignment = xlCenter
Range("F2:L100").VerticalAlignment = xlCenter
Range("K2:K100").WrapText = True




End Sub


Sub ClearSearchResults()
    Range("F3:M1000").Delete 'clears the results table
End Sub


Private Sub btnSoftwareSheet_Click()
    Worksheets("SOFTWARE").Activate
End Sub
