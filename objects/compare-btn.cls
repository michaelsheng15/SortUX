VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnClear_Click()
    Range("A2:F1000").EntireRow.Delete
End Sub

Private Sub btnInputSheet1_Click()
    Worksheets("INPUT").Activate
End Sub

Private Sub btnSoftwareSheet1_Click()
    Worksheets("SOFTWARE").Activate
End Sub


