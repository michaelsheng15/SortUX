VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnCompare_Click()
    Worksheets("COMPARE").Activate
End Sub

Private Sub btnEnter_Click()
    Worksheets("INPUT").Activate
End Sub

Private Sub btnOutput_Click()
    Worksheets("SOFTWARE").Activate
End Sub
