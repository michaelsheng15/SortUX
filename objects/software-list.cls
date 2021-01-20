VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnAdd_Click()
    frmAddSoftware.Show
End Sub


Private Sub btnC0mpare_Click()
    Worksheets("COMPARE").Activate
End Sub

Private Sub btnInput_Click()
    Worksheets("INPUT").Activate
End Sub
