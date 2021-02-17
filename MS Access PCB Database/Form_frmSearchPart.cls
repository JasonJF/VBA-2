VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSearchPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Combo3_AfterUpdate()

End Sub

Private Sub cboPartNumber_AfterUpdate()
DoCmd.OpenQuery "qrySearchPartOnly", acViewNormal, acEdit

'DoCmd.Close acForm, "frmSearchPart"
End Sub

Private Sub Form_AfterUpdate()

End Sub

Private Sub Run_Query_Click()
DoCmd.OpenQuery "qrySearchPartOnly", acViewNormal, acEdit
End Sub

Private Sub txtPart_Enter()
DoCmd.OpenQuery "Query1", acViewNormal, acEdit

'DoCmd.Close acForm, "frmSearchPart"

End Sub

