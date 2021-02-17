VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIncomingPCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnClearForm_Click()
Dim frm As Form
Dim ctl As Control
Dim ctype
Set frm = Me

'Loop through all text boxes and
'Comboboxes on form and set valie to nothing
For Each ctl In frm.Controls
ctype = ctl.Name
Debug.Print ctype
If ctype = "txtPartDescription" Then

Else
If ctl.ControlType = acTextBox _
Or ctl.ControlType = acComboBox Then
ctl.Value = ""
End If
End If
Next ctl
'clearForm (thisForm)

End Sub

Private Sub generate_inspection_report_Click()

Dim frmArray()
frmArray = Array(txtOrderNumber, cboContractors, cboPartNumber, txtPartDescription, txtQuantity, txtSerialNumbers, cboTesters, txtDate)
openInspectionReport (frmArray)

End Sub

Private Sub Save_Record_Click()

   Dim strSQL As String
    Dim frmArray()
    Dim valuesArray()
    Dim msgString, valueString
    
    'Read form values into array
    frmArray = Array(txtOrderNumber, cboContractors, cboPartNumber, txtPartDescription, txtQuantity, txtSerialNumbers, cboTesters, txtDate, txtInspectionNumber) 'if the combobox outputs a number, check what it is bound to in the form properties
    valuesArray = ArrayToSQLValue(frmArray)
    msgString = Join(valuesArray, ",")
    
    'Debug.Print msgString
    
    
     strSQL = "INSERT INTO tblIncomingPCB (JobNumber, Contractor, PartNumber, PartDescription, QTY, SerialNumbers, " & _
     "Tester, DateReceived, IncomingInspection)" & _
     "VALUES (" & msgString & ");"
     'Debug.Print strSQL
     
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    'DoCmd.GoToRecord , , acNewRec

    
End Sub

 

