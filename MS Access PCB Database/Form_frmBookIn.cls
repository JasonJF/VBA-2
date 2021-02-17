VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBookIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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
    frmArray = Array(txtOrderNumber, cboContractors, cboPartNumber, txtPartDescription, txtQuantity, txtSerialNumbers, cboTesters, txtDate) 'if the combobox outputs a number, check what it is bound to in the form properties
    valuesArray = ArrayToSQLValue(frmArray)
    msgString = Join(valuesArray, ",")
    
    'Debug.Print msgString
    
    
     strSQL = "INSERT INTO tblBookIn (OrderNumber, Contractor, PartNumber, PartDescription, QTY, SerialNumbers, " & _
     "Tester, DateCompleted)" & _
     "VALUES (" & msgString & ");"
     'Debug.Print strSQL
     
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True

    
End Sub

 

