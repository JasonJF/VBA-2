Attribute VB_Name = "pcbDatabaseModule"
Option Compare Database
Option Explicit

Public Sub StringDemoSQL()
    Dim sql As String
    sql = "SELECT * FROM tblOrder WHERE OrderStatus = ""In Progress"""
End Sub

Public Function ArrayToSQLValue(valueArray)

Dim newValueArray()
Dim stringplaceholder
Dim Value, i

i = 0
For Each Value In valueArray

    stringplaceholder = "'" & Value & "'"
    'Debug.Print stringplaceholder
    ReDim Preserve newValueArray(i)
    newValueArray(i) = stringplaceholder
    i = i + 1
    
    
    Next
    
ArrayToSQLValue = newValueArray
End Function
Public Sub openInspectionReport(xArray)


Dim appXL As Object
Dim wb As Object
Dim wks As Object
Dim xlf As String
Dim rs As DAO.Recordset
Dim order, contractor, part, description, quantity, serials, tester, datec
'Set cell numbers for variables
order = "B7"
contractor = "F7"
part = "B8"
description = "F8"
quantity = "B9"
serials = "F9"
tester = ""
datec = "I56"

Const xlMinimized As Long = 1
Const xlMaximized As Long = 2

'Debug.Print xArray(0)



xlf = "U:\Spreadsheets\Insp_Report_dB_ver2.xls" 'Full path to Excel file

'Set rs = CurrentDb.OpenRecordset("Query1") 'Replace Query1 with real query name
Set appXL = CreateObject("Excel.Application")
'appXL.WindowState = xlMaximized
appXL.Visible = True

Set wb = appXL.Workbooks.Open(xlf)
Set wks = wb.Worksheets("Inspection") ' Sheet name

'set a values in excel sheet

wks.Range(order).Value = "" & xArray(0)
wks.Range(contractor).Value = "" & xArray(1)
wks.Range(part).Value = "" & xArray(2)
wks.Range(description).Value = "" & xArray(3)
wks.Range(quantity).Value = "" & xArray(4)
wks.Range(serials).Value = "" & xArray(5)
'wks.Range(tester).Value = "" & xArray(6)
wks.Range(datec).Value = "" & xArray(7)

'wb.Open
'wb.Save
'wb.Close
'appXL.Quit
Set wb = Nothing
'rs.Close
Set rs = Nothing

End Sub

Public Function clearForm(thisForm)
    Set frm = thisForm

'Loop through all text boxes and
'Comboboxes on form and set valie to nothing
For Each ctl In frm.Controls
If ctl.ControlType = acTextBox _
Or ctl.ControlType = acComboBox Then
ctl.Value = ""
End If
Next ctl



End Function
