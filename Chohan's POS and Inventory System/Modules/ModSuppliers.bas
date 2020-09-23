Attribute VB_Name = "ModSuppliers"
Public cnSUP As New ADODB.Connection
Public rsSUP, rsgSUP As New ADODB.Recordset

Public Sub conSup()
Set cnSUP = New ADODB.Connection
cnSUP.CursorLocation = adUseClient
cnSUP.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Data.mdb;Persist Security Info=False;Jet OLEDB:Database Password=hamzas007;"
End Sub

Public Sub recSup()
Set rsSUP = New ADODB.Recordset
rsSUP.Open "SELECT * FROM Suppliers", cnSUP, 3, 2
End Sub

Public Sub recUpSup()
Set rsSUP = New ADODB.Recordset
rsSUP.Open "SELECT * FROM Suppliers WHERE SupplierID like'" & frmProducts.cboSupplier.Text & "'", cnSUP, 3, 2
End Sub

Public Sub recGrpSup()
Set rsgSUP = New ADODB.Recordset
rsgSUP.Open "SELECT SupplierName FROM Suppliers GROUP BY SupplierName", cnSUP, 3, 2
End Sub

