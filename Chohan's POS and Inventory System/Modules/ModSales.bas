Attribute VB_Name = "ModSales"
Public cnSales As New ADODB.Connection
Public rsSales As New ADODB.Recordset
Public rsQSales As New ADODB.Recordset
Public Sub conSales()
Set cnSales = New ADODB.Connection
cnSales.CursorLocation = adUseClient
cnSales.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Data.mdb;Persist Security Info=False;Jet OLEDB:Database Password=hamzas007;"
End Sub
Public Sub recDelSales()
Set rsSales = New ADODB.Recordset
rsSales.Open "SELECT * FROM Sales WHERE ProductName LIKE'" & frmSales.lvSale.SelectedItem.SubItems(1) & "'", cnSales, 3, 2
End Sub

Public Sub recSales()
Set rsSales = New ADODB.Recordset
rsSales.Open "SELECT * FROM Sales", cnSales, 3, 2
End Sub

Public Sub recQSales()
Set rsQSales = New ADODB.Recordset
rsQSales.Open "SELECT * FROM Sales WHERE RecieptNumber LIKE'" & frmSales.RecieptNumber & "'", cnSales, 3, 2
End Sub

Public Sub recDSales()
Dim tDate
tDate = Format(Date$, "m") & "/" & Format(Date$, "dd") & "/" & Format(Date$, "yyyy")
Set rsSales = New ADODB.Recordset
rsSales.Open "SELECT * FROM Sales WHERE SalesDate LIKE'" & tDate & "'", cnSales, 3, 2
End Sub


