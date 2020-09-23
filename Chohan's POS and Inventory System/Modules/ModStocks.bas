Attribute VB_Name = "ModStocks"
Public cnStock As New ADODB.Connection
Public rsStock As New ADODB.Recordset
Public rsQStock As New ADODB.Recordset

Public Sub conStock()
Set cnStock = New ADODB.Connection
cnStock.CursorLocation = adUseClient
cnStock.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Data.mdb;Persist Security Info=False;Jet OLEDB:Database Password=hamzas007;"
End Sub

Public Sub recStock()
Set rsStock = New ADODB.Recordset
rsStock.Open "SELECT * FROM Products", cnStock, 3, 2
End Sub

Public Sub recUpStock()
Set rsStock = New ADODB.Recordset
rsStock.Open "SELECT * FROM Products WHERE SalesID like'" & frmStocks.lvStocks.SelectedItem.Text & "'", cnStock, 3, 2
End Sub

Public Sub recQstock()
Set rsQStock = New ADODB.Recordset
rsQStock.Open "SELECT * FROM Products WHERE SerialNumber like'" & frmSales.Text1.Text & "'", cnStock, 3, 2
End Sub

Public Sub recCStock()
Set rsStock = New ADODB.Recordset
rsStock.Open "SELECT * FROM Categories", cnStock, 3, 2
End Sub
