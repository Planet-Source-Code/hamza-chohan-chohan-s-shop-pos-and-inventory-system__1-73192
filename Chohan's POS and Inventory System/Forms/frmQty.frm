VERSION 5.00
Begin VB.Form frmQty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Quantity"
   ClientHeight    =   1965
   ClientLeft      =   5880
   ClientTop       =   5040
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public amt, subtot, tot As Currency
Public nQTY As Integer
Private Sub Form_Load()
SendKeys "{HOME}+{END}"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim slist As ListItem
If KeyAscii = 13 Then

Set slist = frmSales.lvSale.ListItems.Add(, , Text1.Text)
slist.SubItems(1) = frmSales.Text2.Text
slist.SubItems(2) = frmSales.Text3.Text

tot = frmSales.Text3.Text * Val(Text1.Text)

slist.SubItems(3) = Format(tot, "#,###.00")
subtot = tot
amt = amt + subtot
frmSales.Label1.Caption = Format(amt - frmSales.vivocalibuyo, "#,###.00")
nQTY = nQTY + Text1.Text

'add to sales table
conSales
recSales
With rsSales
    .AddNew
    !ProductCode = frmSales.SC
    !ProductName = frmSales.Text2.Text
    !SalesID = frmSales.Text1.Text
    !UnitPrice = frmSales.SP
    !Quantity = Text1.Text
    !TotalAmount = Text1.Text * frmSales.SP
    !RecieptNumber = frmSales.Text5.Text
    !SalesDate = Format(Date$, "m") & "/" & Format(Date$, "dd") & "/" & Format(Date$, "yyyy")
    .Update
End With
recQstock
With rsQStock
    !Quantity = !Quantity - Text1.Text
    .Update
End With
Unload Me
frmSales.Text1.SetFocus
SendKeys "{HOME}+{END}"
End If
End Sub
