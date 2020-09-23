VERSION 5.00
Begin VB.Form frmCash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accept Payment"
   ClientHeight    =   1350
   ClientLeft      =   5880
   ClientTop       =   5625
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public amount, cash, change As Currency

Private Sub Form_Load()
SendKeys "{HOME}+{END}"
conSales
recSales
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cash = Format(Text1.Text, "#,###.00")
    amount = Format(Val(Me.Text1.Text), "#,###.00") - Val(frmSales.Label1.Caption)
    change = amount
    frmSales.Label1.Caption = Format(Val(amount), "#,###.00")
    Unload Me
    MsgBox "This Will Print the Reciept by Number " & frmSales.Text5.Text, vbInformation, "Print Receipt"
    ORrep
End If
End Sub

Private Sub ORrep()
conSales
recQSales
Set rptOR.DataSource = rsQSales
rptOR.Sections("section2").Controls("label5").Caption = Val(frmSales.Text5.Text)
rptOR.Sections("section5").Controls("label15").Caption = Format(change, "#,###.00")
rptOR.Sections("section5").Controls("label13").Caption = Format(cash, "#,###.00")
rptOR.Sections("section5").Controls("label7").Caption = Format(frmSales.AB, "#,###.00")
rptOR.Sections("section5").Controls("label11").Caption = Format(frmSales.AB, "#,###.00")
rptOR.Sections("section5").Controls("label17").Caption = frmQty.nQTY - frmSales.lessQTY
rptOR.Show 1
frmSales.Text1.Text = ""
frmSales.Text2.Text = ""
frmSales.Text3.Text = ""
frmSales.lvSale.ListItems.Clear
frmSales.Text5.Text = frmSales.Text5.Text + 1
frmSales.Label1.Caption = "0.00"
frmQty.tot = "0"
frmQty.amt = "0"
frmQty.subtot = "0"
frmQty.nQTY = "0"
amount = "0"
cash = "0"
change = "0"
End Sub



