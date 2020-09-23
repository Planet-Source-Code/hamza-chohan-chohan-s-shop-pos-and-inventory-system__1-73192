VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVButtons.ocx"
Begin VB.Form frmSales 
   Caption         =   "Point of Sales"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   13905
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin LVbuttons.LaVolpeButton Command4 
      Height          =   615
      Left            =   12240
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   8
      TX              =   "Process"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSales.frx":08CA
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   11400
      Picture         =   "frmSales.frx":08E6
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   5535
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5880
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvSale 
      Height          =   6615
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   11668
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   5400
      TabIndex        =   8
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "Product Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Product Name :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Reciept Number"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SP, AB As Currency
Public SC, RecieptNumber As Integer
Dim slist As ListItem
Public vivocalibuyo, lessQTY

Private Sub Command4_Click()
AB = Label1.Caption
frmCash.Show
End Sub



Private Sub Form_Load()
conSales
recSales
rsSales.MoveLast
Text5.Text = rsSales!RecieptNumber + 1
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmCash
Unload frmQty
End Sub

Private Sub lvSale_KeyDown(KeyCode As Integer, Shift As Integer)


Dim ans
Select Case KeyCode
    Case vbKeyF3
    ans = MsgBox("Are you Sure to Delete Item?" & lvSale.SelectedItem.SubItems(1), vbQuestion + vbYesNo, "Conformation")
    If ans = vbYes Then
        conSales
        recDelSales
        vivocalibuyo = lvSale.SelectedItem.SubItems(3)
        lessQTY = rsSales!Quantity
            rsSales.Delete
            lvSale.ListItems.Clear
                If lvSale.ListItems.Count = 0 Then
                    Label1.Caption = Format(Label1.Caption, "0.00")
                    lvSale.ListItems.Clear
                End If
        Set rsSales = Nothing
        recQSales
        With rsQSales
            Do Until .EOF
            Set slist = lvSale.ListItems.Add(, , !qty)
            slist.SubItems(1) = !ProductName
            slist.SubItems(2) = !UnitPrice
            slist.SubItems(3) = Format(!Quantity * slist.SubItems(2), "#,###.00")
            Label1.Caption = Format(Label1.Caption - vivocalibuyo, "#,###.00")
            .MoveNext
            Loop
        End With
    End If
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    conStock
    recQstock
    On Error Resume Next
    If rsQStock!SalesID = Text1.Text Then
    Text2.Text = rsQStock!ProductName
    Text3.Text = rsQStock!UnitPrice
    SP = rsQStock!UnitPrice
    SC = rsQStock!SerialNumber
    RecieptNumber = Text5.Text
    End If

    frmQty.Show
End If
End Sub

