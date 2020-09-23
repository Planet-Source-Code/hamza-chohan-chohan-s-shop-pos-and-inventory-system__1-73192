VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000004&
   Caption         =   "Chohan's Shop Sales and Inventory System"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11010
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":058A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   14640
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   10950
      TabIndex        =   1
      Top             =   0
      Width           =   11010
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Chohan Sweets, Bakeres and Chinese Restorent"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14880
         TabIndex        =   2
         Top             =   120
         Width           =   7095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6210
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:34 PM"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuLegders 
      Caption         =   "&Legders"
      Begin VB.Menu mnuProducts 
         Caption         =   "Products"
         Begin VB.Menu mnuManageProducts 
            Caption         =   "Manage Products"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuProductsCatagory 
            Caption         =   "Products Catagory"
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customers Book"
         Shortcut        =   ^C
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "Suppliers"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalesman 
         Caption         =   "Salesman Book"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuMakeSales 
         Caption         =   "Make Sales"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuSecurity 
      Caption         =   "&Security"
      Begin VB.Menu mnuUsers 
         Caption         =   "Manage Users"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuSalesReport 
         Caption         =   "Total Sales Report"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProductsReport 
         Caption         =   "Total Products Report"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuCreater 
         Caption         =   "Program Creater"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Need Help!"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
StatusBar.Panels(1) = "You are LOG in as: " & frmLogin.nTYPE
StatusBar.Panels(2) = "You are: " & frmLogin.nNAME

If frmLogin.nTYPE = "Administrator" Then
mnuLegders.Enabled = True
mnuSecurity.Enabled = True
Else
mnuLegders.Enabled = False
mnuSecurity.Enabled = False
End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are You Sure To Exit?", vbQuestion + vbYesNo, "Exit") = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub mnuCreater_Click()
frmAbout.Show
End Sub

Private Sub mnuCustomer_Click()
frmCustomers.Show
End Sub



Private Sub mnuExit_Click()
If MsgBox("Are You Sure To Exit?", vbQuestion + vbYesNo, "Exit") = vbYes Then
End
Else
Exit Sub
End If
End Sub


Private Sub mnuHelp_Click()
MsgBox "If You Need Any Help About This Program or Other Problem So Just Send Email to My ID, email: hamzajhang@yahoo.com or Contact me at, Phone: +92-334-6320905", vbInformation + vbOKOnly, "Need Help!"
End Sub



Private Sub mnuMakeSales_Click()
frmSales.Show
End Sub

Private Sub mnuManageProducts_Click()
frmProducts.Show
End Sub

Private Sub mnuProductsCatagory_Click()
frmCatagoryProducts.Show
End Sub

Private Sub mnuProductsReport_Click()
rptProducts.Show
End Sub

Private Sub mnuSalesman_Click()
frmSalesman.Show
End Sub

Private Sub mnuSalesReport_Click()
rptTotalSales.Show
End Sub

Private Sub mnuSuppliers_Click()
frmSuppliers.Show
End Sub

Private Sub mnuUsers_Click()
frmUserAccounts.Show
End Sub

Private Sub Timer1_Timer()
Label.Left = Label.Left - 100

If Label.Left + Label.Width <= 0 Then
Label.Left = Picture1.Width
End If

End Sub
