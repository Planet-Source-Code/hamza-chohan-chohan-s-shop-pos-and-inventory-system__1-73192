VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVButtons.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Login"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Niagara Solid"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc ADOUser 
      Height          =   375
      Left            =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database\Data.mdb;Persist Security Info=False;Jet OLEDB:Database Password=hamzas007"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database\Data.mdb;Persist Security Info=False;Jet OLEDB:Database Password=hamzas007"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Users"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Niagara Solid"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtPass 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "/"
      TabIndex        =   5
      ToolTipText     =   "Type Password Here..."
      Top             =   2160
      Width           =   2535
   End
   Begin VB.ComboBox cboUser 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin LVbuttons.LaVolpeButton cmdLogin 
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Niagara Solid"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmLogin.frx":09EA
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
   Begin LVbuttons.LaVolpeButton cmdCancel 
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Can&cel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Niagara Solid"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmLogin.frx":0A06
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   0
      Picture         =   "frmLogin.frx":0A22
      Top             =   0
      Width           =   1965
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Username As String
Public Password As String
Public Login As Boolean
Public nUSER, nNAME, nLOG, nTYPE


Private Sub cmdCancel_Click()
    End
End Sub
Private Sub cmdLogin_Click()
If Login = True Then
    If cboUser.Text = Username And txtPass.Text = Password Then
        txtPass.Text = ""
        cboUser.SetFocus
        Login = False
    Else
        MsgBox "Sorry! Username or Password is Wrong", vbCritical + vbOKOnly, "Sales and Inventory System"
        txtPass.SetFocus
        SendKeys "{Home}+{End}"
    End If
Else
    ADOUser.Recordset.Filter = "Username = '" & cboUser.Text & "'"
        If txtPass = ADOUser.Recordset!Password Then
            Username = cboUser.Text
            Password = txtPass.Text
            Username = StrConv(ADOUser.Recordset!Username, vbUpperCase)
            nUSER = ADOUser.Recordset!Username
            nNAME = ADOUser.Recordset!FullName
            nTYPE = ADOUser.Recordset!UserType
            frmMain.Show
            Unload Me
        Else
            MsgBox "Sorry! Username or Password is Wrong", vbCritical + vbOKOnly, "Sales and Inventory System"
            txtPass.SetFocus
            SendKeys "{Home}+{End}"
        End If
End If
End Sub

Private Sub Form_Load()
ADOUser.Refresh
Do While Not ADOUser.Recordset.EOF
    cboUser.AddItem ADOUser.Recordset!Username
    ADOUser.Recordset.MoveNext
Loop
End Sub
