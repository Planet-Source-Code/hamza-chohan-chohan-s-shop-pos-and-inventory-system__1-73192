VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVButtons.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSalesman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salesman"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   Icon            =   "frmSalesman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSalesman 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3135
      Left            =   3600
      TabIndex        =   8
      Top             =   360
      Width           =   4935
      Begin VB.TextBox txtSalary 
         DataField       =   "Salary"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "ADOSalesman"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtContactNumber 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ContactNo"
         DataSource      =   "ADOSalesman"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Address"
         DataSource      =   "ADOSalesman"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         DataField       =   "SalesmanName"
         DataSource      =   "ADOSalesman"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "SalesmanID"
         DataSource      =   "ADOSalesman"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraTools 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   9135
      Begin LVbuttons.LaVolpeButton cmdDelete 
         Height          =   495
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
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
         MICON           =   "frmSalesman.frx":0EE2
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
      Begin LVbuttons.LaVolpeButton cmdClose 
         Height          =   495
         Left            =   7800
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
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
         MICON           =   "frmSalesman.frx":0EFE
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
      Begin LVbuttons.LaVolpeButton cmdRefresh 
         Height          =   495
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
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
         MICON           =   "frmSalesman.frx":0F1A
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
      Begin LVbuttons.LaVolpeButton cmdEdit 
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
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
         MICON           =   "frmSalesman.frx":0F36
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
      Begin LVbuttons.LaVolpeButton cmdSave 
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Save"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
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
         MICON           =   "frmSalesman.frx":0F52
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
      Begin LVbuttons.LaVolpeButton cmdNew 
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
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
         MICON           =   "frmSalesman.frx":0F6E
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
      Begin LVbuttons.LaVolpeButton cmdFind 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Find"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
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
         MICON           =   "frmSalesman.frx":0F8A
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
   End
   Begin MSAdodcLib.Adodc ADOSalesman 
      Height          =   450
      Left            =   240
      Top             =   4680
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   794
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
      BackColor       =   -2147483633
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
      RecordSource    =   "Salesman"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Niagara Solid"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman ID:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman Name:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "frmSalesman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are You Sure To Delete Record?", vbYesNo + vbQuestion, "Delete") = vbYes Then
ADOSalesman.Recordset.Delete
MsgBox "Record Succesfully Deleted.", vbInformation, "Delete Record"
ADOSalesman.Recordset.MoveFirst
End If
End Sub

Private Sub cmdEdit_Click()
    fraSalesman.Enabled = True
    txtName.SetFocus
    SendKeys "{home}+{end}"
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
inp = InputBox("Enter Salesman Name...")
ADOSalesman.Refresh
ADOSalesman.Recordset.Find "SalesmanName = '" & inp & "'"
End Sub

Private Sub cmdNew_Click()
    ADOSalesman.Refresh
    ADOSalesman.Recordset.AddNew
    fraSalesman.Enabled = True
    txtName.SetFocus
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
    txtID.Text = ADOSalesman.Recordset.RecordCount
End Sub

Private Sub cmdRefresh_Click()
    ADOSalesman.Refresh
    fraSalesman.Enabled = False
    cmdNew.Enabled = True
    cmdSave.Enabled = False
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
End Sub

Private Sub cmdSave_Click()
    ADOSalesman.Recordset.Update
    fraSalesman.Enabled = False
    cmdSave.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
End Sub
