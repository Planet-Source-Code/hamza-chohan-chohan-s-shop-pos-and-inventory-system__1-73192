VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVButtons.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Chohan Shop's Sales and Inventory"
   ClientHeight    =   5100
   ClientLeft      =   240
   ClientTop       =   630
   ClientWidth     =   9195
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   720
      Picture         =   "frmAbout.frx":1CCA
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin LVbuttons.LaVolpeButton cmdOK 
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "&OK"
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
      MICON           =   "frmAbout.frx":A94B
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
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Designed for Chohan Sweets, Bakers and Chinese Restorent Management Sales &  Inventory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   690
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   3285
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Chohan Shop's Sales and Inventory System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5520
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5520
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmAbout.frx":A967
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9255
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":16411
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   5655
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version :  1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   6720
      TabIndex        =   3
      Top             =   2880
      Width           =   1725
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Designed for Comprehensive Approach for Managing Activities in School."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   690
      Index           =   0
      Left            =   5520
      TabIndex        =   2
      Top             =   1920
      Width           =   3285
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "School Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8880
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   9240
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":16560
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

