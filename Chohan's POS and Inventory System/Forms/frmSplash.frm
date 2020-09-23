VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "School Management System"
   ClientHeight    =   5040
   ClientLeft      =   2955
   ClientTop       =   2730
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1440
      TabIndex        =   5
      Top             =   4440
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1320
      TabIndex        =   4
      Top             =   4800
      Width           =   6615
   End
   Begin VB.Timer lblcoltim 
      Interval        =   200
      Left            =   3840
      Top             =   1440
   End
   Begin VB.Timer tmProgress 
      Interval        =   30
      Left            =   4440
      Top             =   1440
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   1560
      TabIndex        =   0
      Top             =   4560
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6360
      TabIndex        =   6
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblScltag 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait While Loading"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1560
      TabIndex        =   2
      Top             =   4080
      Width           =   2265
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   3600
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   2040
      Picture         =   "frmSplash.frx":6062
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5415
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer



Private Sub lblcoltim_Timer()
'lblScltag.Left = lblScltag.Left - 100 'code to move the newlife tag in the form

'If lblScltag.Left + lblScltag.Width <= 0 Then
'lblScltag.Left = frmSplash.Width

'End If

'=====================================================
'Now to set the different color on the newlife tag
If lblScltag.ForeColor = vbBlack Then
lblScltag.ForeColor = vbRed
ElseIf lblScltag.ForeColor = vbRed Then
lblScltag.ForeColor = vbGreen
ElseIf lblScltag.ForeColor = vbGreen Then
lblScltag.ForeColor = vbYellow
ElseIf lblScltag.ForeColor = vbYellow Then
lblScltag.ForeColor = vbBlue
ElseIf lblScltag.ForeColor = vbBlue Then
lblScltag.ForeColor = vbBlack
End If

' to blink the loading text
'==========================
If Label3.Visible = True Then
Label3.Visible = False
Else
Label3.Visible = True
End If

End Sub






Private Sub tmProgress_Timer()

' timer code for loading the project as well as increment in the progress bar value
'===========================================
a = a + 1
Label1.Caption = CStr(a) & "% " & "Completed"
ProgressBar1.Value = a
If a = 100 Then
Unload Me
frmLogin.Show


End If
End Sub
