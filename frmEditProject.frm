VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditProject 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Project"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1853
      TabIndex        =   7
      Top             =   1710
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3173
      TabIndex        =   6
      Top             =   1710
      Width           =   1245
   End
   Begin VB.TextBox txtEXE 
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   900
      Width           =   4725
   End
   Begin VB.TextBox txtVBP 
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   300
      Width           =   4725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4890
      TabIndex        =   1
      Top             =   855
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4890
      TabIndex        =   0
      Top             =   255
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4830
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Current EXE File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   660
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Project:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   1350
   End
End
Attribute VB_Name = "frmEditProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mProj As Integer

Private Sub Command1_Click()
  CD.Filter = "VBP files (*.vbp)|*.VBP"
End Sub

Private Sub Command2_Click()
  CD.Filter = "EXE files (*.exe)|*.EXE"
End Sub

Private Sub Command3_Click()
  frmMain.EditResult = vbCancel
  Unload Me
End Sub

Private Sub Command4_Click()
  frmMain.EditResult = vbOK
  mProjects(mProj).ProjectFullPath = txtVBP
  mProjects(mProj).ExeFullPath = txtEXE
  mProjects(mProj).ProjectName = GetFileName(mProjects(mProj).ProjectFullPath)
  mProjects(mProj).EXEName = GetFileName(mProjects(mProj).ExeFullPath)
  Unload Me
End Sub

Private Sub Form_Load()
  txtEXE = mProjects(mProj).ExeFullPath
  txtVBP = mProjects(mProj).ProjectFullPath
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then frmMain.EditResult = vbCancel
End Sub
