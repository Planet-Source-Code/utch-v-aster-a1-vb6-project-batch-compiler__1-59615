VERSION 5.00
Begin VB.Form frmFindVB 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Searching for Visual Basic v6.0"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindVB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      Picture         =   "frmFindVB.frx":058A
      ScaleHeight     =   900
      ScaleWidth      =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   6420
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   105
      TabIndex        =   1
      Top             =   1050
      Width           =   6210
   End
End
Attribute VB_Name = "frmFindVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScanDrive As String

Private Sub Form_Load()
  
  Visible = True
  Do Until Visible
    DoEvents
  Loop
  
  If Command$ <> "" Then
    ScanDrive = Command$
  Else
    ScanDrive = "c"
  End If

  VBPath = FindFile(ScanDrive & ":\", "vb6.exe", Me)
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Cancel = True
  End If
End Sub
