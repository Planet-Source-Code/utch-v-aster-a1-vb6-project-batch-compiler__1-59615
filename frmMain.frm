VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   " Batch VB6 Compiler"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framCompile 
      Height          =   705
      Left            =   330
      TabIndex        =   1
      Top             =   810
      Visible         =   0   'False
      Width           =   645
      Begin VB.Frame frmSlide 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2895
         Left            =   330
         TabIndex        =   2
         Top             =   60
         Width           =   7815
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   2895
            Left            =   0
            ScaleHeight     =   2835
            ScaleWidth      =   705
            TabIndex        =   4
            Top             =   0
            Width           =   765
            Begin VB.Image Image1 
               Height          =   480
               Left            =   120
               Picture         =   "frmMain.frx":058A
               Top             =   0
               Width           =   480
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   150
               Picture         =   "frmMain.frx":08BE
               Top             =   1770
               Width           =   480
            End
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   285
            Left            =   1380
            TabIndex        =   3
            Top             =   1800
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
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
            Left            =   1470
            TabIndex        =   8
            Top             =   30
            Width           =   1350
         End
         Begin VB.Label lblProject 
            AutoSize        =   -1  'True
            Caption         =   "-------"
            Height          =   195
            Left            =   1470
            TabIndex        =   7
            Top             =   240
            Width           =   420
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
            Left            =   1470
            TabIndex        =   6
            Top             =   630
            Width           =   1350
         End
         Begin VB.Label lblEXE 
            AutoSize        =   -1  'True
            Caption         =   "-------"
            Height          =   195
            Left            =   1470
            TabIndex        =   5
            Top             =   840
            Width           =   420
         End
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7620
      Top             =   2670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "BPR files (*.bpr)|*.BPR"
   End
   Begin MSComctlLib.ListView lstProj 
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1349
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EXE Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Batch Profile"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Batch Profile..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Batch Profile..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Batch Profile &As..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuCompile 
      Caption         =   "&Compile All"
   End
   Begin VB.Menu mnuQSave 
      Caption         =   "&Quick Save"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EditResult As Integer
Dim SaveLocation As String
Dim mChanged As Boolean


Private Sub Form_Load()
  Dim Answer As Integer
  
  'Check for VB Path in Registry
  VBPath = GetSetting("BatchCompile", "Settings", "VBPath")
  
  'If it doesnt exist, do a scan.
  If VBPath = "" Then
    
    If Command$ <> "" Then
      
      'if a drive was passed, do not display message.
      Answer = vbYes
      
    Else
      
      Answer = MsgBox("NOTE: If VB6 is not installed on your C: Drive, This scan will fail." & vbCrLf & vbCrLf & _
                     "Please compile the EXE file, and pass the drive VB6 is installed on" & vbCrLf & _
                     "as the command$ FOR THE FIRST RUN ONLY. Once VB6 is found," & vbCrLf & "you can run this code anyway you'd like." & vbCrLf & vbCrLf & _
                     "Example (installed on S: Drive):" & vbCrLf & _
                     "c:\vbcode\batchcompiler.exe s" & vbCrLf & vbCrLf & "Continue to Scan C: Drive?", vbYesNo + vbQuestion, "Scanning for VB6.exe")
    
    End If
    
    'Continue with Scan on C:?
    If Answer = vbYes Then
      
      frmFindVB.Show
      Unload frmFindVB
      
      If VBPath = "" Then
        MsgBox "VB Path Not Found."
        End
      Else
        'Save VBPath to Registry
        SaveSetting "BatchCompile", "Settings", "VBPath", VBPath
        MsgBox "To add a *.VBP Project to your current Batch Profile," & vbCrLf & "simply drag the *.VBP file into the list.", vbInformation
      End If
    Else
      'Dont want to continue Scan on C:? go here.
      End
    End If
  End If
  
  'Ensure form is visible
  Visible = True
  Do Until Visible
    DoEvents
  Loop
  
  'Load Last profile.
  Dim LastProfile As String
  LastProfile = GetSetting("BatchCompile", "Settings", "LastProfile")
  If LastProfile <> "" Then
    Call OpenProfile(LastProfile)
    CD.InitDir = GetDirectoryName(LastProfile)
  End If
  
End Sub

Private Sub Form_Resize()
  'Resizeing Controls and column headers
  lstProj.Move 0, 0, ScaleWidth, ScaleHeight
  framCompile.Move 0, -60, ScaleWidth, ScaleHeight + 30
  frmSlide.Move 60, 240, framCompile.Width - 120
  
  Dim X As Integer
  For X = 1 To lstProj.ColumnHeaders.Count
    lstProj.ColumnHeaders(X).Width = (lstProj.Width - 120) / lstProj.ColumnHeaders.Count
  Next
End Sub

Private Sub lstProj_DblClick()
  
  'Load a new Project Editor Form
  Dim Another As New frmEditProject
  
  'Set the mProj to the project arrays index
  Another.mProj = Val(Replace(LCase$(lstProj.SelectedItem.Key), "key:", ""))
  
  Another.Show 1
  
  If EditResult = vbOK Then
    DoChanged (True)
    PopulateProjects
  End If
  
End Sub

Private Sub lstProj_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim I As Integer
  
  'To Delete a project
  If KeyCode = 46 Then
    Dim Answer As Integer
    I = Val(Replace(LCase$(lstProj.SelectedItem.Key), "key:", ""))
    Answer = MsgBox("Are you sure you want to delete this project from the profile?" & vbCrLf & vbCrLf & lstProj.SelectedItem.Text, vbYesNo + vbQuestion, "Delete Project From Profile?")
    If Answer = vbYes Then
      Dim X As Integer
      
      'move entries down to fill deleted spot
      For X = I To UBound(mProjects) - 1
        mProjects(X).ExeFullPath = mProjects(X + 1).ExeFullPath
        mProjects(X).EXEName = mProjects(X + 1).EXEName
        mProjects(X).ProjectFullPath = mProjects(X + 1).ProjectFullPath
        mProjects(X).ProjectName = mProjects(X + 1).ProjectName
      Next
      
      'Clip off the last array slot
      ReDim Preserve mProjects(UBound(mProjects) - 1)
      
      'Refresh List
      PopulateProjects
      
      'Set CHANGED to true
      DoChanged (True)
    End If
  End If
End Sub

Private Sub lstProj_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Data.GetFormat(vbCFFiles) Then
    Dim tFile
    'Receives Dragged/Dropped Files
    For Each tFile In Data.Files
      'Make sure its a VBP file
      If Right(LCase(tFile), 4) = ".vbp" Then
        Call DoChanged(True)
        Call AddProject(tFile)
      End If
    Next
    Call PopulateProjects
  End If
End Sub

Sub AddProject(vFile)
  On Error GoTo Err
  Dim pFile As String
  pFile = vFile
  
  
  Dim Readln As String
  Dim Num As Integer
  Dim Path32 As String
  Dim ExeName32 As String
  Num = FreeFile()
  
  'Retreive Compile information from VBP file
  Open vFile For Input As #Num
    Do While Not EOF(Num)
      Line Input #Num, Readln
      If Left(Readln, 9) = "ExeName32" Then
        ExeName32 = Replace(Readln, "ExeName32=" & Chr(34), "")
        ExeName32 = Left$(ExeName32, Len(ExeName32) - 1)
      ElseIf Left(Readln, 6) = "Path32" Then
        Path32 = Replace(Readln, "Path32=" & Chr(34), "")
        Path32 = Left$(Path32, Len(Path32) - 1)
      End If
    Loop
  Close #Num
  
  'Project hasnt been compiled yet?
  If Path32 = "" Or ExeName32 = "" Then
    MsgBox "Please compile this project manually in VB First."
    Exit Sub
  End If
  
  'Get Realtive Path to the project file
  Path32 = GetRelativePath(Path32, Left$(pFile, Len(pFile) - (Len(GetFileName(pFile))) - 1))
    
  Dim EXEName As String
  EXEName = Path32 & "\" & ExeName32
  EXEName = Replace(EXEName, "\\", "\")
  
  'Check to see if the project already exists in the list
  If UBound(mProjects()) > 0 Then
    Dim X As Integer
    For X = 0 To UBound(mProjects)
      If LCase$(mProjects(X).ProjectFullPath) = LCase(pFile) Then
        Exit Sub
      End If
    Next
  End If
  
  GoTo 10
  
Err:
  
  'This indicates an empty array
  ReDim Preserve mProjects(0)
  
  GoTo 20
  
10:
  
  If lstProj.ListItems.Count = 0 Then
    ReDim Preserve mProjects(0)
  Else
    ReDim Preserve mProjects(UBound(mProjects) + 1)
  End If
  
20:

  mProjects(UBound(mProjects)).ProjectFullPath = pFile
  mProjects(UBound(mProjects)).ProjectName = GetFileName(mProjects(UBound(mProjects)).ProjectFullPath)
  
  mProjects(UBound(mProjects)).ExeFullPath = EXEName
  mProjects(UBound(mProjects)).EXEName = LCase$(GetFileName(mProjects(UBound(mProjects)).ExeFullPath))
  
End Sub

Function GetRelativePath(findPath As String, startPath As String) As String
  
  Dim L As Integer
  Dim I As Integer
  Dim Backs As Integer
  
  'Find out how many BackDirs (..\) there are
  L = Len(findPath)
  findPath = Replace(findPath, "..\", "")
  Backs = (L - Len(findPath)) / 3
  
  'Back up BACKS BackDirs
  For I = 1 To Backs
    If I = 1 Then
      L = InStrRev(startPath, "\")
    Else
      L = InStrRev(startPath, "\", L - 1)
    End If
    startPath = Left(startPath, L - 1)
  Next
  
  GetRelativePath = startPath & "\" & findPath
  
End Function

Sub Dostatus(vProjName As String, vEXEName As String, vPercent As Single)
  If Not (framCompile.Visible) Then Exit Sub
  
  lblProject = vProjName
  lblEXE = vEXEName
  ProgressBar1.Value = Int(vPercent)
End Sub

Private Sub mnuCompile_Click()
  Dim X As Integer
  Dim I As Integer
  Dim CMD As String
  
  'Disable Controls
  AllControlsEnabled (False)
  
  DoEvents

  'Go Through List
  For X = 1 To lstProj.ListItems.Count
    
    'Find Array's Index Value
    I = Val(Replace(LCase$(lstProj.ListItems(X).Key), "key:", ""))
    
    'Build the command to shell
    CMD = VBPath & " /make " & Chr(34) & mProjects(I).ProjectFullPath & Chr(34) & " " & Chr(34) & mProjects(I).ExeFullPath & Chr(34)
    
    'Update Status Frame
    Call Dostatus(mProjects(I).ProjectFullPath, mProjects(I).ExeFullPath, (X / lstProj.ListItems.Count) * 100)
    
    'Shell the command
    Call Shell(CMD)
    
    DoEvents
  Next

  'Re-Enable Controls
  AllControlsEnabled (True)

End Sub

Sub AllControlsEnabled(vEnabled As Boolean)
  mnuFile.Enabled = vEnabled
  mnuCompile.Enabled = vEnabled
  lstProj.Visible = vEnabled
  framCompile.Visible = Not (vEnabled)
  If vEnabled = False Then
    mnuQSave.Enabled = False
  Else
    mnuQSave.Enabled = mChanged
  End If
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuNew_Click()
  
  SaveLocation = ""
  
  'Clear List
  lstProj.ListItems.Clear
  
  'Reset Caption
  DoCaption ("")
  
  'Clear Projects Array
  ReDim mProjects(0)
  
  'Set Changed=False
  Call DoChanged(False)
  
End Sub

Private Sub mnuOpen_Click()
  On Error GoTo Err
  CD.CancelError = True
  CD.ShowOpen
  Call OpenProfile(CD.FileName)
Err:
End Sub

Sub OpenProfile(vFile As String)
  On Error GoTo Err
  
  If Dir(vFile) = "" Then Exit Sub
  
  Dim Count As Integer
  Dim fNum As Integer
  fNum = FreeFile()
  
  'Simple File read to populate Data
  Dim Readln(0 To 2) As String
  Open vFile For Input As #fNum
    Do While Not EOF(fNum)
      
      If Count = 0 Then
        ReDim mProjects(Count)
      Else
        ReDim Preserve mProjects(Count)
      End If
      
      Line Input #1, mProjects(Count).ProjectFullPath
      mProjects(Count).ProjectName = GetFileName(mProjects(Count).ProjectFullPath)
      
      Line Input #1, mProjects(Count).ExeFullPath
      mProjects(Count).EXEName = GetFileName(mProjects(Count).ExeFullPath)
      
      Count = Count + 1
      
    Loop
  Close #fNum
  
  Call PopulateProjects
  
  'Update Caption
  DoCaption Left(GetFileName(vFile), Len(GetFileName(vFile)) - 4)
  
  'Save Settings
  SaveLocation = vFile
  SaveSetting "BatchCompile", "Settings", "LastProfile", vFile
  
  Call DoChanged(False)
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Sub SaveProfile(Optional vSaveAs As Boolean)
  On Error GoTo Err
  If SaveLocation = "" Or vSaveAs Then
    CD.CancelError = True
    CD.ShowSave
    SaveLocation = CD.FileName
  End If
  
  Dim fNum As String
  fNum = FreeFile()
  
  'Check for over-write
  Dim Answer As Integer
  If vSaveAs And Dir(SaveLocation) <> "" Then
    Answer = MsgBox("The file already exists. OK to over-write?", vbYesNo + vbQuestion, "Over-Write?")
  ElseIf Not vSaveAs Then
    Answer = vbYes
  End If
  
  'Should we delete the existing file?
  If Answer = vbYes Then
    If Dir(SaveLocation) <> "" Then Kill SaveLocation
  Else
    Exit Sub
  End If
   
  'Simple File Write to save settings
  Dim X As Integer
  Open SaveLocation For Output As #fNum
    For X = 0 To UBound(mProjects)
      Print #fNum, mProjects(X).ProjectFullPath
      Print #fNum, mProjects(X).ExeFullPath
    Next
  Close #fNum
  
  'Update Caption
  DoCaption Left(GetFileName(SaveLocation), Len(GetFileName(SaveLocation)) - 4)
  
  'SaveSettings
  SaveSetting "BatchCompile", "Settings", "LastProfile", SaveLocation
  
  Call DoChanged(False)
Err:
End Sub

Sub DoChanged(vChanged As Boolean)
  mChanged = vChanged
  mnuQSave.Enabled = vChanged
End Sub

Sub PopulateProjects()
  Dim X As Integer
  Dim LI As ListItem
  
  lstProj.ListItems.Clear
  
  'populate listview with array data
  For X = 0 To UBound(mProjects())
    Set LI = lstProj.ListItems.Add(X + 1, "Key:" & X, mProjects(X).ProjectName)
    LI.SubItems(1) = mProjects(X).EXEName
  Next
End Sub

Private Sub mnuQSave_Click()
  mnuSave_Click
End Sub

Private Sub mnuSave_Click()
  SaveProfile
End Sub

Private Sub mnuSaveAs_Click()
  SaveProfile (True)
End Sub

Sub DoCaption(vText As String)
  If vText = "" Then
    Caption = " Batch VB6 Compiler"
  Else
    Caption = " Batch VB6 Compiler - " & vText
  End If
End Sub
