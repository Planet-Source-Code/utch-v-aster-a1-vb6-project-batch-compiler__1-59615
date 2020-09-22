Attribute VB_Name = "mCompiler"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public mProjects() As vProjects
Dim fso As New FileSystemObject
Dim fld As Folder

Public VBPath As String
Public CancelVBSearch As Boolean
Public VBSearchDone As Boolean
Public mCancelFileGet As Boolean

Public Type vProjects
   ProjectName     As String
   ProjectFullPath As String
   EXEName         As String
   ExeFullPath     As String
End Type

Public Function FindFile(sDir As String, sFilename As String, sForm As Form) As String
    
    Dim sCurPath As String
    Dim sName As String
    Dim colFiles As Collection
    Dim sLastDir As String
    
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
  
    mCancelFileGet = False
    sCurPath = sDir
    sName = vbNullString
    sName = Dir(sCurPath, vbDirectory)
    
    Do While Not (mCancelFileGet) And sName <> vbNullString  'A nullstring is returned when no files/dirs are left
        
      DoEvents
      If (GetAttr(sCurPath & sName) And vbDirectory) = vbDirectory Then
      
        If sName <> "." And sName <> ".." Then
          sLastDir = sName
          FindFile = FindFile(sCurPath & sName & "\", sFilename, sForm)
          sName = Dir(sCurPath, vbDirectory)
          While sName <> sLastDir And VBPath = ""
              sName = Dir
          Wend
        End If
        
      Else
      
        sForm.Label1 = sCurPath & sName
        If LCase$(sName) = LCase$(sFilename) Then
          mCancelFileGet = True
          VBPath = sCurPath & sName
          FindFile = VBPath
          Exit Function
        End If
        
      End If
      sName = Dir   'Get next file/dir in the loop
      
   Loop
   
End Function

Function GetFileName(vPath As String) As String
  Dim Spot As Integer
  Dim lSpot As Integer
  
  Spot = InStr(Spot + 1, vPath, "\")
  Do Until Spot = 0
    lSpot = Spot
    Spot = InStr(Spot + 1, vPath, "\")
  Loop
  
  If lSpot <= 0 Then
    If Len(Trim(vPath)) = 0 Then
      GetFileName = ""
    Else
      GetFileName = vPath
    End If
  Else
    GetFileName = Mid(vPath, lSpot + 1)
  End If
End Function

Function GetDirectoryName(vPath As String) As String
  Dim Spot As Integer
  Dim lSpot As Integer
  
  Spot = InStr(Spot + 1, vPath, "\")
  Do Until Spot = 0
    lSpot = Spot
    Spot = InStr(Spot + 1, vPath, "\")
  Loop
  
  If lSpot <= 0 Then
    GetDirectoryName = vPath
  Else
    GetDirectoryName = Left(vPath, lSpot + 1)
  End If
End Function

Public Sub Pause(Seconds As Single)
    
    Dim T As Single
    Dim T2 As Single
    Dim Num As Single
    
    Num = Seconds * 1000
    
    T = GetTickCount()
    T2 = GetTickCount()
    
    Do Until T2 - T >= Num
        DoEvents
        T2 = GetTickCount()
    Loop
    
End Sub
