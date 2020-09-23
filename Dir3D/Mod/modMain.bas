Attribute VB_Name = "modMain"
Option Explicit

Dim idx                     As Integer

Global bp                   As Long

Dim chrsout                 As String
Dim chrsin                  As String

Global flDirPath            As String
Global m_Mediadir           As String
Global m_strIn              As String
Global b_Path               As New Collection
Global c_Path               As String

Global i_sOver              As Boolean

Global m_BakMesh            As New CD3DMesh

Global m_dframe()           As CD3DFrame

Global DriveInf             As New drvRecords

Global Animation            As CD3DAnimation
Global m_Object             As CD3DMesh

Global drObject             As Boolean
Global dirObject            As Boolean
Global drvObject            As Boolean
Global gHom                 As Boolean
Global gBak                 As Boolean
Global unKey                As Boolean

Global nStrtX               As Single
Global spcMade              As Single

Function CheckSub(foldername As String) As String
       Dim nStr As String
       Dim Path As String
       
       On Error GoTo aired
       
       If Right(foldername, 1) <> "\" Then foldername = foldername & "\"
       
       Dim FSO, f, fc, f1
       Set FSO = CreateObject("Scripting.FileSystemObject")
       Set f = FSO.GetFolder(foldername)
       Set fc = f.SubFolders
       
       nStr = ","
       
       For Each f1 In fc
           nStr = nStr & foldername & f1.Name & ","
       Next
       
       CheckSub = Mid(nStr, 2)
       
       If CheckSub = "" Then CheckSub = foldername
       
       Set fc = Nothing
       Set f = Nothing
       Set FSO = Nothing
       
       Exit Function
       
aired:
       Set fc = Nothing
       Set f = Nothing
       Set FSO = Nothing
       
       MsgBox "Error: (modMain - CheckSub) doing: " & foldername
       
End Function

Sub Wait(WaitSeconds As Single)

Dim StartTime As Single

StartTime = Timer

Do While Timer < StartTime + WaitSeconds
DoEvents
Loop
End Sub

Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    If Right(strDirName, 1) <> "\" Then strDirName = strDirName & "\"
    
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
End Function

Function doNewStart(Drive As Integer) As Single

   Select Case Drive
     Case 2
      doNewStart = -3
      
     Case 3
      doNewStart = -4
      
     Case 4
      doNewStart = -5
    
     Case 5
      doNewStart = -6
     
     Case 6
      doNewStart = -8
      
     Case 7
      doNewStart = -10
      
     Case 8
      doNewStart = -12
      
     Case 9
      doNewStart = -14
      
     Case 10
      doNewStart = -16
      
     Case Else
      doNewStart = -18
     
    End Select

End Function

Function doNewLeft(dCount As Variant)
  
  Select Case dCount
     Case 0 To 9
       nStrtX = -9
     frmMain.makeBase 20, 20
    
     Case 10 To 14
       nStrtX = -12.5
     frmMain.makeBase 30, 20
     
     Case 15 To 20
       nStrtX = -18
     frmMain.makeBase 45, 20
     
     Case 21 To 25
       nStrtX = -22
     frmMain.makeBase 45, 20
     
     Case 26 To 30
     MsgBox "under 30"
       nStrtX = -26.5
     frmMain.makeBase 50, 20
     
     Case 31 To 35
     MsgBox "under 35"
       nStrtX = -30.5
     frmMain.makeBase 60, 20
     
     Case 36 To 40
     MsgBox "under 40"
       nStrtX = -35.5
     frmMain.makeBase 70, 20
     
     Case 41 To 45
     nStrtX = -41.5
     frmMain.makeBase 85, 20
     
     Case 46 To 50
     MsgBox "under 50"
     nStrtX = -46.5
     frmMain.makeBase 92, 20
     
     Case 51 To 55
     MsgBox "under 55"
     nStrtX = -50
     frmMain.makeBase 100, 20
     
     Case 56 To 60
     MsgBox "under 60"
     nStrtX = -55
     frmMain.makeBase 110, 20
     
     Case 61 To 65
     MsgBox "under 65"
     nStrtX = -60
     frmMain.makeBase 120, 20
     
     Case Else
     nStrtX = -70
     frmMain.makeBase 140, 20
     
    End Select
    
  
End Function

Function makeSpace(Inn As Long)
 Dim tNum As Single
 
 tNum = (Inn / 10) / 2
 
 Select Case Inn
  Case 0 To 9
   spcMade = 2
  Case 10 To 14
   spcMade = 2.05
  Case 15 To 19
   spcMade = 2.1
  Case 20 To 24
   spcMade = 2.15
  Case 25 To 29
   spcMade = 2.28
  Case 30 To 34
   spcMade = 2.35
  Case 35 To 39
   spcMade = 2.5
  Case 40 To 44
   spcMade = 2.7
  Case 45 To 49
   spcMade = 2.9
  Case 50 To 54
   spcMade = 3
  Case 55 To 59
   spcMade = 3.1
  Case 60 To 64
   spcMade = 3.2
  Case 65 To 69
   spcMade = 3.3
  Case 70 To 74
   spcMade = 3.4
  Case 75 To 79
   spcMade = 3.5
  Case 80 To 84
   spcMade = 3.6
  Case 85 To 89
   spcMade = 3.7
  Case 90 To 94
   spcMade = 3.8
  Case 95 To 100
   spcMade = 4
End Select
 
End Function
