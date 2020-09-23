VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8700
   ClientLeft      =   2025
   ClientTop       =   510
   ClientWidth     =   10515
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   701
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu xx 
      Caption         =   "hid"
      Visible         =   0   'False
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuRot 
         Caption         =   "Rotate"
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpeed 
         Caption         =   "Speed it Up"
         Checked         =   -1  'True
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHdr 
      Caption         =   "Drv"
      Visible         =   0   'False
      Begin VB.Menu mnuGo 
         Caption         =   "Explore"
      End
      Begin VB.Menu mnuTip 
         Caption         =   "Tip"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type D3DBLENDVERTEX
    v As D3DVECTOR
    blend As Single
    n As D3DVECTOR
    tu As Single
    tv As Single
End Type

Const D3DFVF_BLENDVERTEX = (D3DFVF_XYZB1 Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Private Enum CONST_D3DPOOL
    D3DPOOL_DEFAULT = 0
    D3DPOOL_MANAGED = 1
    D3DPOOL_SYSTEMMEM = 2
End Enum

Dim m_Showtip            As Boolean
Dim m_binit              As Boolean
Dim m_bGraphInit         As Boolean
Dim m_bMinimized         As Boolean
Dim m_bRot               As Boolean
Dim m_bShowBase          As Boolean
Dim m_drawtextEnable     As Boolean
Dim m_drawpropEnable     As Boolean
Dim m_bMouseDown         As Boolean
Dim lBut                 As Boolean
Dim rBut                 As Boolean
Dim m_bKey(256)          As Boolean
Dim m_Faster             As Boolean
Dim gExit                As Boolean
Dim StartIT              As Boolean

Dim m_maxX               As Double
Dim m_minX               As Double
Dim m_maxY               As Double
Dim m_minY               As Double
Dim m_maxZ               As Double
Dim m_minZ               As Double
Dim m_maxsize            As Double
Dim m_minSize            As Double
Dim m_extX               As Double
Dim m_extY               As Double
Dim m_extZ               As Double
Dim m_extSize            As Double

Dim m_scalex             As Single
Dim m_scaley             As Single
Dim m_scalez             As Single
Dim m_scalesize          As Single
Dim m_lastX              As Single
Dim m_lasty              As Single
Dim m_fYawVelocity       As Single
Dim m_fPitchVelocity     As Single
Dim m_fElapsedTime       As Single
Dim m_fYaw               As Single
Dim m_fPitch             As Single

Dim m_HDFrame()          As CD3DFrame
Dim m_CDFrame()          As CD3DFrame
Dim m_FPFrame()          As CD3DFrame
Dim m_DR1Frame()         As CD3DFrame
Dim m_DR2Frame()         As CD3DFrame
Dim m_CTRLFrame()        As CD3DFrame
Dim m_drFrame            As CD3DFrame

Dim m_graphroot          As CD3DFrame
Dim m_fileroot           As CD3DFrame
Dim m_plor               As CD3DFrame
Dim m_door               As CD3DFrame

Dim m_XZPlaneFrame       As CD3DFrame
Dim m_XZDriveFrame       As CD3DFrame
Dim m_XZDriveSpcFrame    As CD3DFrame
Dim o_Frame              As CD3DFrame
Dim m_LabelX             As CD3DFrame
Dim m_LabelY             As CD3DFrame
Dim m_LabelZ             As CD3DFrame
Dim m_Name()             As CD3DFrame

Dim m_drawtext           As String
Dim m_drawprop           As String
Dim m_xHeader            As String
Dim m_yHeader            As String
Dim m_zHeader            As String
Dim m_Time               As String
Dim m_sizeHeader         As String
Dim m_sizex              As Single
Dim m_sizez              As Single

Dim m_drawtextpos        As RECT
Dim m_drawproppos        As RECT
Dim m_drawtimepos        As RECT

Dim m_data               As Collection

Dim m_hwnd               As Long
Dim m_font2height        As Long

Dim fc                   As Integer
Dim hc                   As Integer
Dim cc                   As Integer
Dim dc                   As Integer
Dim dR1                  As Integer
Dim dR2                  As Integer
Dim Cntc                 As Integer

Dim m_font               As D3DXFont
Dim m_font2              As D3DXFont

Dim m_vbfont             As IFont
Dim m_vbfont2            As IFont

'
Dim m_Tex                As Direct3DTexture8   ' watch for this
'
Dim m_meshdrivespace     As D3DXMesh
Dim m_meshplane          As D3DXMesh
Dim m_meshdriveplane     As D3DXMesh
Dim m_meshdrawerplane    As D3DXMesh

Dim m_meshdrvboxplane    As CD3DMesh
Dim m_meshdrivespace2    As D3DXMesh

Dim m_vPosition          As D3DVECTOR
Dim m_vVelocity          As D3DVECTOR

Dim m_matView            As D3DMATRIX
Dim m_matOrientation     As D3DMATRIX

Const kdx = 256&
Const kdy = 256&
Const kScale = 8

Dim m_labelmesh()        As CD3DMesh
Dim m_LabelTex()         As Direct3DTexture8

Const D3DFVF_VERTEX = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Function doLook()
    Dim whrAt As D3DVECTOR
    Dim vT As D3DVECTOR, vTemp As D3DVECTOR
    
If i_sOver And drObject Then
    
    m_fPitch = m_fPitch + 1
    
    whrAt = o_Frame.GetPosition
    
    If (whrAt.Y < 1) Then whrAt.Y = 1
    
    

    Dim qR As D3DQUATERNION, det As Single
    D3DXQuaternionRotationYawPitchRoll qR, m_fYaw, m_fPitch, 0
    D3DXMatrixAffineTransformation m_matOrientation, 1.25, vec3(0, 0, 0), qR, m_vPosition
    D3DXMatrixInverse m_matView, det, m_matOrientation
    
        'set new view matrix
    g_dev.SetTransform D3DTS_VIEW, m_matView
    
    m_vPosition = vec3(whrAt.X, (whrAt.Y + 1.7), (whrAt.z - 1.6))
End If

End Function

Function MakeDrawer(h As Single, w As Single, d As Single)
  Set m_meshdrawerplane = CreateBoxWithTextureCoords(w, h, d, False)
End Function

Function goHome()

    m_fYaw = 0
    m_fPitch = 0
    
    m_vPosition = vec3(-1, 6, -42)
    
End Function

Sub DestroyDeviceObjects()

    Set m_graphroot = Nothing
    Set m_XZPlaneFrame = Nothing
    Set m_XZDriveFrame = Nothing
    
End Sub


Friend Sub Init(hwnd As Long, font As IFontDisp, font2 As IFontDisp)
    Dim i As Long
    
    '
    '  drive counts
    fc = 0
    hc = 0
    cc = 0
    dR1 = 0
    dR2 = 0
    Cntc = 0
    bp = 0
    
    'Save hwnd
    m_hwnd = hwnd
    
    'convert IFontDisp to Ifont
    Set m_vbfont = font
    Set m_vbfont2 = font2
    
    'initialized d3d
    m_binit = D3DUtil_Init(hwnd, True, 0, 0, D3DDEVTYPE_HAL, Nothing)
        
    'exit if initialization failed
    If m_binit = False Then End
     
    m_bRot = False
    
    D3DXMatrixTranslation m_matOrientation, 0, 0, 0
    
    m_xHeader = Space(17) & "HomePlay Entertainment"
    m_yHeader = ""
    m_zHeader = "Used Space"
        
    m_vPosition = vec3(-1, 6, -42)

    m_sizex = 1
    m_sizez = 1

    m_bShowBase = True
    
    DeleteDeviceObjects
    InitDeviceObjects 30, 20
    RestoreDeviceObjects
    BuildBase
    BuildDrives
    
    'Sound3.Play DSBPLAY_DEFAULT
    
    DoEvents

    'Initialze camera matrices
    g_dev.GetTransform D3DTS_VIEW, m_matView
  
End Sub

Public Sub DrawDir()
    Dim hr As Long
    Dim rc As RECT
        
    If m_binit = False Then Exit Sub
    
    'See what state the device is in.
    hr = g_dev.TestCooperativeLevel
    If hr = D3DERR_DEVICENOTRESET Then
        g_dev.Reset g_d3dpp
        RestoreDeviceObjects
    End If
             
    'Clear the previous render with the backgroud color
    'We clear to grey but notice that we are using a hexidecimal
    'number to represent Alpha Red Green and blue
    D3DUtil_ClearAll &HFF808080
    
    m_graphroot.UpdateFrames
 
    'set the ambient lighting level
    g_dev.SetRenderState D3DRS_AMBIENT, &HFFC0C0C0
    
    g_dev.BeginScene
    
    
    DrawAxisNameSquare 0
    DrawAxisNameSquare 2

    'draw time text
    m_drawtimepos.Top = 0: m_drawtimepos.bottom = 100: m_drawtimepos.Left = (Me.ScaleWidth - 70): m_drawtimepos.Right = 300
    m_font.Begin
        g_d3dx.DrawText m_font, &HFF000000, format(time, "h:mm AM/PM"), m_drawtimepos, 0
    m_font.End
    
    If m_drawtextEnable Then
        m_font.Begin
        g_d3dx.DrawText m_font, &HFF000000, m_drawtext, m_drawtextpos, 0
        m_font.End
    End If
    
    'draw pop up property text
    If m_drawpropEnable Then
        m_font.Begin
        g_d3dx.DrawText m_font, &HFF000000, m_drawprop, m_drawproppos, 0
        m_font.End
    End If


    'render the xzplane with transparency
    If m_bShowBase Then
        m_XZPlaneFrame.Enabled = True
        m_graphroot.Render g_dev
    End If
    
       ' m_XZDriveFrame.Enabled = True
        m_graphroot.Render g_dev
        g_dev.EndScene
        
        
    Dim surf As Direct3DSurface8
    Dim rts As D3DXRenderToSurface
    Dim rtsviewport As D3DVIEWPORT8
    Dim d3ddm As D3DDISPLAYMODE
    
    Set surf = m_Tex.GetSurfaceLevel(0)
  
    rtsviewport.height = kdx
    rtsviewport.width = kdy
    rtsviewport.MaxZ = 1
    
    

    Call g_dev.GetDisplayMode(d3ddm)
    Set rts = g_d3dx.CreateRenderToSurface(g_dev, kdx, kdy, d3ddm.format, 1, D3DFMT_D16)
  
    rts.BeginScene surf, rtsviewport
    g_dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFFC0C0C0, 1, 0
    
    rc.Top = m_font2height * 0: rc.Left = 0: rc.bottom = 0: rc.Right = 0
    g_d3dx.DrawText m_font2, &HFF000000, m_xHeader, rc, DT_CALCRECT Or DT_NOCLIP
    g_d3dx.DrawText m_font2, &HFF000000, m_xHeader, rc, 0
    
    rts.EndScene
    
    D3DUtil_PresentAll m_hwnd

End Sub

Public Sub BuildBase()
   If Not m_binit Then Exit Sub
   
   Dim whitematerial As D3DMATERIAL8
   Dim m_Textr As Direct3DTexture8
   Dim material As D3DMATERIAL8
   Dim d3ddm As D3DDISPLAYMODE
   
   Dim newFrame As CD3DFrame
   Dim tFrame As CD3DFrame
   Dim frameMesh As CD3DMesh
   Dim rc As RECT
   Dim surf As Direct3DSurface8
   Dim rts As D3DXRenderToSurface
   Dim rtsviewport As D3DVIEWPORT8
   Dim i As Long, j As Long
   Dim w As Single, h As Single
   Dim sv As Single, ev As Single
   Dim su As Single, eu As Single
   Dim b As Boolean
   Dim drvStr As String
   
   Set m_graphroot = Nothing
    
    'Create rotatable root object
   Set m_graphroot = D3DUtil_CreateFrame(Nothing)
      
    'Create XZ plane for reference
      material.diffuse = LONGtoD3DCOLORVALUE(&H6FC0C0C0)
      material.Ambient = material.diffuse
    
    Set m_XZPlaneFrame = D3DUtil_CreateFrame(m_graphroot)
      m_XZPlaneFrame.AddD3DXMesh(m_meshplane).SetMaterialOverride material
      m_XZPlaneFrame.SetOrientation D3DUtil_RotationAxis(1, 0, 0, 90)
      m_XZPlaneFrame.MeshNumber = 2
      m_XZPlaneFrame.ObjectType = isBase
      m_XZPlaneFrame.ObjectName = "MAIN"
    
    Set m_XZDriveFrame = D3DUtil_CreateFrame(m_graphroot)
      m_XZDriveFrame.AddD3DXMesh(m_meshdriveplane).SetMaterialOverride material
      m_XZDriveFrame.SetPosition vec3(-20, 6, 3)
      m_XZDriveFrame.MeshNumber = 3
      m_XZDriveFrame.ObjectName = "Drive Properties"
      m_XZDriveFrame.ObjectType = isBase
      
    MakeComponent 6, "Show Drives"
    
     m_CTRLFrame(0).SetScale 1
     m_CTRLFrame(0).SetPosition vec3(-16, -3.5, -7)
     m_CTRLFrame(0).SetOrientation D3DUtil_RotationAxis(0, -1, 0, 30)
     m_graphroot.AddChild m_CTRLFrame(0)
     
    MakeComponent 6, "Back"
    
     m_CTRLFrame(1).SetScale 1
     m_CTRLFrame(1).SetPosition vec3(-10, -3.5, -11.5)
     m_CTRLFrame(1).SetOrientation D3DUtil_RotationAxis(0, 1, 0, 90)
     m_graphroot.AddChild m_CTRLFrame(1)
     
    MakeComponent 7, ""
    
     m_plor.SetScale 1
     m_plor.SetPosition vec3(0, 15, 9)
     m_graphroot.AddChild m_plor
     
    MakeComponent 8, "Exit"
    
     m_door.SetScale 1
     m_door.SetPosition vec3(12, -3.5, -11.5)
     m_graphroot.AddChild m_door
     
      Dim nStrt As Single

    If hDrvCount = 1 Then
     nStrt = -19
    ElseIf hDrvCount = 2 Then
     nStrt = -20.5
    ElseIf hDrvCount = 3 Then
     nStrt = -22
    ElseIf hDrvCount = 4 Then
     nStrt = -23.5
    End If
    
    For i = 1 To DriveInf.count
    
    If DriveInf.item(i).dType = 3 Then
    makeSpace DriveInf.item(i).UsedPercent
    DoEvents
    
    material.diffuse = LONGtoD3DCOLORVALUE(-5098522)
    material.Ambient = material.diffuse
    Set m_XZDriveSpcFrame = D3DUtil_CreateFrame(m_graphroot)
    m_XZDriveSpcFrame.AddD3DXMesh(m_meshdrivespace).SetMaterialOverride material
    m_XZDriveSpcFrame.SetOrientation D3DUtil_RotationAxis(0, 0, 0, 90)
    m_XZDriveSpcFrame.SetPosition vec3(nStrt, spcMade, 2.5)
    m_XZDriveSpcFrame.MeshNumber = 4
    m_XZDriveSpcFrame.ObjectName = DriveInf.item(i).Name
    m_XZDriveSpcFrame.ObjectType = isDrive
    Set m_XZDriveSpcFrame = Nothing
    
    material.diffuse = LONGtoD3DCOLORVALUE(-7444444)
    material.Ambient = material.diffuse
    
    Set m_XZDriveSpcFrame = D3DUtil_CreateFrame(m_graphroot)
    m_XZDriveSpcFrame.AddD3DXMesh(m_meshdrivespace2).SetMaterialOverride material
    m_XZDriveSpcFrame.SetOrientation D3DUtil_RotationAxis(0, 0, 0, 90)
    m_XZDriveSpcFrame.SetPosition vec3(nStrt, (6.5), 2.7)
    m_XZDriveSpcFrame.MeshNumber = 5
    m_XZDriveSpcFrame.ObjectType = isNone
    m_XZDriveSpcFrame.ObjectName = ""
    Set m_XZDriveSpcFrame = Nothing
    
    End If
    
    nStrt = nStrt + 1.5
    
    Next
     
     

    Call g_dev.GetDisplayMode(d3ddm)
    Set rts = g_d3dx.CreateRenderToSurface(g_dev, kdx, kdy, d3ddm.format, 1, D3DFMT_D16)
    rtsviewport.height = kdx
    rtsviewport.width = kdy
    rtsviewport.MaxZ = 1
    
    Set surf = m_Tex.GetSurfaceLevel(0)
          
    rts.BeginScene surf, rtsviewport
    g_dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFFC0C0C0, 1, 0
    
    rc.Top = 0: rc.Left = 0: rc.bottom = 0: rc.Right = 0
    g_d3dx.DrawText m_font2, &HFF000000, "XXX", rc, DT_CALCRECT ' DT_SINGLELINE Or DT_VCENTER Or DT_CENTER Or DT_CENTER

    m_font2height = rc.bottom
    
    rts.EndScene
    
    Set m_LabelX = D3DUtil_CreateFrame(m_graphroot)
    m_LabelX.SetPosition vec3(0, 0, -11)
    
    'Set m_LabelZ = D3DUtil_CreateFrame(m_graphroot)
    'm_LabelZ.SetPosition vec3(-18, 3.3, -1)
    
End Sub

Public Sub BuildDrives()
    Dim material As D3DMATERIAL8
    Dim rtnStr() As String
    Dim rtnStr2() As String
    Dim StrIn2 As String
    Dim mNewStart As Single
    Dim q As Integer, i As Long, h As Integer
    
    nStrtX = -5.5
    
    fc = 0
    hc = 0
    cc = 0
    dR1 = 0
    dR2 = 0
    h = 0
    
    If Not m_binit Then Exit Sub
    
    ' kill any frames showing
    On Error Resume Next
    For i = 0 To UBound(m_dframe) - 1
      If Not m_dframe(i) Is Nothing Then m_dframe(i).Destroy
    Next
    
     ' kill any drawers showing
    On Error Resume Next
    For i = 0 To UBound(m_Name) - 1
      If Not m_Name(i) Is Nothing Then m_Name(i).Destroy
    Next
    
    
    ReDim m_dframe(0 To (DriveInf.count - 1) * 2.5)
    ReDim m_Name(0 To (DriveInf.count - 1) * 2.5)
    
    '
    ' try to center drives
    '
    mNewStart = doNewStart(DriveInf.count)
    
    Dim c As Long
    c = 0
    
    For i = 0 To DriveInf.count - 1
    
    mNewStart = mNewStart + 2

      m_strIn = DriveInf.item(i + 1).Name
      StrIn2 = DriveInf.item(i + 1).dType
      Select Case StrIn2
       
       Case 2  ' flop
        MakeComponent 0, m_strIn
        m_FPFrame(fc - 1).SetOrientation D3DUtil_RotationAxis(0, 1, 0, 180)
        Set m_dframe(c) = m_FPFrame(fc - 1)
        m_dframe(c).SetPosition vec3(mNewStart, 0.75, -9.75)
        mNewStart = mNewStart - 0.5
      
       Case 3   ' hd
        MakeComponent 1, m_strIn
        Set m_dframe(c) = m_HDFrame(hc - 1)
        m_dframe(c).SetPosition vec3(mNewStart, 0.3, -7.75)
        h = h + 1
      
       Set m_Name(c) = CreateSheetWithTextureCoords(1, 0.5, 0, 1, 0, 1, Nothing)
      m_Name(c).SetPosition vec3(mNewStart + 0.8, 0.9, -9.9)
      m_Name(c).ObjectName = m_strIn
      m_Name(c).ObjectType = isNamePlate
      m_graphroot.AddChild m_Name(c)
       
       Case 5   ' cd
        MakeComponent 2, m_strIn
        mNewStart = mNewStart + 0.5
        Set m_dframe(c) = m_CDFrame(cc - 1)
        m_dframe(c).SetPosition vec3(mNewStart, 0.3, -7.5)
      
       Set m_Name(c) = CreateSheetWithTextureCoords(1, 0.5, 0, 1, 0, 1, Nothing)
      m_Name(c).SetPosition vec3(mNewStart + 0.75, 0.9, -9.9)
      m_Name(c).ObjectName = m_strIn
      m_Name(c).ObjectType = isNamePlate
      m_graphroot.AddChild m_Name(c)
       
      End Select
    
        m_dframe(c).SetScale 1
        m_dframe(c).MeshNumber = 6
        m_graphroot.AddChild m_dframe(c)
    
        If StrIn2 = 2 Then mNewStart = mNewStart + 0.5
        If StrIn2 = 5 Then mNewStart = mNewStart - 0.5
        
        c = c + 1
    
    Next
End Sub

Function BuildDir()
    If Not m_binit Then Exit Function
    
    If m_strIn = "" Then Exit Function
    
    Screen.MousePointer = vbHourglass
    
    fc = 0
    hc = 0
    cc = 0
    dR1 = 0
    dR2 = 0
    
    Dim rtnStr() As String
    Dim rtnStr2() As String
    Dim StrIn2 As String
    Dim mNewStart As Single
    Dim q As Integer, i As Long
    
    ' kill any frames showing
    On Error Resume Next
    For i = 0 To UBound(m_dframe) - 1
      m_dframe(i).Destroy
    Next
    
    ' kill any drawers showing
    On Error Resume Next
    For i = 0 To UBound(m_Name) - 1
      m_Name(i).Destroy
    Next
    
    
    i = 0
    
    rtnStr() = Split(m_strIn, ",")
    
    '
    ' builds the base and tris to center dirs
    doNewLeft UBound(rtnStr())
    
    ReDim m_DR1Frame(0)
    ReDim m_DR2Frame(0)
    
    ReDim m_DR1Frame(0 To 400)
    ReDim m_DR2Frame(0 To 100)
    
    q = 0
    
    Dim nxt As Long
    nxt = -1
    
    '
    '  first off,  i tried, like you would not beleive
    '  to redim the m_dframe, on each new entry, for days... , an just got errors
    '  i figure the best way,  the only way i could figure, to clear all dirs and respective
    '  drawers was in a frame array ?
    '  if the count ends up to more than 500 directories and sub dirs, it'll crap out
    '  mine went to 379 on one drive, i figure 500 safe
    '
    ReDim m_dframe(0)
    ReDim m_dframe(0 To 500)
    
    ReDim m_Name(0)
    ReDim m_Name(0 To 500)
    
    DoEvents
    
    Do Until q = UBound(rtnStr())
    
    If UBound(rtnStr()) = 0 Then GoTo strtSubdir
    
    nxt = nxt + 1
    
    StrIn2 = CheckSub(rtnStr(q))
    
    nStrtX = nStrtX + 1.75
    
    rtnStr2() = Split(StrIn2, ",")
    
     MakeComponent 4, rtnStr(q)
     
     Set m_dframe(nxt) = m_DR2Frame(dR2 - 1)
        m_dframe(nxt).SetScale 1
        m_dframe(nxt).SetPosition vec3(nStrtX, 0.27, -8)
        m_graphroot.AddChild m_dframe(nxt)
        
     If InStr(UCase(rtnStr(q)), "RECYCLE") Then GoTo misd ' skip giving the recycle bin a drawer
     Set m_Name(nxt) = CreateSheetWithTextureCoords(1, 0.5, 0, 1, 0, 1, Nothing)
      m_Name(nxt).SetPosition vec3(nStrtX + 0.7, 0.73, -9.75)
      m_Name(nxt).ObjectName = rtnStr(q)
      m_Name(nxt).ObjectType = isNamePlate
      m_graphroot.AddChild m_Name(nxt)
     
       nxt = nxt + 1
       
misd:

  If m_Faster = False Then
            DrawDir
            Sound7.Play DSBPLAY_DEFAULT
          End If
    
     ' first floor drawer door
    
strtSubdir:
        
        For i = 0 To UBound(rtnStr2()) - 1
        
         If i = 0 Then
           mNewStart = 1.1125
         ElseIf i = 1 Then
           mNewStart = 1.95
         Else
           mNewStart = mNewStart + 0.66
         End If
        
          MakeComponent 3, rtnStr2(i)
          ' dir
        Set m_dframe(nxt) = m_DR1Frame(dR1 - 1)
          m_dframe(nxt).SetPosition vec3(nStrtX, mNewStart, -8)
          m_graphroot.AddChild m_dframe(nxt)
          
          ' drawer
        Set m_Name(nxt) = CreateSheetWithTextureCoords(1, 0.5, 0, 1, 0, 1, Nothing)
          m_Name(nxt).SetPosition vec3(nStrtX + 0.7, mNewStart + 0.4, -9.75)
          m_Name(nxt).ObjectName = rtnStr2(i)
          m_Name(nxt).ObjectType = isNamePlate
          m_graphroot.AddChild m_Name(nxt)
          
          nxt = nxt + 1
        mNewStart = mNewStart + 0.192
          
          If m_Faster = False Then
            DrawDir
            Sound7.Play DSBPLAY_DEFAULT
          End If
          
          Next
    
      q = q + 1
      
       If m_Faster = False Then
        DrawDir
       End If
      Loop
      
      Screen.MousePointer = vbDefault
      
    If m_Faster Then
       Sound6.Play DSBPLAY_DEFAULT
    End If
    
    ReDim rtnStr(0)
    ReDim rtnStr2(0)
    
End Function

Sub DeleteDeviceObjects()
    Set m_font = Nothing
    Set m_font2 = Nothing
End Sub

Sub RestoreDeviceObjects()

    g_lWindowWidth = Me.ScaleWidth
    g_lWindowHeight = Me.ScaleHeight
    D3DUtil_SetupDefaultScene
    
    D3DUtil_SetupCamera vec3(0, 5, -20), vec3(0, 0, 0), vec3(0, 1, 0)
    
    'allow the application to show both sides of all surfaces
    g_dev.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    'turn on min filtering since our text is often smaller
    'than original size
    g_dev.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    
    Set m_font = g_d3dx.CreateFont(g_dev, m_vbfont.hFont)
    Set m_font2 = g_d3dx.CreateFont(g_dev, m_vbfont2.hFont)
    
End Sub

Public Function makeBase(l As Single, w As Single)
 On Error Resume Next
 
 Dim material As D3DMATERIAL8
 
      material.diffuse = LONGtoD3DCOLORVALUE(&H6FC0C0C0)
      material.Ambient = material.diffuse
      
  Set m_meshplane = Nothing
  If Not m_XZPlaneFrame Is Nothing Then m_XZPlaneFrame.Destroy
  
  Set m_meshplane = g_d3dx.CreateBox(g_dev, l, w, 0.1, Nothing)
  
   Set m_XZPlaneFrame = D3DUtil_CreateFrame(m_graphroot)
      m_XZPlaneFrame.AddD3DXMesh(m_meshplane).SetMaterialOverride material
      m_XZPlaneFrame.SetOrientation D3DUtil_RotationAxis(1, 0, 0, 90)
      m_XZPlaneFrame.MeshNumber = 2
      m_XZPlaneFrame.ObjectType = isBase
      m_XZPlaneFrame.ObjectName = "MAIN"
      
End Function

Public Sub InitDeviceObjects(l As Single, w As Single)
    
    Dim d3ddm As D3DDISPLAYMODE
    
    If m_binit = False Then Exit Sub
    
    Set m_meshplane = Nothing
    Set m_meshdriveplane = Nothing
    Set m_meshdrivespace = Nothing
    Set m_Tex = Nothing
    Set m_font = Nothing
    Set m_font2 = Nothing
    
    doDrvInf
    
    Dim rc As RECT
    
    ReDim m_LabelTex(0 To 1)
    ReDim m_CTRLFrame(0 To 1)
    ReDim m_dframe(0 To 1)
    ReDim m_Name(0 To 1)
    ReDim m_FPFrame(0 To 4)
    ReDim m_HDFrame(0 To 10)
    ReDim m_CDFrame(0 To 4)
    ReDim m_DR1Frame(0 To 400)
    ReDim m_DR2Frame(0 To 400)
    
    m_Faster = True
    
    Call g_dev.GetDisplayMode(d3ddm)
    
    makeBase l, w
    
    Select Case DriveInf.count
     Case 1
      Set m_meshdriveplane = g_d3dx.CreateBox(g_dev, 2.5, 7.5, 0.1, Nothing)

     Case 2
      Set m_meshdriveplane = g_d3dx.CreateBox(g_dev, 3.5, 7.5, 0.1, Nothing)
      
     Case 3
      Set m_meshdriveplane = g_d3dx.CreateBox(g_dev, 4, 7.5, 0.1, Nothing)
      
     Case 4
      Set m_meshdriveplane = g_d3dx.CreateBox(g_dev, 4.5, 7.5, 0.1, Nothing)
      
     Case 5
      Set m_meshdriveplane = g_d3dx.CreateBox(g_dev, 5, 7.5, 0.1, Nothing)
      
     Case Else
      Set m_meshdriveplane = g_d3dx.CreateBox(g_dev, 7, 7.5, 0.1, Nothing)
     End Select
        
    Set m_meshdrivespace2 = g_d3dx.CreateBox(g_dev, 0.7, 5.75, 0.1, Nothing)
    Set m_Tex = g_d3dx.CreateTexture(g_dev, kdx, kdx, 0, 0, d3ddm.format, D3DPOOL_MANAGED)
    Set m_font = g_d3dx.CreateFont(g_dev, m_vbfont.hFont)
    Set m_font2 = g_d3dx.CreateFont(g_dev, m_vbfont2.hFont)
End Sub

Sub DrawAxisNameSquare(i As Long)
  Dim verts(4) As D3DVERTEX
    Dim w As Single
    Dim h As Single
    Dim mat As D3DMATERIAL8
    Dim sv As Single
    Dim ev As Single
    
    On Error Resume Next
       
    w = 2:    h = 0.3

    sv = (m_font2height * (i) / kdy)
    ev = (m_font2height * (i + 1) / kdy)

    Select Case i
        Case 0
            mat.diffuse = LONGtoD3DCOLORVALUE(&H6FC0C0C0)
            mat.Ambient = mat.diffuse
            g_dev.SetTransform D3DTS_WORLD, m_LabelX.GetUpdatedMatrix
            GoTo 1
        Case 1
            
            Exit Sub
        Case 2
            mat.diffuse = LONGtoD3DCOLORVALUE(&H6FC0C0C0)
            mat.Ambient = mat.diffuse
            w = 0.3
            
    End Select
    
    Exit Sub
    
1
    g_dev.SetTexture 0, m_Tex
    g_dev.SetMaterial mat
    
    With verts(0): .X = -w * 7: .Y = -h - 0.2: .tu = 0: .tv = ev: .nz = -1: End With
    With verts(1): .X = w * 7: .Y = -h - 0.2: .tu = 1: .tv = ev: .nz = -1: End With
    With verts(2): .X = w * 7: .Y = h - 0.2: .tu = 1: .tv = sv: .nz = -1: End With
    With verts(3): .X = -w * 7: .Y = h - 0.2: .tu = 0: .tv = sv: .nz = -1: End With
    g_dev.SetVertexShader D3DFVF_VERTEX
    g_dev.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, verts(0), Len(verts(0))
    
    Exit Sub
    
2
    g_dev.SetTexture 0, m_Tex
    g_dev.SetMaterial mat
    
    With verts(0): .X = -w * 4: .Y = -h - 0.2: .tu = 0: .tv = ev: .nz = 1: End With
    With verts(1): .X = w * 4: .Y = -h - 0.2: .tu = 1: .tv = ev: .nz = 1: End With
    With verts(2): .X = w * 4: .Y = h - 0.2: .tu = 1: .tv = sv: .nz = 1: End With
    With verts(3): .X = -w * 4: .Y = h - 0.2: .tu = 0: .tv = sv: .nz = 1: End With
    g_dev.SetVertexShader D3DFVF_VERTEX
    g_dev.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, verts(0), Len(verts(0))
End Sub

Function CreateBoxWithTextureCoords(width As Single, height As Single, depth As Single, Optional WithTop As Boolean = True) As D3DXMesh
    Dim mesh As CD3DMesh
    Dim retd3dxMesh As D3DXMesh
    Dim vertexbuffer As Direct3DVertexBuffer8
    Dim verts(28) As D3DVERTEX
    Dim indices(36) As Integer
    Dim w As Single, d As Single, h1 As Single, h2 As Single
    w = width / 2
    h2 = height / 2
    h1 = -height / 2
    d = depth / 2
    
    'Create an empty d3dxmesh with room for 12 vertices and 12
    Set retd3dxMesh = g_d3dx.CreateMeshFVF(4 * 6, 6 * 6, D3DXMESH_MANAGED, D3DFVF_VERTEX, g_dev)
    
    
    'front face
    
    'add vertices
    With verts(0): .X = -w: .Y = h2: .z = -d: .nz = 1: .tu = 0: .tv = 0: End With
    With verts(1): .X = w: .Y = h2: .z = -d: .nz = 1: .tu = 1: .tv = 0: End With
    With verts(2): .X = w: .Y = h1: .z = -d: .nz = 1: .tu = 1: .tv = 1: End With
    With verts(3): .X = -w: .Y = h1: .z = -d: .nz = 1: .tu = 0: .tv = 1: End With
    
    'connect verices to make 2 triangles per face
    indices(0) = 0: indices(1) = 1: indices(2) = 2
    indices(3) = 0: indices(4) = 2: indices(5) = 3
    
    'back face
    With verts(4): .X = -w: .Y = h1: .z = d: .nz = -1: .tu = 0: .tv = 1: End With
    With verts(5): .X = w: .Y = h1: .z = d: .nz = -1: .tu = 1: .tv = 1: End With
    With verts(6): .X = w: .Y = h2: .z = d: .nz = -1: .tu = 1: .tv = 0: End With
    With verts(7): .X = -w: .Y = h2: .z = d: .nz = -1: .tu = 0: .tv = 0: End With
    indices(6) = 4: indices(7) = 5: indices(8) = 6
    indices(9) = 4: indices(10) = 6: indices(11) = 7
    
    'right face
    With verts(8): .X = w: .Y = h1: .z = -d: .nx = -1: .tu = 0: .tv = 0: End With
    With verts(9): .X = w: .Y = h1: .z = d: .nx = -1: .tu = 1: .tv = 0: End With
    With verts(10): .X = w: .Y = h2: .z = d: .nx = -1: .tu = 1: .tv = 1: End With
    With verts(11): .X = w: .Y = h2: .z = -d: .nx = -1: .tu = 0: .tv = 1: End With
    indices(12) = 8: indices(13) = 9: indices(14) = 10
    indices(15) = 8: indices(16) = 10: indices(17) = 11
    
    'left face
    With verts(16): .X = -w: .Y = h2: .z = -d: .nx = 1: .tu = 0: .tv = 1: End With
    With verts(17): .X = -w: .Y = h2: .z = d: .nx = 1: .tu = 1: .tv = 1: End With
    With verts(18): .X = -w: .Y = h1: .z = d: .nx = 1: .tu = 1: .tv = 0: End With
    With verts(19): .X = -w: .Y = h1: .z = -d: .nx = 1: .tu = 0: .tv = 0: End With
    indices(18) = 16: indices(19) = 17: indices(20) = 18
    indices(21) = 16: indices(22) = 18: indices(23) = 19
    
    '
    ' for making drawers
    If WithTop Then
    
    'top face
    With verts(20): .X = -w: .Y = h2: .z = -d: .ny = -1: .tu = 0: .tv = 0: End With
    With verts(21): .X = -w: .Y = h2: .z = d: .ny = -1: .tu = 1: .tv = 0: End With
    With verts(22): .X = w: .Y = h2: .z = d: .ny = -1: .tu = 1: .tv = 1: End With
    With verts(23): .X = w: .Y = h2: .z = -d: .ny = -1: .tu = 0: .tv = 1: End With
    indices(24) = 20: indices(25) = 21: indices(26) = 22
    indices(27) = 20: indices(28) = 22: indices(29) = 23
    
    End If
        
    'bottom  face
    With verts(24): .X = w: .Y = h1: .z = -d: .ny = 1: .tu = 0: .tv = 1: End With
    With verts(25): .X = w: .Y = h1: .z = d: .ny = 1: .tu = 1: .tv = 1: End With
    With verts(26): .X = -w: .Y = h1: .z = d: .ny = 1: .tu = 1: .tv = 0: End With
    With verts(27): .X = -w: .Y = h1: .z = -d: .ny = 1: .tu = 0: .tv = 0: End With
    indices(30) = 24: indices(31) = 25: indices(32) = 26
    indices(33) = 24: indices(34) = 26: indices(35) = 27
        
    
    D3DXMeshVertexBuffer8SetData retd3dxMesh, 0, Len(verts(0)) * 28, 0, verts(0)
    D3DXMeshIndexBuffer8SetData retd3dxMesh, 0, Len(indices(0)) * 36, 0, indices(0)
        
        
    
    Set CreateBoxWithTextureCoords = retd3dxMesh
End Function

Function doLoop()

  Do
   
    FrameMove
   
      DoEvents
  
    DrawDir
   
    D3DUtil_PresentAll frmMain.hwnd
    
  Loop Until Terminate = True
  
End Function

Private Sub DrawLines(quad As Long)
    
    g_dev.SetTransform D3DTS_WORLD, m_graphroot.GetMatrix
    
    DrawLine vec3(-5, 0.1, 0), vec3(5, 0.1, 0), &HFF0&
    DrawLine vec3(0, 0.1, -5), vec3(0, 0.1, 5), &HFF0&
    
End Sub

Private Sub DrawLine(v1 As D3DVECTOR, v2 As D3DVECTOR, color As Long)
    
    Dim mat As D3DMATERIAL8
    mat.diffuse = LONGtoD3DCOLORVALUE(color)
    mat.Ambient = mat.diffuse
    g_dev.SetMaterial mat
    
    Dim dataOut(2) As D3DVERTEX
    LSet dataOut(0) = v1
    LSet dataOut(1) = v2
    g_dev.SetVertexShader D3DFVF_VERTEX
    g_dev.DrawPrimitiveUP D3DPT_LINELIST, 2, dataOut(0), Len(dataOut(0))
    
End Sub

Public Sub MouseOver(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If m_binit = False Then Exit Sub
        
    
    Dim pick As New CD3DPick
    Dim frame As CD3DFrame
    Dim nid As Long
    
    On Error Resume Next
    
    'remove the XZ plane from consideration for pick
    m_XZPlaneFrame.Enabled = False
    
    
    pick.ViewportPick m_graphroot, X, Y
    
    nid = pick.FindNearest()
    
    If nid < 0 Then
        i_sOver = False
        gBak = False
        gHom = False
        gExit = False
        drObject = False
        drvObject = False
        dirObject = False
        m_drawtextEnable = False
        m_drawpropEnable = False
        m_xHeader = Space(17) & "HomePlay Entertainment"
        tPath = ""
        Exit Sub
    End If
        
i_sOver = True

    Set frame = pick.GetFrame(nid)
    If frame.MeshNumber = 9 Or frame.ObjectType = isNone Then GoTo miss
    
    tPath = frame.ObjectName
    
    If frame.ObjectType = isNamePlate Then
     drObject = True
     drvObject = False
     dirObject = False
     gExit = False
     gHom = False
     gBak = False
     Set o_Frame = frame
    End If
    
    If frame.ObjectType = isDrawer Then
     drObject = False
     drvObject = False
     dirObject = False
     gExit = False
     gHom = False
     gBak = False
     Set o_Frame = frame
    End If
    
    If frame.ObjectType = isDirectory Then
     dirObject = True
     drvObject = False
     drObject = False
     gExit = False
     gHom = False
     gBak = False
     Set o_Frame = frame
    End If
    
    If frame.ObjectType = isDrive Then
     drvObject = True
     drObject = False
     dirObject = False
     gExit = False
     gHom = False
     gBak = False
     Set o_Frame = frame
    End If
    
    If frame.ObjectName = "Show Drives" Then
      gBak = False
      gExit = False
      m_strIn = ""
      gHom = True
      drvObject = False
      drObject = False
      dirObject = False
    End If
    
    If frame.ObjectName = "Back" Then
      gHom = False
      gExit = False
      gBak = True
      drvObject = False
      drObject = False
      dirObject = False
    End If
    
    If frame.ObjectName = "Exit" Then
      gHom = False
      gBak = False
      gExit = True
      drvObject = False
      drObject = False
      dirObject = False
    End If
    
    'due some math to get position of item in screen space
    Dim viewport As D3DVIEWPORT8
    Dim projmatrix As D3DMATRIX
    Dim viewmatrix As D3DMATRIX
    Dim vOut As D3DVECTOR
    
    g_dev.GetViewport viewport
    g_dev.GetTransform D3DTS_PROJECTION, projmatrix
    g_dev.GetTransform D3DTS_VIEW, viewmatrix
    D3DXVec3Project vOut, vec3(0, 0, 0), viewport, projmatrix, viewmatrix, frame.GetUpdatedMatrix

    Dim destRect As RECT, i As Integer
   
    m_drawtextpos.Left = X - 20
    m_drawtextpos.Top = Y - 70
    
    m_drawproppos.Left = 10
    m_drawproppos.Top = 5
    
    If m_Showtip Then
    
    If m_drawtextpos.Left < 0 Then m_drawtextpos.Left = 1
    If m_drawtextpos.Top < 0 Then m_drawtextpos.Top = 1
    
     m_drawtext = frame.ObjectName
     m_drawtextEnable = True
     
    End If
    
     tPath = frame.ObjectName
     m_xHeader = frame.ObjectName
    
    For i = 1 To DriveInf.count
     If UCase(DriveInf.item(i).Name) = UCase(frame.ObjectName) And DriveInf.item(i).dType = 3 Then
       m_drawprop = "Size:  " & DriveInf.item(i).FullSize & " Gigs" & vbCrLf & "Free:  " & DriveInf.item(i).Freesize & " Gigs" & vbCrLf & "Used:  " & DriveInf.item(i).UsedSize & " Gigs"
       m_drawpropEnable = True
     End If
    Next
      
miss:
m_XZPlaneFrame.Enabled = True
m_XZDriveFrame.Enabled = True

End Sub

Sub FrameMove()

    'for camera movement
    m_fElapsedTime = DXUtil_Timer(TIMER_GETELLAPSEDTIME) * 1.3
    If m_fElapsedTime < 0 Then Exit Sub
        
        
    If m_bRot And m_bMouseDown = False Then
        m_graphroot.AddRotation COMBINE_BEFORE, 0, 1, 0, (g_pi / 40) * m_fElapsedTime
    End If
        
        
    ' Slow things down for the REF device
    If (g_devType = D3DDEVTYPE_REF) Then m_fElapsedTime = 0.05

    Dim fSpeed As Single
    Dim fAngularSpeed
    
    fSpeed = 1.5 * m_fElapsedTime
    fAngularSpeed = 1 * m_fElapsedTime

    ' Slowdown the camera movement
    D3DXVec3Scale m_vVelocity, m_vVelocity, 0.9
    m_fYawVelocity = m_fYawVelocity * 0.9
    m_fPitchVelocity = m_fPitchVelocity * 0.9

    ' Process keyboard input, play sound when moving, form keyup provides stop for loop
    If (m_bKey(vbKeyRight)) Then
      m_vVelocity.X = m_vVelocity.X + fSpeed        '  Slide Right
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyLeft)) Then
      m_vVelocity.X = m_vVelocity.X - fSpeed         '  Slide Left
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyUp)) Then
      m_vVelocity.z = m_vVelocity.z + fSpeed           '  Move up
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyDown)) Then
      m_vVelocity.z = m_vVelocity.z - fSpeed         '  Move down
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyW)) Then
      m_vVelocity.Y = m_vVelocity.Y + fSpeed            '  Move Forward
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyS)) Then
      m_vVelocity.Y = m_vVelocity.Y - fSpeed            '  Move Backward
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyE)) Then
      m_fYawVelocity = m_fYawVelocity + fSpeed          '  Yaw right
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyQ)) Then
      m_fYawVelocity = m_fYawVelocity - fSpeed          '  Yaw left
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyZ)) Then
      m_fPitchVelocity = m_fPitchVelocity + fSpeed      '  turn down
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyA)) Then
      m_fPitchVelocity = m_fPitchVelocity - fSpeed      '  turn up
      Sound4.Play DSBPLAY_LOOPING
    ElseIf (m_bKey(vbKeyR)) Then
      Sound3.Play DSBPLAY_DEFAULT
      'mnuRest_Click
    ElseIf (m_bKey(vbKeyF)) Then
      Call D3DUtil_ResetWindowed
      RestoreDeviceObjects
      Call D3DUtil_ResizeWindowed(frmMain.hwnd)
      
      RestoreDeviceObjects
    ElseIf (m_bKey(vbKeyEscape)) Then
     Terminate = True: End
     
    End If

    ' Update the position vector
    Dim vT As D3DVECTOR, vTemp As D3DVECTOR
    D3DXVec3Scale vTemp, m_vVelocity, fSpeed
    D3DXVec3Add vT, vT, vTemp
    D3DXVec3TransformNormal vT, vT, m_matOrientation
    D3DXVec3Add m_vPosition, m_vPosition, vT

    ' Update the yaw-pitch-rotation vector
    m_fYaw = m_fYaw + fAngularSpeed * m_fYawVelocity
    m_fPitch = m_fPitch + fAngularSpeed * m_fPitchVelocity

    Dim qR As D3DQUATERNION, det As Single
    D3DXQuaternionRotationYawPitchRoll qR, m_fYaw, m_fPitch, 0
    D3DXMatrixAffineTransformation m_matOrientation, 1.25, vec3(0, 0, 0), qR, m_vPosition
    D3DXMatrixInverse m_matView, det, m_matOrientation
    
        'set new view matrix
    g_dev.SetTransform D3DTS_VIEW, m_matView

End Sub

Function CreateSheetWithTextureCoords(width As Single, height As Single, su As Single, eu As Single, sv As Single, ev As Single, texture As Direct3DTexture8) As CD3DFrame
    Dim frame As CD3DFrame
    Dim mesh As CD3DMesh
    Dim retd3dxMesh As D3DXMesh
    Dim vertexbuffer As Direct3DVertexBuffer8
    Dim verts(8) As D3DVERTEX
    Dim indices(12) As Integer
    Dim w As Single, d As Single, h1 As Single, h2 As Single
    
    w = width / 2
    h2 = height / 2
    h1 = -height / 2
    d = 0.01
    
    Dim material As D3DMATERIAL8
    material.diffuse = LONGtoD3DCOLORVALUE(&H6FC0C0C0)
    material.Ambient = material.diffuse
        
    'Create an empty d3dxmesh with room for 12 vertices and 12
    Set retd3dxMesh = g_d3dx.CreateMeshFVF(8, 12, D3DXMESH_MANAGED, D3DFVF_VERTEX, g_dev)
    
    
    'front face
    
    'add vertices
    With verts(0): .X = -w: .Y = h2: .z = -d: .nz = 1: .tu = su: .tv = sv: End With
    With verts(1): .X = w: .Y = h2: .z = -d: .nz = 1: .tu = eu: .tv = sv: End With
    With verts(2): .X = w: .Y = h1: .z = -d: .nz = 1: .tu = eu: .tv = ev: End With
    With verts(3): .X = -w: .Y = h1: .z = -d: .nz = 1: .tu = su: .tv = ev: End With
    
    'connect verices to make 2 triangles per face
    indices(0) = 0: indices(1) = 1: indices(2) = 2
    indices(3) = 0: indices(4) = 2: indices(5) = 3
    
    'back face
    With verts(4): .X = -w: .Y = h1: .z = d: .nz = -1: .tu = eu: .tv = ev: End With
    With verts(5): .X = w: .Y = h1: .z = d: .nz = -1: .tu = su: .tv = ev: End With
    With verts(6): .X = w: .Y = h2: .z = d: .nz = -1: .tu = su: .tv = sv: End With
    With verts(7): .X = -w: .Y = h2: .z = d: .nz = -1: .tu = eu: .tv = sv: End With
    indices(6) = 4: indices(7) = 5: indices(8) = 6
    indices(9) = 4: indices(10) = 6: indices(11) = 7
    
        
    
    D3DXMeshVertexBuffer8SetData retd3dxMesh, 0, Len(verts(0)) * 8, 0, verts(0)
    D3DXMeshIndexBuffer8SetData retd3dxMesh, 0, Len(indices(0)) * 12, 0, indices(0)
        
    Set frame = New CD3DFrame
    Set mesh = frame.AddD3DXMesh(retd3dxMesh)
    
    mesh.bUseMaterials = True
    mesh.SetMaterialCount 1
    mesh.SetMaterial 0, material
    mesh.SetMaterialTexture 0, texture
    
    Set CreateSheetWithTextureCoords = frame
End Function

Sub DrawSheet(w1 As Single, w2 As Single, h1 As Single, h2 As Single, su As Single, eu As Single, sv As Single, ev As Single)
    Dim verts(4) As D3DVERTEX

    g_dev.SetTexture 0, Nothing
    
    With verts(0): .X = w1: .Y = h1: .tu = su: .tv = ev: .nz = -1: End With
    With verts(1): .X = w2: .Y = h1: .tu = eu: .tv = ev: .nz = -1: End With
    With verts(2): .X = w2: .Y = h2: .tu = eu: .tv = sv: .nz = -1: End With
    With verts(3): .X = w1: .Y = h2: .tu = su: .tv = sv: .nz = -1: End With
    
    
    With verts(0): .z = 0.01: .X = w2: .Y = h1: .tu = su: .tv = ev: .nz = 1: End With
    With verts(1): .z = 0.01: .X = w1: .Y = h1: .tu = eu: .tv = ev: .nz = 1: End With
    With verts(2): .z = 0.01: .X = w1: .Y = h2: .tu = eu: .tv = sv: .nz = 1: End With
    With verts(3): .z = 0.01: .X = w2: .Y = h2: .tu = su: .tv = sv: .nz = 1: End With

End Sub

Private Sub Form_Click()
   Dim trCord As D3DVECTOR
   Dim material As D3DMATERIAL8
   
   material.diffuse = LONGtoD3DCOLORVALUE(-1900000)
   material.Ambient = material.diffuse

' if drawer is clicked on, open then go look at it
' if is already open, then close and go home

  If lBut And drObject Then
  
   If Not o_Frame.isOpen Then
   
    Sound1.Play DSBPLAY_DEFAULT
  
    doLook
    
     trCord = o_Frame.GetPosition
     trCord.X = trCord.X
     trCord.Y = trCord.Y
     trCord.z = trCord.z - 0.9
     o_Frame.isOpen = True
     o_Frame.SetPosition trCord
     '
     ' got to make it check file number and size drawer accordingly
     MakeDrawer 0.5, 0.75, 0.9 ' <--
     
     trCord = o_Frame.GetPosition
     trCord.X = trCord.X
     trCord.Y = trCord.Y
     trCord.z = trCord.z + (0.9 / 2)  ' <--
     
     Set m_drFrame = D3DUtil_CreateFrame(m_graphroot)
        material.diffuse = LONGtoD3DCOLORVALUE(&H6FC0C0C0)
        material.Ambient = material.diffuse
        m_drFrame.AddD3DXMesh(m_meshdrawerplane).SetMaterialOverride material
        m_drFrame.SetOrientation D3DUtil_RotationAxis(0, 0, 0, 90)
        m_drFrame.SetPosition trCord '  vec3(mNewStart, 0.25, -4.13)
        m_drFrame.MeshNumber = 5
        m_drFrame.ObjectType = isDrawer
        m_drFrame.ObjectName = ""
        m_graphroot.AddChild m_drFrame
   
   ElseIf o_Frame.isOpen Then
   
     Sound2.Play DSBPLAY_DEFAULT
   
     trCord = o_Frame.GetPosition
     trCord.X = trCord.X
     trCord.Y = trCord.Y
     trCord.z = trCord.z + 0.9
     o_Frame.isOpen = False
     o_Frame.SetPosition trCord
     
     On Error Resume Next  ' if drawer closer errors
     m_drFrame.Destroy
     
     Set m_meshdrawerplane = Nothing
     Set m_drFrame = Nothing
     
     goHome
   
   End If
   
   End If
End Sub

Private Sub Form_DblClick()
   Dim i As Long
   On Error Resume Next

  'If tPath = "" Then Exit Sub
  
  If flDirPath = "" And Right(tPath, 1) = "\" Then
   flDirPath = tPath
  End If
  
  
  If drvObject Then
  
    If c_Path <> "" Then b_Path.Add c_Path
     
     m_strIn = CheckSub(o_Frame.ObjectName)
     If Mid(m_strIn, 1, 1) = "," Then m_strIn = Mid(m_strIn, 2)
     
     c_Path = m_strIn
     
        BuildDir
     
   ElseIf dirObject Then
   
     If c_Path <> "" Then b_Path.Add c_Path
     
     m_strIn = CheckSub(o_Frame.ObjectName)
   
     If Mid(m_strIn, 1, 1) = "," Then m_strIn = Mid(m_strIn, 2)
     
     c_Path = m_strIn

        BuildDir
     
   ElseIf gHom Then
   
    For i = 1 To b_Path.count
     b_Path.Remove (i)
    Next
    
    c_Path = ""
    
    BuildDrives
    
    Sound5.Play DSBPLAY_DEFAULT
     
  ElseIf gBak Then
  
     If b_Path.count = 0 Then
     
       BuildDrives
       
       Sound5.Play DSBPLAY_DEFAULT
       
      Exit Sub
     End If
  
     For i = b_Path.count To 1 Step -1
       If b_Path(i) <> "" Then
       
         m_strIn = b_Path(i)
         BuildDir
         b_Path.Remove (i)
         
       ElseIf b_Path(i) = "" Then
       
         BuildDrives
       
       End If
       
     Next
     
  ElseIf gExit Then
  
   Terminate = True
   D3DUtil_Destory
   
   End
    
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    m_bKey(KeyCode) = True
    unKey = False

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo muststart ' the sound will error when starting, not initialized yet
 If KeyCode = 83 Then StartIT = True

   m_bKey(KeyCode) = False
   unKey = True
   
   Sound4.Stop
muststart:
End Sub

Private Sub Form_Load()
  Dim pX As Long, pY As Long
  StartIT = False
  Me.Show
  
    pX = 10
    pY = 10
    
    Me.ForeColor = vbWhite
    Me.FontSize = 8
    Me.CurrentX = pX
    Me.CurrentY = pY
    Print "Arrows Move You" & vbCrLf & vbCrLf & "  W = UP" & vbCrLf & "  S = DOWN" & vbCrLf & "  A = LOOK UP" & vbCrLf & "  Z = LOOK DOWN" & vbCrLf & vbCrLf & "   Click on a Drawer to Open or Close" & vbCrLf & "   Double Click on Drive or Directory to Explore" & vbCrLf & vbCrLf & "  S    To Start"
    
    Do
     DoEvents
    Loop Until StartIT
    
    Me.Cls
  
    pX = (Me.ScaleWidth / 2) - 130
    pY = 280
    
    Me.ForeColor = vbWhite
    Me.FontSize = 24
    Me.CurrentX = pX
    Me.CurrentY = pY
    Print "Loading 3d-Explor..."

    DoEvents

    m_Mediadir = App.Path & "\Media\"
    D3DUtil_SetMediaPath m_Mediadir
    
    Me.ForeColor = vbBlack
    Me.FontSize = 10
    
    Init Me.hwnd, Me.font, Label1.font
    
    'Start the timers and callbacks
    Call DXUtil_Timer(TIMER_start)
    
    Call Wait(0.125)
    
    doLoop

End Sub

'- Rotate Track ball
'  given a point on the screen the mouse was moved to
'  simulate a track ball
Private Sub RotateTrackBall(X As Integer, Y As Integer, rFrame As CD3DFrame)

    
    Dim delta_x As Single, delta_y As Single
    Dim delta_r As Single, radius As Single, denom As Single, angle As Single
    
    ' rotation axis in camcoords, worldcoords, sframecoords
    Dim axisC As D3DVECTOR
    Dim wc As D3DVECTOR
    Dim axisS As D3DVECTOR
    Dim base As D3DVECTOR
    Dim origin As D3DVECTOR
    
    delta_x = X - m_lastX
    delta_y = Y - m_lasty
    m_lastX = X
    m_lasty = Y

            
     delta_r = Sqr(delta_x * delta_x + delta_y * delta_y)
     radius = 50
     denom = Sqr(radius * radius + delta_r * delta_r)
    
    If (delta_r = 0 Or denom = 0) Then Exit Sub
    angle = (delta_r / denom)

    axisC.X = (-delta_y / delta_r)
    axisC.Y = (-delta_x / delta_r)
    axisC.z = 0


    'transform camera space vector to world space
    'm_largewindow.m_cameraFrame.Transform wc, axisC
    g_dev.GetTransform D3DTS_VIEW, g_viewMatrix
    D3DXVec3TransformCoord wc, axisC, g_viewMatrix
    
    
    'transform world space vector into Model space
    rFrame.UpdateFrames
    axisS = rFrame.InverseTransformCoord(wc)
        
    'transform origen camera space to world coordinates
    'm_largewindow.m_cameraFrame.Transform  wc, origin
    D3DXVec3TransformCoord wc, origin, g_viewMatrix
    
    'transfer cam space origen to model space
    base = rFrame.InverseTransformCoord(wc)
    
    axisS.X = axisS.X - base.X
    axisS.Y = axisS.Y - base.Y
    axisS.z = axisS.z - base.z
    
    rFrame.AddRotation COMBINE_BEFORE, axisS.X, axisS.Y, axisS.z, angle
    
End Sub

Function SpinFrame()
 Dim i As Single
 
  For i = 90 To 450
   m_CTRLFrame(1).SetOrientation D3DUtil_RotationAxis(0, 1, 0, i)
   Sound8.Play DSBPLAY_LOOPING
   DrawDir
   i = i + 29
  Next
  
Sound8.Stop
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        rBut = True
        lBut = False
        
        If drObject Or dirObject Or drvObject Then GoTo oder
    Else
        rBut = False
        lBut = True
    
        '- save our current position
        m_bMouseDown = True
        m_lastX = X
        m_lasty = Y
        
    End If
    
    Exit Sub
    
oder:
 PopupMenu mnuHdr
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    If m_binit = False Then Exit Sub
    
    FrameMove
    
    If m_bMouseDown = False Then
        Call MouseOver(Button, Shift, X, Y)
    Else
        '- Rotate the object
        If i_sOver And m_bMouseDown = False And Button = 2 Then
         PopupMenu mnuHdr
        Exit Sub
        End If
        RotateTrackBall CInt(X), CInt(Y), m_graphroot
    End If
    
    DrawDir
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_bMouseDown = False
    
    If Button = 2 And Not drObject And Not dirObject And Not drvObject Then
      Me.PopupMenu xx
    Exit Sub
    End If
End Sub

Private Sub Form_Paint()
  If Not m_binit Then Exit Sub
  If Not m_bGraphInit Then Exit Sub
    'DrawDir
End Sub

Private Sub Form_Resize()
     ' If D3D is not initialized then exit
    If Not m_binit Then Exit Sub
    
    ' If we are in a minimized state stop the timer and exit
    If Me.WindowState = vbMinimized Then
        DXUtil_Timer TIMER_STOP
        m_bMinimized = True
        Exit Sub
        
    ' If we just went from a minimized state to maximized
    ' restart the timer
    Else
        If m_bMinimized = True Then
            DXUtil_Timer TIMER_start
            m_bMinimized = False
        End If
    End If
        
     ' Dont let the window get too small
    If Me.ScaleWidth < 10 Then
        Me.width = Screen.TwipsPerPixelX * 10
        Exit Sub
    End If
    
    If Me.ScaleHeight < 10 Then
        Me.height = Screen.TwipsPerPixelY * 10
        Exit Sub
    End If
    
    'remove references to FONTs
    DeleteDeviceObjects
    
    'reset and resize our D3D backbuffer to the size of the window
    D3DUtil_ResizeWindowed Me.hwnd
    
    'All state get losts after a reset so we need to reinitialze it here
    RestoreDeviceObjects
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Terminate = True
 
  D3DUtil_Destory
 
    End
End Sub

Function MakeComponent(Which As Integer, Name As String)
  '
  ' This is... well a fix
  ' the only way it'll work for me is if i make each x file again and
  ' again, for each drive, dir and so. if i just make an array of frames
  ' and make the x file component once and set one of the array to it, no good, stumped me
  ' if i make the component in an array, and the frames in an array, it works
  '
  If InStr(UCase(Name), "RECYCLE") Then
    Which = 5
  End If
  
  
  Select Case Which
     Case 0
      Set m_FPFrame(fc) = D3DUtil_LoadFromFile(m_Mediadir & "X\fp.x", m_graphroot, Nothing, Name, isFloppy)
       fc = fc + 1
     
     Case 1
      Set m_HDFrame(hc) = D3DUtil_LoadFromFile(m_Mediadir & "X\hd.x", m_graphroot, Nothing, Name, isDrive)
       hc = hc + 1
     
     Case 2
      Set m_CDFrame(cc) = D3DUtil_LoadFromFile(m_Mediadir & "X\cd.x", m_graphroot, Nothing, Name, isCD)
       cc = cc + 1
     
     Case 3
      Set m_DR1Frame(dR1) = D3DUtil_LoadFromFile(m_Mediadir & "X\dir1.x", m_graphroot, Nothing, Name, isDirectory)
       dR1 = dR1 + 1
       
     Case 4
      Set m_DR2Frame(dR2) = D3DUtil_LoadFromFile(m_Mediadir & "X\dir2.x", m_graphroot, Nothing, Name, isDirectory)
       dR2 = dR2 + 1
     
     Case 5
      Set m_DR2Frame(dR2) = D3DUtil_LoadFromFile(m_Mediadir & "X\rec.x", m_graphroot, Nothing, Name, isDirectory)
       dR2 = dR2 + 1
     
     Case 6
      If Cntc = 2 Then Exit Function
      If Cntc = 0 Then
       Set m_CTRLFrame(Cntc) = D3DUtil_LoadFromFile(m_Mediadir & "X\comp.x", m_graphroot, Nothing, Name, isBase)
      Else
      Set m_CTRLFrame(Cntc) = D3DUtil_LoadFromFile(m_Mediadir & "X\up.x", m_graphroot, Nothing, Name, isBase)
      End If
      
     Cntc = Cntc + 1
     
     Case 7
      Set m_plor = D3DUtil_LoadFromFile(m_Mediadir & "X\plor.x", m_graphroot, Nothing, Name, isBase)
      
     Case 8
      Set m_door = D3DUtil_LoadFromFile(m_Mediadir & "X\door.x", m_graphroot, Nothing, Name, isBase)
      
  End Select
  
  'DoEvents
  
End Function

Private Sub mnuExit_Click()
   Terminate = True
   D3DUtil_Destory
   
  End
End Sub

Private Sub mnuGo_Click()
    m_strIn = CheckSub(o_Frame.ObjectName)
   
     If Mid(m_strIn, 1, 1) = "," Then m_strIn = Mid(m_strIn, 2)

        BuildDir
     
End Sub

Private Sub mnuReset_Click()
    m_graphroot.SetMatrix g_identityMatrix
    m_vPosition = vec3(-1, 6, -42)
    m_fYaw = 0
    m_fPitch = 0

    Call D3DXMatrixTranslation(m_matOrientation, 0, 0, 0)
    D3DUtil_SetupDefaultScene
    g_dev.GetTransform D3DTS_VIEW, m_matView
End Sub

Private Sub mnuRot_Click()
   m_bRot = Not m_bRot
End Sub

Private Sub mnuSpeed_Click()
   m_Faster = Not m_Faster
   
   If m_Faster Then mnuSpeed.Checked = True
   
   If Not m_Faster Then mnuSpeed.Checked = False
   
End Sub

Private Sub mnuTip_Click()
    m_Showtip = Not m_Showtip
  
  If m_Showtip Then mnuTip.Checked = True
  
  If Not m_Showtip Then mnuTip.Checked = False

End Sub

Function makeSpace(Inn As Long)
 Dim tNum As Single
 
 tNum = (Inn / 10) / 2
 
 Select Case Inn
  Case 0 To 9
   spcMade = 4
  Case 10 To 14
   spcMade = 4.05
  Case 15 To 19
   spcMade = 4.1
  Case 20 To 24
   spcMade = 4.15
  Case 25 To 29
   spcMade = 4.28
  Case 30 To 34
   spcMade = 4.35
  Case 35 To 39
   spcMade = 4.5
  Case 40 To 44
   spcMade = 4.7
  Case 45 To 49
   spcMade = 4.9
  Case 50 To 54
   spcMade = 5
  Case 55 To 59
   spcMade = 5.1
  Case 60 To 64
   spcMade = 5.2
  Case 65 To 69
   spcMade = 5.3
  Case 70 To 74
   spcMade = 5.4
  Case 75 To 79
   spcMade = 5.5
  Case 80 To 84
   spcMade = 5.6
  Case 85 To 89
   spcMade = 5.7
  Case 90 To 94
   spcMade = 5.8
  Case 95 To 100
   spcMade = 6
End Select
 
 Set m_meshdrivespace = g_d3dx.CreateBox(g_dev, 0.5, tNum, 0.1, Nothing)
  
End Function
