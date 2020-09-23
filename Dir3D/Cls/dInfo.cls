VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sID   As String
Private fSize   As Long
Private frSize  As Long
Private uSize   As Long
Private uType   As Long
Private nName   As String
Private nPer    As Long


Public Property Let id(ByVal sID As String)
Static bAlreadySet As Boolean

    If Not bAlreadySet Then
        m_sID = sID
        bAlreadySet = True
    End If
End Property
Public Property Get id() As String
    id = m_sID
End Property

Public Property Get FullSize() As Long
    FullSize = fSize
End Property
Public Property Let FullSize(ByVal size As Long)
    fSize = size
End Property

Public Property Get Freesize() As Long
    Freesize = frSize
End Property
Public Property Let Freesize(ByVal size As Long)
    frSize = size
End Property

Public Property Get UsedSize() As Long
    UsedSize = uSize
End Property
Public Property Let UsedSize(ByVal Used As Long)
    uSize = Used
End Property

Public Property Get dType() As Long
    dType = uType
End Property
Public Property Let dType(ByVal New_Type As Long)
    uType = New_Type
End Property

Public Property Get Name() As String
    Name = nName
End Property
Public Property Let Name(ByVal New_Name As String)
    nName = New_Name
End Property

Public Property Get UsedPercent() As Long
    UsedPercent = nPer
End Property
Public Property Let UsedPercent(ByVal New_Percent As Long)
    nPer = New_Percent
End Property

Public Sub Clear()
    nPer = 0
    uSize = 0
    frSize = 0
    fSize = 0
    nName = ""
End Sub





