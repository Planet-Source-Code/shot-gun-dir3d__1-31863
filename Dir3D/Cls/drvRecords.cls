VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "drvRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim newTask As dInfo
Private m_Records As Collection

Public Property Get count() As Long
    count = m_Records.count
End Property

Private Sub Class_Initialize()
    Set m_Records = Nothing
    Set m_Records = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Records = Nothing
End Sub

Public Function Add(ByVal Name As Variant, ByVal FullSize As Variant, ByVal Freesize As Variant, ByVal UsedSize As Variant, ByVal Percentused As Variant, ByVal uType As Variant) As dInfo
                    
On Error GoTo ERR_ROUTINE

Set newTask = New dInfo

    ' Create the instance
    With newTask
        ' Set a unique ID
        .id = NextID
        .Name = Name
        .FullSize = FullSize
        .Freesize = Freesize
        .UsedSize = UsedSize
        .UsedPercent = Percentused
        .dType = uType
        ' Add the member to the collection
        m_Records.Add newTask, .id
    End With

    ' Return the one that was added
    Set Add = newTask
Exit Function

ERR_ROUTINE:
    ' Pass any error up to the calling class from the lower level object
    'Err.Raise Err.Number, Err.Source, Err.Description
End Function
        

' Clear the collection
Public Sub Clear()
    Set m_Records = New Collection
End Sub

' Displays the defined member
' Parameters:
'   vkey         key or index of member to get
Public Function item(ByVal vKey As Variant) As dInfo
On Error GoTo ERR_ROUTINE
    Set item = m_Records.item(vKey)
    
Exit Function
ERR_ROUTINE:
    'Err.Raise tasksInvalidIndexError, "clsStudent::Item", _
        "Collection key or index is invalid"
End Function

' Exposes the enumeration object
' to allow the For Each...Next syntax
Public Function NewEnum() As IUnknown
    Set NewEnum = m_Records.[_NewEnum]
End Function

' Provides the next available ID
' Returns:
'   NextID      Next unique Task ID
Private Function NextID() As String
Static iID As Integer

    iID = iID + 1
    NextID = "T" & format$(iID, "00000")

End Function

' Remove the member from the collection.
' Parameters:
'   vkey         key or index of member to delete
Public Sub Remove(ByVal vKey As Variant)
On Error GoTo ERR_ROUTINE

    ' Remove the member from the collection
    m_Records.Remove vKey
    
Exit Sub
ERR_ROUTINE:
    'Err.Raise tasksInvalidIndexError, "clsStudent::Remove", _
        "Collection key or index is invalid"
End Sub

