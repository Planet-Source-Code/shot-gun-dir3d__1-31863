Attribute VB_Name = "modDrives"
Option Explicit
        
Declare Function GetDiskFreeSpace_FAT32 Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long
Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
        (ByVal nDrive As String) As Long
       Public Const DRIVE_REMOVABLE = 2
       Public Const DRIVE_FIXED = 3
       Public Const DRIVE_REMOTE = 4
       Public Const DRIVE_CDROM = 5
       Public Const DRIVE_RAMDISK = 6
       Declare Function GetDiskFreeSpace_FAT16 Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
       Dim SectorsPerCluster&, BytesPerSector&, NumberOfFreeClusters&, TotalNumberOfClusters&

Declare Function GetVolumeInformation Lib _
                "kernel32" Alias "GetVolumeInformationA" _
                (ByVal lpRootPathName As String, _
                ByVal lpVolumeNameBuffer As String, _
                ByVal nVolumeNameSize As Long, _
                lpVolumeSerialNumber As Long, _
                lpMaximumComponentLength As Long, _
                lpFileSystemFlags As Long, _
                ByVal lpFileSystemNameBuffer As String, _
                ByVal nFileSystemNameSize As Long) As Long
                
Global DrvCount As Integer
Global hDrvCount As Integer
Global currDrive As String

Function doDrvInf()

        Dim allDrives As String
        Dim strOut As String
        
        Dim drv() As String
        Dim SZstrfl As Long
        Dim SZstrfr As Long
        Dim perc As Long
        
        DrvCount = 0
        hDrvCount = 0
        
        allDrives$ = VBGetLogicalDriveStrings()
        
        '
        ' with drive string figure out full, free and used space
        ' and percentage used then save to collection
        ' this must be called early to figure out base size and display
        '
        Do Until allDrives$ = Chr$(0)
          currDrive$ = StripNulls$(allDrives$)
          perc = 0
          
          DrvCount = DrvCount + 1
          
          If GetDriveType(currDrive$) = 3 Then hDrvCount = hDrvCount + 1
          '
          ' figures percentage used
          '
            '
            ' the 'ELSE' is to avoid the delay of checking the floppy when starting
            '
            With DriveInf
             If GetDriveType(currDrive$) <> 2 Then
                perc = lPercent(GetDriveUsedSpace(currDrive$), GetDriveSpace(currDrive$))
                .Add currDrive$, GetDriveSpace(currDrive$), GetDriveFreeSpace(currDrive$), GetDriveUsedSpace(currDrive$), perc, GetDriveType(currDrive$)
             Else
                .Add currDrive$, 0, 0, 0, 0, 2
             End If
            End With
          
        Loop
        
End Function

Function lPercent(Valin As Long, ValMax As Long) As Long
On Error GoTo aired

lPercent& = Int(Valin * 100 / ValMax)

Exit Function

aired:
lPercent = 0
Err = 0
End Function

Private Function VBGetLogicalDriveStrings() As String

       '      'returns a single string of available drive
       '      'letters, each separated by a chr$(0)
        Dim r As Long
        Dim tmp As String
        
        tmp$ = Space$(64)
        
        r& = GetLogicalDriveStrings(Len(tmp$), tmp$)
        
        VBGetLogicalDriveStrings = Trim$(tmp$)
End Function
        
Private Function rgbGetDriveType(RootPathName$) As Integer

       '      'Passed is the drive to check.
       '      'Returned is the type of drive.
        
        Dim r As Long
        
        r& = GetDriveType(RootPathName$)
        
        Select Case r&
       Case 0: rgbGetDriveType = 0
       Case 1: rgbGetDriveType = 1
       Case DRIVE_REMOVABLE:
        Select Case Left$(RootPathName, 1)
        Case "a", "b": rgbGetDriveType = 2
        End Select
       Case DRIVE_FIXED: rgbGetDriveType = 3
       Case DRIVE_REMOTE: rgbGetDriveType = 4
       Case DRIVE_CDROM: rgbGetDriveType = 5
       Case DRIVE_RAMDISK: rgbGetDriveType = 6
        End Select
        
End Function

Private Function StripNulls(startStrg$) As String
        Dim c As Integer
        Dim item As String
        
        c% = 1
        
        Do

              If Mid$(startStrg$, c%, 1) = Chr$(0) Then
                      
                      item$ = Mid$(startStrg$, 1, c% - 1)
                      startStrg$ = Mid$(startStrg$, c% + 1, Len(startStrg$))
                      StripNulls$ = item$
                      Exit Function
              End If

       c% = c% + 1
        Loop
End Function

Function GetDriveSpace(sDrive As String) As Variant
Dim r As Long
Dim BytesFreeToCalller As Currency
Dim TotalBytes As Currency
Dim TotalFreeBytes As Currency
Dim TotalBytesUsed As Currency
Dim sUsage As Variant
    
    'get the drive's disk parameters
    Call GetDiskFreeSpaceEx(sDrive, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    GetDriveSpace = format$(TotalBytes * 10000, "###,###,###,##0")

    ' By Bytes
    If Len(GetDriveSpace) = 5 Or Len(GetDriveSpace) >= 5 Then
        If Len(GetDriveSpace) = 5 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 1)
            GetDriveSpace = sUsage
        ElseIf Len(GetDriveSpace) = 6 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 2)
            GetDriveSpace = sUsage
        ElseIf Len(GetDriveSpace) = 7 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 3)
            GetDriveSpace = sUsage
        End If
    End If

    ' By Megabytes
    If Len(GetDriveSpace) = 9 Or Len(GetDriveSpace) >= 9 Then
        If Len(GetDriveSpace) = 9 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 1)
            GetDriveSpace = sUsage
        ElseIf Len(GetDriveSpace) = 10 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 2)
            GetDriveSpace = sUsage
        ElseIf Len(GetDriveSpace) = 11 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 3)
            GetDriveSpace = sUsage
        End If
    End If

    ' By Gigabytes
    If Len(GetDriveSpace) = 13 Or Len(GetDriveSpace) >= 13 Then
        If Len(GetDriveSpace) = 13 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 1)
            GetDriveSpace = sUsage
        ElseIf Len(GetDriveSpace) = 14 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 2)
            GetDriveSpace = sUsage
        ElseIf Len(GetDriveSpace) = 15 Then
            GetDriveSpace = CStr(GetDriveSpace)
            sUsage = Left(GetDriveSpace, 3)
            GetDriveSpace = sUsage
        End If
    End If
End Function

Function GetDriveFreeSpace(sDrive As String) As Variant
Dim r As Long
Dim BytesFreeToCalller As Currency
Dim TotalBytes As Currency
Dim TotalFreeBytes As Currency
Dim TotalBytesUsed As Currency
Dim sUsage As Variant

    'get the drive's disk parameters
    Call GetDiskFreeSpaceEx(sDrive, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    GetDriveFreeSpace = format$(BytesFreeToCalller * 10000, "###,###,###,##0")
    
    ' By Bytes
    If Len(GetDriveFreeSpace) = 5 Or Len(GetDriveFreeSpace) >= 5 Then
        If Len(GetDriveFreeSpace) = 5 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 1)
            GetDriveFreeSpace = sUsage
        ElseIf Len(GetDriveFreeSpace) = 6 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 2)
            GetDriveFreeSpace = sUsage
        ElseIf Len(GetDriveFreeSpace) = 7 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 3)
            GetDriveFreeSpace = sUsage
        End If
    End If

    ' By Megabytes
    If Len(GetDriveFreeSpace) = 9 Or Len(GetDriveFreeSpace) >= 9 Then
        If Len(GetDriveFreeSpace) = 9 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 1)
            GetDriveFreeSpace = sUsage
        ElseIf Len(GetDriveFreeSpace) = 10 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 2)
            GetDriveFreeSpace = sUsage
        ElseIf Len(GetDriveFreeSpace) = 11 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 3)
            GetDriveFreeSpace = sUsage
        End If
    End If

    ' By Gigabytes
    If Len(GetDriveFreeSpace) = 13 Or Len(GetDriveFreeSpace) >= 13 Then
        If Len(GetDriveFreeSpace) = 13 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 1)
            GetDriveFreeSpace = sUsage
        ElseIf Len(GetDriveFreeSpace) = 14 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 2)
            GetDriveFreeSpace = sUsage
        ElseIf Len(GetDriveFreeSpace) = 15 Then
            GetDriveFreeSpace = CStr(GetDriveFreeSpace)
            sUsage = Left(GetDriveFreeSpace, 3)
            GetDriveFreeSpace = sUsage
        End If
    End If
End Function

Function GetDriveUsedSpace(sDrive As String) As Variant
Dim r As Long
Dim BytesFreeToCalller As Currency
Dim TotalBytes As Currency
Dim TotalFreeBytes As Currency
Dim TotalBytesUsed As Currency
Dim sUsage As Variant

    'get the drive's disk parameters
    Call GetDiskFreeSpaceEx(sDrive, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    GetDriveUsedSpace = format$((TotalBytes - TotalFreeBytes) * 10000, "###,###,###,##0")
    
    ' By Bytes
    If Len(GetDriveUsedSpace) = 5 Or Len(GetDriveUsedSpace) >= 5 Then
        If Len(GetDriveUsedSpace) = 5 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 1)
            GetDriveUsedSpace = sUsage
        ElseIf Len(GetDriveUsedSpace) = 6 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 2)
            GetDriveUsedSpace = sUsage
        ElseIf Len(GetDriveUsedSpace) = 7 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 3)
            GetDriveUsedSpace = sUsage
        End If
    End If

    ' By Megabytes
    If Len(GetDriveUsedSpace) = 9 Or Len(GetDriveUsedSpace) >= 9 Then
        If Len(GetDriveUsedSpace) = 9 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 1)
            GetDriveUsedSpace = sUsage
        ElseIf Len(GetDriveUsedSpace) = 10 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 2)
            GetDriveUsedSpace = sUsage
        ElseIf Len(GetDriveUsedSpace) = 11 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 3)
            GetDriveUsedSpace = sUsage
        End If
    End If

    ' By Gigabytes
    If Len(GetDriveUsedSpace) = 13 Or Len(GetDriveUsedSpace) >= 13 Then
        If Len(GetDriveUsedSpace) = 13 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 1)
            GetDriveUsedSpace = sUsage
        ElseIf Len(GetDriveUsedSpace) = 14 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 2)
            GetDriveUsedSpace = sUsage
        ElseIf Len(GetDriveUsedSpace) = 15 Then
            GetDriveUsedSpace = CStr(GetDriveUsedSpace)
            sUsage = Left(GetDriveUsedSpace, 3)
            GetDriveUsedSpace = sUsage
        End If
    End If
End Function

