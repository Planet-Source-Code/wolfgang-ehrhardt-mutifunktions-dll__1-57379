VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal lngDrive As Long, ByVal lngCapacity As Long, ByVal lngFormatType As Long) As Long
Private Declare Function GetDriveTypeB Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Function SHGetDiskFreeSpace Lib "shell32" Alias "SHGetDiskFreeSpaceA" (ByVal pszVolume As String, pqwFreeCaller As Currency, pqwTot As Currency, pqwFree As Currency) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, hTemplateFile As Long) As Long

Private Type SCSI_PASS_THROUGH_DIRECT
    length             As Integer
    ScsiStatus         As Byte
    PathId             As Byte
    TargetId           As Byte
    Lun                As Byte
    CdbLength          As Byte
    SenseInfoLength    As Byte
    DataIn             As Byte
    DataTransferLength As Long
    TimeOutValue       As Long
    DataBuffer         As Long
    SenseInfoOffset    As Long
    Cdb(15)            As Byte
    Fill(2)            As Byte
End Type

Private Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As Long
    bInheritHandle       As Long
End Type

Private Type SCSI_PASS_THROUGH_DIRECT_WITH_BUFFER
    spt             As SCSI_PASS_THROUGH_DIRECT
    Fill            As Long
    SenseBuffer(31) As Byte
End Type

Private Type OVERLAPPED
    Internal     As Long
    InternalHigh As Long
    offset       As Long
    OffsetHigh   As Long
    hEvent       As Long
End Type

Private Type MMC
    PageCode          As Byte
    PageLen           As Byte
    rsvd2(7)          As Byte
    ReadSupported     As Byte
    WriteSupported    As Byte
    misc(3)           As Byte
    MaxReadSpeed(1)   As Byte
    NumVolLevels(1)   As Byte
    BufferSize(1)     As Byte
    CurrReadSpeed(1)  As Byte
    rsvd              As Byte
    misc2             As Byte
    MaxWriteSpeed(1)  As Byte
    CurrWriteSpeed(1) As Byte
    Rest(20)          As Byte
End Type

Public Enum FormatOptions
    foCAPACITY_DEFAULT = 0
    foCAPACITY_360 = 3
    foCAPACITY_720 = 5
End Enum

Public Enum DriveType
    dtREMOVABLE = 0
    dtHARDDISK = 1
    dtNETDRIVE = 2
    dtCDROM = 3
    dtRAMDISK = 4
    dtUNKNOWN = -1
    dtINVALIDDRIVE = -3
End Enum

Public Enum DiskSpaceType
    dsTOTAL = 1
    dsFREE = 2
    dsUSED = 3
End Enum

Public Enum CDRomCommand
    cdOPEN = 1
    cdCLOSE = 2
End Enum

Public Enum vdVIRTUALDRIVECOMMAND
    vdREMOVE = 0
    vdCREATE = 1
End Enum

Private Enum eDrvErrs
    ERR_SPTD_FAIL
    ERR_SCSI_CMD
    ERR_NT_NEEDED
    ERR_OK
End Enum

Public Enum eDrvType
    CDTYPE_CDROM
    CDTYPE_CDWRITER
    CDTYPE_CDRWWRITER
    CDTYPE_DVDROM
    CDTYPE_DVDWRITER
End Enum

Private Const CB_DELETESTRING = &H144
Private Const CB_SETCURSEL = &H14E

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const INVALID_HANDLE_VALUE = -1

Private Const IOCTL_SCSI_PASS_THROUGH_DIRECT = &H4D014

Private Const OPEN_EXISTING = 3

Private Const SCSI_IOCTL_DATA_IN = 1

Private Const SHFD_FORMAT_QUICK = 0
Private Const SHFD_FORMAT_FULL = 1
Public Function Get_DriveSerial(DriveLetter As String) _
                                    As String
    Dim No As Long
    Dim S As String * 256
    
    If Len(DriveLetter) = 1 Then _
        DriveLetter = DriveLetter & ":\"

    Call GetVolumeInformation(DriveLetter, S, 256, _
                              No, 0&, 0&, S, 256)
    
    Get_DriveSerial = CStr(No)

End Function
Public Function Get_DriveLabel(ByVal DriveLetter As String) _
                                    As String
    Dim Result As Long, SerN As Long, PathL As Long, flags As Long
    Dim VolN As String * 256, FileS As String * 256
  
    If Len(DriveLetter) = 1 Then DriveLetter = DriveLetter & ":\"
    
    Result = GetVolumeInformation(DriveLetter, VolN, 256, SerN, _
                                  PathL, flags, FileS, 256)
    
    If Result Then _
        Get_DriveLabel = Left$(VolN, InStr(VolN, Chr$(0)) - 1)
    
End Function
Public Function Set_DriveLabel(ByVal DriveLetter As String, _
                               ByVal Label As String) _
                                    As Boolean
    
    If Len(DriveLetter) = 1 Then DriveLetter = DriveLetter & ":\"
  
    Call SetVolumeLabel(DriveLetter, Label)

    Set_DriveLabel = (Get_DriveLabel(DriveLetter) = Label)
    
End Function
Public Sub CDROM(Command As CDRomCommand)
    Dim R$
    
    Select Case Command
        Case cdCLOSE: R$ = "closed"
        Case cdOPEN: R$ = "open"
    End Select
    
    Call mciExecute("Set CDaudio door " & R$)
    
End Sub
Public Sub Format(ByVal hWnd As Long, _
                  ByVal DriveLetter As String, _
                  FormatCapacity As FormatOptions, _
                  Optional QuickFormat As Boolean = False)
    Dim Drive As Long, lngFormatType As Long
    
    If Len(DriveLetter) = 1 Then
        DriveLetter = UCase(DriveLetter)
        Drive = Asc(DriveLetter) - 65
        
        lngFormatType = IIf(QuickFormat, SHFD_FORMAT_QUICK, _
                                         SHFD_FORMAT_FULL)
        
        Call SHFormatDrive(hWnd, Drive, FormatCapacity, _
                                        lngFormatType)
    End If

End Sub
Public Function VirtualDrive(Command As vdVIRTUALDRIVECOMMAND, _
                             ByVal TargetDriveLetter As String, _
                             Optional SourceDir As String) _
                                    As Boolean
    Dim strShell As Variant
           
    Select Case Command
        Case vdCREATE
            If Len(TargetDriveLetter) = 1 Then
                If Not DriveExist(TargetDriveLetter) Then
                    SourceDir = F.Get_DOSfileName(SourceDir)
        
                    strShell = "subst " & TargetDriveLetter & _
                               ": " & SourceDir
                   
                    Call Shell(strShell, vbHide)
            
                    If DI.Exist(TargetDriveLetter & ":\") Then _
                        VirtualDrive = True
                End If
            End If
        Case vdREMOVE
            If Len(TargetDriveLetter) = 1 Then
                If DriveExist(TargetDriveLetter) Then
                    strShell = "subst /D " & TargetDriveLetter & ":"
            
                    Call Shell(strShell, vbHide)
                            
                    If Not DI.Exist(TargetDriveLetter & ":\") Then _
                        VirtualDrive = True
                End If
            End If
    End Select

End Function
Public Function Get_NextFreeDrive() As String
    Dim P As Long
    Dim A As String
    
    A = UCase$(Get_AllDrives)
    
    For P = Asc("C") To Asc("Z")
        If InStr(A, Chr$(P)) = 0 Then _
            Get_NextFreeDrive = Chr$(P) & ":\": _
            Exit For
    Next P
    
End Function
Public Function Get_AllDrives() As String
    Dim L As Long, Result As Long
    Dim Puffer As String, Drives As String
    
    Puffer = Space(64)
    L = 64
    
    Result = GetLogicalDriveStrings(L, Puffer)
    Drives = Left$(Puffer, Result)
    
    Drives = Trim$(Drives)
    
    Drives = Replace(Drives, ":", "")
    Drives = Replace(Drives, "\", "")
    Drives = Replace(Drives, Chr$(0), "")
    
    Get_AllDrives = Drives
    
End Function
Public Function Get_DriveType(ByVal DriveLetter As String) _
                                    As DriveType
    Dim Typ As Long
    Dim Drive As String, AllDrives As String

    If Len(DriveLetter) <> 1 Then _
        Get_DriveType = dtINVALIDDRIVE: _
        Exit Function
    
    Drive = UCase$(DriveLetter)
    AllDrives = Get_AllDrives
    
    If InStr(AllDrives, Drive) = 0 Then
        Get_DriveType = dtINVALIDDRIVE
    Else
        Drive = Drive & ":\"
        Typ = GetDriveTypeB(Drive)

        Select Case Typ
            Case 2: Get_DriveType = dtREMOVABLE
            Case 3: Get_DriveType = dtHARDDISK
            Case 4: Get_DriveType = dtNETDRIVE
            Case 5: Get_DriveType = dtCDROM
            Case 6: Get_DriveType = dtRAMDISK
            
            Case Else
                Get_DriveType = dtUNKNOWN
        End Select
    End If
    
End Function
Public Function Get_DiskSpace(ByVal DriveLetter As String, _
                              WhatSpace As DiskSpaceType) _
                                    As String
    Dim FreeCaller As Currency, Tot As Currency, Free As Currency
    
    If Len(DriveLetter) = 1 Then DriveLetter = DriveLetter & ":\"

    Call SHGetDiskFreeSpace(DriveLetter, FreeCaller, Tot, Free)
    
    Select Case WhatSpace
        Case dsTOTAL
            Get_DiskSpace = Tot
        Case dsFREE
            Get_DiskSpace = Free
        Case dsUSED
            Get_DiskSpace = Tot - Free
        Case Else: Exit Function
    End Select

End Function
Public Function isDriveReady(ByVal DriveLetter As String) _
                                As Boolean
    Dim OldDrive As String
    Dim Result As Boolean

    On Local Error Resume Next
  
    OldDrive = CurDir()
    ChDir DriveLetter
    Result = (Err = 0)
    ChDir OldDrive
    
    isDriveReady = Result

End Function
Public Function Get_FirstCDdrive() As String
    
    Get_FirstCDdrive = Get_DriveByType(dtCDROM)
    
    If Len(Get_FirstCDdrive) > 1 Then _
        Get_FirstCDdrive = Left$(Get_FirstCDdrive, 1)

End Function
Public Function Get_DriveByType(WhatType As DriveType) As String
    Dim DRV As String
    Dim P As Long
    
    DRV = Get_AllDrives
        
    For P = 1 To Len(DRV)
        If Get_DriveType(Mid$(DRV, P, 1)) = WhatType Then _
            Get_DriveByType = Get_DriveByType & Mid$(DRV, P, 1)
    Next P
    
End Function
Public Function DriveBox_Select(DrvListBox As Object, _
                                DriveType As DriveType) As Boolean
    Dim Drive As String
    Dim X As Long
    
    On Local Error GoTo Quit
    
    DrvListBox.Refresh
    
    Do While X < DrvListBox.ListCount
        Drive = Left$(DrvListBox.List(X), 1)
        
        If Get_DriveType(Drive) <> DriveType Then
            Call SendMessage(DrvListBox.hWnd, CB_DELETESTRING, _
                             ByVal X, ByVal CLng(0))
        Else
            X = X + 1
        End If
    Loop
    
    If DrvListBox.ListCount Then _
        Call SendMessage(DrvListBox.hWnd, _
                         CB_SETCURSEL, CLng(0), CLng(1))
    
    DriveBox_Select = True

Quit:
End Function
Public Function DriveExist(ByVal DriveLetter As String) _
                                As Boolean
    Dim A As String
            
    If Len(DriveLetter) > 1 Then Exit Function
    DriveLetter = UCase$(DriveLetter)
    
    A = Get_AllDrives
    DriveExist = (InStr(A, DriveLetter))
    
End Function
Public Function Get_CDROM(CDtype As eDrvType) As String
    Dim D As String, CD As String, R$
    Dim P As Long
    Dim DRVtype As eDrvType
    
    If Not SYS.isWindowsNT Then Exit Function
    
    D = DRV.Get_DriveByType(dtCDROM)
    
    For P = 1 To Len(D)
        CD = Mid$(D, P, 1)
        
        Call GetCDROMType(CD, DRVtype)
        
        If CDtype = DRVtype Then R$ = R$ & CD
    Next P
    
    Get_CDROM = R$
    
End Function
Private Function GetDriveHandle(ByVal DRV As String, ByRef fh As Long) As Boolean
    Dim sd As SECURITY_ATTRIBUTES
    Dim flags As Long

    flags = GENERIC_READ
    flags = flags Or GENERIC_WRITE

    fh = CreateFile("\\.\" & DRV & ":", flags, _
                    FILE_SHARE_READ Or FILE_SHARE_WRITE, sd, _
                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)

    If fh = INVALID_HANDLE_VALUE Then Exit Function

    GetDriveHandle = True
    
End Function
Private Function IsBitSet(ByVal InByte As Byte, ByVal Bit As Byte) As Boolean
    IsBitSet = ((InByte And (2 ^ Bit)) > 0)
End Function
Private Function GetCDROMType(ByVal sDrv As String, _
                              ByRef DRVtype As eDrvType) _
                                    As eDrvErrs
    Dim status As Long, hf As Long
    Dim pswb As SCSI_PASS_THROUGH_DIRECT_WITH_BUFFER
    Dim OL As OVERLAPPED
    Dim length As Long, returned As Long, i As Long
    Dim MMCaps As MMC

    If Not GetDriveHandle(sDrv, hf) Then Exit Function

    With pswb
        .spt.SenseInfoOffset = Len(.spt) + Len(.Fill)
        .spt.DataTransferLength = Len(MMCaps)
        .spt.DataBuffer = VarPtr(MMCaps)
        .spt.DataIn = SCSI_IOCTL_DATA_IN
        .spt.SenseInfoLength = 14
        .spt.TimeOutValue = 9999
        .spt.CdbLength = 10
        .spt.length = 44
    End With

    With pswb.spt
        .Cdb(0) = &H5A
        .Cdb(2) = &H2A And &H3F
        .Cdb(8) = Len(MMCaps)
    End With

    length = Len(pswb)
    status = DeviceIoControl(hf, IOCTL_SCSI_PASS_THROUGH_DIRECT, _
                             pswb, length, pswb, length, _
                             returned, OL)

    If status = 1 Then
        If pswb.spt.ScsiStatus = 0 Then

            If IsBitSet(MMCaps.ReadSupported, 0) Then _
                DRVtype = CDTYPE_CDROM

            If IsBitSet(MMCaps.ReadSupported, 3) Then _
                DRVtype = CDTYPE_DVDROM

            If IsBitSet(MMCaps.WriteSupported, 0) Then _
                DRVtype = CDTYPE_CDWRITER

            If IsBitSet(MMCaps.WriteSupported, 1) Then _
                DRVtype = CDTYPE_CDRWWRITER

            If IsBitSet(MMCaps.WriteSupported, 4) Then _
                DRVtype = CDTYPE_DVDWRITER

            GetCDROMType = ERR_OK

        Else
            GetCDROMType = ERR_SCSI_CMD
        End If
    Else
        GetCDROMType = ERR_SPTD_FAIL
    End If

    CloseHandle hf
    
End Function

