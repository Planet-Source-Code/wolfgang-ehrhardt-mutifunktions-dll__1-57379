VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type SHFILEINFO
    hIcon         As Long
    iIcon         As Long
    dwAttributes  As Long
    szDisplayName As String * MAX_PATH
    szTypeName    As String * 80
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32.dll" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ShellExExInfo As SHELLEXECUTEINFO) As Long
Private Declare Function SHGetFileInfo2 Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, riid As CLSIdType, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function ExtractFiles Lib "advpack.dll" (ByVal pszCabName As String, ByVal pszExpandDir As String, ByVal dwFlags As Long, ByVal pszFileList As String, lpReserved As Any, ByVal dwReserved As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal lpszPath As String, ByVal dx As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)

Public Enum ReadAs
    raDEC = 1
    raBIN = 2
    raHEX = 3
End Enum

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Public Enum FileOperationNFO
    foCOPY = 1
    foMOVE = 2
    foRENAME = 3
    foDELETE = 4
End Enum

Public Enum TimeInfoNFO
    tiCREATED = 1
    tiLASTACCESS = 2
    tiLASTCHANGED = 3
End Enum

Private Type VS_FIXEDFILEINFO
    dwSignature         As Long
    dwStrucVersionl     As Integer
    dwStrucVersionh     As Integer
    dwFileVersionMSl    As Integer
    dwFileVersionMSh    As Integer
    dwFileVersionLSl    As Integer
    dwFileVersionLSh    As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask     As Long
    dwFileFlags         As Long
    dwFileOS            As Long
    dwFileType          As Long
    dwFileSubtype       As Long
    dwFileDateMS        As Long
    dwFileDateLS        As Long
End Type

Private Type SHELLEXECUTEINFO
    cbSize       As Long
    fMask        As Long
    hWnd         As Long
    lpVerb       As String
    lpFile       As String
    lpParameters As String
    lpDirectory  As String
    nShow        As Long
    hInstApp     As Long
    lpIDList     As Long
    lpClass      As String
    hkeyClass    As Long
    dwHotKey     As Long
    hIcon        As Long
    hProcess     As Long
End Type

Private Type IconType
    cbSize  As Long
    picType As PictureTypeConstants
    hIcon   As Long
End Type

Private Type CLSIdType
    id(16) As Byte
End Type

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    Flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Type udtFileInfo
    CompanyName      As String
    FileDescription  As String
    FileVersion      As String
    InternalName     As String
    LegalCopyright   As String
    OriginalFilename As String
    ProductName      As String
    ProductVersion   As String
    Comments         As String
    LegalTrademarks  As String
End Type

Private Type DROPFILES
    pFiles As Long
    pt     As POINTAPI
    fNC    As Long
    fWide  As Long
End Type

Private Const OF_READ As Long = &H0&
Private Const OF_WRITE As Long = &H1&

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Private Const OFS_MAXPATHNAME As Long = 128&

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Enum SCnfo
    SC_AUTOSTART = 0
    SC_DESKTOP = 1
    SC_PROGRAMS = 2
    SC_FROMFILE = 3
    SC_SENDTO = 4
    SC_QUICKLAUNCH = 5
End Enum

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Const SHGFI_TYPENAME = &H400&
Public Function Exist(ByVal File As String) As Boolean
    Dim FN As Integer

    On Local Error GoTo Quit
    
    FN = FreeFile
    
    Open File For Input As FN
    Close FN

    Exist = True
    
Quit:
End Function
Public Function Get_FileVersion(ByVal File As String) As String
    Dim Buff() As Byte
    Dim L As Long, BuffL As Long, Pointer As Long
    Dim Version As VS_FIXEDFILEINFO
 
    L = GetFileVersionInfoSize(File, 0&)
    
    If L < 1 Then Exit Function
    
    ReDim Buff(L)
      
    Call GetFileVersionInfo(File, 0&, L, Buff(0))
    Call VerQueryValue(Buff(0), "\", Pointer, BuffL)
    Call MoveMemory(Version, Pointer, Len(Version))
      
    Get_FileVersion = Format$(Version.dwFileVersionMSh) & "." & _
                      Format$(Version.dwFileVersionMSl) & "." & _
                      Format$(Version.dwFileVersionLSh) & "." & _
                      Format$(Version.dwFileVersionLSl)
    
End Function
Public Function TempFile_Create(Optional Directory As String) _
                                    As String
    Dim Result As Long
    Dim Buff As String, Dir As String
            
    On Local Error GoTo Quit
    
    Dir = Directory
    
    If Dir = "" Then
        Dir = SYS.SystemDirectory(sdWindowsTempDir)
    Else
        If Len(Dir) = 1 Then Dir = Dir & ":\"
    End If
    
    If Dir <> "" Then
        Buff = Space$(MAX_PATH)
    
        If GetTempFileName(Dir, "~", 0&, Buff) <> 0 Then
            Result = InStr(1, Buff, Chr$(0))
            
            If Result Then _
                TempFile_Create = Buff: _
                TempFile_Create = Left$(Buff, Result - 1): _
                TempFile_Create = Get_LongFileName(TempFile_Create)
        End If
    End If
    
Quit:
    If Err.Number <> 0 Then TempFile_Create = ""
    
End Function
Public Function Get_DOSfileName(ByVal Path As String) As String
    Dim Result As Long
    Dim aa As String
    
    aa = Space$(255)
    Result = GetShortPathName(Path, aa, Len(aa))
    Get_DOSfileName = Mid$(aa, 1, Result)

End Function
Public Function Get_LongFileName(ByVal Path As String) As String
    Dim Result As Long
    Dim aa As String
    
    aa = Space$(255)
    Result = GetLongPathName(Path, aa, Len(aa))
    Get_LongFileName = Mid$(aa, 1, Result)

End Function
Public Function Get_FileDescription(ByVal File As String) As String
    Dim L As Integer
    Dim SHInf As SHFILEINFO
  
    File = GetDir(File)
    
    Call SHGetFileInfo(File, 0, SHInf, Len(SHInf), SHGFI_TYPENAME)
    
    L = InStr(1, SHInf.szTypeName, Chr$(0)) - 1
    Get_FileDescription = Left$(SHInf.szTypeName, L)
    
End Function
Public Function Show_FilePropertiesDialog(ByVal hWnd As Long, _
                                          ByVal File As String) _
                                                As Boolean
    Dim ShInfo As SHELLEXECUTEINFO
  
    On Local Error GoTo Quit
  
    With ShInfo
        .cbSize = Len(ShInfo)
        .fMask = SEE_MASK_NOCLOSEPROCESS _
                 Or SEE_MASK_INVOKEIDLIST _
                 Or SEE_MASK_FLAG_NO_UI
        .hWnd = hWnd
        .lpVerb = "properties"
        .lpFile = File
        .lpIDList = 0
    End With

    Show_FilePropertiesDialog = ShellExecuteEx(ShInfo)
    
Quit:
End Function
Public Function Get_LinkedIconFromFile(ByVal File As String, _
                                       Optional SmallIcon As Boolean = False) _
                                             As IPictureDisp
    Dim SIZE As Long
    Dim Unkown As IUnknown
    Dim iCon As IconType
    Dim CLSID As CLSIdType
    Dim ShellInfo As SHFILEINFO
    
    SIZE = IIf(SmallIcon, &H101, &H100)
    
    Call SHGetFileInfo2(File, 0, ShellInfo, Len(ShellInfo), SIZE)
 
    iCon.cbSize = Len(iCon)
    iCon.picType = vbPicTypeIcon
    iCon.hIcon = ShellInfo.hIcon
    
    CLSID.id(8) = &HC0
    CLSID.id(15) = &H46
    
    Call OleCreatePictureIndirect(iCon, CLSID, 1, Get_LinkedIconFromFile)
    
End Function
Public Function Show_FileDialogOpen(ByVal hWnd As Long, _
                                    Optional Filter As String = "*.*") _
                                            As String
    Dim Buffer As String
    Dim Flags As Long
    Dim ComDlgOpenFileName As OPENFILENAME
  
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
    
    Buffer = String$(128, 0)
   
    With ComDlgOpenFileName
        .lStructSize = Len(ComDlgOpenFileName)
        .hwndOwner = hWnd
        .Flags = Flags
        .nFilterIndex = 1
        .nMaxFile = Len(Buffer)
        .lpstrFile = Buffer
        .lpstrFilter = Filter
    End With
    
    If GetOpenFileName(ComDlgOpenFileName) Then _
        Show_FileDialogOpen = Left$(ComDlgOpenFileName.lpstrFile, _
                                    InStr(ComDlgOpenFileName.lpstrFile, _
                                    Chr$(0)) - 1)
    
End Function
Public Function Show_FileDialogSave(ByVal hWnd As Long, _
                                    Optional Filter As String = "*.*", _
                                    Optional FileName As String) _
                                            As String
    Dim Buffer As String
    Dim ComDlgOpenFileName As OPENFILENAME
  
    Buffer = FileName & String$(128 - Len(FileName), 0)
    
    With ComDlgOpenFileName
        .lStructSize = Len(ComDlgOpenFileName)
        .hwndOwner = hWnd
        .Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
        .nFilterIndex = 1
        .nMaxFile = Len(Buffer)
        .lpstrFile = Buffer
        .lpstrFilter = Filter
    End With

    If GetSaveFileName(ComDlgOpenFileName) Then _
        Show_FileDialogSave = Left$(ComDlgOpenFileName.lpstrFile, _
                                    InStr(ComDlgOpenFileName.lpstrFile, _
                                    Chr$(0)) - 1)
    
End Function
Public Function Create_File(ByVal File As String, _
                            Optional DeleteExisting As Boolean = False) _
                                As Boolean
    Dim FN As Integer
        
    On Local Error GoTo Quit
    
    If F.Exist(File) Then
        If Not DeleteExisting Then Exit Function
        Call F.fKill(File)
    End If
    
    FN = FreeFile
    
    Open File For Output As FN
    Close FN

    Create_File = (Err.Number = 0)

Quit:
End Function
Public Function Get_FileLen(ByVal File As String) As Long
    Get_FileLen = FileLen(File)
End Function
Public Function Find(ByVal FileToFind As String, _
                     ByVal StartSearchPath As String, _
                     Optional SearchAllDrives As Boolean = False, _
                     Optional ShowProgress As Boolean = False) _
                            As String
    Dim P As Long
    Dim sDRV As String, cDrv As String
    
    If StartSearchPath = "" And Not SearchAllDrives Then _
        Exit Function
    
    If ShowProgress Then pModule.LoadProgbar = True
    
    If SearchAllDrives Then
        sDRV = DRV.Get_AllDrives
                
        sDRV = Replace(sDRV, "A", "")
        sDRV = Replace(sDRV, "B", "")
        
        For P = 1 To Len(sDRV)
            cDrv = Mid$(sDRV, P, 1) & ":\"
            Find = GetAllFiles(cDrv, FileToFind, False, ShowProgress)
            If Len(Find) Then Exit For
        Next P
    Else
        Find = GetAllFiles(StartSearchPath, _
                           FileToFind, False, ShowProgress)
    End If
    
    If ShowProgress Then Unload pModule.PBar: _
                         Set pModule.PBar = Nothing
    
End Function
Public Function Compare(ByVal File1 As String, _
                        ByVal File2 As String) As Boolean
        
    On Local Error Resume Next
    
    If FileLen(File1) = FileLen(File2) Then _
        Compare = (TXT.TextFile_Read(File1) = TXT.TextFile_Read(File2))

End Function
Public Function isInUse(ByVal Path As String) As Boolean
    Dim FN As Integer
  
    On Local Error Resume Next
    
    FN = FreeFile

    Open Path For Binary Access Write As #FN
    Close #FN
    
    isInUse = (Err.Number <> 0)
    
End Function
Public Function ShredderFile(ByVal FilePath As String, _
                             Optional WriteLoop As Integer = 3) _
                                    As Boolean
    Dim A As String, B As String, Key As String
    Dim FN As Integer
    Dim L As Long, X As Long, P As Long
        
    On Local Error GoTo Quit
    
    L = FileLen(FilePath)
        
    If WriteLoop = 0 Then WriteLoop = 1
    If WriteLoop < 0 Then WriteLoop = WriteLoop * (-1)

    For P = 1 To WriteLoop
        Key = ""
            
        For X = 1 To 10
            A = CStr(MISC.Get_Randomize(0, 9))
            B = Chr$(MISC.Get_Randomize(Asc("A"), Asc("E")))
        
            Key = Key & A & B
        Next X
            
        FN = FreeFile

        Open FilePath For Binary As #FN
            For X = 1 To (L / 5)
                Put #FN, , Key
                Call FlushFileBuffers(1)
            Next X
        Close #FN
    Next P
    
    ShredderFile = (fKill(FilePath) = 1)
            
Quit:
End Function
Public Function SFV_GetSFVFromFile(ByVal File As String) _
                                        As String
    
    SFV_GetSFVFromFile = SFV(CVar(File))

End Function
Public Function SFV_Compare(ByVal File1 As String, _
                            ByVal File2 As String) _
                                As Boolean
    SFV_Compare = (SFV(File1) = SFV(File2))
End Function
Public Function SFV_Validate(ByVal File As String, _
                             SFVs As String) As Boolean
    SFV_Validate = (SFVs = SFV(File))
End Function
Public Function SFV_CreateSFVfile(ByVal File As Variant, _
                                  SFVfile As String) As Boolean
    Dim H$, L$, sSFV As String, A$
    Dim F() As Variant
    Dim P As Long
    
    If Not isArray(File) Then
        ReDim F(0)
        F(0) = File
    Else
        F = File
    End If
    
    H$ = "; Generated by MFDLL " & _
         App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
         "; (Compatible: WIN-SFV32)" & vbCrLf & vbCrLf
         
    For P = 0 To UBound(F)
        L$ = Get_FileFromPath(F(P)) & " " & SFV(F(P)) & vbCrLf
        A$ = A$ & L$
    Next P
    
    A$ = H$ & A$
    
    SFV_CreateSFVfile = TXT.TextFile_WriteText(SFVfile, A$, True)
    
End Function
Public Function isEXERunning(ByVal ExeName As String) _
                                As Boolean
    Dim E() As String
    Dim I() As Long
        
    If SYS.Get_ProcessList(E, I) Then _
        isEXERunning = (ARR.Find(E, ExeName, , , False) > -1)
        
End Function
Public Function Get_FileFromPath(ByVal Path As String) As String
    
    On Local Error Resume Next
    
    If Right$(Path, 1) <> "\" Then _
        Get_FileFromPath = Mid$(Path, SplitPath(Path) + 1)
        
End Function
Public Function Get_DirectoryFromPath(ByVal Path As String) As String
    On Local Error Resume Next
    
    If Right$(Path, 1) = "\" Then
        Get_DirectoryFromPath = Path
    Else
        Get_DirectoryFromPath = GetDir(Mid$(Path, 1, SplitPath(Path) - 1))
    End If
    
End Function
Public Function FileOperation(Operation As FileOperationNFO, _
                              ByVal Source As String, _
                              ByVal Target As String, _
                              Optional MoveToRecycleBin As Boolean = False) _
                                    As Boolean
                                    
    FileOperation = SH_FileOperation(CLng(0), Operation, _
                                     Source, Target, _
                                     False, False, False, _
                                     False, MoveToRecycleBin)
End Function
Public Function FileOperationUsingShell _
                    (Operation As FileOperationNFO, _
                     ByVal hWnd As Long, _
                     ByVal Source As String, _
                     ByVal Target As String, _
                     Optional Confirm As Boolean = True, _
                     Optional ShowDialogs As Boolean = True, _
                     Optional ShowProgress As Boolean = True, _
                     Optional MoveToRecycleBin As Boolean = False) _
                                As Boolean
    
    FileOperationUsingShell = SH_FileOperation(hWnd, Operation, _
                                     Source, Target, _
                                     False, Confirm, _
                                     ShowDialogs, _
                                     ShowProgress, _
                                     MoveToRecycleBin)

End Function
Public Function Get_RootFromPath(ByVal Path As String) As String
    Dim Y As Long, X As Long
    Dim R$
    
    On Local Error GoTo Quit
    
    X = InStr(Path, ":\")
    If X = 0 Then Exit Function
    
    Y = 1
    
    If Len(Path) > X + 1 Then _
        If Mid$(Path, X + 2, 1) = "\" Then Y = 2
    
    R$ = Mid$(Path, 1, X + Y)
    
    Get_RootFromPath = R$
    
Quit:
End Function
Private Function CalcFTime(FTime As SYSTEMTIME) As Date
    Dim Datum As String, Zeit As String, aa As String
    Dim mm As String, ss As String, DT As Date
    Dim Da As String, Mo As String, Ye As String

    With FTime
        Da = .wDay
        If Len(Da) < 2 Then Da = "0" & Da
    
        Mo = .wMonth
        If Len(Mo) < 2 Then Mo = "0" & Mo
    
        Ye = CStr(.wYear)
        Datum = Da & "." & Mo & "." & Ye
    
        mm = Trim$(CStr(.wMinute))
        If Len(mm) < 2 Then mm = "0" & mm
    
        ss = Trim$(CStr(.wSecond))
        If Len(ss) < 2 Then ss = "0" & ss
    
        Zeit = .wHour & ":" & mm & ":" & ss
        DT = CDate(Datum & " " & Zeit)
        CalcFTime = DT
    End With

End Function
Public Function CABfile_Extract(ByVal CABFile As String, _
                                ByVal SourceFile As String, _
                                ByVal DestPath As String) _
                                    As Boolean
    On Local Error Resume Next
    
    If ExtractFiles(CABFile, DestPath, 0, SourceFile, 0, 0) = 0 Then _
        CABfile_Extract = True

End Function
Public Function ShortCut_Create(Where As SCnfo, _
                                SC_SourceFile As String, _
                                SC_Text As String, _
                                Optional SC_TargetFile As String) _
                                    As Boolean
    Dim Path As String, FullName As String, Str As String
    Dim fso As New FileSystemObject
    Dim WSH As New IWshShell_Class
    Dim SCut, DestFolder As Folder
    Dim X As Long
    
    On Local Error GoTo Quit
    
    ShortCut_Create = True
    
    Select Case Where
        Case SC_AUTOSTART
            Path = ".\Autostart"
        Case SC_DESKTOP
            Path = "..\..\Desktop"
        Case SC_PROGRAMS
            Path = "."
        Case SC_SENDTO, SC_QUICKLAUNCH
            Select Case Where
                Case SC_SENDTO
                    X = sdSendToDirectory
                Case SC_QUICKLAUNCH
                    X = sdUserQuickLaunchDir
            End Select
            
            Path = SYS.SystemDirectory(X)
            
            Str = GetBlankName(SC_SourceFile) & ".lnk"
            Path = Path & Str
            
            ShortCut_Create = ShortCut_Create(SC_FROMFILE, SC_SourceFile, _
                                              SC_Text, Path)
            
            Exit Function
        Case SC_FROMFILE
            If Not fso.FileExists(SC_TargetFile) Then _
                Set SCut = WSH.CreateShortcut(SC_TargetFile): _
                SCut.TargetPath = SC_SourceFile: _
                SCut.Save
        
            Set SCut = Nothing
            Set fso = Nothing
            Set WSH = Nothing
            
            Exit Function
    End Select
    
    Call CreateShellLink(Path, SC_Text, SC_SourceFile, "", _
                         True, "$(Programs)")
    
Quit:
    ShortCut_Create = (Err.Number = 0)
    
End Function
Public Function Split(ByVal File As String, _
                      ByVal KBperFile As Long, _
                      Optional OutputFile As String, _
                      Optional CreateBatch As Boolean = False) _
                            As Long
    Dim Zieln As String, Batch As String, R$, F As String
    Dim Max As Long, P As Long, I As Long, L As Long
    Dim Found As Boolean
        
    On Local Error GoTo Quit
  
    Max = KBperFile * 1024
    
    If Max < 1 Then Exit Function
    If Not Exist(File) Then Split = -1: _
                            Exit Function
  
    L = FileLen(File)
    If L < Max Then Split = -2: _
                    Exit Function
    
    If OutputFile <> vbNullString Then
        OutputFile = OutputFile & "."
    Else
        For P = Len(File) To 1 Step -1
            If Mid$(File, P, 1) = "." Then _
                Found = True: _
                Exit For
        Next P
        
        If Found Then
            OutputFile = Mid$(File, 1, P)
        Else
            OutputFile = OutputFile & "."
        End If
    End If
        
    Batch = OutputFile & "Bat"
        
    R$ = ""
    
    Open File For Binary As 1
        R$ = R$ & "copy /b "
        I = 1
      
        For P = 0 To L Step Max
            Zieln = OutputFile & Format$(I, "000")
                   
            If Exist(Zieln) Then Kill Zieln
        
            Open Zieln For Binary As 3
                If L - P < Max Then
                    ByteCopy L - P, 1, 3
                Else
                    ByteCopy Max, 1, 3
                End If
            Close 3
    
            If I > 1 Then R$ = R$ & " + "
            R$ = R$ & Get_FileFromPath(Zieln)

            I = I + 1
        Next P
  
        R$ = R$ & " " & Get_FileFromPath(File)
    Close 1
    
    If CreateBatch Then _
        If Not TXT.TextFile_WriteText(Batch, R$, True) Then _
            Split = -3: _
            Exit Function
            
    Split = 1
   
Quit:
    If Err.Number <> 0 Then Split = Err.Number
    
End Function
Private Sub ByteCopy(Bytes, Quelle, Ziel)
    Dim I As Long
    Dim Puffer() As Byte
  
    ReDim Puffer(1 To 32768)
  
    For I = 0 To Bytes - 1 Step 32768
        If Bytes - I < 32768 Then ReDim Puffer(1 To Bytes - I)
        Get Quelle, , Puffer
        Put Ziel, , Puffer
    Next I

End Sub
Public Function Join(ByVal StartFile As String, _
                     Optional OutputFile As String) As Long
    Dim L As Long, Max As Long, P As Long, I As Long
    Dim Target As String, Source As String, R$
        
    On Local Error GoTo Quit
    
    If Not Exist(StartFile) Then Join = -1: _
                                 Exit Function
                                
    If Right$(StartFile, 4) <> ".001" Then Join = -2: _
                                          Exit Function
                                      
    If FileLen(StartFile) = 0 Then Join = -3: _
                                   Exit Function
                
    If OutputFile = vbNullString Then _
        OutputFile = Left$(StartFile, Len(StartFile) - 3) & "out": _
        R$ = Left$(StartFile, Len(StartFile) - 3)
    
    If Exist(OutputFile) Then Kill OutputFile

    If Not Exist(R$ & Format$(2, "000")) Then Join = -4: _
                                              Exit Function
    
    I = 1
    Max = 1024000
    
    Source = R$ & Format$(I, "000")
    
    Open OutputFile For Binary As 1
        Do While Exist(Source)
            Open Source For Binary As 3
                L = FileLen(Source)

                For P = 0 To L Step Max
                    If L - P < Max Then
                        ByteCopy L - P, 3, 1
                    Else
                        ByteCopy Max, 3, 1
                    End If
                    
                    DoEvents
                Next P
            Close 3
            
            I = I + 1
            Source = R$ & Format$(I, "000")
        Loop
    Close 1
    
    Join = 1
    
Quit:
    If Err.Number <> 0 Then Join = Err.Number
    
End Function
Public Function Get_FileTime(ByVal File As String, _
                             ByVal FileInfo As TimeInfoNFO, _
                             Optional TimeOnly As Boolean = False, _
                             Optional DateOnly As Boolean = False) _
                                    As Date
    Dim DTformat As String, A() As String
    Dim hFile As Long
    Dim OFS As OFSTRUCT
    Dim cTime As FILETIME, lTime As FILETIME, lwTime As FILETIME
    Dim sTime As SYSTEMTIME
     
    On Local Error Resume Next
    
    OFS.cBytes = Len(OFS)
    hFile = OpenFile(File, OFS, OF_READ)
    
    If hFile > 0 Then
        DTformat = "dd.mm.yyyy hh:mm:ss"
        Call GetFileTime(hFile, cTime, lTime, lwTime)
        
        Select Case FileInfo
            Case tiCREATED
                Call FileTimeToSystemTime(cTime, sTime)
                Get_FileTime = Format$(CalcFTime(sTime), DTformat)
            Case tiLASTACCESS
                Call FileTimeToSystemTime(lTime, sTime)
                Get_FileTime = Format$(CalcFTime(sTime), DTformat)
            Case tiLASTCHANGED
                Call FileTimeToSystemTime(lwTime, sTime)
                Get_FileTime = Format$(CalcFTime(sTime), DTformat)
            Case Else: Exit Function
        End Select
    
        Call CloseHandle(hFile)
                
        If TimeOnly Or DateOnly Then _
            A = VBA.Split(Get_FileTime, Chr$(32)): _
            Get_FileTime = A(0): _
            If TimeOnly Then Get_FileTime = A(1)
    Else
        Get_FileTime = CDate(0)
    End If
    
Quit:
End Function
Public Function Binary_Read(ByVal File As String, _
                            ByVal ReadLenght As Long, _
                            ReadAs As ReadAs, _
                            Optional ByVal Offset As Long = 1) _
                                As String()
    Dim A() As Byte, B() As Byte
    Dim P As Long, Max As Long, Count As Long
    Dim Str() As Variant
    Dim FN As Integer
    
    On Local Error GoTo Quit
    
    Offset = Offset - 1
    
    If Not Exist(File) Then Exit Function
    
    If ReadLenght Then
        ReadLenght = ReadLenght - 1
    Else
        ReadLenght = FileLen(File)
        Offset = 1
    End If
    
    If Offset + ReadLenght > FileLen(File) + 1 Then Exit Function
    
    ReDim A(ReadLenght)
    ReDim B(10000)
    
    Max = Offset / 10000
    FN = FreeFile
    
    Open File For Binary As FN
        For P = 1 To Max
            Get FN, , B
        Next P
        
        If Offset Then _
            ReDim B((P - 1) * Max + Offset): _
            Get FN, , B
    
        Get FN, , A
    Close FN

    ReDim Str(UBound(A))
    
    For P = 0 To UBound(A)
        Select Case ReadAs
            Case raBIN: Str(P) = MA.DecToBin(A(P))
                        Str(P) = String(8 - Len(Str(P)), "0") & Str(P)
            Case raDEC: Str(P) = CStr(A(P))
                        If Len(Str(P)) < 2 Then _
                            Str(P) = String(2 - Len(Str(P)), "0") & Str(P)
            Case raHEX: Str(P) = Hex(A(P))
                        Str(P) = String(2 - Len(Str(P)), "0") & Str(P)
        End Select
        
        Count = Count + 1
        
        If Count > 99 Then Count = 0: _
                           DoEvents
    Next P
    
    Binary_Read = ARR.VarriantToString(Str)
    
Quit:
    Erase Str: Erase A: Erase B
    
End Function
Public Function BinCopy(ByVal SourceFile As String, _
                        ByVal TargetDirectory As String, _
                        Optional OverwriteExist As Boolean = True, _
                        Optional ShowProgress As Boolean = False, _
                        Optional MessageText As String = "") _
                                As Long
    Dim X As Long, Max As Long, L As Long, P As Long
    Dim Target As String
    Dim frm As New ProgBar
            
    On Local Error GoTo Quit
    
    If Not Exist(SourceFile) Then BinCopy = 0: _
                                  GoTo Quit
    
    TargetDirectory = GetDir(TargetDirectory)
        
    If Not DI.Exist(TargetDirectory, False) Then _
        If Not DI.CreatePath(TargetDirectory) Then _
            BinCopy = -1: _
            GoTo Quit
        
    X = SplitPath(SourceFile)
    Target = TargetDirectory & Mid$(SourceFile, X + 1)
    
    If Exist(Target) Then
        If Not OverwriteExist Then BinCopy = -2: _
                                   GoTo Quit
        If fKill(Target) <> 1 Then BinCopy = -3: _
                                   GoTo Quit
    End If
    
    Max = 1024000
    L = FileLen(SourceFile)

    If ShowProgress Then
        Load frm
        frm.Tag = "BinCopy"
            
        Call frm.SetOption(0, L, MessageText, True)
    End If
        
    X = 0
        
    Open SourceFile For Binary As 3
    Open Target For Binary As 1

        For P = 0 To L Step Max
            If ShowProgress And P = 0 Then _
                frm.Message.Caption = SourceFile & " -> " & Target
                
            If L - P < Max Then
                X = X + (L - P)
                ByteCopy L - P, 3, 1
            Else
                X = X + Max
                ByteCopy Max, 3, 1
            End If
                
            If ShowProgress Then Call frm.SetValue(X)
          
            DoEvents
        Next P
            
    Close 3
    Close 1
    
    BinCopy = 1
    
Quit:
    If ShowProgress Then Unload frm
    Set frm = Nothing

    If Err.Number <> 0 Then BinCopy = Err.Number
    
End Function
Public Function fKill(ByVal File As String) As Long
    
    On Local Error Resume Next
    
    If Not Exist(File) Then Exit Function
    
    fKill = 1
        
    Call SetAttr(File, vbNormal)
    Err.Clear
        
    Kill File
    If Err.Number <> 0 Then fKill = -1
    
End Function
Public Function Get_PrefixFromPath(ByVal Path As String) As String
    Dim P As Long
        
    If Right$(Path, 1) = "." Then Exit Function
    
    For P = Len(Path) To 1 Step -1
        If Mid$(Path, P, 1) = "." Then _
            Get_PrefixFromPath = Mid$(Path, P + 1): _
            Exit For
    Next P
    
End Function
Public Function CompactPath(ByVal Path As String, _
                            Obj As Object) As String
    Dim dx As Long, X As Long
    
    On Local Error GoTo Quit
    
    dx = Obj.Width / Screen.TwipsPerPixelX
        
    Call PathCompactPath(Obj.Parent.hDC, Path, dx)
    
    X = InStr(1, Path, Chr$(0))
    
    If X <> 0 Then CompactPath = Left$(Path, X - 1)
    
Quit:
    If Err.Number <> 0 Then CompactPath = ""
    
End Function
Public Function Binary_Write(ByVal File As String, _
                             sByte() As Byte, _
                             Optional OverwriteExist As Boolean = False) _
                                    As Boolean
    Dim FN As Integer
    
    On Local Error GoTo Quit
    
    If Not OverwriteExist Then _
        If Exist(File) Then Exit Function
    
    Call fKill(File)
    
    FN = FreeFile
    
    Open File For Binary As FN
        Put FN, , sByte
    Close FN
    
    Binary_Write = True
    
Quit:
End Function
Public Function Get_ExtendedFileInfo(ByVal File As String) As Variant
    Dim A() As Variant
    Dim uInfo As udtFileInfo
    
    On Local Error GoTo Quit
    
    If Not Exist(File) Then Exit Function
    
    uInfo = GetExtendedFileInfo(File)

    ReDim A(9)

    With uInfo
        A(0) = .CompanyName
        A(1) = .FileDescription
        A(2) = .FileVersion
        A(3) = .InternalName
        A(4) = .LegalCopyright
        A(5) = .OriginalFilename
        A(6) = .ProductName
        A(7) = .ProductVersion
        A(8) = .Comments
        A(9) = .LegalTrademarks
    End With
    
    Get_ExtendedFileInfo = A
    
Quit:
End Function
Private Function GetExtendedFileInfo(ByVal sFile As String) As udtFileInfo
    Dim nSize As Long
    Dim sBuffer As String
    
    nSize = GetFileVersionInfoSize(sFile, 0&)
    
    If nSize = 0 Then Exit Function
  
    sBuffer = Space$(nSize)
  
    If GetFileVersionInfo(sFile, 0&, nSize, ByVal strptr(sBuffer)) = 0 Then _
        Exit Function
  
    With GetExtendedFileInfo
        .Comments = ExtractData(sBuffer, "Comments")
        .CompanyName = ExtractData(sBuffer, "CompanyName")
        .FileDescription = ExtractData(sBuffer, "FileDescription")
        .FileVersion = ExtractData(sBuffer, "FileVersion")
        .InternalName = ExtractData(sBuffer, "InternalName")
        .LegalCopyright = ExtractData(sBuffer, "LegalCopyright")
        .LegalTrademarks = ExtractData(sBuffer, "Trademarks")
        .OriginalFilename = ExtractData(sBuffer, "OriginalFilename")
        .ProductName = ExtractData(sBuffer, "ProductName")
        .ProductVersion = ExtractData(sBuffer, "ProductVersion")
    End With

End Function
Private Function ExtractData(ByVal sBuffer As String, _
                             ByVal sLookFor As String) As String
    Dim nPos As Long
    Dim sData As String
  
    nPos = InStr(sBuffer, Chr$(1) + sLookFor)
    
    If nPos > 0 Then
        sData = Mid$(sBuffer, nPos + Len(sLookFor) + 2)
        If Left$(sData, 1) = vbNullChar Then sData = Mid$(sData, 2)
        nPos = InStr(sData, vbNullChar)
        If nPos > 0 Then sData = Left$(sData, nPos - 1)
    End If
  
    ExtractData = sData

End Function
Public Function FilesToClipboard(FileList As Variant) As Boolean
    Dim sFiles As String, sPath As String
    Dim pTmp As Long, hTmp As Long, P As Long
    Dim F() As Variant
    Dim DROPFILES As DROPFILES
 
    On Local Error GoTo Quit
 
    If Not isArray(FileList) Then
        ReDim F(0)
        F(0) = FileList
    Else
        F = FileList
    End If
        
    For P = LBound(F) To UBound(F)
        sFiles = sFiles & F(P) & vbNullChar
    Next P
    
    sFiles = sFiles & vbNullChar
    
    If OpenClipboard(0&) Then
        EmptyClipboard
 
        hTmp = GlobalAlloc((GMEM_MOVEABLE Or GMEM_ZEROINIT), _
                            Len(DROPFILES) + Len(sFiles))
 
        If hTmp Then
            pTmp = GlobalLock(hTmp)
 
            DROPFILES.pFiles = Len(DROPFILES)
            Call CopyMemory(ByVal pTmp, DROPFILES, Len(DROPFILES))
            Call CopyMemory(ByVal (pTmp + Len(DROPFILES)), ByVal sFiles, Len(sFiles))
      
            Call GlobalUnlock(hTmp)
 
            If SetClipboardData(CF_HDROP, hTmp) Then _
                FilesToClipboard = True
        End If
    End If

Quit:
    Call CloseClipboard
    
End Function
Public Function ReadAsByte(ByVal File As String) As Byte()
    Dim FN As Integer
    Dim B() As Byte
    Dim L As Long
    Dim A As String
    
    On Local Error GoTo Quit
    
    If Not Exist(File) Then Exit Function
    
    FN = FreeFile
    
    Open File For Binary As #FN
        A = Input(LOF(1), #FN)
    Close FN
    
    L = Len(A) - 1
    ReDim B(0 To L)
    
    Call CopyMemory(B(0), ByVal A, L + 1)

Quit:
    If Err.Number <> 0 Then Erase B
    
    ReadAsByte = B
    
End Function
Public Function SaveAsByte(ByRef B() As Byte, _
                           ByVal File As String, _
                           Optional OverwriteExist As Boolean = False) _
                                As Boolean
    Dim FN As Integer
    
    On Local Error GoTo Quit
    
    If Exist(File) Then
        If Not OverwriteExist Then Exit Function
        Call fKill(File)
    End If
  
    FN = FreeFile
    
    Open File For Binary Access Write As #FN
        Put #FN, , B()
    Close #FN
    
Quit:
    SaveAsByte = (Err.Number) = 0
    
End Function
Public Function RegisterExtension(ByVal AppName As String, ByVal AppPath As String, _
                                  ByVal AppReg As String, ByVal AppExt As String, _
                                  Optional IconFile As String = "", _
                                  Optional IconId As Long) As Boolean
    
    If Not REG.RegKeyCreate(HKEY_CLASSES_ROOT_, AppReg, AppName) Then _
        Exit Function
        
    If Len(IconFile) > 0 Then
        If Not IsMissing(IconId) Then IconFile = RTrim$(IconFile) & "," & IconId

        If Not REG.RegKeyCreate(HKEY_CLASSES_ROOT_, AppReg & "\DefaultIcon", IconFile) Then _
            Exit Function
    End If

    If Not REG.RegKeyCreate(HKEY_CLASSES_ROOT_, "." & AppExt, AppReg) Then _
        Exit Function
    
    If Not REG.RegKeyCreate(HKEY_CLASSES_ROOT_, AppReg & "\shell\open\command", AppPath & " %1") Then _
        Exit Function
        
    RegisterExtension = True
    
End Function
