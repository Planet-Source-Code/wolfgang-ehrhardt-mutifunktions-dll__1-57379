VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type LUID
    LowPart  As Long
    HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT        As LUID
    Attributes     As Long
End Type

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LockWorkStationA Lib "user32.dll" Alias "LockWorkStation" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As Any, ByRef ReturnLength As Any) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValueA Lib "advapi32" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlgas As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function NetUserGetInfo Lib "netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function PdhOpenQuery Lib "PDH.DLL" (ByVal Reserved As Long, ByVal dwUserData As Long, ByRef hQuery As Long) As Long
Private Declare Function PdhCloseQuery Lib "PDH.DLL" (ByVal hQuery As Long) As Long
Private Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As Long
Private Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As Long
Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Private Declare Function NetLocalGroupAddMembers Lib "netapi32.dll" (ByVal servername As String, ByVal GroupName As String, ByVal Level As Long, ByVal buf As String, ByVal totalentries As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Const TH32CS_SNAPPROCESS As Long = 2&

Public Type ProcessNFO
    EXE As String
    id  As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Type PROCESSENTRY32
    dwSize              As Long
    cntUsage            As Long
    th32ProcessID       As Long
    th32DefaultHeapID   As Long
    th32ModuleID        As Long
    cntThreads          As Long
    th32ParentProcessID As Long
    pcPriClassBase      As Long
    dwFlags             As Long
    szexeFile           As String * MAX_PATH
End Type

Private Type LOCALGROUP_MEMBERS_INFO_3
    lgrmi3_domainandname As Long
End Type

Private Type USER_INFO
    usri1_name         As Long
    usri1_password     As Long
    usri1_password_age As Long
    usri1_priv         As Long
    usri1_home_dir     As Long
    usri1_comment      As Long
    usri1_flags        As Long
    usri1_script_path  As Long
End Type

Public Enum USER_PRIV
    USER_PRIV_GUEST = 0
    USER_PRIV_USER = 1
    USER_PRIV_ADMIN = 2
End Enum

Public Enum oS
    osUNKNOWN = 0
    osWindows32s = 1
    osWindows95 = 2
    osWindows98 = 3
    osWindows98SE = 4
    osWindowsNT35x = 5
    osWindowsNT40 = 6
    osWindowsMillenium = 7
    osWindows2000 = 8
    osWindowsXP = 9
    osWindows2003 = 10
    osWindowsNET = osWindows2003
End Enum

Public Enum ShutDownType
    sdSHUTDOWN = 1
    sdREBOOT = 2
    sdLOGOFF = 3
End Enum

Public Enum DLLType
    dtNODLL = 0
    dtSTANDARDDLL = 1
    dtACTIVEXDLL = 2
End Enum

Private Type ITEMID
    CB   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ITEMID
End Type

Public Enum SystemDirNFO
    sdCommonProgramFilesDir = 1
    sdAllUsersProfileDir = 2
    sdUserProfileDir = 4
    sdWindowsDirectory = 5
    sdWindowsSystemDir = 6
    sdWindowsTempDir = 7
    sdStartmenuDir = 8
    sdDesktopDirectory = 9
    sdProgramFilesDir = 10
    sdPersonalDirectory = 11
    sdFavoritDirectory = 12
    sdAutoStartDirectory = 13
    sdRecentDocumentDir = 14
    sdSendToDirectory = 15
    sdNewShellDir = 16
    sdNetworkDirectory = 17
    sdFontDirectory = 18
    sdTemporaryInternetFilesDir = 19
    sdUserDesktopDir = 20
    sdUserOwnMusicDir = 21
    sdUserOwnPicsDir = 22
    sdUserCookiesDir = 23
    sdAllUserDesktopDir = 24
    sdUserAppDataDir = 25
    sdUserQuickLaunchDir = 26
End Enum

Public Enum BootModeConstants
    bmNormal = 0
    bmSafeMode = 1
    bmSafeModeNetwork = 2
    bmUNKNOWN = -1
End Enum

Public Enum SecurityOption
    TASKMANAGER = 0
    REGEDIT = 1
    LOGOFF = 2
    WINDOWSUPDATE = 3
    COMMANDCOM = 4
    SAVEDESKTOP = 5
    CONTROLPANEL = 6
    NETWORKICON = 7
End Enum

Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const EWX_POWEROFF = 8

Private Const BITSPIXEL = 12
Private Const PLANES = 14

Private Const PDH_CSTATUS_VALID_DATA = &H0
Private Const PDH_CSTATUS_NEW_DATA = &H1

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Private Const SE_PRIVILEGE_ENABLED = &H2

Private Const SM_CLEANBOOT = &H43

Private Const NOERROR = 0

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private hQuery As Long
Private hCounter As Long
Private Retval As Long

Private Stack() As Long
Private StackPointer As Long
Private m_StackSize As Long
Private m_Sum As Long

Private m_BandWidth As Long
Public Function Get_DOScommandPath() As String
    Get_DOScommandPath = Environ("comspec")
End Function
Public Function Get_SystemDrive() As String
    Get_SystemDrive = GetDir(Environ("systemdrive"))
End Function
Public Function Get_LocalUserName() As String
    Dim L As Long, Result As Long
    Dim User As String, UserName As String
  
    User = Space(255)
    L = 255
    Result = GetUserName(User, L)
    
    If Result Then UserName = Left$(User, L - 1): _
                   Get_LocalUserName = Trim$(UserName)

End Function
Public Function WindowsIsRunning() As String
    Dim H As Single, M As Single, S As Single, MS As Single
    Dim strH As String, strM As String, strS As String
    Dim strMS As String
    
    MS = GetTickCount()
    MS = MS / 1000
    
    H = Int(MS / 3600)
    MS = MS - H * 3600
    M = Int(MS / 60)
    MS = MS - M * 60
    S = Int(MS)
    MS = Int((MS - S) * 10)
    
    strH = CStr(H)
    strM = Format(CStr(M), "##00")
    strS = Format(CStr(S), "##00")
    strMS = CStr(MS)
    
    WindowsIsRunning = strH & ":" & strM & ":" & strS & ":" & strMS

End Function
Public Function Get_WinVer() As oS
    Dim WinVer As String
    Dim udtOS As OSVERSIONINFO

    udtOS.dwOSVersionInfoSize = Len(udtOS)

    Call GetVersionEx(udtOS)

    Select Case udtOS.dwPlatformId
        Case VER_PLATFORM_WIN32_NT
            If udtOS.dwMajorVersion = 3 Then
                Get_WinVer = osWindowsNT35x
        ElseIf udtOS.dwMajorVersion = 4 Then
            Get_WinVer = osWindowsNT40
        ElseIf udtOS.dwMajorVersion = 5 Then
            If udtOS.dwMinorVersion = 0 Then
                Get_WinVer = osWindows2000
            ElseIf udtOS.dwMinorVersion = 1 Then
                Get_WinVer = osWindowsXP
            ElseIf udtOS.dwMinorVersion = 2 Then
                Get_WinVer = osWindows2003
            End If
        End If
    Case VER_PLATFORM_WIN32_WINDOWS
        If udtOS.dwMinorVersion = 0 Then
             Get_WinVer = osWindows95
        ElseIf udtOS.dwMinorVersion = 10 Then
             Get_WinVer = osWindows98
        ElseIf udtOS.dwMinorVersion = 90 Then
            Get_WinVer = osWindowsMillenium
        End If
    Case VER_PLATFORM_WIN32s
        Get_WinVer = osWindows32s
    Case Else
        Get_WinVer = osUNKNOWN
    End Select

End Function
Public Function Lock_WorkStation() As Boolean
    Lock_WorkStation = LockWorkStationA
End Function
Public Function Get_SystemColor() As Variant
    Dim ColorPixel As Long, hDC As Long
    
    hDC = GetDC(0)
    
    ColorPixel = GetDeviceCaps(hDC, BITSPIXEL) * _
                 GetDeviceCaps(hDC, PLANES)
    
    Get_SystemColor = 2 ^ ColorPixel
    
End Function
Public Function IsAdmin() As Boolean
    IsAdmin = IsNTAdmin(ByVal 0&, ByVal 0&)
End Function
Public Function isWindowsNT() As Boolean
    Dim OSInfo As OSVERSIONINFO

    With OSInfo
        .dwOSVersionInfoSize = Len(OSInfo)
        GetVersionEx OSInfo
        isWindowsNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT)
    End With
  
End Function
Public Function is32BitWindows() As Boolean
    is32BitWindows = isWindowsNT
End Function
Public Function is16BitWindows() As Boolean
    is16BitWindows = Not isWindowsNT
End Function
Public Function System_Shutdown(Optional Reboot As Boolean, _
                                Optional force As Boolean = False) _
                                        As Boolean
    Dim Flags As Long, Token As Long, X As Long
    Dim TP As TOKEN_PRIVILEGES
    
    If GetVersion() >= 0 Then
        Call OpenProcessToken(GetCurrentProcess(), _
                              TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, Token)
        Call LookupPrivilegeValueA("", "SeShutdownPrivilege", TP.LuidUDT)
        
        TP.PrivilegeCount = 1
        TP.Attributes = SE_PRIVILEGE_ENABLED
    
        Call AdjustTokenPrivileges(Token, False, TP, 0, ByVal 0&, ByVal 0&)
    End If
  
    Flags = EWX_SHUTDOWN Or IIf(Reboot, EWX_REBOOT, EWX_SHUTDOWN)
    Flags = IIf(force, Flags Or EWX_FORCE, Flags)
    
    If ExitWindowsEx(Flags, &HFFFF) Then System_Shutdown = True
    
End Function
Public Function RegisterComponents(ByVal ComponentPath As String, _
                                   Optional Unregister As Boolean = False) _
                                        As Boolean
    Dim sProc As String
    Dim Lib As Long, r1 As Long, r2 As Long, Thread As Long

    On Local Error GoTo Quit

    Lib = LoadLibrary(ComponentPath)
  
    If Lib Then
        sProc = IIf(Unregister, "DllUnregisterServer", _
                                "DllRegisterServer")
        r1 = GetProcAddress(Lib, sProc)
        
        If r1 Then
            Thread = CreateThread(ByVal 0, 0, ByVal r1, _
                                  ByVal 0, 0, r2)
            If Thread Then
                r2 = WaitForSingleObject(Thread, 10000)
                
                If r2 Then _
                    Call FreeLibrary(Lib): _
                    r2 = GetExitCodeThread(Thread, r2): _
                    Call ExitThread(r2): _
                    Exit Function
            
                Call CloseHandle(Thread)
            End If
        End If
        Call FreeLibrary(Lib)
    End If
  
    RegisterComponents = True
  
Quit:
End Function
Public Function TaskIsActive(ByVal ExeName As String) As Boolean
    Dim E() As Variant, A() As Variant
        
    If Get_ProcessList(E, A, False) Then _
        If ARR.isArray(E) Then _
            TaskIsActive = (ARR.FindPhrase(E, ExeName, , , False, False) > -1)
            
End Function
Public Function Get_ProcessList(aEXE As Variant, _
                                aID As Variant, _
                                Optional GetFullPath As Boolean = False) _
                                    As Boolean
    Dim lSnapshot As Long, nResult As Long, X As Long
    Dim hProcess As Long, nSize As Long, lRet As Long
    Dim uProcess As PROCESSENTRY32
    Dim P As String, Str As String, ModuleName As String, GetProcessName As String
    
    On Local Error GoTo Quit
    
    X = -1
    
    lSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

    If isWindowsNT Then
        If lSnapshot Then
            uProcess.dwSize = Len(uProcess)
            nResult = ProcessFirst(lSnapshot, uProcess)
    
            Do Until nResult = 0
                Str = Mid$(uProcess.szexeFile, 1, _
                          InStr(uProcess.szexeFile, Chr$(0)) - 1)
            
                If GetFullPath Then
                    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                                           Or PROCESS_VM_READ, 0, uProcess.th32ProcessID)
                    ModuleName = Space(MAX_PATH)
                    nSize = 500
                
                    lRet = GetModuleFileNameExA(hProcess, 0, _
                                                ModuleName, nSize)
                
                    GetProcessName = Left(ModuleName, lRet)
                    lRet = CloseHandle(hProcess)
                    
                    If Len(GetProcessName) Then Str = GetProcessName
                End If
            
                Call ARR.Add(aEXE, Str, ADDTOEND)
                Call ARR.Add(aID, uProcess.th32ProcessID, ADDTOEND)
            
                nResult = ProcessNext(lSnapshot, uProcess)
            Loop
        End If
    Else
        If lSnapshot Then
            uProcess.dwSize = Len(uProcess)
            
            nResult = Process32First(lSnapshot, uProcess)
            
            Do While nResult
                Str = StrZToStr(uProcess.szexeFile)
            
                Call ARR.Add(aEXE, Str, ADDTOEND)
                Call ARR.Add(aID, uProcess.th32ProcessID, ADDTOEND)
            
                nResult = Process32Next(lSnapshot, uProcess)
            Loop
        End If
    End If
    
    Call CloseHandle(lSnapshot)

    Get_ProcessList = True
    
Quit:
    If Err Then Erase aEXE: _
                Erase aID
    
End Function
Private Function StrZToStr(S As String) As String
    StrZToStr = Left$(S, Len(S) - 1)
    StrZToStr = Replace(StrZToStr, Chr$(0), "")
End Function
Public Function SystemDirectory(SysInfo As SystemDirNFO) As String
    Dim temp As String * 256, Str As String
    Dim Result As Long, V As Long
    Dim M As Integer
    Dim WshShell As Object
    Dim idl As ITEMIDLIST
    
    M = 5
    
    Select Case SysInfo
        Case sdCommonProgramFilesDir
            M = 1: Str = "CommonProgramFiles"
        Case sdAllUsersProfileDir
            M = 1: Str = "allusersprofile"
        Case sdUserProfileDir
            M = 1: Str = "userprofile"
        Case sdWindowsDirectory, _
             sdWindowsSystemDir, _
             sdWindowsTempDir
            M = SysInfo - 3
        Case sdStartmenuDir: V = &HB
        Case sdDesktopDirectory: V = &H10
        Case sdProgramFilesDir: V = &H2
        Case sdPersonalDirectory: V = &H5
        Case sdFavoritDirectory: V = &H6
        Case sdAutoStartDirectory: V = &H7
        Case sdRecentDocumentDir: V = &H8
        Case sdSendToDirectory: V = &H9
        Case sdNewShellDir: V = &H15
        Case sdNetworkDirectory: V = &H13
        Case sdFontDirectory: V = &H14
        Case sdTemporaryInternetFilesDir: V = &H20
        Case sdUserDesktopDir: V = &H0
        Case sdUserOwnMusicDir: V = &HD
        Case sdUserOwnPicsDir: V = &H27
        Case sdUserCookiesDir: V = &H21
        Case sdAllUserDesktopDir: V = &H19
        Case sdUserAppDataDir, _
             sdUserQuickLaunchDir: M = 6
    End Select
    
    Select Case M
        Case 1
            SystemDirectory = Environ(Str)
        Case 2
            Result = GetWindowsDirectory(temp, Len(temp))
            SystemDirectory = Left$(temp, Result)
        Case 3
            Result = GetSystemDirectory(temp, Len(temp))
            SystemDirectory = Left$(temp, Result)
        Case 4
            Str = String(MAX_PATH, 0)
            Result = GetTempPath(MAX_PATH, Str)
            
            If Result Then _
                SystemDirectory = Left$(Str, InStr(Str, Chr$(0)) - 1)
        Case 5
            Result = SHGetSpecialFolderLocation(CLng(0), V, idl)
    
            If Result = NOERROR Then _
                Str = Space$(512): _
                Result = SHGetPathFromIDList(ByVal idl.mkid.CB, _
                                             ByVal Str): _
                If Result Then SystemDirectory = Trim$(Str)
        Case 6
            Set WshShell = CreateObject("WScript.Shell")
            
            SystemDirectory = WshShell.SpecialFolders.Item("AppData")
            
            If SysInfo = sdUserQuickLaunchDir Then _
                SystemDirectory = SystemDirectory & _
                                  "\Microsoft\Internet Explorer\Quick Launch\"
        End Select
        
        SystemDirectory = GetDir(Replace(SystemDirectory, Chr$(0), ""))

End Function
Public Function AutoRun(ByVal Field As String, _
                        ByVal EXEpath As String) _
                            As Boolean
        
    Const Key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    
    AutoRun = REG.RegValueSet(HKEY_LOCAL_MACHINE, _
                              Key, Field, EXEpath, True)
                            
End Function
Public Function Get_BootMode() As BootModeConstants
    
    Get_BootMode = bmUNKNOWN
    
    If CBool(GetSystemMetrics(SM_CLEANBOOT) = bmNormal) Then
        Get_BootMode = bmNormal
    Else
        If CBool(GetSystemMetrics(SM_CLEANBOOT) = bmSafeMode) Then
            Get_BootMode = bmSafeMode
        Else
            If CBool(GetSystemMetrics(SM_CLEANBOOT) = bmSafeModeNetwork) Then
                Get_BootMode = bmSafeModeNetwork
            End If
        End If
    End If
        
End Function
Public Function Get_UserLogInState() As USER_PRIV
    Dim nResult As Long, nBuffer As Long
    Dim UserInfo As USER_INFO
    Dim bServer() As Byte, bUser() As Byte
    
    If Not SYS.isWindowsNT Then _
        Get_UserLogInState = USER_PRIV_ADMIN: _
        Exit Function
        
    bUser = SYS.Get_LocalUserName & vbNullChar
    bServer = "" & vbNullChar
  
    nResult = NetUserGetInfo(bServer(0), bUser(0), &H1, nBuffer)
  
    If nResult = 0 Then
        CopyMemory UserInfo, ByVal nBuffer, Len(UserInfo)
        NetApiBufferFree nBuffer
    
        Get_UserLogInState = UserInfo.usri1_priv
    End If
  
End Function
Public Function Get_DLLType(DLLpath As String) As DLLType
    Dim nHandle As Long
    Dim sProcedure As String
  
    Get_DLLType = dtNODLL
    nHandle = LoadLibrary(DLLpath)
    
    If nHandle Then
        Get_DLLType = dtSTANDARDDLL
        
        If CBool(GetProcAddress(nHandle, "DllRegisterServer")) Then _
            Get_DLLType = dtACTIVEXDLL
    
        Call FreeLibrary(nHandle)
    End If

End Function
Public Function Uninstall(ByVal DestinationDir As String, _
                          Optional FileList As Variant, _
                          Optional RemoveDestinationDir As Boolean = False) _
                                As Boolean
    Dim Bat As String, T As String, H As String, Z As String
    Dim F() As String, D() As String, File As String, BatStr As String
    Dim Key As String, Field As String, Value As String
    Dim P As Long
                
    On Local Error GoTo Quit
    
    DestinationDir = GetDir(DestinationDir)
    Bat = DestinationDir & "Uninstall.bat"
    
    BatStr = "echo Set Fileattribut " & Bat & vbCrLf
    BatStr = BatStr & "Attrib -R -A -S -H " & Bat & vbCrLf
    BatStr = BatStr & "echo Delete " & Bat & vbCrLf
    BatStr = BatStr & "Del " & Bat & vbCrLf

    If RemoveDestinationDir Then
        If Not DI.ReadDirectory(DestinationDir, D, F, True, True, True) Then Exit Function
        
        ReDim FileList(UBound(F))
        
        For P = LBound(F) To UBound(F)
            FileList(P) = F(P)
        Next P
    End If
    
    T$ = ""
    
    If Not ARR.isArray(FileList) Then
        File = DestinationDir & FileList
        
        T = "echo Set Fileattribut " & File & vbCrLf
        T = T & "Attrib -R -A -S -H " & File & vbCrLf
        T = "echo Delete " & File & vbCrLf
        T = T & "Del " & File & vbCrLf
        
        T = T & BatStr
    Else
        For P = LBound(FileList) To UBound(FileList)
            File = DestinationDir & FileList(P)
            
            T = T & "echo Set Fileattribut " & File & vbCrLf
            T = T & "Attrib -R -A -S -H " & File & vbCrLf
            T = T & "echo Delete " & File & vbCrLf
            T = T & "Del " & File & vbCrLf
        Next P
        
        T = T & BatStr
        
        If RemoveDestinationDir Then _
            DestinationDir = Left$(DestinationDir, Len(DestinationDir) - 1): _
            T = T & "Delete Directory " & DestinationDir & vbCrLf: _
            T = T & "RmDir " & DestinationDir & vbCrLf
    End If
    
    H = "@echo off" & vbCrLf & "cls" & vbCrLf
    
    H = H & "echo Uninstall created by MfDLL" & vbCrLf & vbCrLf
    H = H & "Created on " & Date & " / " & Time & _
            vbCrLf & _
            vbCrLf & _
            vbCrLf
    H = H & "echo Uninstall in progress....." & vbCrLf & vbCrLf
    
    H = H & T
    H = H & "echo." & vbCrLf & _
            "echo." & vbCrLf & _
            "echo." & vbCrLf & _
            "echo Uninstall done." & vbCrLf & _
            "echo Please close this window" & _
            vbCrLf & vbCrLf & _
            "Exit"
            
                   
    If Not TXT.TextFile_WriteText(Bat, H, True) Then Exit Function
    
    Z = Str.Get_RandomString(8, True, False, True, False, False)
    
    Key = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    Field = "MfDLL Uninstall (" & Z & ")"
    Value = Bat
            
    If Not REG.RegValueSet(HKEY_LOCAL_MACHINE_, _
                           Key, Field, Value, True) Then Exit Function
     
    Uninstall = True
    
Quit:
    Erase F
    Erase D
    
End Function
Public Function XPstyleSupported() As Boolean
    Dim hLib As Long

    On Error Resume Next
    
    hLib = LoadLibrary("uxtheme.dll")
    If hLib Then FreeLibrary hLib
    XPstyleSupported = Not (hLib = 0)

End Function
Public Function Font_Register(ByVal FontPath As String, _
                              Optional Unregister As Boolean = False) _
                                    As Boolean
    
    If Not Unregister Then
        Font_Register = (AddFontResource(FontPath) <> 0)
    Else
        Font_Register = (RemoveFontResource(FontPath) <> 0)
    End If
    
End Function
Public Function Get_ServicePack() As String
    
    Const Path As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

    If isWindowsNT Then _
        Get_ServicePack = REG.RegValueGet(HKEY_LOCAL_MACHINE_, _
                                          Path, "CSDVersion")
        
End Function
Public Function IsClassRegistered(ByVal Class As String) _
                                    As Boolean
    
    IsClassRegistered = REG.RegKeyExist(HKEY_CLASSES_ROOT_, Class)

End Function
Public Function Get_DLLpath(ByVal Class As String) _
                                    As String
    Dim strCLSID As String, strFilename As String
    
    If IsClassRegistered(Class) Then
        strCLSID = REG.RegValueGet(HKEY_CLASSES_ROOT_, _
                                   Class & "\CLSID", "")
        strFilename = REG.RegValueGet(HKEY_CLASSES_ROOT_, _
                                      "CLSID\" & strCLSID & _
                                      "\InprocServer32", "")
    
        Get_DLLpath = strFilename
    End If
    
End Function
Public Function CPUload() As Long
    Dim V As Long, pdhStatus As Long
    Dim dblValue As Double

    Const KDyn& = &H80000006

    Static Count As Long, hK As Long
    Static isNT As Boolean
    Static sK As String
    
    On Local Error GoTo Quit
    
    If Count = 0 Then
        isNT = isWindowsNT
        
        m_StackSize = 5
        ReDim Stack(0 To m_StackSize - 1)
        StackPointer = 0
       
        If isNT Then
            Retval = PdhOpenQuery(0, 1, hQuery)
        
            If Retval = 0 Then
                Retval = PdhVbAddCounter(hQuery, "\Prozessor(0)\Prozessorzeit (%)", hCounter)
                If Retval <> 0 Then PdhCloseQuery hQuery
            End If
        End If

        m_StackSize = 10
        ReDim Stack(0 To m_StackSize - 1)
        StackPointer = 0
    
        Count = 1
    End If

    If isNT Then
        PdhCollectQueryData hQuery
        dblValue = PdhVbGetDoubleCounterValue(hCounter, pdhStatus)
        
        If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or _
           (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
                PushBandWidth CLng(dblValue)
        End If
    Else
        'Muß noch unter 98 getestet werden !!!
        sK = IIf(hK = 0, "PerfStats\StartStat", "PerfStats\StatData")
        
        If Not REG.RegKeyExist(KDyn, sK) Then Exit Function

        V = REG.RegValueGet(KDyn, sK & "\KERNEL", "CPUUsage")
               
        PushBandWidth V
    End If

    CPUload = m_BandWidth

Quit:
    If Err.Number <> 0 Then CPUload = -1
    
End Function
Private Sub PushBandWidth(ByRef NewBandWidth As Long)
    
    Static U As Long
    
    U = UBound(Stack)
  
    If StackPointer <= U Then
        Stack(StackPointer) = NewBandWidth
        m_Sum = m_Sum + NewBandWidth
        StackPointer = StackPointer + 1
    Else
        m_Sum = m_Sum - Stack(0) + NewBandWidth
        Call CopyMemory(Stack(0), Stack(1), U * 4)
        Stack(U) = NewBandWidth
    End If
    
    m_BandWidth = m_Sum / StackPointer

End Sub
Public Sub TempDirectotyClear()
    Dim D As String
    Dim A() As Variant
    Dim DI As New Directory
    
    D = SystemDirectory(sdWindowsTempDir)
    
    Call DI.DirectoryDeleteFiles(D, A, True)
    Call DI.DirectoryDeleteSubDirs(D, A, True)
    
End Sub
Public Function Get_ComputerName() As String
    Get_ComputerName = NET.Get_LocalNetworkName
End Function
Public Function Set_ComputerName(ByVal NewName As String) As Boolean
    Set_ComputerName = NET.Set_LocalNetworkName(NewName)
End Function
Public Sub Get_RegOwner(User As String, _
                        Organization As String)
    Dim Path As String
    
    Path = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    
    User = REG.RegValueGet(HKEY_LOCAL_MACHINE_, Path, "RegisteredOwner")
    Organization = REG.RegValueGet(HKEY_LOCAL_MACHINE_, Path, "RegisteredOrganization")

End Sub
Public Function User_LogOff() As Boolean
    If ExitWindowsEx(EWX_LOGOFF, &HFFFF) Then User_LogOff = True
End Function
Public Function Security(sOption As SecurityOption, _
                         ByVal Enable As Boolean) _
                            As Boolean
    Dim Root As Long, Value As Long
    Dim Key As String, Field As String

    On Local Error GoTo Quit

    Root = Declar.HKEY_CURRENT_USER
    Key = "Software\Microsoft\Windows\CurrentVersion\Policies\System"

    Value = IIf(Enable, 0, 1)

    Select Case sOption
        Case TASKMANAGER
            If isWindowsNT Then
                Field = "DisableTaskMgr"
            Else
                If SystemParametersInfo(97, CBool(Value), "1", 0) Then _
                    Security = True: _
                    Exit Function
            End If
        Case REGEDIT
            Field = "DisableRegistryTools"
        Case LOGOFF
            Field = "NoLogOff"
        Case WINDOWSUPDATE
            Field = "DisableWindowsUpdateAccess"
        Case COMMANDCOM
            Key = "Software\Policies\Microsoft\Windows\System"
            If Value = 1 Then Value = 2
            Field = "DisableCMD"
        Case SAVEDESKTOP
            Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
            Field = "NoSaveSettings"
        Case CONTROLPANEL
            Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
            Field = "NoControlPanel"
        Case NETWORKICON
            Key = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
            Field = "NoNetHood"
    End Select
    
    Security = REG.RegValueSet(Root, Key, Field, Value, True)

Quit:
End Function
