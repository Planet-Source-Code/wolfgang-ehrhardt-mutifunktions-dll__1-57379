VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Network"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function InetIsOffline Lib "url.dll" (ByVal dwFlags As Long) As Long
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpSFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WinSocketDataType) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal HostName$, ByVal HostLen%) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal HostName$) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (ByVal addr$, ByVal Laenge%, ByVal Typ%) As Long
Private Declare Function IsValidURLA Lib "URLMON.DLL" Alias "IsValidURL" (ByVal pbc As Long, ByVal szURL As String, ByVal dwReserved As Long) As Long
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, ByVal lpszConiID As String, ByVal dwFlags As Long, ByRef hCon As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Declare Function SHAutoComplete Lib "shlwapi" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function InternetAttemptConnect Lib "wininet" (ByVal dwReserved As Long) As Long
Private Declare Function NetRemoteTOD Lib "netapi32.dll" (UncServerName As Any, lpBuffer As Long) As Long
Private Declare Function GetIfTable Lib "iphlpapi.dll" (ByRef pIfTable As Any, ByRef pdwSize As Long, ByVal border As Long) As Long
Private Declare Function GetIpNetTable Lib "Iphlpapi" (pIpNetTable As Byte, pdwSize As Long, ByVal border As Long) As Long
Private Declare Function GetIpForwardTable Lib "iphlpapi.dll" (ByRef pIpForwardTable As Any, ByRef pdwSize As Long, ByVal border As Long) As Long
Private Declare Function NetShareEnum Lib "netapi32.dll" (ByVal servername As Long, ByVal Level As Long, ByVal bufptr As Long, ByVal prefmaxlen As Long, ByRef entriesread As Long, ByRef totalentries As Long, ByRef resume_handle As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal strptr As Long) As Long
Private Declare Function RasEnumEntries Lib "RasApi32.DLL" Alias "RasEnumEntriesA" (ByVal Reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal length As Long)
Private Declare Sub CopyMemoryA Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Const RAS95_MaxEntryName As Long = 256

Private Type RASENTRYNAME95
    dwSize                          As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type

Private Type tShare
    Name         As String
    Typ          As Long
    Beschreibung As String
End Type

Private Type NETRESOURCE
    dwScope       As Long
    dwType        As Long
    dwDisplayType As Long
    dwUsage       As Long
    pLocalName    As Long
    pRemoteName   As Long
    pComment      As Long
    pProvider     As Long
End Type

Private Type NETRESOURCE_REAL
    dwScope       As Long
    dwType        As Long
    dwDisplayType As Long
    dwUsage       As Long
    sLocalName    As String
    sRemoteName   As String
    sComment      As String
    sProvider     As String
End Type

Private Type HostDeType
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Private Const MAXLEN_PHYSADDR = 8

Private Type MIB_IPNETROW
    dwIndex                             As Long
    dwPhysAddrLen                       As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAddr                              As Long
    dwType                              As Long
End Type

Private Type PMIB_IPFORWARDROW
    dwForwardDest      As Long
    dwForwardMask      As Long
    dwForwardPolicy    As Long
    dwForwardNextHop   As Long
    dwForwardIfIndex   As Long
    dwForwardType      As Long
    dwForwardProto     As Long
    dwForwardAge       As Long
    dwForwardNextHopAS As Long
    dwForwardMetric1   As Long
    dwForwardMetric2   As Long
    dwForwardMetric3   As Long
    dwForwardMetric4   As Long
    dwForwardMetric5   As Long
End Type

Private Type PMIB_IPFORWARDTABLE
    dwNumEntries As Long
    table        As PMIB_IPFORWARDROW
End Type

Public Enum BroadcastReturnCode
    bcINVALIDMESSAGE = -1
    bcINVALIDFILE = -2
    bcSEND = 0
    bcSENDCRC = 1
End Enum

Public Enum icDevice
    icLAN = 0
    icMODEM = 1
    icPROXY = 2
    icNONE = 3
End Enum

Public Enum DefaultBrowser
    dbINTERNETEXPLORER = 0
    dbNETSCAPE = 1
    dbOPERA = 2
    dbUNKNOWN = 3
End Enum

Private Type TIME_OF_DAY_INFO
    tod_elapsedt  As Long
    tod_msecs     As Long
    tod_hours     As Long
    tod_mins      As Long
    tod_secs      As Long
    tod_hunds     As Long
    tod_timezone  As Long
    tod_tinterval As Long
    tod_day       As Long
    tod_month     As Long
    tod_year      As Long
    tod_weekday   As Long
End Type

Private Const MAX_INTERFACE_NAME_LEN = 256
Private Const MAXLEN_IFDESCR = 256

Private Type MIB_IFROW
    wszName(0 To (MAX_INTERFACE_NAME_LEN - 1) * 2) As Byte
    dwIndex              As Long
    dwType               As Long
    dwMtu                As Long
    dwSpeed              As Long
    dwPhysAddrLen        As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAdminStatus        As Long
    dwOperStatus         As Long
    dwLastChange         As Long
    dwInOctets           As Long
    dwInUcastPkts        As Long
    dwInNUcastPkts       As Long
    dwInDiscards         As Long
    dwInErrors           As Long
    dwInUnknownProtos    As Long
    dwOutOctets          As Long
    dwOutUcastPkts       As Long
    dwOutNUcastPkts      As Long
    dwOutDiscards        As Long
    dwOutErrors          As Long
    dwOutQLen            As Long
    dwDescrLen           As Long
    bDescr(0 To MAXLEN_IFDESCR - 1) As Byte
End Type

Private Const MIN_SOCKETS_REQD = 1

Private Const S_FALSE = &H1
Private Const S_OK = &H0
Private Const SOCKET_ERROR = -1

Private Const WS_VERSION_REQD = &H101
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type WinSocketDataType
    wversion                               As Integer
    wHighVersion                           As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets                            As Integer
    iMaxUdpDg                              As Integer
    lpszVendorInfo                         As Long
End Type

Private Const FLAG_ICC_FORCE_CONNECTION = &H1

Private Const GENERIC_WRITE = &H40000000

Private Const ERROR_MORE_DATA = 234&

Private Const FILE_SHARE_READ = &H1
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Const INTERNET_CONNECTION_LAN = &H2
Private Const INTERNET_CONNECTION_MODEM = &H1
Private Const INTERNET_CONNECTION_PROXY = &H4
Private Const INTERNET_CONNECTION_MODEM_BUSY = &H8

Private Const INTERNET_DIAL_UNATTENDED = &H8000
Private Const INTERNET_DIAL_FORCE_PROMPT = &H2000
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Const MAX_COMPUTERNAME_LENGTH = 15

Private Const MIB_IF_TYPE_OTHER = 1
Private Const MIB_IF_TYPE_ETHERNET = 6
Private Const MIB_IF_TYPE_TOKENRING = 9
Private Const MIB_IF_TYPE_FDDI = 15
Private Const MIB_IF_TYPE_PPP = 23
Private Const MIB_IF_TYPE_LOOPBACK = 24
Private Const MIB_IF_TYPE_SLIP = 28

Private Const NO_ERROR = 0&

Private Const OPEN_EXISTING = 3

Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF&

Private Const RESDTYPE_DIRECTORY As Long = &H9&
Private Const RESDTYPE_DOMAIN As Long = &H1&
Private Const RESDTYPE_FILE As Long = &H4&
Private Const RESDTYPE_GENERIC As Long = &H0&
Private Const RESDTYPE_GROUP As Long = &H5&
Private Const RESDTYPE_NETWORK As Long = &H6&
Private Const RESDTYPE_ROOT As Long = &H7&
Private Const RESDTYPE_SERVER As Long = &H2&
Private Const RESDTYPE_SHARE As Long = &H3&
Private Const RESDTYPE_SHAREADMIN As Long = &H8&

Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&

Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000

Private Const NCBNAMSZ = 16
    
Private Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major          As Byte
    reserved0          As Byte
    adapter_type       As Byte
    rev_minor          As Byte
    duration           As Integer
    frmr_recv          As Integer
    frmr_xmit          As Integer
    iframe_recv_err    As Integer
    xmit_aborts        As Integer
    xmit_success       As Long
    recv_success       As Long
    iframe_xmit_err    As Integer
    recv_buff_unavail  As Integer
    t1_timeouts        As Integer
    ti_timeouts        As Integer
    Reserved1          As Long
    free_ncbs          As Integer
    max_cfg_ncbs       As Integer
    max_ncbs           As Integer
    xmit_buf_unavail   As Integer
    max_dgram_size     As Integer
    pending_sess       As Integer
    max_cfg_sess       As Integer
    max_sess           As Integer
    max_sess_pkt_size  As Integer
    name_count         As Integer
End Type

Private Type NAME_BUFFER
    Name       As String * NCBNAMSZ
    name_num   As Integer
    name_flags As Integer
End Type
Public Sub OpenURL(ByVal URL As String)
    Call ShellExecute(0, "Open", URL, "", "", 1)
End Sub
Public Sub SendNewMail(ByVal MailTo As String, _
                       ByVal Subject As String, _
                       ByVal Body As String)
    Dim Buff As String
    
    Buff = "mailto:" & MailTo & "?Subject=" & _
           Subject & "&Body=" & Body
    
    Call ShellExecute(0&, "Open", Buff, "", "", 1)

End Sub
Public Function Get_DefaultBrowser() As DefaultBrowser
    Dim File As String, BrExe As String, Browser As String
    Dim Result As Long, X As Long
    Dim FN As Integer
  
    On Local Error Resume Next
    
    File = "~temp.html"
    FN = FreeFile
    
    Open App.Path & "\" & File For Output As #FN
    Close FN
    
    BrExe = String(256, Chr$(0))
    Result = FindExecutable(File, App.Path, BrExe)
    BrExe = Left$(BrExe, InStr(1, BrExe, Chr$(0)) - 1)
    
    For X = Len(BrExe) To 1 Step -1
        If Mid$(BrExe, X, 1) = "\" Then Exit For
    Next X
    
    BrExe = Mid$(BrExe, X + 1)
    
    Select Case BrExe
        Case "iexplore.exe"
            Get_DefaultBrowser = dbINTERNETEXPLORER
        Case "netscape.exe"
            Get_DefaultBrowser = dbNETSCAPE
        Case "opera.exe"
            Get_DefaultBrowser = dbOPERA
        Case Else
            Get_DefaultBrowser = dbUNKNOWN
    End Select
    
    Kill App.Path & "\" & File

End Function
Public Function isDFÜconnected() As Boolean
    isDFÜconnected = (InternetConnectedBy = icMODEM)
End Function
Public Function Set_LocalNetworkName(Name As String) As Boolean
    Set_LocalNetworkName = SetComputerName(Name)
End Function
Public Function Get_LocalNetworkName() As String
    Dim Buffer As String, CName As String
    Dim Result As Long, L As Long

    L = MAX_COMPUTERNAME_LENGTH + 1
    
    Buffer = Space$(L)
    Result = GetComputerName(Buffer, L)
    
    If Result = 1 Then _
        CName = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1): _
        Get_LocalNetworkName = Buffer: _
        Get_LocalNetworkName = Left$(Get_LocalNetworkName, L)
    
End Function
Public Function InternetConnectedBy() As icDevice
    Dim icFlags As Long
    
    Call InternetGetConnectedState(icFlags, 0&)
    
    If icFlags And INTERNET_CONNECTION_LAN Then
        InternetConnectedBy = icLAN
    Else
        If icFlags And INTERNET_CONNECTION_MODEM _
        Or icFlags And INTERNET_CONNECTION_MODEM_BUSY Then
            InternetConnectedBy = icMODEM
        Else
            If icFlags And INTERNET_CONNECTION_PROXY Then
                InternetConnectedBy = icPROXY
            Else
                InternetConnectedBy = icNONE
            End If
        End If
    End If

End Function
Public Function Get_HTMLfile(URL As String) As String
    Dim Buffer As String
    Dim L As Long, hOpen As Long, hFile As Long, Result As Long
    
    On Local Error GoTo Quit
    
    L = 50000
    Buffer = Space(L)
    
    hOpen = InternetOpen("Wininet Test", INTERNET_OPEN_TYPE_DIRECT, _
                         vbNullString, vbNullString, 0)
  
    hFile = InternetOpenUrl(hOpen, URL, vbNullString, _
                            ByVal 0&, INTERNET_FLAG_RELOAD, _
                            ByVal 0&)
                            
    Call InternetReadFile(hFile, Buffer, L, Result)
    
    Get_HTMLfile = Left$(Buffer, Result)
    
Quit:
    Call InternetCloseHandle(hFile)
    Call InternetCloseHandle(hOpen)
    
    If Err Then Get_HTMLfile = ""
    
End Function
Public Function TAPIdialNumber() As Boolean
    TAPIdialNumber = (tapiRequestMakeCall&("", "", "", "") = 0)
End Function
Public Function IE_GetStartSite() As String
    Dim Key As String
        
    Key = "Software\Microsoft\Internet Explorer\Main"
    
    IE_GetStartSite = REG.RegValueGet(HKEY_CURRENT_USER_, _
                                      Key, "Start Page")

End Function
Public Function IE_SetStartSite(ByVal URL As String, _
                                ByVal Title As String) _
                                    As Boolean
    Dim Key As String
        
    Key = "Software\Microsoft\Internet Explorer\Main"
    
    If REG.RegValueSet(HKEY_CURRENT_USER_, _
                       Key, "Start Page", URL, True) Then
        If REG.RegValueSet(HKEY_CURRENT_USER_, _
                           Key, "Window Title", Title) Then _
                                IE_SetStartSite = True
    End If
    
End Function
Public Function Create_InternetShortcut(ByVal URL As String, _
                                        ByVal Linkfile As String) _
                                            As Boolean
    Dim Link As String
    
    Link = "[InternetShortcut]" & vbCrLf & "URL = " & URL
    
    Create_InternetShortcut = TXT.TextFile_WriteText(Linkfile, _
                                                     Link, True)
    
End Function
Public Function isPortInUse(ByVal PortNumber As Integer) _
                                    As Boolean
    Dim TempWinsock As Object
    Dim Flag As Boolean
  
    On Error Resume Next

    Set TempWinsock = CreateObject("MSWinsock.Winsock.1")
    
    TempWinsock.LocalPort = PortNumber
    TempWinsock.Listen

    Flag = Err.Number = 10048
    TempWinsock.Close
    Set TempWinsock = Nothing
    
    isPortInUse = Flag

End Function
Public Function Get_LocalIP() As String
    Dim Host As String
    Dim X As Integer

    Call InitSocketAPI
    Host = MyHostName
         
    Get_LocalIP = HostByName(Host, X)
       
    Call CleanSockets

End Function
Private Sub InitSocketAPI()
    Dim Result As Integer
    Dim LoBy As Integer, HiBy As Integer
    Dim SocketData As WinSocketDataType
  
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
  
End Sub
Private Function MyHostName() As String
    Dim HostName As String * 256
  
    MyHostName = NextChar(Trim$(HostName), Chr$(0))
End Function
Private Function HostByName(Name As String, _
                            Optional X As Integer = 0) _
                                    As String
    Dim MemIp() As Byte
    Dim Y As Integer
    Dim HostDeAddress As Long, HostIp As Long
    Dim IpAddress As String
    Dim Host As HostDeType
  
    HostDeAddress = gethostbyname(Name)
    
    If HostDeAddress = 0 Then _
        HostByName = "": _
        Exit Function
    
    Call RtlMoveMemory(Host, HostDeAddress, LenB(Host))
    
    For Y = 0 To X
        Call RtlMoveMemory(HostIp, Host.hAddrList + 4 * Y, 4)
      
        If HostIp = 0 Then _
            HostByName = "": _
            Exit Function
    Next Y
    
    ReDim MemIp(1 To Host.hLength)
    
    Call RtlMoveMemory(MemIp(1), HostIp, Host.hLength)
    
    IpAddress = ""
    
    For Y = 1 To Host.hLength
        IpAddress = IpAddress & MemIp(Y) & "."
    Next Y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    
    HostByName = IpAddress

End Function
Private Sub CleanSockets()
    Call WSACleanup
End Sub
Private Function NextChar(Text As String, _
                          Char As String) As String
    Dim Pos As Integer
    
    Pos = InStr(1, Text, Char)
    
    If Pos = 0 Then
        NextChar = Text
        Text = ""
    Else
        NextChar = Left$(Text, Pos - 1)
        Text = Mid$(Text, Pos + Len(Char))
    End If
    
End Function
Private Function toHex(ByVal Number As Long, ByVal NumLen As Integer, _
                       Optional ByVal UpperCase As Boolean) As String
    Dim C As String
    
    C = Hex(Number)

    If Len(C) < NumLen Then _
        C = String(NumLen - Len(C), "0") & C
    
    toHex = IIf(UpperCase, C, LCase$(C))
    
End Function
Public Function URLToHTML(ByVal URL As String) As String
    Dim Char As String, R$
    Dim P As Long, A As Long
    
    Const Str As String = Num & lABC & uABC
    
    For P = 1 To Len(URL)
        Char = Mid$(URL, P, 1)
        
        If InStr(Str, Char) <> 0 Then
            R$ = R$ & Char
        Else
            If Char = " " Then
                R$ = R$ & "+"
            Else
                A = Asc(Char)
                R$ = R$ & "%" & toHex(A, 2)
            End If
        End If
    Next P
    
    URLToHTML = R$
    
End Function
Public Function HTMLtoURL(ByVal HTML As String) As String
    Dim Char As String, R$
    Dim P As Long
    
    Const Str As String = Num & lABC & uABC
    
    For P = 1 To Len(HTML)
        Char = Mid$(HTML, P, 1)
        
        If Char = "+" Then
            R$ = R$ & " "
        Else
            If Char = "%" Then
                R$ = R$ & Chr$(Val("&H" & Mid$(HTML, P + 1, 1) & Mid$(HTML, P + 2, 1)))
                P = P + 2
            Else
                R$ = R$ & Char
            End If
        End If
    Next P
    
    HTMLtoURL = R$
    
End Function
Public Function IsValidURL(ByVal URL As String) As Boolean
    
    If Not CBool(IsValidURLA(ByVal 0&, _
                             StrConv(URL, vbUnicode), 0)) Then _
                                IsValidURL = True
    
End Function
Public Sub InternetDialUP(hWnd As Long, _
                          Optional Connection As String)
    
    Call InternetDial(hWnd, Connection, 0, 0, 0)
    
End Sub
Public Sub InternetHangup()
    Call InternetAutodialHangup(CLng(0))
End Sub
Public Function IE_ShowAutoComplete(TextBox As Object) As Boolean
    On Local Error Resume Next
    IE_ShowAutoComplete = (SHAutoComplete(TextBox.hWnd, CLng(0)) = 0)
End Function
Public Function IE_GetVersion() As String
    Dim Key As String
        
    Key = "Software\Microsoft\Internet Explorer\Version Vector"
    
    IE_GetVersion = REG.RegValueGet(HKEY_LOCAL_MACHINE_, Key, "IE")
    
End Function
Public Function Hyperlink_Create(Label As Object, _
                                 ByVal LabelColor As Long, _
                                 ByVal LabelColorMouseOver As Long) _
                                    As Boolean
    Dim frm As New MyControls
    Dim X As Long
    
    X = -1
    
    Do While X < 0
        X = FRMS.Find_Form(Forms, "MyControls", , "Hyperlink")
        
        If X = -1 Then _
            Load frm: _
            frm.Tag = "Hyperlink"
    Loop
    
    Call Forms(X).Create_Hyperlink(Label.Parent, Label, _
                                   LabelColor, _
                                   LabelColorMouseOver)

End Function
Public Function Adapter_GetMACadress(ByVal AdapterIndex As Long) As String
    Dim A() As String
    
    A = NET.Adapter_GetInfo(AdapterIndex)
    
    If ARR.isArray(A) Then Adapter_GetMACadress = A(0)
        
End Function
Public Function isConnectedToInternet() As Boolean
    isConnectedToInternet = (InternetConnectedBy <> icNONE)
End Function
Public Function BroadcastMessage(ByVal msgFrom As String, _
                                 ByVal SendTo As String, _
                                 ByVal Message As String) _
                                    As BroadcastReturnCode
    Dim hFile As Long, byteswritten As Long
    Dim Buff As String, sSlotName As String
   
    If Len(Message) = 0 Then BroadcastMessage = bcINVALIDMESSAGE: _
                             Exit Function
                             
    If Len(SendTo) = 0 Then SendTo = "*"
    If Len(msgFrom) = 0 Then msgFrom = "Mailslot System Message"
    
    sSlotName = "\\" & SendTo & "\mailslot\messngr"
    
    Buff = msgFrom & vbNullChar & _
           SendTo & vbNullChar & _
           Message & vbNullChar & vbNullChar
    
    hFile = CreateFile(sSlotName, GENERIC_WRITE, _
                       FILE_SHARE_READ, 0&, OPEN_EXISTING, _
                       FILE_ATTRIBUTE_NORMAL, 0&)
    
    If hFile Then
        If WriteFile(hFile, Buff, Len(Buff), _
                     byteswritten, 0) Then
            If (Len(Buff) = byteswritten) Then
                BroadcastMessage = bcSEND
            Else
                BroadcastMessage = bcSENDCRC
            End If
        Else
            BroadcastMessage = bcINVALIDFILE
        End If
        
        Call CloseHandle(hFile)
    End If
    
End Function
Public Function INetFile_Read(ByVal URL As String) As String
    Dim Path As String
    
    On Local Error Resume Next
    
    Path = F.TempFile_Create("C:\")
    
    If INetFile_Save(URL, Path, False) Then _
        INetFile_Read = TXT.TextFile_Read(Path): _
        Kill Path
        
End Function
Public Function INetFile_Save(ByVal URL As String, _
                              Optional ByVal LocalFilename As String, _
                              Optional ShowInternetExplorer As Boolean = False) _
                                    As Boolean
    If ShowInternetExplorer Then
        Call DoFileDownload(StrConv(URL, vbUnicode))
        INetFile_Save = True
    Else
        If URLDownloadToFile(0, URL, LocalFilename, 0, 0) = 0 Then _
            INetFile_Save = True
    End If
    
End Function
Public Function DOS_Netstat() As String
    DOS_Netstat = SH.DOS_Execute("Netstat /a")
    DOS_Netstat = Replace(DOS_Netstat, Chr$(13), "")
End Function
Public Function DOS_IPconfig() As String
    DOS_IPconfig = SH.DOS_Execute("ipconfig /all")
    DOS_IPconfig = Replace(DOS_IPconfig, Chr$(13), "")
End Function
Public Function IPtoName(ByVal IP As String) As String
    Dim aa As String
    
    If Not isConnectedToInternet Then Exit Function
    
    Call InitSocketAPI
    aa = HostByAddress(IP)
    Call CleanSockets
    
    IPtoName = aa
    
End Function
Public Function NameToIP(ByVal Name As String) As String
    Dim aa As String
    
    If isConnectedToInternet Then
        Call InitSocketAPI
        aa = HostByName(Name)
        Call CleanSockets
    
        NameToIP = aa
    End If
    
End Function
Public Function Get_HostName() As String
    Dim HostName As String * 256
    
    Call InitSocketAPI
    
    If gethostname(HostName, 256) <> SOCKET_ERROR Then
        Get_HostName = NextChar(Trim$(HostName), Chr$(0))
        Call CleanSockets
    End If
    
End Function
Private Function Dynamic() As Variant
    Dim IP As String, DNS As String, Host As String
    Dim A(1) As String
    
    If Not isConnectedToInternet Then Exit Function
    
    Call InitSocketAPI
    Host = Get_HostName
         
    A(0) = HostByName(Host, 0)
    A(1) = HostByAddress(A(0))
    Dynamic = A
    
    Call CleanSockets
    
    Erase A
    
End Function
Public Function Get_DynamicDNS() As String
    Dim A() As String
    
    A = Dynamic
    Get_DynamicDNS = A(1)
    
    Erase A
    
End Function
Public Function Get_DynamicIP() As String
    Dim A() As String
    
    A = Dynamic
    Get_DynamicIP = A(0)
    
    Erase A
    
End Function
Public Function IE_AddFavorite(ByVal SiteName As String, _
                               ByVal URL As String) As Boolean
    Dim Path As String
        
    On Local Error GoTo Quit

    Path = SYS.SystemDirectory(sdFavoritDirectory)
    Path = Path & SiteName & ".Url"
    
    IE_AddFavorite = Create_InternetShortcut(URL, Path)

Quit:
End Function
Private Function Netsuche() As Variant

    Const MAX_RESOURCES = 2048
    Const NOT_A_CONTAINER = -1
    
    Dim bFirstTime As Boolean, Läufer As Long, lRet As Long, hEnum As Long
    Dim lCnt As Long, lMin As Long, lLen As Long, lBufSize As Long
    Dim lLastIx As Long, L As Long, NetAusgabe As String, X As Long
    Dim R$, rPC() As String

    Dim uNetApi(0 To MAX_RESOURCES) As NETRESOURCE
    Dim uNet() As NETRESOURCE_REAL
    
    bFirstTime = True
    
    X = 0
    
    Do
        If bFirstTime Then
            lRet = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                                RESOURCEUSAGE_ALL, ByVal 0&, hEnum)
            bFirstTime = False
        Else
            If uNet(lLastIx).dwUsage And RESOURCEUSAGE_CONTAINER Then
                lRet = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                                RESOURCEUSAGE_ALL, uNet(lLastIx), hEnum)
        Else
            lRet = NOT_A_CONTAINER
            hEnum = 0
        End If
        
            lLastIx = lLastIx + 1
        End If
        
        If lRet = NO_ERROR Then
            lCnt = RESOURCE_ENUM_ALL
            
            Do
                lBufSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
                lRet = WNetEnumResource(hEnum, lCnt, uNetApi(0), lBufSize)
                
                If lCnt > 0 Then
                
                    ReDim Preserve uNet(0 To lMin + lCnt - 1) _
                                   As NETRESOURCE_REAL
                    
                    For L = 0 To lCnt - 1
                        uNet(lMin + L).dwScope = uNetApi(L).dwScope
                        uNet(lMin + L).dwType = uNetApi(L).dwType
                        uNet(lMin + L).dwDisplayType = uNetApi(L).dwDisplayType
                        uNet(lMin + L).dwUsage = uNetApi(L).dwUsage
                        
                        If uNetApi(L).pLocalName Then
                            lLen = lstrlen(uNetApi(L).pLocalName)
                            uNet(lMin + L).sLocalName = Space$(lLen)
                            CopyMemory ByVal uNet(lMin + L).sLocalName, _
                                ByVal uNetApi(L).pLocalName, lLen
                        End If
                        
                        If uNetApi(L).pRemoteName Then
                            lLen = lstrlen(uNetApi(L).pRemoteName)
                            uNet(lMin + L).sRemoteName = Space$(lLen)
                            
                            CopyMemory ByVal uNet(lMin + L).sRemoteName, _
                                ByVal uNetApi(L).pRemoteName, lLen
                        End If
                        
                        If uNetApi(L).pComment Then
                            lLen = lstrlen(uNetApi(L).pComment)
                            uNet(lMin + L).sComment = Space$(lLen)
                            
                            CopyMemory ByVal uNet(lMin + L).sComment, _
                                ByVal uNetApi(L).pComment, lLen
                        End If
                        
                        If uNetApi(L).pProvider Then
                            lLen = lstrlen(uNetApi(L).pProvider)
                            uNet(lMin + L).sProvider = Space$(lLen)
                            
                            CopyMemory ByVal uNet(lMin + L).sProvider, _
                                ByVal uNetApi(L).pProvider, lLen
                        End If
                        DoEvents
                    Next L
                End If
            
                lMin = lMin + lCnt
                DoEvents
            Loop While lRet = ERROR_MORE_DATA
        End If
        If hEnum Then L = WNetCloseEnum(hEnum)
        DoEvents
    Loop While lLastIx < lMin
    
    If UBound(uNet) + 1 = 0 Then Exit Function
    
    For Läufer = 0 To UBound(uNet)
        If uNet(Läufer).dwDisplayType = RESDTYPE_SERVER Then _
              R$ = LCase(Replace(uNet(Läufer).sRemoteName, "\", "")): _
              ReDim Preserve rPC(X): _
              rPC(X) = LCase(Replace(uNet(Läufer).sRemoteName, "\", "")): _
              X = X + 1
              
    Next Läufer

    Netsuche = rPC
    Erase rPC
    
End Function
Public Function Get_LocalPCs(Optional AddOwnPC As Boolean = True) _
                                As String()
    Dim n As String, PClist() As String
    Dim X As Long
        
    On Local Error GoTo Quit
    
    PClist = Netsuche
    
    If (Not AddOwnPC) Then
        If isArray(PClist) Then
            n = Get_LocalNetworkName
        
            X = ARR.Find(PClist, n, , , False)
            If X > -1 Then Call ARR.DeleteIndex(PClist, X)
        End If
    End If
      
Quit:
    If Err.Number <> 0 Then Erase PClist
    Get_LocalPCs = PClist
    
End Function
Private Function HostByAddress(ByVal Addresse$) As String
    Dim X As Integer
    Dim HostDeAddress As Long
    Dim aa As String, bb As String * 5
    Dim Host As HostDeType
  
    aa = Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(Addresse))
    
    HostDeAddress = gethostbyaddr(aa, Len(aa), 2)
    
    If HostDeAddress = 0 Then HostByAddress = "": _
                              Exit Function
    
    Call RtlMoveMemory(Host, HostDeAddress, LenB(Host))
 
    aa = ""
    X = 0
    
    Do
       Call RtlMoveMemory(ByVal bb, Host.hName + X, 1)
       If Left$(bb, 1) = Chr$(0) Then Exit Do
       aa = aa + Left$(bb, 1)
       X = X + 1
    Loop
    
    HostByAddress = aa

End Function
Public Function HTTP_CanConnect() As Boolean
    HTTP_CanConnect = (InternetAttemptConnect(ByVal 0&) = 0)
End Function
Public Function Host_isReachable(ByVal Host As String) _
                                        As Boolean
    
    Host_isReachable = (InternetCheckConnection( _
                        Host, FLAG_ICC_FORCE_CONNECTION, 0&) = 0)
    
End Function
Public Function HTTP_isOnline() As Boolean
     HTTP_isOnline = (InetIsOffline(0) = 0)
End Function
Public Function Get_NetTime(ByVal Server As String, _
                            ServerDate As Date, _
                            ServerTime As Date, _
                            Optional SetLocalDate As Boolean = False) _
                                As Boolean
    Dim nBuffer As Long, nResult As Long
    Dim uTime As TIME_OF_DAY_INFO
    Dim bServer() As Byte
        
    On Local Error GoTo Quit
    
    If Trim$(Server) = "" Then
        nResult = NetRemoteTOD(vbNullString, nBuffer)
    Else
        bServer = Server & vbNullChar
        nResult = NetRemoteTOD(bServer(0), nBuffer)
    End If
  
    If nResult = 0 Then
        CopyMemory uTime, nBuffer, Len(uTime)
        NetApiBufferFree nBuffer
    
        With uTime
            ServerDate = DateSerial(.tod_year, .tod_month, .tod_day)
            ServerTime = TimeSerial(.tod_hours, .tod_mins - .tod_timezone, .tod_secs)
        End With
        
        Get_NetTime = True
        
        If SetLocalDate Then
            On Local Error Resume Next
            
            Date = ServerDate
            Time = ServerTime
            
            If Err.Number <> 0 Then Get_NetTime = False
        End If
    End If
    
Quit:
    If Err.Number <> 0 Then ServerDate = CDate(0): _
                            ServerTime = CDate(0): _
                            Get_NetTime = False
    
End Function
Public Function isValidMail(Mail As String) As Boolean
    Dim bGoodAdress As Boolean
    Dim sTopLevelDomainsArray() As String, sTopLevelDomains As String
    Dim sValidChars As String, eMailSplices() As String
    Dim i As Integer
    
    Mail = LCase(Mail)
    
    sTopLevelDomains = "com,net,edu,arpa,org,gov,museum," & _
                       "biz,info,pro,name,aero,coop,mil,ac,ad,ae,af,ag,ai,al," & _
                       "am,an,ao,aq,ar,as,at,au,aw,az,ba,bb,bd,be,bf,bg," & _
                       "bh,bi,bj,bm,bn,bo,br,bs,bt,bv,bw,by,bz,ca,cc,cd," & _
                       "cf,cg,ch,ci,ck,cl,cm,cn,co,cr,cu,cv,cx,cy,cz,de," & _
                       "dj,dk,dm,do,dz,ec,ee,eg,eh,er,es,et,fi,fj,fk,fm," & _
                       "fo,fr,ga,gd,ge,gf,gg,gh,gi,gl,gm,gn,gp,gq,gr,gs," & _
                       "gt,gu,gw,gy,hk,hm,hn,hr,ht,hu,id,ie,il,im,in,io," & _
                       "iq,ir,is,it,je,jm,jo,jp,ke,kg,kh,ki,km,kn,kp,kr," & _
                       "kw,ky,kz,la,lb,lc,li,lk,lr,ls,lt,lu,lv,ly,ma,mc," & _
                       "md,mg,mh,mk,ml,mm,mn,mo,mp,mq,mr,ms,mt,mu,mv,mw," & _
                       "mx,my,mz,na,nc,ne,nf,ng,ni,nl,no,np,nr,nu,nz,om," & _
                       "pa,pe,pf,pg,ph,pk,pl,pm,pn,pr,ps,pt,pw,py,qa,re," & _
                       "ro,ru,rw,sa,sb,sc,sd,se,sg,sh,si,sj,sk,sl,sm,sn," & _
                       "so,sr,st,sv,sy,sz,tc,td,tf,tg,th,tj,tk,tm,tn,to," & _
                       "tp,tr,tt,tv,tw,tz,ua,ug,uk,um,us,uy,uz,va,vc,ve," & _
                       "vg,vi,vn,vu,wf,ws,ye,yt,yu,za,zm,zr,zw"
  
    sValidChars = lABC & Num & ".-_@"
  
    sTopLevelDomainsArray = Split(sTopLevelDomains, ",")

    eMailSplices = Split(Mail, "@")
  
    If UBound(eMailSplices) <> 1 Then Exit Function
  
    eMailSplices = Split(eMailSplices(1), ".")
  
    If UBound(eMailSplices) < 1 Then Exit Function
    
    For i = 0 To UBound(sTopLevelDomainsArray)
        If eMailSplices(UBound(eMailSplices)) = sTopLevelDomainsArray(i) Then _
            bGoodAdress = True: _
            Exit For
    Next i
  
    If Not bGoodAdress Then Exit Function
    
    For i = 1 To Len(Mail)
        If InStr(1, sValidChars, LCase$(Mid$(Mail, i, 1))) = 0 Then _
            Exit Function
    Next i
  
    isValidMail = True

End Function
Private Sub NetAnalyze(Cards() As Variant, _
                       ReceivedBytes() As Variant, _
                       SendedBytes() As Variant)
    Dim IPInterfaceRow As MIB_IFROW
    Dim Buff() As Byte
    Dim cbRequired As Long, nStructSize As Long, nRows As Long, cnt As Long
    Dim C$, R$, S$
    Dim D As Double, X As Double
    
    Call GetIfTable(ByVal 0&, cbRequired, 1)
    
    If cbRequired > 0 Then
        ReDim Buff(0 To cbRequired - 1)
        
        If GetIfTable(Buff(0), cbRequired, 1) = ERROR_SUCCESS Then
            nStructSize = LenB(IPInterfaceRow)
            
            CopyMemoryA nRows, Buff(0), 4
                        
            For cnt = 1 To nRows
                CopyMemoryA IPInterfaceRow, Buff(4 + (cnt - 1) * nStructSize), nStructSize
                
                If IPInterfaceRow.dwType = MIB_IF_TYPE_ETHERNET Then
                    C$ = GetName(IPInterfaceRow.bDescr)
                    
                    X = IPInterfaceRow.dwMtu / 1024
                    D = IPInterfaceRow.dwInUcastPkts * X
                    
                    R$ = FormatNumber(D, 0)
                    
                    D = IPInterfaceRow.dwOutUcastPkts _
                                     + IPInterfaceRow.dwOutNUcastPkts _
                                     * IPInterfaceRow.dwMtu / 1024
                    S$ = FormatNumber(D, 0)
                    
                    R$ = Replace(R$, ".", "")
                    S$ = Replace(S$, ".", "")
                    
                    Call ARR.Add(Cards, C$, ADDTOEND)
                    Call ARR.Add(ReceivedBytes, R$, ADDTOEND)
                    Call ARR.Add(SendedBytes, S$, ADDTOEND)
                End If
            Next cnt
        End If
    End If

End Sub
Private Function GetName(ByRef AR) As String
    Dim i As Integer
    Dim Str As String
    
    For i = 0 To UBound(AR)
        Str = Str + Chr(AR(i))
    Next i
    
    Str = Replace(Str, Chr$(0), "")
    
    GetName = Str
    
End Function
Public Function Adapter_GetInstalled() As String()
    Dim C() As Variant, R() As Variant, S() As Variant
    Dim T() As String
    Dim P As Long
    
    Call NetAnalyze(C, R, S)
    
    If ARR.isArray(C) Then
        ReDim T(UBound(C))
        
        For P = 0 To UBound(C)
            T(P) = CStr(C(P))
        Next P
    End If
    
    Adapter_GetInstalled = T
    
End Function
Public Function Adapter_Count() As Long
    Dim C() As Variant, R() As Variant, S() As Variant
    
    Adapter_Count = 0
    
    Call NetAnalyze(C, R, S)

    If ARR.isArray(C) Then Adapter_Count = UBound(C) + 1
    
End Function
Public Sub Adapter_GetTransferredBytes(Cards As Variant, _
                                       ReceivedBytes As Variant, _
                                       SendedBytes As Variant)
    Dim C() As Variant, R() As Variant, S() As Variant
    
    Call NetAnalyze(C, R, S)
    
    Cards = C
    ReceivedBytes = R
    SendedBytes = S
    
End Sub
Public Function IE_GetURL() As String()
    Dim B() As Variant
    Dim Parent As Long, Child As Long, P As Long, A() As Long
    Dim URL As String
    
    A = WIN.Get_AllHwndByClass("IEFrame")
    
    If ARR.isArray(A) Then
        For P = 0 To UBound(A)
            Parent = A(P)
            
            Child = FindWindowEx(Parent, 0, "WorkerW", vbNullString)
            Child = FindWindowEx(Child, 0, "ReBarWindow32", vbNullString)
            Child = FindWindowEx(Child, 0, "ComboBoxEx32", vbNullString)
            Child = FindWindowEx(Child, 0, "ComboBox", vbNullString)
            Child = FindWindowEx(Child, 0, "Edit", vbNullString)
            
            URL = WIN.Get_Text(Child)
            
            If URL <> "" Then Call ARR.Add(B, URL, ADDTOEND)
        Next P
    End If
    
    IE_GetURL = ARR.VarriantToString(B)
    
End Function
Public Sub IE_CloseAll()
    Dim P As Long, A() As Long
    
    A = WIN.Get_AllHwndByClass("IEFrame")
    
    If ARR.isArray(A) Then
        For P = 0 To UBound(A)
            Call WIN.CloseWindow(A(P))
        Next P
    End If
    
End Sub
Public Function Adapter_GetInfo(ByVal AdapterIndex As Long) _
                                    As String()
    Dim oAdapters As Object, oAdapter As Object
    Dim A() As String
    Dim X As Long, P As Long, C As Long
    
    On Local Error Resume Next

    If Adapter_Count < AdapterIndex Then Exit Function
    
    Set oAdapters = GetObject("winmgmts:").execquery( _
                    "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    If Err.Number <> 0 Then Exit Function
    
    On Local Error GoTo Quit
    
    ReDim A(6)
    
    For Each oAdapter In oAdapters
        C = C + 1
        
        If C = AdapterIndex Then
            With oAdapter
                A(0) = .MACAddress
                A(1) = Join(.IpAddress)
                A(2) = Join(.IPSubnet)
                If Not IsNull(.DefaultIPGateway) Then A(3) = Join(.DefaultIPGateway)
                If Not IsNull(.DNSHostName) Then A(4) = .DNSHostName
                If Not IsNull(.WINSPrimaryServer) Then A(5) = .WINSPrimaryServer
                If Not IsNull(.WINSSecondaryServer) Then A(6) = .WINSSecondaryServer
            End With
        
            Exit For
        End If
    Next
    
Quit:
    If Err.Number <> 0 Then Erase A
    Adapter_GetInfo = A
      
End Function
Private Function pLongIP2Dotted(ByVal LongIP As Variant) As String
    Dim i As Integer
    Dim Num As Byte
        
    On Error GoTo ErrLongIP2Dotted
    
    If LongIP = "" Then GoTo ErrLongIP2Dotted
    If LongIP < 0 Then LongIP = 4294967296# + LongIP
    If LongIP < 0 Then GoTo ErrLongIP2Dotted
    
    For i = 3 To 0 Step -1
        Num = Int(LongIP / 256 ^ (i))
        LongIP = LongIP - (Num * 256 ^ (i))
        pLongIP2Dotted = Num & "." & pLongIP2Dotted
    Next i
    
    pLongIP2Dotted = Left$(pLongIP2Dotted, Len(pLongIP2Dotted) - 1)
    
    Exit Function

ErrLongIP2Dotted:
    pLongIP2Dotted = "0.0.0.0"

End Function
Public Function Get_Gateway() As String
    Dim Siz As Long, X As Long
    Dim IPTab As PMIB_IPFORWARDTABLE
    Dim IPRow As PMIB_IPFORWARDROW
    Dim bBuffer() As Byte
    
    X = GetIpForwardTable(ByVal 0&, Siz, False)
    If X <> 0 And X <> 122 Then Exit Function
    
    ReDim bBuffer(Siz)
    
    If GetIpForwardTable(bBuffer(0), Siz, False) = 0 Then
        CopyMemoryA IPTab, bBuffer(0), Len(IPTab)
        Siz = IPTab.dwNumEntries - 1
        
        For X = 0 To Siz
            CopyMemoryA IPTab, bBuffer(X * (Len(IPTab) - 4)), Len(IPTab)
            
            If IPTab.table.dwForwardDest = 0 Then _
                IPRow = IPTab.table: _
                Get_Gateway = pLongIP2Dotted(IPTab.table.dwForwardNextHop)
        Next X
    End If
    
End Function
Private Function GetShares(ByVal Server As String, Shares() As tShare) As Boolean
    Dim Gelesen As Long, Puffer As Long, Resultat As Long
    Dim ResumeHandle As Long, Total As Long, i As Long
    
    Resultat = NetShareEnum(strptr(Server), 1, VarPtr(Puffer), -1, Gelesen, Total, ResumeHandle)
        
    If Resultat = 0 And Gelesen > 0 Then
        ReDim Shares(0 To Gelesen - 1)
        
        For i = 0 To Gelesen - 1
            With Shares(i)
                .Name = Dereferenz_String(Dereferenz_Long(Puffer + 12 * i))
                .Typ = Dereferenz_Long(Puffer + 12 * i + 4)
                .Beschreibung = Dereferenz_String(Dereferenz_Long(Puffer + 12 * i + 8))
            End With
        Next i
        
        GetShares = True
    End If
    
    If CBool(Puffer) Then NetApiBufferFree Puffer
    
End Function
Private Function Dereferenz_Long(ByVal Pointer As Long) As Long
    Call CopyMemory(Dereferenz_Long, ByVal Pointer, 4)
End Function
Private Function Dereferenz_String(ByVal Pointer As Long) As String
    Dim Laenge As Long
    Dim Puffer() As Byte
    
    If CBool(Pointer) Then
        Laenge = lstrlenW(Pointer) * 2
        
        If CBool(Laenge) Then
            ReDim Puffer(Laenge - 1) As Byte
            CopyMemory Puffer(0), ByVal Pointer, Laenge
            Dereferenz_String = Puffer
        End If
    End If
    
End Function
Public Function Get_SharedDirectory(ByVal Server As String) As String()
    Dim A() As Variant
    Dim TS() As tShare
    Dim P As Long
    
    On Local Error GoTo Quit
    
    Call GetShares(Server, TS)
    
    For P = 0 To UBound(TS)
        If TS(P).Typ = 0 Then Call ARR.Add(A, TS(P).Name, ADDTOEND)
    Next P
    
    Get_SharedDirectory = ARR.VarriantToString(A)
    
Quit:
End Function
Public Function Get_DialUPlist() As String()
    Dim lngBuffer As Long, lngEntries As Long, i As Long
    Dim strRASConName As String, A() As String
    Dim udtRASCon(255) As RASENTRYNAME95

    On Local Error GoTo Quit
    
    udtRASCon(0).dwSize = 264
    lngBuffer = 256 * udtRASCon(0).dwSize

    Call RasEnumEntries(vbNullString, vbNullString, udtRASCon(0), lngBuffer, lngEntries)

    ReDim A(lngEntries - 1)
    
    For i = 0 To lngEntries - 1
        strRASConName = StrConv(udtRASCon(i).szEntryName(), vbUnicode)
        strRASConName = Replace(strRASConName, Chr$(0), "")
    
        A(i) = strRASConName
    Next i

Quit:
    If Err.Number <> 0 Then Erase A
    
    Get_DialUPlist = A
    
End Function

