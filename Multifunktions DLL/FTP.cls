VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

Public Enum IE_OpenType
    INTERNET_OPEN_TYPE_PRECONFIG = 0
    INTERNET_OPEN_TYPE_DIRECT = 1
    INTERNET_OPEN_TYPE_PROXY = 3
    INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4
End Enum

Public Enum TransferType
    FTP_TRANSFER_TYPE_UNKNOWN = &H0
    FTP_TRANSFER_TYPE_ASCII = &H1
    FTP_TRANSFER_TYPE_BINARY = &H2
End Enum

Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_FLAG_PASSIVE = &H8000000

Private iCon() As String
Private hOpen As Long
Public Function Disconnect(ByVal Server As String) As Long
    Dim Index As Long, P As Long
    
    Index = ConnectionExist(Server)
    
    If Index > -1 Then
        If InternetCloseHandle(CLng(iCon(Index, 2))) Then _
            iCon(Index, 1) = "": _
            Disconnect = 1
    Else
        Disconnect = -1
    End If
    
End Function
Public Function DisconnectAll() As Boolean
    Dim P As Long, X As Long, Y As Long
    
    On Local Error Resume Next
    
    If hOpen = 0 Then Exit Function
    
    X = 0
    
    For P = 0 To UBound(iCon)
        If iCon(P, 1) <> "" Then
            Y = Y + 1
            If InternetCloseHandle(CLng(iCon(P, 2))) Then _
                iCon(P, 1) = "": _
                X = X + 1
        End If
    Next P
    
    If X = Y Then
        Erase iCon
        If InternetCloseHandle(hOpen) Then DisconnectAll = True
    End If
    
End Function
Public Function Connect(IEopen As IE_OpenType, _
                        ByVal Server As String, _
                        Optional UserName As String = "anonymous", _
                        Optional Password As String = "guest", _
                        Optional Port As Long = 21, _
                        Optional PassiveConnection As Boolean = True) _
                            As Long
    Dim hConnection As Long, P As Long
    Dim Found As Boolean
    
    If ConnectionExist(Server) > -1 Then _
        Connect = -1: _
        Exit Function
    
    If hOpen = 0 Then _
        hOpen = InternetOpen("MfDLL", IEopen, _
                             vbNullString, vbNullString, 0)
    
    hConnection = InternetConnect(hOpen, Server, Port, _
                                  UserName, Password, _
                                  INTERNET_SERVICE_FTP, _
                                  IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)

    On Local Error Resume Next
    
    If hConnection Then
        P = UBound(iCon)
        
        If Err.Number Then
            Err.Clear
            ReDim iCon(0, 2)
            P = 0
        Else
            For P = 0 To UBound(iCon)
                If iCon(P, 1) = "" Then Found = True: _
                                        Exit For
            Next P
        
            If Not Found Then _
                ReDim Preserve iCon(UBound(iCon) + 1): _
                P = UBound(iCon)
        End If
        
        iCon(P, 1) = LCase(Server)
        iCon(P, 2) = hConnection
            
        Connect = 1
    End If
    
End Function
Public Sub GetLastResponse(ErrorCode As String, _
                           ErrorStr As String)
    Dim lErr As Long, lenBuf As Long
    Dim sErr As String, R$
    
    Call InternetGetLastResponseInfo(lErr, sErr, lenBuf)

    sErr = String(lenBuf, 0)
    
    Call InternetGetLastResponseInfo(lErr, sErr, lenBuf)
    
    ErrorCode = 0
    ErrorStr = ""
    
    If lErr Then ErrorCode = lErr: _
                 R$ = "Error " & CStr(ErrorCode) & ": "
    
    ErrorStr = R$ & sErr
    
End Sub
Private Function ConnectionExist(ByVal Server As String) As Long
    Dim P As Long
    
    On Local Error GoTo Quit
        
    For P = 0 To UBound(iCon)
        If iCon(P, 1) = LCase(Server) Then ConnectionExist = P: _
                                           Exit Function
    Next P
        
Quit:
    ConnectionExist = -1

End Function
Public Function GetCurrentDirectory(ByVal Server As String) As String
    Dim Index As Long, hConnection As Long
    Dim sOrgPath As String

    Index = ConnectionExist(Server)
    If Index = -1 Then Exit Function
    
    hConnection = iCon(hConnection, 2)
        
    sOrgPath = String(MAX_PATH, 0)
    
    Call FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
    
    GetCurrentDirectory = Replace(sOrgPath, Chr$(0), "")
    
End Function
Public Function DirectoryCreate(ByVal Server As String, _
                                ByVal Directory As String) _
                                      As Long
    Dim Index As Long
    
    Index = ConnectionExist(Server)
    
    If Index = -1 Then
        DirectoryCreate = -1
    Else
        If FtpCreateDirectory(CLng(iCon(Index, 2)), Directory) Then _
            DirectoryCreate = 1
    End If
    
End Function
Public Function DirectoryRemove(ByVal Server As String, _
                                ByVal Directory As String) _
                                    As Long
    Dim Index As Long
    
    Index = ConnectionExist(Server)
    
    If Index = -1 Then
        DirectoryRemove = -1
    Else
        If FtpRemoveDirectory(CLng(iCon(Index, 2)), Directory) Then _
            DirectoryRemove = 1
    End If
    
End Function
Public Function SetCurrentDirectory(ByVal Server As String, _
                                    ByVal Directory As String) _
                                        As Long
    Dim Index As Long
    
    Index = ConnectionExist(Server)
    
    If Index = -1 Then
        SetCurrentDirectory = -1
    Else
        If FtpSetCurrentDirectory(CLng(iCon(Index, 2)), Directory) Then _
            SetCurrentDirectory = 1
    End If
    
End Function
Public Function Upload(ByVal Server As String, _
                       ByVal LocalFile As String, _
                       ByVal TransferModus As TransferType, _
                       Optional RemoteFile As String) _
                            As Long
    Dim Index As Long
        
    Index = ConnectionExist(Server)
    
    If Index = -1 Then
        Upload = -1
    Else
        If RemoteFile = "" Then _
            RemoteFile = F.Get_FileFromPath(LocalFile)
    
        If FtpPutFile(CLng(iCon(Index, 2)), _
                      LocalFile, RemoteFile, _
                      TransferModus, 0) Then Upload = 1
    End If
    
End Function
Public Function Rename(ByVal Server As String, _
                       ByVal Source As String, _
                       ByVal NewName As String) _
                            As Long
    Dim Index As Long
    
    Index = ConnectionExist(Server)
    
    If Index = -1 Then
        Rename = -1
    Else
        If FtpRenameFile(CLng(iCon(Index, 2)), Source, NewName) Then _
            Rename = 1
    End If
    
End Function
Public Function Download(ByVal Server As String, _
                         ByVal RemoteFile As String, _
                         ByVal LocalFile As String, _
                         TransferModus As TransferType) _
                            As Long
    Dim Index As Long
    
    Index = ConnectionExist(Server)
    
    If Index = -1 Then
        Download = -1
    Else
        If FtpGetFile(CLng(iCon(Index, 2)), _
                      RemoteFile, LocalFile, _
                      False, 0, TransferModus, 0) Then _
                            Download = 1
    End If
    
End Function
Public Function FileRemove(ByVal Server As String, _
                           ByVal File As String) _
                                As Long
    Dim Index As Long
    
    Index = ConnectionExist(Server)
    
    If Index = -1 Then
        FileRemove = -1
    Else
        If FtpDeleteFile(CLng(iCon(Index, 2)), File) Then _
            FileRemove = 1
    End If
    
End Function
Private Function EnumFiles(hConnection As Long, _
                           FileList As String, _
                           DirectoryList As String) As Boolean
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long, lRet As Long
    Dim R$
    
    FileList = ""
    DirectoryList = ""
    
    pData.cFileName = String(MAX_PATH, 0)
    
    hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    If hFind = 0 Then Exit Function
    
    R$ = Left$(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1) & Chr$(0)
    
    If pData.dwFileAttributes = vbDirectory Then
        DirectoryList = R$
    Else
        FileList = R$
    End If
    
    Do
        pData.cFileName = String(MAX_PATH, 0)
        
        lRet = InternetFindNextFile(hFind, pData)
        If lRet = 0 Then Exit Do
        
        R$ = Left$(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1) & Chr$(0)
    
        If pData.dwFileAttributes = vbDirectory Then
            DirectoryList = DirectoryList & R$
        Else
            FileList = FileList & R$
        End If
    Loop
    
    If Right$(DirectoryList, 1) = Chr$(0) Then _
        DirectoryList = Mid$(DirectoryList, 1, Len(DirectoryList) - 1)
    If Right$(FileList, 1) = Chr$(0) Then _
        FileList = Mid$(FileList, 1, Len(FileList) - 1)

    Call InternetCloseHandle(hFind)
    
    EnumFiles = True
    
End Function
Public Function GetFileList(ByVal Server As String, _
                            FileList As String, _
                            DirectoryList As String) _
                                As Boolean
    Dim Index As Long
    
    Index = ConnectionExist(Server)
    
    If Index <> -1 Then _
        GetFileList = EnumFiles(CLng(iCon(Index, 2)), FileList, DirectoryList)
    
End Function
