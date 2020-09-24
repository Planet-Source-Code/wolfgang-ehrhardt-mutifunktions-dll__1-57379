Attribute VB_Name = "FileBase"
Option Explicit

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type SHFILEOPSTRUCT
    hWnd                  As Long
    wFunc                 As Long
    pFrom                 As String
    pTo                   As String
    fFlags                As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings         As Long
    lpszProgressTitle     As String
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4

Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_WANTMAPPINGHANDLE = &H20
Public Const FOF_FILESONLY = &H80
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_NOERRORUI As Long = &H400
Public Function SplitPath(ByVal Path As String) As Long
    Dim P As Long
    
    For P = Len(Path) To 1 Step -1
        If Mid$(Path, P, 1) = "\" Then Exit For
    Next P
    
    SplitPath = P
    
End Function
Public Function GetBlankName(ByVal File As String) As String
    Dim X As Integer
    Dim Str As String
    
    X = SplitPath(File)
    File = Mid$(File, X + 1)
            
    For X = Len(File) To 1 Step -1
        If Mid$(File, X, 1) = "." Then _
            GetBlankName = Mid$(File, 1, X - 1): _
            Exit Function
    Next X
    
    GetBlankName = File

End Function
Public Function GetDir(ByVal Directory As String) As String
    GetDir = Directory
    If Right$(GetDir, 1) <> "\" Then GetDir = GetDir & "\"
End Function
Public Function DRVisEqual(ByVal Path1 As String, _
                           ByVal Path2 As String) As Boolean
    
    On Local Error GoTo Quit
    
    Path1 = LCase(GetDir(Path1))
    Path2 = LCase(GetDir(Path2))
    
    DRVisEqual = (Mid$(Path1, 1, 3) = Mid$(Path2, 1, 3))

Quit:

End Function
Public Function DeleteTree(ByVal Directory As String) _
                                As Boolean
    
    If Not DI.Exist(Directory) Then Exit Function
    
    Call DeleteTree_(Directory)
    If Not DI.Exist(Directory) Then DeleteTree = True
    
End Function
Private Sub DeleteTree_(Directory As String)
    Dim sCurrFile As String
    
    On Local Error Resume Next
    
    sCurrFile = GetDir(Directory)
    sCurrFile = Dir(sCurrFile & "*.*", vbDirectory)
    
    Do While Len(sCurrFile) > 0
        If sCurrFile <> "." And sCurrFile <> ".." Then
            If (GetAttr(Directory & "\" & sCurrFile) _
            And vbDirectory) = vbDirectory Then
                Call DeleteTree_(Directory & "\" & sCurrFile)
                sCurrFile = Dir(Directory & "\*.*", vbDirectory)
            Else
                Kill Directory & "\" & sCurrFile
                sCurrFile = Dir
            End If
        Else
            sCurrFile = Dir
        End If
    Loop
    
    RmDir Directory
    
End Sub
Public Function SH_FileOperation(ByVal hWnd As Long, _
                              ByVal Operation As FileOperationNFO, _
                              ByVal pFrom As String, _
                              ByVal pTo As String, _
                              ByVal IncludingSubDirectorys As Boolean, _
                              ByVal Confirm As Boolean, _
                              ByVal ShowDialogs As Boolean, _
                              ByVal ShowProgress As Boolean, _
                              Optional MoveToRecycleBin As Boolean = False) _
                                    As Boolean
    Dim SFO As SHFILEOPSTRUCT
    Dim nPath As String, Path As String
    Dim wFunc As Long, Flags As Long, X As Long
    
    On Local Error GoTo Quit
    
    Select Case Operation
        Case foCOPY
            wFunc = FO_COPY
        Case foMOVE
            wFunc = FO_MOVE
        Case foRENAME
            wFunc = FO_RENAME
            
            If InStr(pTo, "\") = 0 Then _
                X = SplitPath(pFrom): _
                pTo = GetDir(Mid$(pFrom, 1, X - 1)) & pTo
        Case foDELETE
            wFunc = FO_DELETE
            pTo = ""
    End Select

    With SFO
        .hWnd = hWnd
        .wFunc = wFunc
        .pFrom = pFrom & Chr$(0)
        .pTo = pTo & Chr$(0)
        .fFlags = 0
        
        If Not IncludingSubDirectorys Then
            .fFlags = .fFlags Or FOF_FILESONLY
        Else
            If .wFunc = FO_COPY Then _
                Call DI.CreatePath(GetDir(pTo))
        End If
        
        If ShowProgress Then _
            .fFlags = .fFlags Or FOF_SIMPLEPROGRESS
        If Not Confirm Then _
            .fFlags = .fFlags Or FOF_NOCONFIRMATION
        
        If Not ShowDialogs Then _
            .fFlags = .fFlags Or FOF_SILENT: _
            .fFlags = .fFlags Or FOF_NOERRORUI
    
        If Operation = foDELETE And MoveToRecycleBin Then _
            .fFlags = .fFlags Or FOF_ALLOWUNDO
    End With

    If SHFileOperation(SFO) = 0 Then _
        SH_FileOperation = True
        
Quit:
End Function
