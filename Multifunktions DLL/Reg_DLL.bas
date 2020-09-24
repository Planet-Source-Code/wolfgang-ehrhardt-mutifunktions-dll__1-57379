Attribute VB_Name = "Reg_DLL"
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Const SYNCHRONIZE = &H100000

Dim File As String
Public Function RegDLL() As Boolean
    Dim P As Integer
    Dim Path As String
    Dim L As Long, T As Long
    
    If TestDLL Then _
        RegDLL = True: _
        Exit Function
    
    Path = App.Path
    If Right(Path, 1) = "\" Then Path = Mid(Path, 1, Len(Path) - 1)
    
    For P = Len(Path) To 1 Step -1
        If Mid(Path, P, 1) = "\" Then Exit For
    Next P
    
    If P = 0 Then Path = Path & "\": P = Len(Path)
        
    Path = Mid(Path, 1, P) & "WindowAPI.dll"
    
    If Not Exist(Path) Then _
        MsgBox "Can't find '" & Path & "'", _
                vbCritical + vbOKOnly, _
                "Error": _
        End
    
    If Not RegisterComponents(Path) Then _
        MsgBox "Can't register '" & Path & "' !", _
               vbCritical + vbOKOnly, _
               "Error": _
        End

End Function
Private Function Exist(ByVal File As String) As Boolean
    Dim FN As Integer

    On Local Error GoTo Quit
    
    FN = FreeFile
    
    Open File For Input As FN
    Close FN

    Exist = True
    
Quit:
    
End Function
Private Function RegisterComponents(ByVal ComponentPath As String) _
                                        As Boolean
    Dim Result As Long, TaskID As Long, Handle As Long
    Dim Path As String, Param As String
    
    If Not Exist(ComponentPath) Then Exit Function
    
    Path = Dir_GetWindowsSystemDir & "\REGSVR32.EXE"
    
    If Not Exist(Path) Then _
        Path = Dir_GetWindowsDir & "\REGSVR32.EXE": _
        If Not Exist(Path) Then _
            Call MsgBox("Sorry...Can't find 'Regsvr32.exe'", _
                        vbCritical + vbOKCancel, _
                        "Can't Register DLL"): _
            End
    
    Path = Get_DOSfileName(Path)
    ComponentPath = Get_DOSfileName(ComponentPath)
    
    Param = " /s "
    TaskID = Shell(Path & Param & ComponentPath)
    
    Handle = OpenProcess(SYNCHRONIZE, False, TaskID)
    Result = WaitForSingleObject(Handle, 8000)
    
    If Result Then _
        Call TerminateProcess(Handle, 0): _
        Call CloseHandle(Handle): _
        RegisterComponents = True
    
End Function
Private Function Get_DOSfileName(ByVal Path As String) _
                                    As String
    Dim Result&, aa$
    
    aa = Space$(255)
    Result = GetShortPathName(Path, aa, Len(aa))
    Get_DOSfileName = Mid$(aa, 1, Result)

End Function
Private Function Dir_GetWindowsDir() As String
    Dim TEMP As String * 256
    Dim Result As Integer
    
    Result = GetWindowsDirectory(TEMP, Len(TEMP))
    Dir_GetWindowsDir = Left$(TEMP, Result)

End Function
Private Function Dir_GetWindowsSystemDir() As String
    Dim TEMP As String * 256
    Dim Result As Integer
    
    Result = GetSystemDirectory(TEMP, Len(TEMP))
    Dir_GetWindowsSystemDir = Left$(TEMP, Result)
    
 End Function
Public Function TestDLL() As Boolean
    Dim wAPI As New Window
    
    On Local Error Resume Next
    
    Call wAPI.GetParent(0)
    If Err = 0 Then TestDLL = True

End Function
