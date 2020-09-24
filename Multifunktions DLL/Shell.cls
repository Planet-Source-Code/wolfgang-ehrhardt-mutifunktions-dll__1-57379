VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long

Private Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadID  As Long
End Type

Private Type STARTUPINFO
    CB              As Long
    lpReserved      As Long
    lpDesktop       As Long
    lpTitle         As Long
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As Long
    bInheritHandle       As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&

Private Const PROCESS_QUERY_INFORMATION = &H400&

Private Const STILL_ACTIVE = &H103&
Public Sub Execute(AppPath As String)
    Call ShellExecute(0, "Open", AppPath, "", "", 1)
End Sub
Public Sub Explorer_OpenPath(Path As String)
    Call Execute(Path)
End Sub
Public Sub ExecuteAndWait(ByVal AppPath As String, _
                          Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus)
    Dim ProcId As Long, ProcHnd As Long, Active As Long

    ProcId = Shell(AppPath, WindowStyle)
    ProcHnd = OpenProcess(PROCESS_QUERY_INFORMATION, _
                          True, ProcId)

    Do
        DoEvents
        
        Call GetExitCodeProcess(ProcHnd, Active)
    Loop While Active = STILL_ACTIVE

    Call CloseHandle(ProcHnd)
    
End Sub
Public Function WaitForTask(ByVal WindowCaption As String) As Long
    Dim ProcId As Long, ProcHnd As Long, hWnd As Long
        
    hWnd = WIN.FindWindowByCaption(WindowCaption)
    ProcId = WIN.Get_TaskID(hWnd)
    
    ProcHnd = OpenProcess(PROCESS_QUERY_INFORMATION, _
                          True, ProcId)

    Do
        DoEvents
        
        Call GetExitCodeProcess(ProcHnd, WaitForTask)
    Loop While WaitForTask = STILL_ACTIVE

    Call CloseHandle(ProcHnd)
    
End Function
Public Function DOS_Execute(ByVal Param As String) As String
    Dim Proc As PROCESS_INFORMATION
    Dim Result As Long, bSuccess As Long, hReadPipe As Long
    Dim hWritePipe As Long, L As Long
    Dim Start As STARTUPINFO
    Dim Sa As SECURITY_ATTRIBUTES
    Dim Buffer As String
     
    Sa.nLength = Len(Sa)
    Sa.bInheritHandle = 1&
    Sa.lpSecurityDescriptor = 0&
    
    Result = CreatePipe(hReadPipe, hWritePipe, Sa, 0)
    
    If Result = 0 Then GoTo Quit
    
    Start.CB = Len(Start)
    Start.dwFlags = STARTF_USESTDHANDLES
    Start.hStdOutput = hWritePipe
    
    Result = CreateProcessA(0&, Param, Sa, Sa, 1&, _
                            NORMAL_PRIORITY_CLASS, 0&, _
                            0&, Start, Proc)
 
    If Result <> 1 Then GoTo Quit
    
    Buffer = String(2000, Chr$(0))
    bSuccess = ReadFile(hReadPipe, Buffer, Len(Buffer), L, 0&)
      
    If bSuccess = 1 Then _
        DOS_Execute = Trim(Left$(Buffer, L)): _
        DOS_Execute = Replace(DOS_Execute, Chr$(0), "")
    
Quit:
    Call CloseHandle(Proc.hProcess)
    Call CloseHandle(Proc.hThread)
    Call CloseHandle(hReadPipe)
    Call CloseHandle(hWritePipe)

End Function
Public Function DOS_Help(ByVal DOScommand As String) As String
    Dim Cmd As String
    
    Cmd = IIf(DOScommand = vbNullString, "Help", _
                                         DOScommand & " /?")
    DOS_Help = DOS_Execute(Cmd)

End Function
