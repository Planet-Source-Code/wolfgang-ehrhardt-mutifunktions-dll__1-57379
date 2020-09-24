VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ***********************************************
' * Copyright / Bitte lesen                     *
' ***********************************************
' *                                             *
' * Zusammengetragen, ediert & neu verfasst von *
' *     Wolfgang Ehrhardt                       *
' *         woeh@gmx.de                         *
' *                                             *
' * Dieser Code ist frei verfügbar, solange  er *
' * für PRIVATE Zwecke benutzt wird.            *
' * Wird dieser Code oder Auszüge daraus in ein *
' * Kommerzielles Produkt implantiert, bitte    *
' * ich um Kontaktaufnahme unter obig genannter *
' * Email-Adresse.                              *
' *                                             *
' * Ansonsten ist es erwünscht, das Funktionen  *
' * und Subs in andere Projekte implantiert     *
' * werden (Solange es für private Zwecke ist!) *
' *                                             *
' * Das einzige, worum ich bitte, ist den       *
' * DLL-SourceCode nur komplett & vollständig   *
' * weiterzugeben.                              *
' * (Bitte nicht meine Arbeit zerstückeln)      *
' *                                             *
' * Ansonsten wünsche ich Dir viel Spaß         *
' * mit den Funktionen & Subs                   *
' *                                             *
' * *********************************************
Option Explicit

Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
Private Declare Function IsWindowVisibleA Lib "user32" Alias "IsWindowVisible" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Private Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function IsWindowEnabledA Lib "user32.dll" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
Private Declare Function WindowEnabled Lib "user32.dll" Alias "EnableWindow" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetActiveWindowA Lib "user32" Alias "GetActiveWindow" () As Long
Private Declare Function SetFocusA Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetParentA Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hDC As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As SIZE) As Long

Private Type WINDOWPLACEMENT
    Length           As Long
    Flags            As Long
    showCmd          As Long
    ptMinPosition    As POINTAPI
    ptMaxPosition    As POINTAPI
    rcNormalPosition As RECT
End Type

Private Type FLASHWINFO
    cbSize    As Long
    hWnd      As Long
    dwFlags   As Long
    uCount    As Long
    dwTimeout As Long
End Type

Private Const FLASHW_STOP = 0
Private Const FLASHW_CAPTION = &H1
Private Const FLASHW_TRAY = &H2
Private Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY)
Private Const FLASHW_TIMER = &H4
Private Const FLASHW_TIMERNOFG = &HC

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0

Private Const HWND_NOTOPMOST = -2

Private Const PROCESS_TERMINATE = &H1

Private Const SC_CLOSE = &HF060

Private Const WS_CHILD = &H40000000
Private Const WS_OVERLAPPED = &H0
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
Private Const WS_DLGFRAME = &H400000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Public Function Get_TopParent(ByVal hWnd As Long) As Long
        
    Do Until GetParentA(hWnd) = 0
        hWnd = GetParentA(hWnd)
    Loop
    
    Get_TopParent = hWnd
    
End Function
Public Sub FlashWindow(ByVal hWnd As Long, _
                       Optional FlashCount As Long = 3)
    Dim FlashInfo As FLASHWINFO
    
    With FlashInfo
        .cbSize = Len(FlashInfo)
        .dwFlags = FLASHW_CAPTION Or FLASHW_TIMER
        .dwTimeout = 0
        .hWnd = hWnd
        .uCount = FlashCount
    End With
    
    Call FlashWindowEx(FlashInfo)

End Sub
Public Sub Get_Pos(ByVal hWnd As Long, _
                   Left As Long, _
                   Top As Long, _
                   Right As Long, _
                   Bottom As Long)
    Dim R As RECT
    
    Call GetWindowRect(hWnd, R)
    
    Left = R.Left
    Top = R.Top
    Right = R.Right
    Bottom = R.Bottom
    
End Sub
Public Function StayOnTop(ByVal hWnd As Long, _
                          Optional RemoveFromTop As Boolean = False) _
                                As Boolean
    Dim hWndInsertAfter As Long
    
    hWndInsertAfter = IIf(RemoveFromTop, HWND_NOTOPMOST, _
                                         HWND_TOPMOST)
    
    If SetWindowPos(hWnd, hWndInsertAfter, 0, 0, 0, 0, _
                    SWP_NOSIZE Or SWP_NOMOVE) Then _
                        StayOnTop = True
    
End Function
Public Function Set_Pos(ByVal hWnd As Long, _
                        ByVal Top As Long, _
                        ByVal Left As Long) As Boolean
    Dim R As RECT
    
    Call GetWindowRect(hWnd, R)
    
    If MoveWindow(hWnd, Left, Top, _
                  R.Right - R.Left, _
                  R.Bottom - R.Top, 1) Then _
                        Set_Pos = True

End Function
Public Sub Click(ByVal hWnd As Long)
    Call SendMessageB(hWnd, WM_LBUTTONDOWN, 0, 0&)
    Call SendMessageB(hWnd, WM_LBUTTONUP, 0, 0&)
End Sub
Public Function Get_Class(ByVal hWnd As Long) As String
    Dim Buffer As String
    Dim GetClas As Long
    
    Buffer = Space(250)
    GetClas = GetClassName(hWnd, Buffer, 250)
    Get_Class = Left$(Buffer, GetClas)
    
End Function
Public Function Get_Text(ByVal hWnd As Long) As String
    Dim GetTrim As Long
    Dim TrimSpace As String, GetString As String
    
    GetTrim = SendMessageByNum(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
    TrimSpace = Space$(GetTrim)
    GetString = SendMessageByString(hWnd, WM_GETTEXT, _
                                    GetTrim + 1, TrimSpace)
    
    Get_Text = TrimSpace
        
End Function
Public Function FindChildByClass(ByVal ParentW As Long, _
                                 ByVal ClassToFind As String) As Long
    FindChildByClass = FindWindowEx(ParentW, 0&, ClassToFind, vbNullString)
End Function
Public Sub Set_Text(ByVal hWnd As Long, _
                    ByVal Text As String)
    Call SendMessageByString(hWnd, WM_SETTEXT, 0, Text)
End Sub
Public Function TerminateThread(ByVal hWnd As Long) As Boolean
    Dim Thread As Long
    
    Call GetWindowThreadProcessId(hWnd, Thread)
    Thread = OpenProcess(PROCESS_TERMINATE, 0&, Thread)
    
    TerminateThread = (TerminateProcess(Thread, 1&))
    Call CloseHandle(Thread)

End Function
Public Function IsWindowEnabled(ByVal hWnd As Long) As Boolean
    IsWindowEnabled = (IsWindowEnabledA(hWnd))
End Function
Public Sub EnableWindow(ByVal hWnd As Long, _
                        Optional Disable As Boolean = False)
    Dim fEnable As Long
    
    fEnable = IIf(Disable, 0, 1)
    Call WindowEnabled(hWnd, fEnable)
    
End Sub
Public Sub Set_WindowState(ByVal hWnd As Long, _
                           Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus)
    Call ShowWindow(hWnd, WindowStyle)
End Sub
Public Function Get_TaskID(ByVal hWnd As Long) As Long
    Call GetWindowThreadProcessId(hWnd, Get_TaskID)
End Function
Public Function Get_ActiveWindow() As Long
    Get_ActiveWindow = GetActiveWindowA
End Function
Public Function Set_Focus(ByVal hWnd As Long) As Long
    Set_Focus = SetFocusA(hWnd)
End Function
Public Function CloseWindow(ByVal hWnd As Long) As Boolean
    Dim TaskID As Long
    
    Call SendMessage(hWnd, WM_CLOSE, 0, 0)
    Call SendNotifyMessage(hWnd, WM_SYSCOMMAND, SC_CLOSE, 0)
        
    Call GetWindowThreadProcessId(hWnd, TaskID)
    
    CloseWindow = (Not TaskID)
    
End Function
Public Function isWindowVisible(ByVal hWnd As Long) As Boolean
    isWindowVisible = IsWindowVisibleA(hWnd)
End Function
Public Function isTopParent(ByVal hWnd As Long) As Boolean
    isTopParent = (GetParentA(hWnd) = 0)
End Function
Public Function FindWindowByClass(ByVal ClassName As String) As Long
    FindWindowByClass = FindWindowA(ClassName, vbNullString)
End Function
Public Function Get_Parent(ByVal hWnd As Long) As Long
    Get_Parent = GetParentA(hWnd)
End Function
Public Sub Set_Parent(ByVal hWnd As Long, _
                      ByVal NewParent As Long)
    Call SetParentA(hWnd, NewParent)
End Sub
Public Function FindWindowByCaption(ByVal Caption As String) _
                                        As Long
    FindWindowByCaption = FindWindowA(vbNullString, Caption)
End Function
Public Function Get_ActiveWindowTitle() As String
   Get_ActiveWindowTitle = Get_Text(GetForegroundWindow)
End Function
Public Function isHwnd(ByVal hWnd As Long) As Boolean
    isHwnd = Get_TaskID(hWnd)
End Function
Public Sub LockWindowUpdate(ByVal hWnd As Long, _
                            Optional UnlockWindow As Boolean = False)
    Dim wParam As Long
    
    wParam = IIf(UnlockWindow, 1, 0)
    Call SendMessage(hWnd, WM_SETREDRAW, wParam, 0)
    
End Sub
Public Function Get_Focus() As Long
    Get_Focus = GetFocus
End Function
Public Function tFade(ByVal hWnd As Long, _
                      ByVal FadeStart As Byte, _
                      ByVal FadeStop As Byte, _
                      Optional StartAtCurrentTransparenz As Boolean = False, _
                      Optional FadeStep As Integer = 5, _
                      Optional Speed As Integer = 3) _
                            As Boolean
    Dim I As Integer, T As Integer
    Dim tStart As Integer, tStop As Integer
  
    On Local Error GoTo Quit
    
    If StartAtCurrentTransparenz Then _
        FadeStart = Get_Transparenz(hWnd): _
        If FadeStart = 0 Then FadeStart = 254
        
    If FadeStart = 255 Then FadeStart = 254
    
    If FadeStop < FadeStart Then FadeStep = FadeStep * (-1)
    
    For I = FadeStart To FadeStop Step FadeStep
        Call DoTrans(hWnd, I)
                
        For T = 1 To (Speed * 1000) / 2
            DoEvents
        Next T
    Next I

    tFade = True

Quit:
End Function
Public Function Get_Transparenz(ByVal hWnd As Long) As Byte
    Call GetLayeredWindowAttributes(hWnd, CLng(0), _
                                    Get_Transparenz, CLng(0))
End Function
Public Sub Set_Transparenz(ByVal hWnd As Long, Transparenz As Byte)
    Call DoTrans(hWnd, Transparenz)
End Sub
Public Function Get_ChildWindows(ByVal hWnd As Long) As Long()
        
    Call EnumChild(hWnd)
    
    Get_ChildWindows = tChild
        
    Erase tChild

End Function
Public Function EXEmoduleByHwnd(ByVal hWnd As Long) As String
    Dim id As Long, I() As Long, P As Long
    Dim E() As String
        
    On Local Error GoTo Quit
    
    Call GetWindowThreadProcessId(hWnd, id)
    
    If SYS.Get_ProcessList(E, I) Then
        For P = 0 To UBound(I)
            If I(P) = id Then _
                EXEmoduleByHwnd = E(P): _
                Exit For
        Next P
    End If
        
Quit:
    Erase I: Erase E
    
End Function
Public Function Get_AllHwnd() As Long()
    Dim A() As Long
    
    On Local Error GoTo Quit
    
    Call EnumAll(A)
    
    Get_AllHwnd = A
    
Quit:
    Erase tTask: Erase tChild: Erase A
    
End Function
Public Sub SendKeys(ByVal hWnd As Long, _
                    ByVal Text As String)
    Dim P As Long, KeyCode As Long
    Dim R$
    
    For P = 1 To Len(Text)
        R$ = Mid$(Text, P, 1)
        KeyCode = Asc(R$)
    
        Call SendMessageByNum(hWnd, WM_CHAR, KeyCode, 0)
    Next P
    
End Sub
Public Function Get_WindowState(ByVal hWnd As Long) As VbAppWinStyle
    Dim W As WINDOWPLACEMENT

    If GetWindowPlacement(hWnd, W) = 0 Then
        Get_WindowState = -1
    Else
        Get_WindowState = W.showCmd
    End If

End Function
Public Function Get_WindowInfo(ByVal hWnd As Long) As String()
    Dim A() As String
    Dim F As Long
    
    ReDim A(9)
    
    A(0) = hWnd
    A(1) = EXEmoduleByHwnd(hWnd)
    A(2) = Get_Text(hWnd)
    A(3) = Get_Class(hWnd)
    A(4) = Get_TaskID(hWnd)
    A(5) = IsWindowEnabled(hWnd)
    A(6) = isWindowVisible(hWnd)
    A(7) = Get_Parent(hWnd)
    If A(7) = 0 Then A(7) = hWnd
    A(8) = Get_TopParent(hWnd)
    
    F = Get_Focus
    A(9) = IIf(F = hWnd, True, False)
    
    Get_WindowInfo = A
    
    Erase A
    
End Function
Public Function Get_AllHwndByClass(ByVal Class As String) As Long()
    Dim L() As Long, hWnd As Long
    
    hWnd = FindWindowA(Class, vbNullString)

    If hWnd = 0 Then Exit Function
    
    Call ARR.Add(L, hWnd, ADDTOEND)
    
    Do While hWnd <> 0
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        If Get_Class(hWnd) = Class Then Call ARR.Add(L, hWnd, ADDTOEND)
    Loop

    Get_AllHwndByClass = L
    
End Function
Public Function Get_AllHwndByText(ByVal Text As String, _
                                  Optional FindExact As Boolean = True) _
                                        As Long()
    Dim L() As Long, hWnd As Long, T$
    
    If Not FindExact Then Text = LCase$(Text)
    
    hWnd = FindWindowA(vbNullString, Text)

    If hWnd = 0 Then Exit Function
    
    Call ARR.Add(L, hWnd, ADDTOEND)
    
    Do While hWnd <> 0
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        T$ = Get_Text(hWnd)
        If Not FindExact Then T$ = LCase$(T$)
        
        If T$ = Text Then Call ARR.Add(L, hWnd, ADDTOEND)
    Loop

    Get_AllHwndByText = L
    
End Function
Public Function Get_TextSize(ByVal hWnd As Long, _
                             Width As Long, _
                             Height As Long) As Long
    Dim nText As Long, hObj As Long, hDC As Long, hFont As Long
    Dim cSize As SIZE
    Dim sText As String
    
    Const BUFFER_SIZE As Long = &HFF&
    
    Width = -1
    Height = -1
    
    sText = String(BUFFER_SIZE, 0)
    nText = GetWindowText(hWnd, sText, BUFFER_SIZE)
    
    If nText > 0 Then
        sText = Left$(sText, nText)
        hDC = GetDC(hWnd)
        
        If hDC <> 0 Then
            hFont = SendMessage(hWnd, WM_GETFONT, &H0&, &H0&)
                       
            If hFont <> 0 Then
                hObj = SelectObject(hDC, hFont)
                               
                If hObj > 0 Then _
                    If GetTextExtentPoint(hDC, sText, nText, cSize) <> 0 Then _
                        Width = cSize.cx: _
                        Height = cSize.cy: _
                        Get_TextSize = 1: _
                        Exit Function
                                
                ReleaseDC hWnd, hDC
            End If
        End If
    End If
    
    Get_TextSize = -1
    
End Function
Public Function Set_XPstyle(hWnd As Long, _
                            Optional ApplyStyle As Boolean = True) _
                                As Boolean
            
    If Not SYS.XPstyleSupported Then Set_XPstyle = False: _
                                     Exit Function
    
    On Local Error Resume Next

    If ApplyStyle Then
        Call ActivateWindowTheme(hWnd)
    Else
        Call DeactivateWindowTheme(hWnd)
    End If
        
    Set_XPstyle = True

End Function
