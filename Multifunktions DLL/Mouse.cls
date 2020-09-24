VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowCursorA Lib "user32" Alias "ShowCursor" (ByVal bShow As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Function Clip_Cursor Lib "user32" Alias "ClipCursor" (lpRect As Any) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Private Const SM_CMOUSEBUTTONS = 43

Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4
Public Sub Set_MousePos(ByVal X As Long, ByVal Y As Long)
    Call SetCursorPos(X, Y)
End Sub
Public Sub Click_Left(Optional DoubleClick As Boolean = False)
    Dim P As Long
    
    For P = 1 To IIf(DoubleClick, 2, 1)
        Call mouse_event(MOUSEEVENTF_LEFTDOWN Or _
                         MOUSEEVENTF_LEFTUP, _
                         0&, 0&, CLng(0), CLng(0))
    Next P

End Sub
Public Sub Click_Right()
    Call mouse_event(MOUSEEVENTF_RIGHTDOWN Or _
                     MOUSEEVENTF_RIGHTUP, _
                     0&, 0&, CLng(0), CLng(0))
End Sub
Public Sub Show_Mouse(Optional HideMouse As Boolean = False)
    Call ShowCursorA(IIf(HideMouse, 0, 1))
End Sub
Public Function Get_UnderMouseHwnd() As Long
    Dim nPoint As POINTAPI
    
    Call GetCursorPos(nPoint)
    Get_UnderMouseHwnd = WindowFromPoint(nPoint.X, nPoint.Y)

End Function
Public Function Get_UnderMouseColor() As Long
    Dim Hdcp As Long, hWnd As Long
    Dim nPoint As POINTAPI
    
    Call GetCursorPos(nPoint)
    
    hWnd = WindowFromPoint(nPoint.X, nPoint.Y)
    Hdcp = GetDC(hWnd)
    
    Call ScreenToClient(hWnd, nPoint)
    Get_UnderMouseColor = GetPixel(Hdcp, nPoint.X, nPoint.Y)

End Function
Public Sub Get_MousePos(X As Long, Y As Long)
    Dim Pa As POINTAPI
    
    Call GetCursorPos(Pa)
    
    X = Pa.X
    Y = Pa.Y
    
End Sub
Public Function Get_UnderMouseText() As String
    
    Get_UnderMouseText = WIN.Get_Text(Get_UnderMouseHwnd)
    
End Function
Public Function Get_UnderMouseClass() As String
    
    Get_UnderMouseClass = WIN.Get_Class(Get_UnderMouseHwnd)
    
End Function
Public Function isMouseOverHwnd(ByVal hWnd As Long) As Boolean
    isMouseOverHwnd = (Get_UnderMouseHwnd = hWnd)
End Function
Public Function isUserClick(ByVal hWnd As Long) As Boolean
    Dim nPoint As POINTAPI
    
    Call GetCursorPos(nPoint)
    
    If GetAsyncKeyState(VK_LBUTTON) _
    And WindowFromPoint(nPoint.X, nPoint.Y) = hWnd Then _
            isUserClick = True

End Function
Public Function isMouseOverControl(Control As Object) _
                                        As Boolean
    Dim P1 As POINTAPI, P2 As POINTAPI, Cur As POINTAPI

    On Local Error GoTo Quit

    With Control.Parent
        P1.X = .ScaleX(Control.Left, .ScaleMode, vbPixels)
        P1.Y = .ScaleY(Control.Top, .ScaleMode, vbPixels)
        
        Call ClientToScreen(.hWnd, P1)
    
        P2.X = .ScaleX(Control.Left + Control.Width, .ScaleMode, vbPixels)
        P2.Y = .ScaleY(Control.Top + Control.Height, .ScaleMode, vbPixels)
    
        Call ClientToScreen(.hWnd, P2)
    End With
    
    Call GetCursorPos(Cur)
    
    If (Cur.X >= P1.X And Cur.X <= P2.X) _
    And (Cur.Y >= P1.Y And Cur.Y <= P2.Y) Then _
        isMouseOverControl = True
    
Quit:
End Function
Public Function Get_ButtonStatus() As MouseButtonConstants
    
    If GetAsyncKeyState(VK_LBUTTON) Then
        Get_ButtonStatus = vbLeftButton
    ElseIf GetAsyncKeyState(VK_RBUTTON) Then
            Get_ButtonStatus = vbRightButton
        ElseIf GetAsyncKeyState(VK_MBUTTON) Then _
                Get_ButtonStatus = vbMiddleButton
    End If
    
End Function
Public Sub BlockUserInput(Optional UnBlock As Boolean = False)
    Dim fBlock As Long
    
    fBlock = IIf(UnBlock, 0, 1)
    Call BlockInput(fBlock)
    
End Sub
Public Function Get_MouseButtons() As Long
    Get_MouseButtons = GetSystemMetrics(SM_CMOUSEBUTTONS)
End Function
Public Function Center(ByVal hWnd As Long, _
                       Optional Ctrl As Object) As Boolean
    Dim P As POINTAPI
    Dim R As RECT
  
    On Local Error GoTo Quit
    
    If hWnd = 0 Then
        Call GetWindowRect(Ctrl.hWnd, R)
    Else
        Call GetWindowRect(hWnd, R)
    End If
    
    P.X = R.Left + (R.Right - R.Left) / 2
    P.Y = R.Top + (R.Bottom - R.Top) / 2
    
    Call SetCursorPos(P.X, P.Y)
    
    Center = True
     
Quit:
End Function
Public Function ClipCursor(hWnd As Long, _
                           Optional UnClip As Boolean = False) _
                                As Boolean
    Dim R As RECT
   
    On Local Error GoTo Quit
   
    If UnClip Then
        ClipCursor = Clip_Cursor(ByVal vbNullString)
    Else
        Call GetWindowRect(hWnd, R)
        ClipCursor = Clip_Cursor(R)
    End If
       
Quit:
End Function
Public Function isMoving() As Boolean
    Dim P1 As POINTAPI, P2 As POINTAPI
    
    Static P3 As POINTAPI
    
    If P3.X = 0 And P3.Y = 0 Then
        Call Get_MousePos(P1.X, P1.Y)
    Else
        P1.X = P3.X
        P1.Y = P3.Y
    End If
    
    Call Get_MousePos(P2.X, P2.Y)
        
    isMoving = ((P1.X <> P2.X) Or (P1.Y <> P2.Y))
    
    P3.X = P2.X
    P3.Y = P2.Y
    
End Function
