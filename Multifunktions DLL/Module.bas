Attribute VB_Name = "Module"
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long

Private Win2Find As String
Private wHwnd As Long
Private Function WndEnumProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim nw As New Window
    
    If nw.GetText(hWnd) = Win2Find Then _
        wHwnd = hWnd: _
        WndEnumProc = 0: _
        Exit Function
        
    WndEnumProc = 1
            
End Function
Public Function FindEnum(ByVal ToFind As String) As Long
    
    Win2Find = ToFind
    wHwnd = 0
    
    Call EnumWindows(AddressOf WndEnumProc, CLng(0))

    FindEnum = wHwnd
    
End Function
