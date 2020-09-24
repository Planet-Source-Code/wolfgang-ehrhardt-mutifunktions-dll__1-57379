Attribute VB_Name = "Declar"
Public Const MAX_PATH = 260

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageB Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function RemoveMenuA Lib "user32" Alias "RemoveMenu" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwnewLong As Long) As Long
Public Declare Function GetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal lngX As Long, ByVal lngY As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetParentA Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function PlgBlt Lib "gdi32.dll" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal lpBuffer As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ActivateWindowTheme Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
Public Declare Function DeactivateWindowTheme Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long

'Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type SIZE
    cx As Long
    cy As Long
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type vLBNFO
    List() As String
End Type

Public Const CF_HDROP = 15

Public Const DT_CALCRECT = &H400

Public Const EM_GETLINECOUNT = &HBA

Public Const ERROR_SUCCESS = 0

Public Const GW_HWNDNEXT = 2

Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_STYLE = -16
Public Const GWL_WNDPROC = (-4&)

Public Const HWND_TOPMOST = -1

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_REMOVE = &H1000&

Public Const pi180 = 3.14159265358979 / 180

Public Const RGN_OR = 2

Public Const S_OK As Long = &H0

Public Const SRCCOPY = &HCC0020

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_FRAMECHANGED = &H20

Public Const SYNCHRONIZE = &H100000

Public Const WM_ACTIVATE = &H6
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_PASTE = &H302
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_GETFONT As Long = &H31&
Public Const WM_USER As Long = &H400

Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_RIGHT As Long = &H1000
Public Const WS_POPUP = &H80000000
Public Const WS_EX_TOPMOST As Long = &H8&

Public Const WS_THICKFRAME = &H40000
Public Const WS_CAPTION = &HC00000
Public Const WS_VISIBLE = &H10000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_BORDER = &H800000

Public Const lABC As String = "abcdefghijklmnopqrstuvwxyz"
Public Const uABC As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const Num As String = "1234567890"

Public ARR   As New WindowAPI.Arry
Public CRTL  As New WindowAPI.Controls
Public DESK  As New WindowAPI.Desktop
Public DLG   As New WindowAPI.Dialog
Public DI    As New WindowAPI.Directory
Public DRV   As New WindowAPI.Drive
Public F     As New WindowAPI.File
Public FRMS  As New WindowAPI.Forms
'Public FTP  As New WindowAPI.FTP
Public GRAFX As New WindowAPI.Grafic
'Public HW   As New WindowAPI.Hardware
Public INI   As New WindowAPI.INI
Public LB    As New WindowAPI.LB_CB
Public MA    As New WindowAPI.Mathematics
Public MNU   As New WindowAPI.Menu
Public misc  As New WindowAPI.misc
Public MOUSE As New WindowAPI.MOUSE
Public Str   As New WindowAPI.mString
Public NET   As New WindowAPI.Network
Public REG   As New WindowAPI.Registry
'Public RTF  As New WindowAPI.RTF
Public SH    As New WindowAPI.Shell
'Public MM    As New WindowAPI.Multimedia
Public SYS   As New WindowAPI.System
Public TXT   As New WindowAPI.Text
Public WIN   As New WindowAPI.Window

Public CBacWork As Boolean
