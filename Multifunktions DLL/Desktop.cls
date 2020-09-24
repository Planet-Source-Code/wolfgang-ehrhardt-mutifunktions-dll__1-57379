VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Desktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SHAddToRecentDocs Lib "shell32.dll" (ByVal dwFlags As Long, ByVal dwData As String) As Long
Private Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnableTheming Lib "UxTheme.dll" (ByVal fEnable As Boolean) As Long
Private Declare Function IsThemeActive Lib "UxTheme.dll" () As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As gRECT) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SystemParametersInfo2 Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function PathToRegion Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function TextOutW Lib "gdi32.dll" (ByVal hDC As Long, ByVal nXStart As Long, ByVal nYStart As Long, ByVal lpString As Long, ByVal cbString As Long) As Long
Private Declare Function GetTextExtentPoint32W Lib "gdi32.dll" (ByVal hDC As Long, ByVal lpString As Long, ByVal cbString As Long, ByRef lpSize As SIZE) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Type gRECT
    Left   As Long
    Top    As Long
    Width  As Long
    Height As Long
End Type

Private Type APPBARDATA
    cbSize           As Long
    hWnd             As Long
    uCallbackMessage As Long
    uEdge            As Long
    rc               As RECT
    lParam           As Long
End Type

Private Type ULARGE_INTEGER
    LowPart  As Long
    HighPart As Long
End Type

Private Type SHQUERYRBINFO
    cbSize      As Long
    i64Size     As ULARGE_INTEGER
    i64NumItems As ULARGE_INTEGER
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Public Enum PanelOptionNFO
    poControlPanel = 1
    poRemove_Software = 2
    poAdd_Software = 3
    poAdd_Remove_WindowsComponents = 4
    poSystemInformation = 5
    poComputerNameSettings = 6
    poHardwareManager = 7
    poTime_Date_Settings = 8
    poRegional_Settings = 9
    poDisplay_Settings = 10
    poInternet_Settings = 11
    poKeyboard_Settings = 12
    poMouse_Settings = 13
    poModem_Settings = 14
    poSound_Settings = 15
    poNew_Printer = 16
    poInput_Assistance = 17
    poCountrie_Attitudes = 18
    poODCB_Settings = 19
    poGame_Settings = 20
    poWindows_Update = 21
End Enum

Public Enum SysDirNfo
    sdMYCOMPUTER = 1
    sdRECYCLEBIN = 2
    sdSYSTEMPANEL = 3
    sdNETWORKCONNECTIONS = 4
    sdNETWORKENVIRONMENT = 5
    sdPRINTER = 6
End Enum

Public Enum dtHwnds
    dtTRAY = 1
    dtDESKTOP = 2
    dtSTARTBUTTON = 3
    dtTASKBAR = 4
    dtCLOCK = 5
    dtSTARTMENU = 6
    dtINTERNETEXPLORER = 7
    dtEXPLORER = 8
    dtOUTLOOK = 9
    dtOUTLOOKEXPRESS = 10
    dtWORD = 11
    dtNOTEPAD = 12
    dtCALCULATOR = 13
    dtWORDPAD = 14
    dtACCESS = 15
End Enum

Public Enum DefaultCusor
    OCR_CROSS = 32515
    OCR_HAND = 32649
    OCR_IBEAM = 32513
    OCR_NO = 32648
    OCR_NORMAL = 32512
    OCR_SIZEALL = 32646
    OCR_SIZENESW = 32643
    OCR_SIZENS = 32645
    OCR_SIZEWE = 32644
    OCR_SIZENWSE = 32642
    OCR_UP = 32516
    OCR_WAIT = 32514
End Enum

Public Enum RecentDocsCommand
    rdADDFILE = 1
    rdCLEAR = 2
End Enum

Public Enum DESKTOPICON
    diRECYCLEBIN = 1
    diMYCOMPUTER = 2
    diSYSTEMPANEL = 3
    diNETNEIGHBOURHOOD = 4
End Enum

Private Const ABM_GETTASKBARPOS = &H5

Private Const SC_SCREENSAVE = 61760

Private Const SHARD_PATH = &H2&

Private Const SHERB_NOPROGRESSUI = &H2
Private Const SHERB_NOCONFIRMATION = &H1
Private Const SHERB_NOSOUND = &H4

Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPI_GETWORKAREA = 48

Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1

Private Const KEYEVENTF_KEYUP = &H2

Private Const VK_WIN = &H5B
Private Const VK_F5 As Long = &H74

Private Const WM_KEYDOWN As Long = &H100
Public Function RecycleBin_isEmpty() As Boolean
    RecycleBin_isEmpty = IIf(RecycleBin_ItemCount = 0, True, False)
End Function
Public Function RecycleBin_ItemCount() As Long
    Dim RBinInfo As SHQUERYRBINFO
    
    RBinInfo.cbSize = Len(RBinInfo)
    
    Call SHQueryRecycleBin(vbNullString, RBinInfo)

    RecycleBin_ItemCount = RBinInfo.i64NumItems.LowPart
    
End Function
Public Function RecycleBin_Size() As Long
    Dim RBinInfo As SHQUERYRBINFO
    
    RBinInfo.cbSize = Len(RBinInfo)
    
    Call SHQueryRecycleBin(vbNullString, RBinInfo)
    RecycleBin_Size = RBinInfo.i64Size.LowPart
    
End Function
Public Function RecycleBin_Clear(ByVal hWnd As Long, _
                                 Optional Confirm As Boolean = False, _
                                 Optional ShowProgress As Boolean = False, _
                                 Optional PlaySound As Boolean = False) _
                                        As Boolean
    Dim dwFlags As Long
    
    If Not Confirm Then dwFlags = dwFlags Or SHERB_NOCONFIRMATION
    If Not ShowProgress Then dwFlags = dwFlags Or SHERB_NOPROGRESSUI
    If Not PlaySound Then dwFlags = dwFlags Or SHERB_NOSOUND
                        
    If SHEmptyRecycleBin(hWnd, vbNullString, dwFlags) = S_OK Then _
        RecycleBin_Clear = True: _
        Call RecycleBin_UpdateIcon
    
End Function
Public Sub RecycleBin_UpdateIcon()
    Call SHUpdateRecycleBinIcon
End Sub
Public Sub Get_DesktopDimensions(dHeight As Long, _
                                 dWidth As Long)
    dWidth = Screen.Width / Screen.TwipsPerPixelX
    dHeight = Screen.Height / Screen.TwipsPerPixelX
End Sub
Public Sub Get_WorkAreaDimensions(wHeight As Long, _
                                  wWidth As Long)
    Dim R As RECT

    Call SystemParametersInfo2(SPI_GETWORKAREA, 0, R, 0)

    wWidth = (R.Right - R.Left) * Screen.TwipsPerPixelX
    wHeight = (R.Bottom - R.Top) * Screen.TwipsPerPixelY

End Sub
Public Sub RecentDocs(Command As RecentDocsCommand, _
                      Optional File As String)
    
    Select Case Command
        Case rdADDFILE
            Call SHAddToRecentDocs(SHARD_PATH, File)
        Case rdCLEAR
            Call SHAddToRecentDocs(SHARD_PATH, vbNullString)
    End Select
    
End Sub
Public Sub Get_TaskbarPos(Left As Long, _
                          Top As Long, _
                          Right As Long, _
                          Bottom As Long)
    Dim TPos As APPBARDATA
    
    Call SHAppBarMessage(ABM_GETTASKBARPOS, TPos)
    
    Left = TPos.rc.Left
    Top = TPos.rc.Top
    Right = TPos.rc.Right
    Bottom = TPos.rc.Bottom

End Sub
Public Function Wallpaper_Change(ByVal File As String) _
                                    As Boolean
    
    If SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, File, _
                            SPIF_UPDATEINIFILE Or _
                            SPIF_SENDWININICHANGE) Then _
                                Wallpaper_Change = True

End Function
Public Function Wallpaper_Clear() As Boolean
    Wallpaper_Clear = Wallpaper_Change("")
End Function
Public Sub ScreenSaver_Activate(Form As Object)
    On Local Error Resume Next
    Call SendMessage(Form.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
Public Function SystemPanel(Run As PanelOptionNFO) As Boolean
    Dim R$
    
    Select Case Run
        Case poControlPanel
            R$ = DESK.SystemDirOpen(sdSYSTEMPANEL)
            Exit Function
        Case poRemove_Software
            R$ = " appwiz.cpl,,0" & Chr$(34) & ", 1"
        Case poAdd_Software
            R$ = " appwiz.cpl,,1" & Chr$(34) & ", 1"
        Case poAdd_Remove_WindowsComponents
            R$ = " appwiz.cpl,,2" & Chr$(34) & ", 1"
        Case poSystemInformation
            R$ = " sysdm.cpl,,0" & Chr$(34) & ", 1"
        Case poComputerNameSettings
            R$ = " sysdm.cpl,,1" & Chr$(34) & ", 1"
        Case poHardwareManager
            R$ = " sysdm.cpl,,2" & Chr$(34) & ", 1"
        Case poTime_Date_Settings
            R$ = " timedate.cpl" & Chr$(34) & ", 1"
        Case poRegional_Settings
            R$ = " intl.cpl,,0" & Chr$(34) & ", 1"
        Case poDisplay_Settings
            R$ = " desk.cpl,,0" & Chr$(34) & ", 1"
        Case poInternet_Settings
            R$ = " inetcpl.cpl,,0" & Chr$(34) & ", 1"
        Case poKeyboard_Settings
            R$ = " main.cpl @1" & Chr$(34) & ", 1"
        Case poMouse_Settings
            R$ = " main.cpl @0" & Chr$(34) & ", 1"
        Case poModem_Settings
            R$ = " modem.cpl" & Chr$(34) & ", 1"
        Case poSound_Settings
            R$ = " mmsys.cpl @1" & Chr$(34) & ", 5"
        Case poNew_Printer
            SystemPanel = Shell("rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", 1)
            Exit Function
        Case poInput_Assistance
            R$ = " access.cpl @1" & Chr$(34) & ", 1"
        Case poCountrie_Attitudes
            R$ = " intl.cpl,,0" & Chr$(34) & ", 1"
        Case poODCB_Settings
            R$ = " odbccp32.cpl @1" & Chr$(34) & ", 1"
        Case poGame_Settings
            R$ = " joy.cpl @1" & Chr$(34) & ", 1"
        Case poWindows_Update
            R$ = " wuaucpl.cpl @1" & Chr$(34) & ", 1"
        Case Else
            Exit Function
    End Select
    
    If Shell("rundll32.exe shell32.dll,Control_RunDLL" & R$) Then _
        SystemPanel = True
    
End Function
Public Sub StartButton_Click()
    Call WIN.Click(DesktopHwnd(dtSTARTBUTTON))
End Sub
Public Function Wallpaper_GetPath() As String
        
    Wallpaper_GetPath = REG.RegValueGet(HKEY_CURRENT_USER_, _
                                        "Control Panel\Desktop", _
                                        "OriginalWallpaper")
        
    If Wallpaper_GetPath = "" Then _
        Wallpaper_GetPath = REG.RegValueGet(HKEY_CURRENT_USER_, _
                                            "Control Panel\Desktop", _
                                            "Wallpaper")
End Function
Public Function Set_WindowsDefaultMouseCursor(MP As DefaultCusor) _
                                    As Boolean
    Dim hcursor As Long
    Static oMP As DefaultCusor
    
    If oMP Then _
        hcursor = LoadCursor(0, oMP): _
        Call SetSystemCursor(hcursor, OCR_NORMAL)

    oMP = MP
    hcursor = LoadCursor(0, MP)
    
    If SetSystemCursor(hcursor, OCR_NORMAL) Then _
        Set_WindowsDefaultMouseCursor = True

End Function
Public Function SysTray_Refresh() As Boolean
    Dim X As Long, Z As Long, hWnd As Long
    Dim rRect As RECT
    
    On Local Error GoTo Quit

    hWnd = FindSysTray
    
    If hWnd = 0 Then Exit Function
    
    X = GetClientRect(hWnd, rRect)

    For Z = 0 To (rRect.Bottom / 20) - 1
        For X = 0 To rRect.Right
            Call SendMessage(hWnd, WM_MOUSEMOVE, 0, _
                             ByVal ((((Z * 20) + 10) * &H10000) Or X))
        Next X
    Next Z
    
    SysTray_Refresh = True

Quit:
End Function
Public Function Windows_Minimize(Optional RestoreWindows As Boolean = False) _
                                        As Boolean
    Dim hWnd As Long, wMsg As Long
        
    wMsg = IIf(RestoreWindows, 416, 419)
    hWnd = WIN.FindWindowByClass("Shell_TrayWnd")
    
    If hWnd Then _
        If PostMessage(hWnd, WM_COMMAND, wMsg, 0&) Then _
            Windows_Minimize = True
        
End Function
Public Function DesktopHwnd(DeskHwnd As dtHwnds) As Long
    Dim Child As Long, Parent As Long
    
    Select Case DeskHwnd
        Case dtSTARTBUTTON
            Parent = FindWindowA("Shell_TrayWnd", vbNullString)
            Child = FindWindowEx(Parent, 0, "Button", vbNullString)
            
            DesktopHwnd = Child
        Case dtDESKTOP
            Parent = FindWindowA("Progman", vbNullString)
            Child = FindWindowEx(Parent, 0, "SHELLDLL_DefView", vbNullString)
            Child = FindWindowEx(Child, 0, "SysListView32", vbNullString)
            
            DesktopHwnd = Child
        Case dtTRAY
            Parent = FindWindowA("Shell_TrayWnd", vbNullString)
            Child = FindWindowEx(Parent, 0, "TrayNotifyWnd", vbNullString)

            DesktopHwnd = Child
        Case dtTASKBAR
            Parent = FindWindowA("Shell_TrayWnd", vbNullString)
            Child = FindWindowEx(Parent, 0, "ReBarWindow32", vbNullString)
            Child = FindWindowEx(Child, 0, "MSTaskSwWClass", vbNullString)
            Child = FindWindowEx(Child, 0, "ToolbarWindow32", vbNullString)

            DesktopHwnd = Child
        Case dtCLOCK
            Parent = FindWindowA("Shell_TrayWnd", vbNullString)
            Child = FindWindowEx(Parent, 0, "TrayNotifyWnd", vbNullString)
            Child = FindWindowEx(Child, 0, "TrayClockWClass", vbNullString)
        
            DesktopHwnd = Child
        Case dtSTARTMENU
            Parent = FindWindowA("BaseBar", vbNullString)
            Parent = GetWindow(Parent, GW_HWNDNEXT)
            Parent = GetWindow(Parent, GW_HWNDNEXT)
            Child = FindWindowEx(Parent, 0, "MenuSite", vbNullString)
            Child = FindWindowEx(Child, 0, "ToolbarWindow32", vbNullString)
            
            DesktopHwnd = Child
        Case dtINTERNETEXPLORER
            DeskHwnd = FindWindowA("IEFrame", vbNullString)
        Case dtEXPLORER
            DesktopHwnd = FindWindowA("ExploreWClass", vbNullString)
        Case dtOUTLOOK
            DesktopHwnd = FindWindowA("rctrl_renwnd32", vbNullString)
        Case dtOUTLOOKEXPRESS
            DesktopHwnd = FindWindowA("Outlook Express Browser Class", vbNullString)
        Case dtWORD
            DesktopHwnd = FindWindowA("OpusApp", vbNullString)
        Case dtNOTEPAD
            DesktopHwnd = FindWindowA("Notepad", vbNullString)
        Case dtCALCULATOR
            DesktopHwnd = FindWindowA("SciCalc", vbNullString)
        Case dtWORDPAD
            DesktopHwnd = FindWindowA("WordPadClass", vbNullString)
        Case dtACCESS
            DesktopHwnd = FindWindowA("OMain", vbNullString)
        Case Else: Exit Function
    End Select
    
End Function
Public Function SystemDirOpen(SystemDir As SysDirNfo) As Boolean
    Dim R$
    
    Select Case SystemDir
        Case sdMYCOMPUTER
            R$ = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        Case sdNETWORKENVIRONMENT
            R$ = "{208D2C60-3AEA-1069-A2D7-08002B30309D}"
        Case sdNETWORKCONNECTIONS
            R$ = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" & _
                 "\::{21EC2020-3AEA-1069-A2DD-08002B30309D}" & _
                 "\::{7007acc7-3202-11d1-aad2-00805fc1270e}"
        Case sdRECYCLEBIN
            R$ = "{645FF040-5081-101B-9F08-00AA002F954E}"
        Case sdSYSTEMPANEL
            R$ = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" & _
                 "\::{21EC2020-3AEA-1069-A2DD-08002B30309D}"
        Case sdPRINTER
            R$ = "{2227A280-3AEA-1069-A2DE-08002B30309D}"
        Case Else: Exit Function
    End Select
    
    If Shell("Explorer.exe ::" & R$, vbNormalFocus) Then _
        SystemDirOpen = True
    
End Function
Public Function isXPthemeActive() As Boolean
    On Local Error Resume Next
    isXPthemeActive = IsThemeActive
End Function
Public Function XPthemeEnable(Optional Disable As Boolean = False) _
                                    As Boolean
    On Local Error Resume Next
    XPthemeEnable = (EnableTheming(IIf(Disable, False, True)) = 0)
End Function
Public Function SnapShot() As StdPicture
    Dim hDC As Long, hWnd As Long, X As Long, Y As Long
    Dim WIN As gRECT
    Dim frm As New MyControls

    On Local Error GoTo Quit
    
    Load frm
    frm.Tag = "SnapShot"
    
    frm.PicTMP.AutoRedraw = True
    frm.PicTMP.AutoSize = False
    
    hWnd = GetDesktopWindow()
    hDC = GetDC(hWnd)
    
    Call GetWindowRect(hWnd, WIN)
    
    X = WIN.Width
    Y = WIN.Height
    
    Call GRAFX.PixelToTwips(X, Y)

    frm.PicTMP.Width = X
    frm.PicTMP.Height = Y
    
    frm.PicTMP.ZOrder
    
    Call StretchBlt(frm.PicTMP.hDC, WIN.Left, WIN.Top, _
                    WIN.Width, WIN.Height, hDC _
                    , 0, 0, WIN.Width, WIN.Height, _
                    SRCCOPY)
                    
    Call ReleaseDC(hWnd, hDC)
         
    Set frm.PicTMP.Picture = frm.PicTMP.Image
    frm.PicTMP.Refresh

    Set SnapShot = frm.PicTMP.Picture
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Sub Show()

    keybd_event VK_WIN, 1, 0, 0
    keybd_event Asc("D"), 1, 0, 0

    keybd_event Asc("D"), 1, KEYEVENTF_KEYUP, 0
    keybd_event VK_WIN, 1, KEYEVENTF_KEYUP, 0

End Sub
Public Function WriteOnDesktop(ByVal Text As String, _
                               ByVal Top As Long, Left As Long, _
                               Optional Center As Boolean, _
                               Optional StayOnTop As Boolean = True, _
                               Optional FontName As String = "Times New Roman", _
                               Optional FontSize As Long = 36, _
                               Optional FontColor As Long = vbRed, _
                               Optional FontBold As Boolean = False, _
                               Optional FontItalic As Boolean = False, _
                               Optional FontUnderline As Boolean = False, _
                               Optional FontStrikethrough As Boolean = False, _
                               Optional Clear As Boolean = False) _
                                    As Boolean
    Dim P As Long, W As Long, H As Long, lH As Long, lW As Long, X As Long
    Dim Pointer As Long, Length As Long, Device As Long, Region As Long
    Dim TextSize As SIZE
    Dim frm As New frmTMP
    
    Static fText() As String
    Static fFrm() As Form
    
    On Local Error GoTo Quit
    
    If Clear Then
        If Not isArray(fText) Then Exit Function
        
        For P = 0 To UBound(fText)
            If fText(P) = Text Then Unload fFrm(P): _
                                    Set fFrm(P) = Nothing: _
                                    Call ARR.DeleteIndex(fFrm, P): _
                                    Call ARR.DeleteIndex(fText, P): _
                                    Exit For
        Next P
    Else
        Load frm
        
        With frm
            If StayOnTop Then Call FRMS.StayOnTop(frm)
            
            .BackColor = FontColor
            .Font.Name = FontName
            .Font.SIZE = FontSize
            .Font.Bold = FontBold
            .Font.Italic = FontItalic
            .Font.Underline = FontUnderline
            .Font.Strikethrough = FontStrikethrough
            
            .LabelTMP.AutoSize = True
            .LabelTMP.Caption = Text
            .LabelTMP.Font.Name = FontName
            .LabelTMP.Font.SIZE = FontSize
            .LabelTMP.Font.Bold = FontBold
            .LabelTMP.Font.Italic = FontItalic
            .LabelTMP.Font.Underline = FontUnderline
            .LabelTMP.Font.Strikethrough = FontStrikethrough
            
            .Width = .LabelTMP.Width * vbTwips
            .Height = .LabelTMP.Height * vbTwips
            
            If Center Then Call Get_WorkAreaDimensions(H, W): _
                           Top = (H / 2) - (.Height / 2): _
                           Left = (W / 2) - (.Width / 2)
            
            .Top = Top
            .Left = Left
            
            Pointer = strptr(Text)
            Length = Len(Text)
            Device = .hDC
            
            GetTextExtentPoint32W Device, Pointer, Length, TextSize

            BeginPath Device
            TextOutW Device, 0, 0, Pointer, Length
            EndPath Device
            Region = PathToRegion(Device)

            SetWindowRgn .hWnd, Region, 1
        
            .Visible = True
        
            Call ARR.Add(fText, Text, ADDTOEND)
            
            On Local Error Resume Next
            
            X = UBound(fFrm)
            
            If Err.Number <> 0 Then
                ReDim fFrm(0)
            Else
                ReDim Preserve fFrm(UBound(fFrm) + 1)
            End If
            
            Set fFrm(UBound(fFrm)) = frm
        End With
    End If

    WriteOnDesktop = True
    
Quit:
End Function
Public Function DesktopIconsShow(Optional Hide As Boolean = False) _
                                    As Boolean
    Dim hWnd As Long, mCmdShow As Long
    
    mCmdShow = IIf(Hide, SW_HIDE, SW_RESTORE)
    hWnd = FindWindowA(vbNullString, "Program Manager")
  
    If hWnd <> 0 Then Call ShowWindow(hWnd, mCmdShow): _
                      DesktopIconsShow = True
  
End Function
Public Function RenameDesktopIcon(iCon As DESKTOPICON, _
                                  ByVal NewName As String) As Boolean
    Dim Path  As String
    Dim hWnd As Long
    
    Select Case iCon
        Case diMYCOMPUTER
            Path = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        Case diNETNEIGHBOURHOOD
            Path = "{208D2C60-3AEA-1069-A2D7-08002B30309D}"
        Case diRECYCLEBIN
            Path = "{645FF040-5081-101B-9F08-00AA002F954E}"
        Case diSYSTEMPANEL
            Path = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    End Select
    
    Path = "CLSID\" & Path
    
    RenameDesktopIcon = REG.RegValueSet(HKEY_CLASSES_ROOT_, _
                                        Path, "", _
                                        NewName, True)
                                        
    Call Update
                            
End Function
Public Sub Update()
    Dim Result As Long
    
    Result = InvalidateRect(0&, 0&, False)
    If Result <> 0 Then Call InvalidateRect(0&, 0&, False)
  
End Sub
