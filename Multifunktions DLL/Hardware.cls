VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hardware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type JOYCAPS
    wMid        As Integer
    wPid        As Integer
    szPname     As String * 32
    wXmin       As Long
    wXmax       As Long
    wYmin       As Long
    wYmax       As Long
    wZmin       As Long
    wZmax       As Long
    wNumButtons As Long
    wPeriodMin  As Long
    wPeriodMax  As Long
End Type

Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As Any, ByVal hWnd As Long, ByVal pPortName As String) As Long
Private Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Private Declare Function joyGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength        As Long
    dwMemoryLoad    As Long
    dwTotalPhys     As Long
    dwAvailPhys     As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual  As Long
    dwAvailVirtual  As Long
End Type

Private Type JOYINFO
    X       As Long
    Y       As Long
    Z       As Long
    Buttons As Long
End Type

Private lpInfo As MEMORYSTATUS

Public Enum PrtCommandNFO
    prtSTOP = 1
    prtRESUME = 2
    prtJOBCANCEL = 3
End Enum

Public Enum prtGetNFO
    PrtName = 1
    prtDRIVER = 2
    prtPORT = 3
End Enum

Private Type PRINTER_DEFAULTS
     pDatatype     As Long
     pDevMode      As Long
     DesiredAccess As Long
End Type

Private Const JOYERR_NOERROR = (0)

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000

Private Const PRINTER_ACCESS_ADMINISTER As Long = &H4
Private Const PRINTER_ACCESS_USE As Long = &H8
Private Const PRINTER_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Private Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Private Const PRINTER_CONTROL_PURGE As Long = 3&
Private Const PRINTER_CONTROL_PAUSE As Long = 1&
Private Const PRINTER_CONTROL_RESUME As Long = 2&

Private Const SC_MONITORPOWER = &HF170&
Public Function Printer_SetCommand(ByVal Command As PrtCommandNFO) _
                                        As Boolean
    Dim pName As String
    Dim hPrinter As Long
    Dim pd As PRINTER_DEFAULTS
  
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    pName = Printer.DeviceName
    
    If OpenPrinter(pName, hPrinter, pd) Then _
        If SetPrinter(hPrinter, 0, vbNull, Command) Then _
            Call ClosePrinter(hPrinter): _
            Printer_SetCommand = True
    
End Function
Public Function Printer_Default() As String()
    Dim Buffer As String, A(2) As String
    Dim R As Long, X As Long, Y As Long
    
    Buffer = Space(8192)
    
    R = GetProfileString("windows", "Device", "", _
                         Buffer, Len(Buffer))
    
    If R Then
        Buffer = Mid$(Buffer, 1, R)
        X = InStr(Buffer, ",")
        Y = InStr(X + 1, Buffer, ",")
        
        A(0) = Mid$(Buffer, 1, X - 1)
        A(1) = Mid$(Buffer, X + 1, Y - X - 1)
        A(2) = Mid$(Buffer, Y + 1)
    End If

End Function
Public Function Printer_List() As String()
    Dim P As Long
    Dim A() As String
    
    On Local Error GoTo Quit
    
    ReDim A(Printers.Count - 1)
    
    For P = 0 To Printers.Count - 1
        A(P) = Printers(P).DeviceName
    Next P

    Printer_List = A
    
Quit:
End Function
Public Function Printer_Select(ByVal PrinterName As String) As Boolean
    Dim P As Long

    For P = 0 To Printers.Count - 1
        If LCase(Printers(P).DeviceName) = LCase(PrinterName) Then _
            Set Printer = Printers(P): _
            Printer_Select = True: _
            Exit Function
    Next P
    
End Function
Public Function Printer_DeleteDocsFromQueue(Optional ByVal PrinterName As String) _
                                            As Boolean
    Dim Result As Long
    Dim hPrinter As Variant
    Dim udtPrinter As PRINTER_DEFAULTS
  
    If Len(PrinterName) = 0 Then _
        PrinterName = Printer.DeviceName

    udtPrinter.DesiredAccess = PRINTER_ALL_ACCESS

    Call OpenPrinter(PrinterName, hPrinter, udtPrinter)
    
    If hPrinter Then _
        Call SetPrinter(hPrinter, 0, vbNull, PRINTER_CONTROL_PURGE): _
        Printer_DeleteDocsFromQueue = True

    Call ClosePrinter(hPrinter)
    
End Function
Public Function Printer_isAvaible() As Boolean
    Dim A() As String
    
    On Local Error Resume Next
    
    A = Printer_Default
    
    Printer_isAvaible = (A(0) <> "")
    
End Function
Public Function COMport_Exist(ByVal Port As Long) As Boolean
    Dim F As Integer
    Dim sCOMParam As String

    On Local Error GoTo Quit

    sCOMParam = "COM" & CStr(Port) & ":9600,n,8,1"
    
    F = FreeFile
  
    Open sCOMParam For Input As #F
    Close #F
    
    COMport_Exist = True

Quit:
End Function
Public Function COMport_ConfigureDLG(ByVal Port As Long, _
                                     Optional hWnd As Long) _
                                            As Boolean
    Dim sPort As String
    
    On Local Error GoTo Quit
    
    If COMport_Exist(Port) Then
        sPort = "COM" & Port & ":"
        Call ConfigurePort(vbNullString, hWnd, sPort)
        COMport_ConfigureDLG = True
    End If
    
Quit:
End Function
Public Function COMport_GetInstalledCOMPorts() As Long()
    Dim Port As Long
    Dim A() As Long
    
    On Local Error Resume Next
    
    For Port = 1 To 16
        If COMport_Exist(Port) Then _
            Call ARR.Add(A, CStr(Port), ADDTOEND)
    Next Port
    
    COMport_GetInstalledCOMPorts = A
    
    Erase A
    
End Function
Public Function Keyboard_CAPSisEnabled() As Boolean
    Keyboard_CAPSisEnabled = (GetKeyState(vbKeyCapital) = 1)
End Function
Public Function Keyboard_NUMisEnabled() As Boolean
    Keyboard_NUMisEnabled = (GetKeyState(vbKeyNumlock) = 1)
End Function
Public Function CPUinfo() As String()
    Dim Key As String, A(3) As String
    Dim P As Long
    
    Key = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
    
    A(0) = REG.RegValueGet(HKEY_LOCAL_MACHINE_, Key, "ProcessorNameString")
    A(1) = REG.RegValueGet(HKEY_LOCAL_MACHINE_, Key, "Identifier")
    A(2) = REG.RegValueGet(HKEY_LOCAL_MACHINE_, Key, "VendorIdentifier")
    A(3) = REG.RegValueGet(HKEY_LOCAL_MACHINE_, Key, "~MHz") & " MHz"
    
    For P = 0 To UBound(A)
        A(P) = Trim$(A(P))
    Next P
    
    CPUinfo = A
    
End Function
Public Function CPUspeed() As Long
    
    CPUspeed = CLng(REG.RegValueGet(HKEY_LOCAL_MACHINE_, _
                               "Hardware\Description\System\CentralProcessor\0", _
                               "~MHz"))
End Function
Public Function FreeMem(Optional ShowProgress As Boolean = False) _
                            As Boolean
    Dim Phys As Long
    Dim frm As New ProgBar
    Dim I As Integer
    Dim At(20) As String

    On Local Error GoTo Quit

    If ShowProgress Then
        Load frm
        frm.Tag = "FreeMem"
        
        Call frm.SetOption(0, 20, "Arbeitsspeicher wird freigeräumt" & _
                           vbCrLf & "(Bitte warten)", True)
    End If
    
    Call GlobalMemoryStatus(lpInfo)

    Phys = lpInfo.dwTotalPhys / 60
    
    For I = 0 To 20
        If ShowProgress Then Call frm.SetValue(I)
        At(I) = Space$(Phys)
    Next I
    
    FreeMem = True
    
Quit:
    Erase At

    If ShowProgress Then _
        Unload frm: _
        Set frm = Nothing

End Function
Public Function Joystick_isPresent() As Boolean
    Joystick_isPresent = (Joystick_Count > 0)
End Function
Public Function Joystick_Count() As Long
    Dim I As Long, J As Long, P As Long, X As Long
    Dim JI As JOYINFO
    
    I = joyGetNumDevs + 1
    
    For P = I To 0 Step -1
        If joyGetPos(P, JI) = JOYERR_NOERROR Then X = X + 1
    Next P
    
    Joystick_Count = X
    
End Function
Public Function JoyStick_GetInfo(ByVal JoyStick As Long) As String()
    Dim A(6) As String
    Dim Jc As JOYCAPS
    Dim JI As JOYINFO
    
    If joyGetDevCaps(JoyStick, Jc, Len(Jc)) <> JOYERR_NOERROR Then _
        Exit Function
        
    A(2) = Jc.wXmax
    A(3) = Jc.wYmax
    A(4) = Jc.wZmax
    A(1) = Jc.wNumButtons
    
    If joyGetPos(JoyStick, JI) <> JOYERR_NOERROR Then Exit Function
    
    A(0) = JI.Buttons
    A(5) = JI.X
    A(6) = JI.Y
    
End Function
