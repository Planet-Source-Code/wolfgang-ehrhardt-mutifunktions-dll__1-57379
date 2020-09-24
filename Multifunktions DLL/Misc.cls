VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub SleepA Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Public Enum msSTPNFO
    msSTOP = 0
    msSTART = 1
    msRESET = 2
End Enum

Public Enum PassCommand
    psCREATEPASSWORD = 1
    psQUERYPASSWORD = 2
    psCHANGEPASSWORD = 3
End Enum

Public Enum STPNFO
    stSTOP = 0
    stSTART = 1
    stCURRENTSTOPTIME = 2
    stSTARTTIME = 3
    stRESET = 4
End Enum

Public Enum DateType
    dtTIMEONLY = 0
    dtDATEONLY = 1
    dtTIMEDATE = 2
    dtUNKOWN = -1
End Enum

Public Enum WinSymbol
    wsArrow_Left = 51
    wsArrow_Right = 52
    wsArrow_Up = 53
    wsArrow_Down = 54
    wsArrow_Spin = 118
    wsArrow_Bottom = 55
    wsArrow_Small_Left = 119
    wsArrow_Small_Right = 56
    wsArrow_Small_Up = 116
    wsArrow_Small_Down = 56
    wsSymbol_Check = 97
    wsSymbol_Bullet = 104
    wsSymbol_Small_Bullet = 105
    wsSymbol_Close = 114
    wsSymbol_Help = 115
    wsSymbol_Minimize = 48
    wsSymbol_Maximize = 49
    wsSymbol_Restore = 50
End Enum

Public Enum BS_Alignment
    bsTOP = &H400
    bsVCENTER = &HC00
    bsBOTTOM = &H800
    bsLEFT = &H100
    bsCENTER = &H300
    bsRIGTH = &H200
End Enum

Public Enum tfFORMATTYPE
    tfDAY = 0
    tfHOUR = 1
    tfMINUTE = 2
    tfSECOND = 3
End Enum

Private Type LARGE_INTEGER
    Lo As Long
    Hi As Long
End Type

Dim Strt As LARGE_INTEGER
Dim Ende As LARGE_INTEGER
Dim Freq As LARGE_INTEGER

Dim Calibr As Double
Public Function Get_Randomize(ByVal Min As Long, _
                              ByVal Max As Long) As Long
    Randomize
    Get_Randomize = Rnd * (Max - Min) + Min
    
End Function
Public Function Get_RandomTrueFalse() As Boolean
    Get_RandomTrueFalse = Get_Randomize(0, 1)
End Function
Public Sub Sleep(ByVal MilliSeconds As Long)
    Call SleepA(MilliSeconds)
End Sub
Public Function Get_WeekDay(datDate As Date) As String
    Get_WeekDay = WeekdayName(Weekday(datDate, vbMonday))
End Function
Public Sub About()
    frmAbout.Show
End Sub
Public Function Convert_DateToUnix(ByVal uDate As Date, _
                                   ByVal uTime As Date) _
                                        As String
    Dim DateTimeEnd As Date
    
    On Local Error Resume Next
    
    DateTimeEnd = FormatDateTime(uDate & " " & uTime)
    
    Convert_DateToUnix = DateDiff("s", "01.01.1970", DateTimeEnd, _
                                  vbUseSystemDayOfWeek, vbUseSystem)

    If Err.Number <> 0 Then Convert_DateToUnix = ""

End Function
Public Function Convert_UnixToDate(ByVal UnixDate As Double) _
                                        As Date
    Dim vStart As Variant
  
    On Local Error GoTo Quit
  
    vStart = DateSerial(1970, 1, 1)
    Convert_UnixToDate = DateAdd("s", UnixDate, vStart)

Quit:
End Function
Private Sub Calibrieren()
    Call QueryPerformanceFrequency(Freq)
    Call QueryPerformanceCounter(Strt)
    Call QueryPerformanceCounter(Ende)
    Calibr = (D(Ende) - D(Strt)) / D(Freq) * 1000
End Sub
Private Function RunTime() As Variant
    RunTime = (D(Ende) - D(Strt)) / D(Freq) * 1000 - Calibr
End Function
Private Function D(X As LARGE_INTEGER) As Double
    Dim L As Double, H As Double
    
    L = X.Lo
    H = X.Hi
    
    If L < 0 Then L = 4294967296# + L + 1
    If H < 0 Then H = 4294967296# + H + 1
    
    D = L + H * 4294967296#

End Function
Public Function msStopClock(Command As msSTPNFO) As Variant
    
    Select Case Command
        Case msSTOP
            Call QueryPerformanceCounter(Ende)
            msStopClock = Format(RunTime, "0.00")
        Case msSTART, msRESET
            Call Calibrieren
            Call QueryPerformanceCounter(Strt)
        Case Else: Exit Function
    End Select
    
End Function
Public Function StopClock(Command As STPNFO) As Date
    Static sDate As Date
    
    Select Case Command
        Case stSTOP
            StopClock = Format$((Now - sDate), "hh:mm:ss")
            sDate = CDate(0)
        Case stSTART, stRESET
            sDate = Time
            StopClock = sDate
        Case stCURRENTSTOPTIME
            StopClock = Format$((Now - sDate), "hh:mm:ss")
        Case stSTARTTIME
            StopClock = sDate
        Case Else: Exit Function
    End Select

End Function
Public Function isLeapYear(ByVal Year As Long) As Boolean
    isLeapYear = IsDate("02/29/" & Year)
End Function
Public Sub VarSwap(Var1 As Variant, Var2 As Variant)
    Dim V As Variant
    
    V = Var1
    
    Var1 = Var2
    Var2 = V

End Sub
Public Function DateIs(ByVal Dat As Date) As DateType
    
    If CBool((CLng(Dat) = CDbl(Dat))) Then
        DateIs = dtDATEONLY
    Else
        If (Not CBool(Fix(CDbl(Dat)))) Then
            DateIs = dtTIMEONLY
        Else
            If (Not CBool((CLng(Dat) = CDbl(Dat)))) Then
                DateIs = dtTIMEDATE
            Else
                DateIs = dtUNKOWN
            End If
        End If
    End If
    
End Function
Public Function LogFile(ByVal File As String, _
                        ByVal LogText As String, _
                        Optional FileSizeLimitBytes As Long = 0) _
                            As String
    Dim sText As String
    Dim R$
    
    sText = TXT.TextFile_Read(File)
    
    If Len(LogText) > 0 Then
        R$ = Date & "/" & Time & " -> " & LogText
        sText = sText & R$ & vbCrLf
    Else
        sText = sText & vbCrLf
        R$ = ""
    End If
    
    If FileSizeLimitBytes > 0 Then _
        sText = Right$(sText, FileSizeLimitBytes - 2)
    
    If TXT.TextFile_WriteText(File, sText, True) Then _
        LogFile = R$
        
End Function
Public Function Get_WinSymbol(Crtl As Object, _
                              Symbol As WinSymbol) As Boolean
    Dim TXT As String
    Dim isText As Boolean
    
    On Local Error Resume Next
    
    TXT = Crtl.Caption
    
    If Err = 438 Then
        Err.Clear
        TXT = Crtl.Text
        If Err = 438 Then Exit Function
        isText = True
    End If
             
    On Local Error GoTo Quit
    
    Crtl.Font.Name = "Marlett"
       
    If isText Then
        Crtl.Text = Chr$(Symbol)
    Else
        Crtl.Caption = Chr$(Symbol)
    End If
    
    Get_WinSymbol = True
    
Quit:
End Function
Public Function DayOfMonthCount(ByVal Month As Long, _
                                ByVal Year As Long) As Long
    Dim P As Long
      
    On Local Error GoTo Quit
    
    For P = 31 To 1 Step -1
        If IsDate(CStr(P) & "." & CStr(Month) & "." & CStr(Year)) Then _
            DayOfMonthCount = P: _
            Exit Function
    Next P

Quit:
    DayOfMonthCount = 0

End Function
Public Function Time_Format(ByVal lngInputValue As Long, _
                            Optional sInputType As tfFORMATTYPE = tfSECOND, _
                            Optional ByVal sFormat As String = "hh:mm:ss", _
                            Optional sMaxValue As tfFORMATTYPE = tfDAY) _
                                    As String
    Dim dd As String, dh As String, dm As String, ds As String
    Dim Textd As String, Texth As String, Textm As String, Texts As String
    Dim D As Long, H As Long, M As Long, S As Long
    Dim iMaxValue As Integer
    
    sFormat = LCase$(sFormat)
    
    sFormat = Replace(sFormat, "d$", "(0)")
    sFormat = Replace(sFormat, "h$", "(1)")
    sFormat = Replace(sFormat, "m$", "(2)")
    sFormat = Replace(sFormat, "s$", "(3)")
    sFormat = Replace(sFormat, "&", "(&)")

    Select Case sInputType
        Case tfMINUTE: lngInputValue = lngInputValue * 60
        Case tfHOUR: lngInputValue = lngInputValue * 3600
        Case tfDAY: lngInputValue = lngInputValue * 86400
        Case tfSECOND
        Case Else: Exit Function
    End Select
    
    iMaxValue = sMaxValue
    
    If iMaxValue = 0 Then D = Int(lngInputValue / 86400): _
                          lngInputValue = lngInputValue - (86400 * D)
    
    If iMaxValue <= 1 Then H = Int(lngInputValue / 3600): _
                           lngInputValue = lngInputValue - (3600 * H)
  
    If iMaxValue <= 2 Then M = Int(lngInputValue / 60): _
                           lngInputValue = lngInputValue - (60 * M)
    
    S = lngInputValue
  
    dd = Format$(D, "00")
    dh = Format$(H, "00")
    dm = Format$(M, "00")
    ds = Format$(S, "00")
  
    sFormat = Replace(sFormat, "dd", dd)
    sFormat = Replace(sFormat, "d", D)
    sFormat = Replace(sFormat, "hh", dh)
    sFormat = Replace(sFormat, "h", H)
    sFormat = Replace(sFormat, "mm", dm)
    sFormat = Replace(sFormat, "m", M)
    sFormat = Replace(sFormat, "ss", ds)
    sFormat = Replace(sFormat, "s", S)
  
    If D = 1 Then Textd = " Tag" Else Textd = " Tage"
    If H = 1 Then Texth = " Stunde" Else Texth = " Stunden"
    If M = 1 Then Textm = " Minute" Else Textm = " Minuten"
    If S = 1 Then Texts = " Sekunde" Else Texts = " Sekunden"
  
    Textd = CStr(dd) & " Tage"
    Texth = CStr(dh) & Texth
    Textm = CStr(dm) & Textm
    Texts = CStr(ds) & Texts

    sFormat = Replace(sFormat, "(0)", Textd)
    sFormat = Replace(sFormat, "(1)", Texth)
    sFormat = Replace(sFormat, "(2)", Textm)
    sFormat = Replace(sFormat, "(3)", Texts)
    sFormat = Replace(sFormat, "(&)", " und ")
  
    Time_Format = sFormat

End Function
Public Sub sDebug(ByVal DebugText As String, _
                  Optional DebugLogFile As String, _
                  Optional LogFileLenght As Long = 50000, _
                  Optional Control As Object, _
                  Optional isError As Boolean = False, _
                  Optional ErrorNumber As Long = 0)
    Dim T$, R$, S$, U$, V$
    
    Static LogFile As String
    
    If LogFile = "" Then
        If DebugLogFile = "" Then Exit Sub
        
        LogFile = DebugLogFile
    End If
    
    If DebugText <> vbCrLf And DebugText <> "" Then
        U$ = CStr(Date) & " / " & CStr(Time) & " -> "
    
        If isError Then
            V$ = "ERROR " & CStr(ErrorNumber) & _
                 " / " & "Description: " & DebugText
        Else
            V$ = DebugText
        End If
    
        T$ = U$ & V$
    Else
        T$ = vbCrLf
    End If
    
    On Local Error Resume Next
    
    Control.Caption = T$
    
    If Err.Number <> 0 Then
        Err.Clear
        
        Control.AddItem T$
        
        If Err.Number <> 0 Then Control.Text = Control.Text & T$ & vbCrLf
    End If
    
    If T$ <> vbCrLf Then
        Call MISC.LogFile(LogFile, V$, LogFileLenght)
    Else
        Call TXT.TextFile_AddText(LogFile, T$, Add_to_End)
    End If
    
End Sub
