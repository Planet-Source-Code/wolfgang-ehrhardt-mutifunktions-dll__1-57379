VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Multimedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Private Declare Function BeepA Lib "kernel32.dll" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

Public Enum pNFO
    mmPLAY = 1
    mmSTOP = 0
    mmPAUSE = 2
    mmRESUME = 3
End Enum
Public Function hasSoundCard() As Boolean
    If waveOutGetNumDevs > 0 _
    Or midiOutGetNumDevs > 0 Then hasSoundCard = True
End Function
Public Function PlayMultimediaFile(Command As pNFO, _
                                   ByVal File As String) _
                                   As Boolean
    Dim oStr As String
    
    File = F.Get_DOSfileName(File)
    
    Select Case Command
        Case mmSTOP
            oStr = "stop"
        Case mmPLAY, mmRESUME
            oStr = "play"
        Case mmPAUSE
            oStr = "pause"
    End Select
    
    oStr = oStr & " " & File
    
    If mciSendString(oStr, 0&, 0, 0) = 0 Then _
        PlayMultimediaFile = True
    
End Function
Public Function GetPlayTime(ByVal File As String) As Long
    Dim mssg As String * 255
        
    File = F.Get_DOSfileName(File)
    
    If mciSendString("set " & File & _
                     " time format ms", 0&, 0, 0) = 0 Then
        If mciSendString("status " & File & _
                         " length", mssg, 255, 0) = 0 Then _
                            GetPlayTime = (Val(mssg) / 1000)
    End If
    
End Function
Public Function isFilePlayable(ByVal File As String) As Boolean
    Dim mssg As String * 255
        
    File = F.Get_DOSfileName(File)

    If mciSendString("capability " & File & " can play", mssg, 255, 0) = 0 Then
        If InStr(LCase$(mssg), "true") Then isFilePlayable = True
    End If
    
End Function
Public Function PCspeaker(ByVal Enabled As Boolean) As Boolean
    Dim Path As String, Key As String, Value As String
        
    Path = "Control Panel\Sound"
    Key = "ExtendedSounds"
    
    Value = IIf(Enabled, "Yes", "No")
    
    PCspeaker = REG.RegValueSet(HKEY_CURRENT_USER_, Path, Key, Value, True)
    
End Function
Public Function PCspeakerIsEnabled() As Boolean
    Dim Path As String, Key As String, Value As String
        
    Path = "Control Panel\Sound"
    Key = "ExtendedSounds"
    
    Value = REG.RegValueGet(HKEY_CURRENT_USER_, Path, Key)
    Value = LCase$(Value)
    
    If Value = "ja" Or Value = "yes" Or _
       Value = "true" Or Val(Value) > 0 Then _
            PCspeakerIsEnabled = True
    
End Function
Public Function Beep(ByVal dwFreq As Long, _
                     ByVal dwDuration As Long) As Boolean
      
    If BeepA(dwFreq, dwDuration) <> 0 Then Beep = True
   
End Function
Public Function SetVolume(ByVal vRight As Long, _
                          ByVal vLeft As Long) As Boolean
    Dim intLeft As Integer, intRight As Integer
    Dim lngVolume As Long

    On Local Error GoTo Quit
    
    intRight = vRight
    intLeft = vLeft
    
    lngVolume = lngVolume Or (intLeft * 2 ^ 16)
    lngVolume = lngVolume Or intRight

    SetVolume = (waveOutSetVolume(0, lngVolume) = 0)

Quit:
End Function
Public Function Video_PlayFullScreen(ByVal Path As String) _
                                        As Boolean
    Dim R As Long
    Dim aa As String
    
    aa = F.Get_DOSfileName(Path)
    
    Video_PlayFullScreen = (mciSendString("play " & aa & _
                                          " fullscreen ", _
                                          0&, 0, 0&) = 0)

End Function
