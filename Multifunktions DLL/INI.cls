VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function DeletePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section As String, ByVal NoKey As Long, ByVal NoSetting As Long, ByVal FileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function Read_INI(ByVal INIpath As String, _
                         ByVal Section As String, _
                         ByVal Abschnitt As String) _
                             As String
   Dim RetStr As String
   
   RetStr = String(255, Chr$(0))
   
   Read_INI = Left$(RetStr, _
                    GetPrivateProfileString(Section, _
                                            Abschnitt, "", _
                                            RetStr, Len(RetStr), _
                                            INIpath))

End Function
Public Function Write_INI(ByVal INIpath As String, _
                          ByVal Section As String, _
                          ByVal Abschnitt As String, _
                          ByVal Value As String) As Boolean
                           
    If WritePrivateProfileString(Section, Abschnitt, _
                                 Value, INIpath) Then Write_INI = True

End Function
Public Sub Remove_Section(ByVal INIpath As String, _
                          ByVal Section As String)
    Call DeletePrivateProfileSection(Section, 0, 0, INIpath)
End Sub
Public Function WriteArray(ByVal INIpath As String, _
                           ByVal Section As String, _
                           sArray() As String) As Boolean
    Dim P As Long
        
    On Local Error GoTo Quit

    Call Remove_Section(INIpath, Section)
    
    For P = LBound(sArray) To UBound(sArray)
        If WritePrivateProfileString(Section, CStr(P), _
                                     CStr(sArray(P)), INIpath) = 0 Then _
                                     Exit Function
    Next P

    WriteArray = True
    
Quit:
End Function
Public Function ReadSection(ByVal INIpath As String, _
                            ByVal Section As String) As String()
    Dim Result As Long, L As Long, P As Long, Z As Long, X As Long
    Dim Buffer As String, A() As String, B() As String
        
    Buffer = Space(32767)
    Result = GetPrivateProfileSection(Section, Buffer, Len(Buffer), INIpath)
    
    Buffer = Left$(Buffer, Result)
    
    If Buffer <> "" Then
        L = 1
        Z = 0
        
        Do While L < Result
            P = InStr(L, Buffer, Chr$(0))
            If P = 0 Then Exit Do
            
            ReDim Preserve A(Z)
            
            A(Z) = Mid$(Buffer, L, P - L)
            
            Z = Z + 1
            L = P + 1
        Loop
    End If

    For P = 0 To UBound(A)
        X = InStr(A(P), "=")
        A(P) = Mid(A(P), X + 1)
        A(P) = Trim(A(P))
    Next P
        
    ReadSection = A

End Function
