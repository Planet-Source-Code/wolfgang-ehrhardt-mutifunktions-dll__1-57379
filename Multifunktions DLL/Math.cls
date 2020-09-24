VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mathematics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ntDateNFO
    ntHHMMSS = 1
    ntMMSS = 2
End Enum

Public Enum bbBit
    bb_4_BIT = 0
    bb_8_BIT = 1
    bb_16_BIT = 2
    bb_31_BIT = 3
End Enum

Public Enum btBYTENFO
    btKiloByte = 0
    btMegaByte = 1
    btGigaByte = 2
End Enum

Public Function DecToBin(ByVal Dec As Long) As String
    Dim X As Integer

    If Dec >= 2 ^ 32 Then _
        DecToBin = "-1": _
        Exit Function
        
    Do
        If (Dec And 2 ^ X) Then
            DecToBin = "1" & DecToBin
        Else
            DecToBin = "0" & DecToBin
        End If
        
        X = X + 1
    Loop Until 2 ^ X > Dec

End Function
Public Function HexToDec(ByVal Hex As String) As Long
    HexToDec = Val("&H" & Hex)
End Function
Public Function DecToHex(ByVal Dec As Long) As String
    DecToHex = CStr(Hex(Dec))
End Function
Public Function isPrime(ByVal n As Long) As Boolean
    Dim I As Long

    If (n <> 2 And (n And 1) = 0) _
    Or (n <> 3 And n Mod 3 = 0) Then Exit Function
    
    For I = 6 To Sqr(n) Step 6
        If (n Mod (I - 1) = 0) _
        Or (n Mod (I + 1) = 0) Then Exit Function
    Next
    
    isPrime = True

End Function
Public Function BinToDec(ByVal Bin As Long) As Long
    Dim sngNumber As Long, tmp As Long, Output As Long
    Dim X As Integer
    
    sngNumber = Bin

    For X = 0 To Len(CStr(sngNumber)) - 1
        tmp = CLng(Right$(CStr(sngNumber), 1))
    
        If tmp = 1 Then tmp = tmp * Pow(2, X)
        
        Output = Output + tmp
        tmp = 1

        If Len(CStr(sngNumber)) > 1 Then
            sngNumber = CLng(Left$(CStr(sngNumber), _
                             Len(CStr(sngNumber)) - 1))
        Else
            sngNumber = 0
        End If
    Next X
    
    BinToDec = Output

End Function
Public Function Fakultaet(ByVal Zahl As Long) As Double
    Dim L As Long
    Dim nFak As Double
    
    On Local Error GoTo Quit
    
    If Zahl = 1 Then Fakultaet = 1: _
                     Exit Function
        
    nFak = 1
            
    For L = 1 To Zahl
        nFak = nFak * CDbl(L)
    Next L
            
    Fakultaet = nFak
    
Quit:
    If Err.Number <> 0 Then Fakultaet = -1
    
End Function
Public Function Quersumme(ByVal Zahl As Long) As Long
    Do While Zahl
        Quersumme = Quersumme + (Zahl Mod 10)
        Zahl = Zahl \ 10
    Loop
End Function
Public Function Pythag(Optional A As Variant, _
                       Optional B As Variant, _
                       Optional Hyp As Variant) As Double
    Dim X As Integer
    
    On Local Error GoTo Quit
    
    If IsMissing(A) Then X = 1
    If IsMissing(B) Then X = X + 1
    If IsMissing(Hyp) Then X = X + 1
    
    If X < 2 Then
        If IsMissing(Hyp) Then Pythag = Sqr((A ^ 2) + (B ^ 2))
        If IsMissing(A) Then Pythag = Sqr((Hyp ^ 2) - (B ^ 2))
        If IsMissing(B) Then Pythag = Sqr((Hyp ^ 2) - (A ^ 2))
    Else
        Pythag = -1
    End If
    
Quit:
    If Err.Number <> 0 Then Pythag = -2

End Function
Public Function LoWord(ByVal nDWord As Long) As Long
    LoWord = CLng("&H" & Left$(Right$("00000000" & Hex$(nDWord), 8), 4))
End Function
Public Function HiWord(ByVal nDWord As Long) As Long
    HiWord = CLng("&H" & Right$(Right$("00000000" & Hex$(nDWord), 8), 4))
End Function
Public Function DWord(ByVal nLow As Integer, _
                      ByVal nHigh As Integer) As Long
  Dim sLow As String, sHigh As String
   
  sLow = String$(4 - Len(Hex$(nLow)), "0") & Hex$(nLow)
  sHigh = String$(4 - Len(Hex$(nHigh)), "0") & Hex$(nHigh)

  DWord = CLng("&H" & sLow & sHigh)

End Function
Public Sub Incr(Var As Long, _
                Optional Value As Long = 1)
    Var = Var + Value
End Sub
Public Sub Decr(Var As Long, _
                Optional Value As Long = 1)
    Var = Var - Value
End Sub
Public Function ByteTo(ByVal Bytes As Variant, _
                       BytesTo As btBYTENFO) As Variant
    
    Select Case BytesTo
        Case btGigaByte
            ByteTo = Round((Bytes / 1048576) / 1024, 2)
        Case btKiloByte
            ByteTo = Round(Bytes / 1024, 2)
        Case btMegaByte
            ByteTo = Round(Bytes / 1048576, 2)
    End Select
    
End Function
