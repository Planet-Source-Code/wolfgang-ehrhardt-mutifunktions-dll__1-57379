VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function OemToCharA Lib "user32.dll" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function CharToOemA Lib "user32.dll" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Function EnCrypt(ByVal Key As String, _
                        ByVal TextToEnCrypt As String) _
                            As String
    EnCrypt = sEnCrypt(TextToEnCrypt, Key)
End Function
Public Function DeCrypt(ByVal Key As String, _
                        ByVal TextToDeCrypt As String) _
                            As String
    DeCrypt = sDeCrypt(TextToDeCrypt, Key)
End Function
Public Function Str_to8ByteHex(Str As String) As String
    Str_to8ByteHex = toHex(8, Str)
End Function
Public Function Str_to16ByteHex(Str As String) As String
    Str_to16ByteHex = toHex(16, Str)
End Function
Public Function isValidDouble(ByVal TestStr As String) As Boolean
    Dim D As Double
    
    On Local Error Resume Next
    D = CDbl(TestStr)
    
    isValidDouble = (Err.Number = 0)
    
End Function
Public Function isValidSingle(ByVal TestStr As String) As Boolean
    Dim S As Single
    
    On Local Error Resume Next
    S = CSng(TestStr)
    
    isValidSingle = (Err.Number = 0)
    
End Function
Public Function isValidNum(ByVal TestStr As String) As Boolean
    isValidNum = IsNumeric(TestStr)
End Function
Public Function isValidAlpha(ByVal TestStr As String, _
                             Optional inclSpace As Boolean = True) _
                                    As Boolean
    Dim P As Long
    Dim V$
    
    V$ = lABC & IIf(inclSpace, " ", "")
    
    TestStr = LCase$(TestStr)

    For P = 1 To Len(TestStr)
        If InStr(V$, Mid$(TestStr, P, 1)) = 0 Then _
            Exit Function
    Next P
    
    isValidAlpha = True

End Function
Public Function Str_toANSI(ByVal Str As String) As String
    Call OemToCharA(Str, Str)
    Str_toANSI = Str
End Function
Public Function Str_toASCII(ByVal Str As String) As String
    Call CharToOemA(Str, Str)
    Str_toASCII = Str
End Function
Public Function isValidDate(ByVal TestStr As String) As Boolean
    isValidDate = IsDate(TestStr)
End Function
Public Function Str_Reverse(ByVal Str As String) As String
    Str_Reverse = StrReverse(Str)
End Function
Public Function Str_toCHR(ByVal Str As String) As String()
    Dim P As Long
    Dim A() As String, R$
    
    ReDim A(Len(Str) - 1)
    
    For P = 1 To Len(Str)
        R$ = Mid$(Str, P, 1)
        A(P - 1) = R$
    Next P
    
    Str_toCHR = A
    
End Function
Public Function CountLinesInString(ByVal Str As String) As Long
    Dim A() As String
    
    A = Split(Str, vbCrLf)
    
    If Not ARR.isArray(A) Then
        CountLinesInString = 0
    Else
        CountLinesInString = UBound(A) + 1
    End If
    
End Function
Public Function GetLineFromString(ByVal Str As String, _
                                  ByVal LineNr As Long) As String
    Dim S() As String
    
    On Local Error GoTo Quit
    
    LineNr = LineNr - 1
    
    S = Split(Str, vbCrLf)
    
    If UBound(S) >= LineNr Then _
        GetLineFromString = S(LineNr)
        
Quit:
    Erase S
    If Err Then GetLineFromString = ""
    
End Function
Public Function CountStringInString(ByVal Text As String, _
                                    ByVal SearchFor As String, _
                                    Optional SearchExact As Boolean = True, _
                                    Optional WholeWordsOnly As Boolean = False) _
                                            As Long
  
    On Local Error Resume Next
    
    If Not SearchExact Then Text = LCase$(Text): _
                            SearchFor = LCase$(SearchFor)
    
    If WholeWordsOnly Then SearchFor = " " & SearchFor & " ": _
                           Text = " " & Text & " "

    CountStringInString = (Len(Text) - Len(Replace(Text, SearchFor, ""))) _
                                    / Len(SearchFor)
    
End Function
Public Function Str_toASC(ByVal Str As String) As String()
    Dim P As Long
    Dim S$, A() As String
    
    ReDim A(Len(Str) - 1)
    
    For P = 1 To Len(Str)
        S$ = CStr(Asc(Mid$(Str, P, 1)))
        A(P - 1) = S$
    Next P
    
    Str_toASC = A
    
End Function
Public Function Str_Cut(ByVal Str As String, _
                        ByVal StrToCut, _
                        Optional CutFromRight As Boolean = False) _
                            As String
    Dim Pos As Long, X As Long
    Dim R$, S$
    
    If CutFromRight Then
        X = 1
        
        Do While InStr(X, Str, StrToCut)
            Pos = InStr(X, Str, StrToCut)
            X = Pos + Len(StrToCut)
        Loop
    Else
        Pos = InStr(Str, StrToCut)
    End If
    
            
    If Pos Then
        R$ = Mid$(Str, 1, Pos - 1)
        S$ = Mid$(Str, Pos + Len(StrToCut))
            
        Str_Cut = R$ & S$
    End If

End Function
Public Function Get_RandomString(ByVal strLenght As Long, _
                                 Optional inclNum As Boolean = True, _
                                 Optional inclLAlpha As Boolean = True, _
                                 Optional inclUAlpha As Boolean = True, _
                                 Optional Space As Boolean = True, _
                                 Optional Other As Boolean = False) _
                                        As String
    Dim A(3) As String, S$, T$
    Dim P As Long, X As Long
    
    If strLenght < 1 Then Exit Function

    A(0) = Num
    A(1) = lABC
    A(2) = uABC
    A(3) = Chr$(32)
    
    If Other Then
        For P = 0 To 255
            S$ = S$ & Chr$(P)
        Next P
    
        If Not inclNum Then S$ = Replace(S$, A(0), "")
        If Not inclLAlpha Then S$ = Replace(S$, A(1), "")
        If Not inclUAlpha Then S$ = Replace(S$, A(2), "")
        If Not Space Then S$ = Replace(S$, A(3), "")
    Else
        If inclNum Then S$ = A(0)
        If inclLAlpha Then S$ = S$ & A(1)
        If inclUAlpha Then S$ = S$ & A(2)
        If Space Then S$ = S$ & A(3)
    End If
        
    For P = 1 To strLenght
        T$ = T$ & Mid$(S$, MISC.Get_Randomize(1, Len(S$)), 1)
    Next P
    
    Get_RandomString = T$
        
End Function
Public Function isValidInteger(ByVal TestStr As String) As Boolean
    Dim I As Integer
    
    On Local Error Resume Next
    I = CInt(TestStr)
    
    isValidInteger = (Err.Number = 0)
    
End Function
Public Function isValidLong(ByVal TestStr As String) As Boolean
    Dim L As Long
    
    On Local Error Resume Next
    L = CLng(TestStr)
    
    isValidLong = (Err.Number = 0)
    
End Function
Public Function GetNum(ByVal Str As String, _
                       ByVal DigitsForDecimal As Long, _
                       ByVal DigitsAfterDecimal As Long) _
                            As String
    Dim X As Long
    Dim R$, S$
    
    On Local Error GoTo Quit
    
    If Not isValidNum(Str) Then Exit Function
    
    R$ = CStr(Round(CDbl(Str), DigitsAfterDecimal))
    X = InStr(R$, ",")
        
    If X = 0 Then
        If DigitsAfterDecimal Then _
            S$ = "," & String(DigitsAfterDecimal, "0")
    Else
        S$ = Right$(R$, Len(R$) - X)
        R$ = Left$(R$, X - 1)
    
        If DigitsAfterDecimal = 0 Then S$ = ""

        If DigitsAfterDecimal Then _
            S$ = S$ & String(DigitsAfterDecimal - Len(S$), "0"): _
            S$ = "," & S$
    End If
    
    X = DigitsForDecimal - Len(R$)
    X = IIf(X < 0, 0, X)
        
    R$ = String(X, "0") & R$
    R$ = R$ & S$
    
    GetNum = R$

Quit:
End Function
Public Function CheckSum(Str As String) As String
    Dim lCtr As Long, X As Long
    Dim DatenASCII() As Byte

    On Local Error GoTo Quit

    DatenASCII = Str
  
    For lCtr = LBound(DatenASCII) To UBound(DatenASCII) Step 2
        X = (X + DatenASCII(lCtr)) And 255
    Next lCtr

    CheckSum = CStr(X)
    
Quit:
    If Err.Number <> 0 Then CheckSum = ""

End Function
