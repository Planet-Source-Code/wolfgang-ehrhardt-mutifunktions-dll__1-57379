VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TextInfoNfo
    tiCHARCOUNT = 1
    tiLINECOUNT = 2
    tiWORDCOUNT = 3
    tiFILELEN = 4
End Enum

Public Enum AddTextNFO
    Add_to_Begin = 1
    Add_to_End = 2
End Enum
Public Function TextFile_WriteText(ByVal TF_TextFile As String, _
                                   ByVal TF_Text As String, _
                                   Optional OverwriteExist As Boolean = False) _
                                         As Boolean
    Dim FN As Integer
        
    On Local Error GoTo Quit
    
    If F.Exist(TF_TextFile) Then
        If Not OverwriteExist Then Exit Function
        
        Call F.fKill(TF_TextFile)
    End If
    
    FN = FreeFile
    
    Open TF_TextFile For Output As FN
        Print #FN, TF_Text
    Close FN

    TextFile_WriteText = True

Quit:
End Function
Public Function TextFile_AddText(ByVal TF_TextFile As String, _
                                 ByVal TF_Text As String, _
                                 Add_To As AddTextNFO) _
                                       As Boolean
    Dim FN As Integer
    Dim Str As String
    
    On Local Error GoTo Quit
    
    Select Case Add_To
        Case Add_to_Begin
            Str = TF_Text & vbCrLf & TextFile_Read(TF_TextFile)
            TextFile_AddText = TextFile_WriteText(TF_TextFile, _
                                                  Str, True)
        Case Add_to_End
            FN = FreeFile
    
            Open TF_TextFile For Append As FN
                Print #FN, TF_Text
            Close FN
            
            TextFile_AddText = True
    End Select
    
Quit:
End Function
Public Function TextFile_GetInfo(ByVal TF_TextFile As String, _
                                 TextInfo As TextInfoNfo) _
                                        As Long
    Dim AstrData() As String, Text As String
        
    On Local Error GoTo Quit
    
    TextFile_GetInfo = -1
    
    If Not F.Exist(TF_TextFile) Then Exit Function
    
    Select Case TextInfo
        Case tiCHARCOUNT, tiFILELEN
             TextFile_GetInfo = FileLen(TF_TextFile)
        Case tiLINECOUNT
            Text = TextFile_Read(TF_TextFile)
            AstrData() = Split(Text, vbCrLf)
            TextFile_GetInfo = Format$(UBound(AstrData) + 1, "#,##0")
        Case tiWORDCOUNT
            Text = TextFile_Read(TF_TextFile)
            AstrData() = Split(Text, " ")
            TextFile_GetInfo = Format$(UBound(AstrData) + 1, "#,##0")
        Case Else
            '
    End Select
        
Quit:
End Function
Public Function TextFile_Read(ByVal TF_TextFile As String) As String
    Dim FN As Integer
    
    On Local Error GoTo Quit
    
    FN = FreeFile()
    
    Open TF_TextFile For Binary Access Read As #FN
        TextFile_Read = Space(LOF(FN))
        Get #FN, , TextFile_Read
    Close #FN
    
    If Right$(TextFile_Read, 2) = vbCrLf Then _
        TextFile_Read = Left$(TextFile_Read, Len(TextFile_Read) - 2)

Quit:
    If Err.Number <> 0 Then TextFile_Read = ""
    
End Function
Public Function TextFile_ReplaceText(ByVal TF_TextFile As String, _
                                     ByVal TF_SearchFor As String, _
                                     ByVal TF_ReplaceWith As String) _
                                            As Boolean
    Dim Text As String
        
    On Local Error GoTo Quit
    
    If TF_SearchFor = "" Or (Not F.Exist(TF_TextFile)) Then _
        Exit Function
        
    Text = TextFile_Read(TF_TextFile)
    Text = Replace(Text, TF_SearchFor, TF_ReplaceWith)
        
    TextFile_ReplaceText = TextFile_WriteText(TF_TextFile, _
                                              Text, True)

Quit:
End Function
Public Function TextFile_LineCount(ByVal TF_TextFile As String) As Long
    Dim S() As String
        
    S = Split(TextFile_Read(TF_TextFile), vbCrLf)
    TextFile_LineCount = UBound(S)
    
End Function
Public Function TextFile_ReadLine(ByVal TF_TextFile As String, _
                                  ByVal LineNr As Long) As String
    Dim S$
        
    S$ = TextFile_Read(TF_TextFile)
    TextFile_ReadLine = Str.GetLineFromString(S$, LineNr)
    
End Function
Public Function TextFile_Print(ByVal TF_TextFile As String) As Boolean
    TextFile_Print = Text_Print(TextFile_Read(TF_TextFile))
End Function
Public Function Text_Print(ByVal Text As String) As Boolean
    
    On Local Error GoTo Quit
    
    If Len(Text) Then Printer.Print Text: _
                      Printer.EndDoc: _
                      Text_Print = True

Quit:
End Function
Public Function Text_PrintBlocksatz(ByVal Text As String, _
                                    Optional LeftMargin As Long = 80, _
                                    Optional RightMargin As Long = 80, _
                                    Optional TopMargin As Long = 80) _
                                        As Boolean
    Dim S As String, T As String, O As String
    Dim AP As Long, L As Long, n As Long, C As Long, ASM As Long
    Dim V As Variant
        
    On Local Error GoTo Quit
    
    With Printer
        ASM = .ScaleMode
        .ScaleMode = vbTwips
        S = Text
        AP = TopMargin
        
        Do While Len(S) > 0
            Do While Len(S) > 0
                If InStr(1, S, " ") > 0 Then
                    T = Left$(S, InStr(1, S, " ") - 1)
                Else
                    T = S
                End If
                
                If .TextWidth(O & " " & T) > (.ScaleWidth - LeftMargin - RightMargin) Then
                    Exit Do
                Else
                    If InStr(1, S, " ") > 0 Then
                        O = O & " " & T
                        S = Mid$(S, Len(T) + 2)
                    Else
                        O = O & " " & S
                        S = vbNullString
                    End If
                End If
            Loop
    
            Do
                C = Len(O)
                O = Replace(O, "  ", " ")
            Loop While C > Len(O)
      
            O = Trim(O)
            V = Split(O, " ")
            O = Replace(O, " ", vbNullString)
      
            Do
                If UBound(V) = 0 And Len(S) > 0 Then
                    O = vbNullString
                
                    For C = 1 To Len(V(0))
                        O = O & " " & Mid$(V(0), C, 1)
                    Next C
          
                    O = Trim(O)
                    V = Split(O, " ")
                    O = Replace(O, " ", vbNullString)
                ElseIf Len(S) > 0 Then
                    n = .ScaleWidth - LeftMargin - RightMargin
                    n = n - .TextWidth(O)
                    n = n / UBound(V)
            
                    Exit Do
                Else
                    n = .TextWidth(" ")
                    Exit Do
                End If
            Loop
      
            L = LeftMargin
        
            For C = 0 To UBound(V)
                .CurrentX = L
                .CurrentY = AP
             
                Printer.Print V(C)
            
                L = L + .TextWidth(V(C)) + n
            Next C
      
            AP = AP + .TextHeight(O)
            O = vbNullString
        Loop
        
        .ScaleMode = ASM
    End With

    Text_PrintBlocksatz = True

Quit:
End Function
Public Function TextFile_ReadLinesInAray(ByVal TF_TextFile As String) _
                                            As String()
    Dim Text As String
    Dim A() As String
        
    On Local Error GoTo Quit
    
    Text = TextFile_Read(TF_TextFile)
    A = Split(Text, vbCrLf)
    
Quit:
    If Err Then Erase A
    TextFile_ReadLinesInAray = A
    
End Function
Public Function TextFile_RemoveLine(ByVal TF_TextFile As String, _
                                    ByVal LineToRemove As Long) _
                                        As Boolean
    Dim R$, A() As String
    Dim P As Long
        
    On Local Error GoTo Quit
    
    A = Split(TextFile_Read(TF_TextFile), vbCrLf)
    
    LineToRemove = LineToRemove - 1
    
    If LineToRemove <= UBound(A) Then
        If ARR.DeleteIndex(A, LineToRemove) Then
            R$ = ARR.ToString(A)
            
            If TextFile_WriteText(TF_TextFile, R$, True) Then _
                TextFile_RemoveLine = True
        End If
    End If

Quit:
End Function
