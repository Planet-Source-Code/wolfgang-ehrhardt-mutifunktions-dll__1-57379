VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum arADDnfo
    ADDTOBEGIN = 1
    ADDTOEND = 2
End Enum
Public Function Compare(sArray1 As Variant, _
                        sArray2 As Variant) As Boolean
    Dim P As Long
    
    On Local Error GoTo Quit
    
    If (UBound(sArray1) <> UBound(sArray2)) Or _
       (LBound(sArray1) <> LBound(sArray2)) Then _
            Exit Function
        
    For P = LBound(sArray1) To UBound(sArray1)
        If sArray1(P) <> sArray2(P) Then Exit Function
    Next P
        
    Compare = True
    
Quit:
End Function
Public Function Find(ByVal sArray As Variant, _
                     ByVal ToFind As Variant, _
                     Optional StartIndex = -1, _
                     Optional EndIndex As Long = -1, _
                     Optional FindExact As Boolean = False) _
                            As Long
    Dim P As Long, sI As Long, eI As Long
    
    On Local Error GoTo Quit
    
    If Not FindExact Then ToFind = LCase$(ToFind)
        
    If StartIndex > -1 Then
        sI = StartIndex
    Else
        sI = LBound(sArray)
    End If
    
    If EndIndex > -1 Then
        eI = EndIndex
    Else
        eI = UBound(sArray)
    End If
        
    For P = sI To eI
        If IIf(FindExact, CStr(sArray(P)), _
                          CStr(LCase$(sArray(P)))) = ToFind Then _
            Find = P: _
            Exit Function
    Next P
    
    Find = -1
    
Quit:
    If Err.Number <> 0 Then Find = -2
    
End Function
Public Function Shuffle(sArray As Variant) As Boolean
    Dim I As Long, X As Long, Y As Long
    Dim A As Variant, B As Variant
        
    On Local Error GoTo Quit
    
    A = sArray
    B = sArray
    
    Y = LBound(A)
        
    Do
        X = MISC.Get_Randomize(UBound(A), LBound(A) + 1)
        
        B(Y) = A(X)
        Y = Y + 1
        
        If X <> UBound(A) Then
            For I = X To UBound(A) - 1
                A(I) = A(I + 1)
            Next I
        End If
        
        If UBound(A) <= LBound(A) Then
            For I = LBound(B) To UBound(B)
                sArray(I) = B(I)
            Next I
            
            Exit Do
        Else
            ReDim Preserve A(UBound(A) - 1)
        End If
    Loop
        
    Shuffle = True
    
Quit:
End Function
Public Function Sort(sArray As Variant, _
                     Optional sBubbleSort As Boolean = False) _
                        As Boolean
    Dim frm As New MyControls
            
    On Local Error GoTo Quit
    
    If sBubbleSort Then
        Sort = BubbleSort(sArray)
    Else
        Load frm
        frm.Tag = "Sort"
    
        If LB.AddArray(frm.sLB, sArray, True) Then _
            If Not LB.ToArray(frm.sLB, sArray, False) Then GoTo Quit
        
        Sort = True
    End If
    
Quit:
    If Not sBubbleSort Then Unload frm
    Set frm = Nothing

End Function
Public Function ToString(sArray As Variant) As String
    Dim P As Long
    Dim R$
    
    For P = LBound(sArray) To UBound(sArray)
        R$ = R$ & sArray(P) & Chr$(0)
    Next P
    
    If Len(R$) Then R$ = Left$(R$, Len(R$) - 1)
    
    ToString = R$
    
End Function
Public Function AddFromLBCB(LBCB As Object, _
                            sArray As Variant) As Boolean
    Dim P As Long
    Dim A() As String
    
    On Local Error GoTo Quit
    
    ReDim A(LBCB.ListCount - 1)
    
    For P = 0 To LBCB.ListCount - 1
        A(P) = LBCB.List(P)
    Next P
    
    sArray = A
    
    AddFromLBCB = True
    
Quit:
    Erase A
    If Err Then sArray = A
    
End Function
Public Function DeleteIndex(sArray As Variant, _
                            ByVal Index As Long) As Boolean
    Dim P As Long
    Dim A As Variant
    
    On Local Error GoTo Quit
    
    A = sArray
    
    If UBound(sArray) = 0 And Index = 0 Then
        Erase A
    Else
        For P = Index + 1 To UBound(sArray)
            A(P - 1) = sArray(P)
        Next P
    
        ReDim Preserve A(UBound(A) - 1)
    End If
    
    sArray = A
    
    DeleteIndex = True
       
Quit:
End Function
Public Function DeleteValue(sArray As Variant, _
                            ByVal Value As Variant, _
                            Optional DeleteAll As Boolean = False) _
                                    As Boolean
    Dim A() As Variant
    Dim X As Long
    
    On Local Error GoTo Quit
    
    Do While X <= UBound(sArray)
        If CStr(sArray(X)) = CStr(Value) Then
            Call DeleteIndex(sArray, X)
            
            If Not DeleteAll Then Exit Do
            
            If X = 0 And UBound(sArray) = 0 Then Erase sArray: _
                                                 Exit Do
        Else
            X = X + 1
        End If
    Loop
    
    DeleteValue = True
    
Quit:
End Function
Public Function AddArray(sArray1 As Variant, _
                         ByVal sArray2 As Variant) As Boolean
    Dim P As Long, X As Long
        
    On Local Error GoTo Quit
    
    X = UBound(sArray1)
    
    ReDim Preserve sArray1(X + UBound(sArray2) + 1)
                
    For P = 0 To UBound(sArray2)
        sArray1(X + P + 1) = sArray2(P)
    Next P

    AddArray = True
    
Quit:
End Function
Public Function CountItem(sArray As Variant, _
                          ByVal Item As String, _
                          Optional FindExact As Boolean = True) _
                                As Long
    Dim P As Long, X As Long
    
    On Local Error GoTo Quit
    
    If Not FindExact Then Item = LCase$(Item)
    
    For P = LBound(sArray) To UBound(sArray)
        If IIf(FindExact, CStr(sArray(P)), CStr(LCase$(sArray(P)))) = Item Then _
            X = X + 1
    Next P
    
    CountItem = X
    
Quit:
    If Err.Number <> 0 Then CountItem = -1
    
End Function
Public Function IsDouble(sArray As Variant, _
                         ByVal Value As String) As Boolean
    Dim P As Long, X As Long
    
    On Local Error GoTo Quit
    
    For P = LBound(sArray) To UBound(sArray)
        If sArray(P) = Value Then _
            X = X + 1: _
            If X > 1 Then IsDouble = True: _
                          Exit Function
    Next P
    
Quit:
End Function
Public Function Exist(sArray As Variant, _
                      ByVal Value As String) As Boolean
    Exist = (Find(sArray, Value, , , True) > -1)
End Function
Public Function Fill(sArray As Variant, _
                     ByVal Value As Variant) As Boolean
    Dim P As Long
    
    On Local Error GoTo Quit
    
    For P = LBound(sArray) To UBound(sArray)
        sArray(P) = Value
    Next P
    
    Fill = True
    
Quit:
End Function
Public Function WriteToFile(sArray As Variant, _
                            ByVal FileName As String) _
                                As Boolean
    Dim R$
    Dim P As Long
    
    On Local Error GoTo Quit
    
    For P = LBound(sArray) To UBound(sArray)
        R$ = R$ & sArray(P) & vbCrLf
    Next P
    
    If Len(R$) Then R$ = Mid$(R$, 1, Len(R$) - 2)
    
    WriteToFile = TXT.TextFile_WriteText(FileName, R$, True)
    
Quit:
End Function
Public Function AddToLBCB(sArray As Variant, _
                          LBCB As Object, _
                          Optional ClearBeforeAdding As Boolean = True) _
                                As Boolean
    
    AddToLBCB = LB.AddArray(LBCB, sArray, ClearBeforeAdding)
    
End Function
Public Function Add(sArray As Variant, ByVal Value As Variant, _
                    AddAr As arADDnfo, _
                    Optional Index As Long = -1) _
                        As Boolean
    Dim X As Long, P As Long
    Dim A() As Variant
    
    On Local Error Resume Next
    
    X = UBound(sArray)
    X = Err.Number
    
    On Local Error GoTo Quit
    
    If X <> 0 Then
        If Index > 0 Then Exit Function
        
        ReDim sArray(0)
        sArray(0) = Value
        
        Exit Function
    End If
    
    If Index = -1 Then
        Select Case AddAr
            Case ADDTOBEGIN
                ReDim Preserve sArray(UBound(sArray) + 1)
                
                For P = UBound(sArray) To LBound(sArray) + 1 Step -1
                    sArray(P) = sArray(P - 1)
                Next P
            
                sArray(LBound(sArray)) = Value
            Case ADDTOEND
                ReDim Preserve sArray(UBound(sArray) + 1)
                sArray(UBound(sArray)) = Value
            Case Else: Exit Function
        End Select
    Else
        If Index > UBound(sArray) + 1 Then Exit Function
        
        ReDim Preserve sArray(UBound(sArray) + 1)
        
        For P = UBound(sArray) To Index Step -1
            sArray(P) = sArray(P - 1)
        Next P
        
        sArray(Index) = Value
    End If
       
    Add = True
    
Quit:
End Function
Public Function Copy(sArray1 As Variant, _
                     sArray2 As Variant) As Boolean
    
    On Local Error GoTo Quit
    
    If ARR.isArray(sArray2) Then Erase sArray2
    sArray2 = sArray1
    
    Copy = True
    
Quit:
    On Local Error Resume Next
    If Err.Number <> 0 Then Erase sArray2
    
End Function
Public Function Get_MaxValue(sArray As Variant, _
                             MaxValue As String, Index As Long) _
                                    As Boolean
    Dim P As Long
    
    On Local Error GoTo Quit
    
    MaxValue = ""
    
    For P = LBound(sArray) To UBound(sArray)
        If CStr(sArray(P)) > MaxValue Then MaxValue = sArray(P): _
                                           Index = P
    Next P
    
    Get_MaxValue = True
    
Quit:
    If Err.Number <> 0 Then MaxValue = "": _
                            Index = -1
    
End Function
Public Function Get_MinValue(sArray As Variant, _
                             MinValue As String, Index As Long) _
                                    As Boolean
    Dim P As Long
    
    On Local Error GoTo Quit
    
    MinValue = sArray(LBound(sArray))
    Index = LBound(sArray)
    
    For P = LBound(sArray) To UBound(sArray)
        If CStr(sArray(P)) < MinValue Then MinValue = sArray(P): _
                                           Index = P
    Next P
    
    Get_MinValue = True
    
Quit:
    If Err.Number <> 0 Then MinValue = "": _
                            Index = -1
    
End Function
Public Function FindPhrase(sArray As Variant, _
                           ByVal Find As String, _
                           Optional StartIndex As Long = -1, _
                           Optional EndIndex As Long = -1, _
                           Optional FindAtStart As Boolean = False, _
                           Optional FindExact As Boolean = True) _
                                As Long
    Dim P As Long, X As Long, sI As Long, eI As Long
    Dim R$
    
    On Local Error GoTo Quit
    
    FindPhrase = -1
    
    If Not FindExact Then Find = LCase$(Find)
    
    If StartIndex > -1 Then
        sI = StartIndex
    Else
        sI = LBound(sArray)
    End If
    
    If EndIndex > -1 Then
        eI = EndIndex
    Else
        eI = UBound(sArray)
    End If
        
    For P = sI To eI
        R$ = IIf(FindExact, sArray(P), LCase$(sArray(P)))
        X = InStr(R$, Find)
        
        If (FindAtStart And X = 1) Or _
           (Not FindAtStart And X > 0) Then FindPhrase = P: _
                                             Exit Function
    Next P
        
Quit:
    If Err.Number <> 0 Then FindPhrase = -2
    
End Function
Public Function FillRandom(sArray As Variant, _
                           ByVal Min As Long, _
                           ByVal Max As Long, _
                           Optional AllowDouble As Boolean = False) _
                                As Boolean
    Dim P As Long, Num As Long
    Dim Ok As Boolean
    
    If (Not AllowDouble And (Max - UBound(sArray)) < Min) Or _
       Min > Max Then Exit Function
       
    For P = LBound(sArray) To UBound(sArray)
        Num = MISC.Get_Randomize(Min, Max)
        
        If AllowDouble Then
            sArray(P) = Num
        Else
            Do While ARR.Exist(sArray, Num)
                Num = MISC.Get_Randomize(Min, Max)
            Loop
            
            sArray(P) = Num
        End If
    Next P
    
    FillRandom = True
    
End Function
Public Function isArray(sArray As Variant) As Boolean
    Dim X As Long
    
    On Local Error Resume Next
    X = LBound(sArray)
    
    isArray = (Err.Number = 0)
    
End Function
Public Function ReadBinary(ByVal File As String) As Variant
    Dim FN As Integer
    Dim avarArray As Variant
    
    On Error GoTo Quit

    If F.Exist(File) Then
        FN = FreeFile()
    
        Open File For Binary As #FN
            avarArray = Space(LOF(FN))
            Get #FN, 1, avarArray
        Close #FN
    End If
    
Quit:
    If Err.Number <> 0 Then Erase avarArray
    
    ReadBinary = avarArray
    
End Function
Public Function SaveBinary(ByVal File As String, _
                           sArray As Variant) As Boolean
    Dim FN As Integer
      
    On Error GoTo Quit
  
    If F.Exist(File) Then _
        If F.fKill(File) <> 1 Then GoTo Quit
        
    FN = FreeFile()
    
    Open File For Binary As #FN
        Put #FN, 1, sArray
    Close #FN
  
    SaveBinary = True
    
Quit:
End Function
Public Function VarriantToString(Var() As Variant) As String()
    Dim P As Long
    Dim A() As String
    
    On Local Error GoTo Quit
    
    ReDim A(UBound(Var))
    
    For P = LBound(Var) To UBound(Var)
        A(P) = CStr(Var(P))
    Next P
    
Quit:
    If Err.Number <> 0 Then Erase A
    VarriantToString = A
    
End Function
