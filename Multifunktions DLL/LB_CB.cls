VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LB_CB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SearchLB Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SearchCB Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetHscroll Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal Hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Private Type COMBOBOXINFO
    cbSize       As Long
    rcItem       As RECT
    rcButton     As RECT
    stateButton  As Long
    hwndCombo    As Long
    hwndEdit     As Long
    hwndList     As Long
End Type

Private Type vLBNFO
    Hwnd As Long
    List() As String
End Type

Private Type acHwndNFO
    Hwnd    As Long
End Type

Public Enum rsCommand
    asTEXT = 0
    asBINARY = 1
End Enum

Public Enum vlbCommand
    vlLOAD = 1
    vlSAVE = 2
    vlCLEAR = 3
End Enum

Public Enum RegReadOption
    rrEnum_KEY = 0
    rrEnum_VALUE = 1
End Enum

Public Enum lbDescendingOption
    lbAscending = 0
    lbDescending = 1
End Enum

Private Const CB_ADDSTRING = &H143
Private Const CB_FINDSTRINGEXACT As Long = &H158
Private Const CB_GETCURSEL = &H147
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_GETCOUNT = &H146
Private Const CB_INSERTSTRING = &H14A
Private Const CB_SETCURSEL = &H14E
Private Const CB_RESETCONTENT = &H14B
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_DELETESTRING = &H144
Private Const CB_FINDSTRING = &H14C

Private Const EM_LIMITTEXT = &HC5

Private Const LB_ADDSTRING = &H180
Private Const LB_GETCOUNT = &H18B
Private Const LB_GETTEXTLEN = &H18A
Private Const LB_GETTEXT = &H189
Private Const LB_GETCURSEL = &H188
Private Const LB_INSERTSTRING = &H181
Private Const LB_SETCURSEL = &H186
Private Const LB_RESETCONTENT = &H184
Private Const LB_DELETESTRING = &H182
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const LB_ITEMFROMPOINT = &H1A9
Private Const LB_SETHORIZONTAL = &H194
Private Const LB_FINDSTRING = &H18F
Private Const LB_GETTOPINDEX = &H18E
Private Const LB_GETITEMHEIGHT = &H1A1
Private Const LB_SETSEL As Long = &H185

Dim vLB() As vLBNFO
Private Function Cis(Hwnd As Long, Command As Long) As Long
    Dim Class As String
        
    Class = LCase(WIN.Get_Class(Hwnd))
    
    If InStr(Class, "listbox") And InStr(Class, "drivelist") = 0 Then
        Select Case Command
            Case 1: Cis = LB_GETCURSEL
            Case 2: Cis = LB_SETCURSEL
            Case 3: Cis = LB_INSERTSTRING
            Case 4: Cis = LB_DELETESTRING
            Case 5: Cis = LB_RESETCONTENT
            Case 6: Cis = LB_FINDSTRINGEXACT
            Case 7: Cis = LB_FINDSTRING
            Case 8: Cis = LB_SETSEL
        End Select
        
        Exit Function
    End If
    
    If InStr(Class, "combobox") Or InStr(Class, "drivelist") Then
        Select Case Command
            Case 1: Cis = CB_GETCURSEL
            Case 2: Cis = CB_SETCURSEL
            Case 3: Cis = CB_INSERTSTRING
            Case 4: Cis = CB_DELETESTRING
            Case 5: Cis = CB_RESETCONTENT
            Case 6: Cis = CB_FINDSTRINGEXACT
            Case 7: Cis = CB_FINDSTRING
        End Select
        
        Exit Function
    End If
    
    Cis = -2
    
End Function
Public Function CB_SetDropDownState(ByVal Hwnd As Long, _
                                    Optional isDropped As Boolean = False) _
                                            As Boolean
    
    Call SendMessage(Hwnd, CB_SHOWDROPDOWN, ByVal CLng(isDropped) * -1, 0)
    
    CB_SetDropDownState = CB_GetDropDownState(Hwnd)
    
End Function
Public Function CB_GetDropDownState(ByVal Hwnd As Long) As Boolean
    Dim Result As Long
    
    Result = SendMessage(Hwnd, CB_GETDROPPEDSTATE, 0, 0)
    CB_GetDropDownState = CBool(Result)
    
End Function
Public Function Get_ListIndex(Hwnd As Long) As Long
    Dim Cmd As Long
    
    Cmd = Cis(Hwnd, 1)
    If Cmd = -2 Then Get_ListIndex = Cmd: _
                     Exit Function
    
    Get_ListIndex = SendMessage(Hwnd, Cmd, ByVal CLng(0), ByVal CLng(0))

End Function
Public Function Set_ListIndex(ByVal Hwnd As Long, _
                              ByVal Index As Integer) _
                                    As Boolean
    Dim Cmd As Long
    
    Cmd = Cis(Hwnd, 2)
    
    If Cmd <> -2 Then _
        Call SendMessageLong(Hwnd, Cmd, Index, 0): _
        Set_ListIndex = (Get_ListIndex(Hwnd) = Index)
        
End Function
Public Sub AddItem(ByVal Hwnd As Long, _
                   ByVal Index As Long, _
                   ByVal Text As String)
    Dim Cmd As Long
    
    Cmd = Cis(Hwnd, 3)
    
    If Cmd <> -2 Then _
        Call SendMessage(Hwnd, Cmd, ByVal Index, ByVal Text)

End Sub
Public Sub DeleteItem(ByVal Hwnd As Long, _
                      ByVal Index As Long)
    Dim Cmd As Long
    
    Cmd = Cis(Hwnd, 4)
    If Cmd = -2 Then Exit Sub
    
    Call SendMessage(Hwnd, Cmd, ByVal Index, ByVal CLng(0))

End Sub
Public Sub Clear(ByVal Hwnd As Long)
    Dim Cmd As Long
    
    Cmd = Cis(Hwnd, 5)
    
    If Cmd <> -2 Then _
        Call SendMessage(Hwnd, Cmd, _
                         ByVal CLng(0), ByVal CLng(0))
    
End Sub
Public Function Get_SystemFonts(LBCB As Object) As Boolean
    Dim X As Integer
    
    On Local Error GoTo Quit
    
    LBCB.Clear
    
    For X = 0 To Screen.FontCount - 1
        LBCB.AddItem Screen.Fonts(X)
    Next X

    Get_SystemFonts = True
    
Quit:
End Function
Public Function FindItem(LBCB As Object, _
                         ByVal Find As String, _
                         Optional FindExact As Boolean = False) _
                                As Long
    Dim P As Long, Cmd As Long
    
    On Local Error GoTo Quit
    
    Cmd = Cis(LBCB.Hwnd, 6)
    If Cmd = -2 Then FindItem = Cmd: Exit Function
                     
    FindItem = SearchLB(LBCB.Hwnd, Cmd, -1, Find)
        
    If Not FindExact Then Find = LCase$(Find)
    
    For P = FindItem To LBCB.ListCount - 1
        If IIf(FindExact, LBCB.List(P), LCase$(LBCB.List(P))) = Find Then _
            FindItem = P: _
            Exit Function
    Next P
        
    FindItem = -1
    
Quit:
    If Err Then FindItem = -3
    
End Function
Public Function Clone(LBCBsource As Object, _
                      LBCBtarget As Object, _
                      Optional ClearSource As Boolean = False) _
                            As Boolean
    Dim P As Integer
    
    On Local Error GoTo Quit

    LBCBtarget.Clear
    
    For P = 0 To LBCBsource.ListCount - 1
        LBCBtarget.AddItem LBCBsource.List(P)
        LBCBtarget.ItemData(P) = LBCBsource.ItemData(P)
    Next P
    
    LBCBtarget.ToolTipText = LBCBsource.ToolTipText

    If ClearSource Then LBCBsource.Clear
    
    Clone = True
    
Quit:
End Function
Public Function Sort(LBCB As Object, _
                     Descending As lbDescendingOption) As Boolean
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    If LBCB.Sorted Then Exit Function
    
    Select Case Descending
        Case lbDescending
            Load frm
            frm.Tag = "Sort"
    
            Sort = frm.SortLB(LBCB)
    
            Unload frm
            Set frm = Nothing
        Case Is = lbAscending
            Sort = LBSortItem(LBCB, True)
    End Select
    
    Sort = True
    
Quit:
End Function
Public Sub Get_MaxValue(LBCB As Object, _
                        MaxStr As String, _
                        MaxIndex As Long)
    Dim P As Long

    MaxStr = ""
    MaxIndex = -1

    For P = 0 To LBCB.ListCount - 1
        If LBCB.List(P) > MaxIndex Then _
            MaxStr = LBCB.List(P): _
            MaxIndex = P
    Next P
    
End Sub
Public Sub Get_MinValue(LBCB As Object, _
                        MinStr As String, _
                        MinIndex As Long)
    Dim P As Long
        
    If LBCB.ListCount Then
        MinStr = LBCB.List(0)
        MinIndex = 0
        
        For P = 1 To LBCB.ListCount - 1
            If LBCB.List(P) < MinStr Then _
                MinStr = LBCB.List(P): _
                MinIndex = P
        Next P
    Else
        MinStr = ""
        MinIndex = -1
    End If
    
End Sub
Public Function Kill_DoubleItems(LBCB As Object, _
                                 Optional FindExact As Boolean = False) _
                                        As Boolean
    Dim P As Long
    Dim LBitem() As String, Find As String
    Dim LBID() As Long, Cmd As Long

    On Local Error GoTo Quit
    
    Cmd = Cis(LBCB.Hwnd, 6)
    If Cmd = -2 Then Exit Function
                     
    If LBCB.ListCount = 0 Then _
        Kill_DoubleItems = True: _
        Exit Function
    
    ReDim LBitem(0)
    ReDim LBID(0)
        
    Do While LBCB.ListCount
        ReDim Preserve LBitem(UBound(LBitem) + 1)
        ReDim Preserve LBID(UBound(LBID) + 1)
        
        LBitem(UBound(LBitem)) = LBCB.List(0)
        LBID(UBound(LBID)) = LBCB.ItemData(0)
        
        Find = LBitem(UBound(LBitem))
        
        LBCB.RemoveItem 0
                
        P = SearchLB(LBCB.Hwnd, Cmd, -1, Find)

        If (Not FindExact) _
        Or (FindExact And LBCB.List(P) = Find) Then
            Do While P > -1
                LBCB.RemoveItem P
                P = SearchLB(LBCB.Hwnd, Cmd, -1, Find)
            Loop
        End If
    Loop
    
    LBCB.Clear
    
    For P = 1 To UBound(LBitem)
        LBCB.AddItem LBitem(P)
        LBCB.ItemData(LBCB.NewIndex) = LBID(P)
    Next P
    
    Erase LBitem
    Erase LBID
    
    Kill_DoubleItems = True
    
Quit:
End Function
Public Sub CB_SetAlignment(Hwnd As Long, _
                           ByVal RightAlignment As Boolean)
    Dim hList As Long, nStyle As Long

    hList = GetComboListHandle(Hwnd)

    If hList Then
        nStyle = GetWindowLong(hList, GWL_EXSTYLE)
      
        nStyle = IIf(RightAlignment, _
                     nStyle Or WS_EX_RIGHT, _
                     nStyle And Not WS_EX_RIGHT)
        
        Call SetWindowLong(hList, GWL_EXSTYLE, nStyle)
    End If
   
End Sub
Private Function GetComboListHandle(Hwnd As Long) As Long

   Dim CBI As COMBOBOXINFO

   CBI.cbSize = Len(CBI)
   Call GetComboBoxInfo(Hwnd, CBI)
   GetComboListHandle = CBI.hwndList

End Function
Public Function LB_MouseMove(LB As Object, _
                             ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single, _
                             Optional SetMousePointer As Boolean = True, _
                             Optional SetToolTip As Boolean = True) _
                                    As Long
    Dim Index As Long, MP As Long
    Dim TT$
    
    On Local Error GoTo Quit
    
    Index = LB_GetIndexUnderMouse(LB, Button, Shift, X, Y)
    
    If Index >= LB.ListCount Then
        Index = -1
        MP = 12
        TT$ = ""
    Else
        MP = 0
        TT$ = LB.List(Index)
    End If
    
    LB_MouseMove = Index
    LB.ListIndex = Index
    
    If SetMousePointer Then LB.MousePointer = MP
    If SetToolTip Then LB.ToolTipText = TT$
    
    Exit Function
    
Quit:
    LB_MouseMove = -2

End Function
Public Function AddFilesFromPath(LBCB As Object, _
                                 ByVal Path As String) _
                                       As Boolean
    Dim D() As String, F() As String
    Dim P As Integer
        
    On Local Error GoTo Quit
    
    LBCB.Clear
    
    Call DI.ReadDirectory(Path, D, F)
    
    AddFilesFromPath = AddArray(LBCB, F)
    
Quit:
End Function
Public Function AddDirsFromPath(LBCB As Object, _
                                ByVal Path As String) _
                                    As Boolean
    Dim D() As String, F() As String
            
    On Local Error GoTo Quit
    
    LBCB.Clear
    
    Call DI.ReadDirectory(Path, D, F, True, True, True)
    
    AddDirsFromPath = AddArray(LBCB, D)
    
Quit:
End Function
Public Function AddArray(LBCB As Object, _
                         sArray As Variant, _
                         Optional ClearBeforeAdding As Boolean = True) _
                                As Boolean
    Dim P As Long
    
    On Local Error GoTo Quit
    
    If ClearBeforeAdding Then LBCB.Clear
    
    For P = LBound(sArray) To UBound(sArray)
        LBCB.AddItem sArray(P)
    Next P

    AddArray = True

Quit:
End Function
Public Function AddProcessList(LBCB As Object, _
                               Optional ClearBeforAdding As Boolean = True) _
                                    As Boolean
    Dim E() As String
    Dim I() As Long, P As Long
    
    On Local Error GoTo Quit
    
    If SYS.Get_ProcessList(E, I) Then _
        AddProcessList = AddArray(LBCB, E, _
                                  ClearBeforAdding)
    
Quit:
End Function
Public Function ToArray(LBCB As Object, _
                        sArray As Variant, _
                        Optional WriteSorted As Boolean = False) _
                                As Boolean
    Dim Lc As Long, P As Long
        
    On Local Error GoTo Quit
    
    Lc = LBCB.ListCount - 1
    
    ReDim sArray(Lc)
    
    For P = 0 To Lc
        sArray(P) = LBCB.List(P)
    Next P
    
    If WriteSorted Then _
        If Not ARR.Sort(sArray) Then Exit Function
    
    ToArray = True
    
Quit:
    If Err Then Erase sArray
    
End Function
Public Function AddItem_WithoutDouble(LBCB As Object, _
                    ByVal StringToAdd As String, _
                    Optional ItemData As Variant, _
                    Optional FindExact As Boolean = True) _
                                    As Long
    
    AddItem_WithoutDouble = FindItem(LBCB, StringToAdd, FindExact)
                                     
    If AddItem_WithoutDouble = -1 Then
        LBCB.AddItem StringToAdd
        AddItem_WithoutDouble = LBCB.NewIndex
        
        If Not IsMissing(ItemData) Then _
            LBCB.ItemData(LBCB.NewIndex) = ItemData
    Else
        AddItem_WithoutDouble = -1
    End If
    
End Function
Public Function Replace(LBCB As Object, _
                        ByVal SearchFor As String, _
                        ByVal ReplaceWith As String, _
                        Optional FindExact As Boolean = True) _
                            As Boolean
    Dim P As Long, Cmd As Long

    On Local Error GoTo Quit
    
    If SearchFor = ReplaceWith Then Exit Function
    
    Cmd = Cis(LBCB.Hwnd, 6)
    If Cmd = -2 Then Exit Function
    
    If Not FindExact Then
        SearchFor = LCase$(SearchFor)
        
        For P = 0 To LBCB.ListCount - 1
            If LCase$(LBCB.List(P)) = SearchFor Then _
                LBCB.List(P) = ReplaceWith
        Next P
    Else
        For P = 0 To LBCB.ListCount - 1
            If LBCB.List(P) = SearchFor Then _
                LBCB.List(P) = ReplaceWith
        Next P
    End If
    
    Replace = True
    
Quit:
End Function
Public Function vListBox(LBCB As Object, _
                         ByVal Command As vlbCommand, _
                         Optional Index As Long = 0) As Long
    Dim Hwnd As Long, P As Long
    Dim Found As Boolean
    
    On Local Error Resume Next
    
    P = UBound(vLB)
    If Err Then ReDim vLB(0): Err.Clear
    
    On Local Error GoTo Quit
    
    If Index = 0 Or (Command = vlSAVE) Then
        For P = 1 To UBound(vLB)
            If vLB(P).Hwnd = LBCB.Hwnd Then Found = True: _
                                       Exit For
        Next P
    
        If Not Found Then _
            ReDim Preserve vLB(UBound(vLB) + 1): _
            P = UBound(vLB)
                
        vListBox = P
    Else
        If Index < 1 Or Index > UBound(vLB) Then GoTo Quit
        vListBox = Index
    End If
    
    Select Case Command
        Case vlSAVE
            vLB(vListBox).Hwnd = LBCB.Hwnd
            
            ReDim vLB(vListBox).List(LBCB.ListCount - 1, 2)
            
            For P = 0 To LBCB.ListCount - 1
                vLB(vListBox).List(P, 1) = LBCB.List(P)
                vLB(vListBox).List(P, 2) = LBCB.ItemData(P)
            Next P
        
        Case vlLOAD
            LBCB.Clear
            P = 0
            
            On Local Error Resume Next
            
            Do
                LBCB.AddItem vLB(vListBox).List(P, 1)
                LBCB.ItemData(LBCB.NewIndex) = CLng(vLB(vListBox).List(P, 2))
                
                If Err Then Exit Do
                
                P = P + 1
            Loop
            
            vListBox = 0
        Case vlCLEAR
            Erase vLB
    End Select
    
    Exit Function

Quit:
    vListBox = -1
    
End Function
Public Function CB_MaxLength(ComboBox As Object, _
                             ByVal Lenght As Long) _
                                As Boolean
    Dim hWndCmb As Long
    Dim R$
    
    On Local Error GoTo Quit
    
    With ComboBox
        hWndCmb = FindWindowEx(.Hwnd, 0&, vbNullString, vbNullString)
  
        R$ = .Text
        .Text = ""
        
        Call SendMessage(hWndCmb, EM_LIMITTEXT, Lenght, 0&)
    
        If Len(R$) Then
            If Len(R$) < (Lenght + 1) Then
                .Text = R$
            Else
                R$ = Mid(R$, 1, Lenght)
            End If
                
            ComboBox.Text = R$
            ComboBox.SelStart = Len(R$)
        End If
    End With
    
    CB_MaxLength = True

Quit:
End Function
Public Sub CB_SetText(ByVal Hwnd As Long, ByVal Text As String)
    Dim hEdit As Long
        
    hEdit = FindWindowEx(Hwnd, 0&, "Edit", vbNullString)
    If hEdit Then Call WIN.Set_Text(hEdit, Text)

End Sub
Public Function ToString(LBCB As Object) As String
    Dim R$
    Dim P As Long
    
    On Local Error GoTo Quit
    
    For P = 0 To LBCB.ListCount - 1
        R$ = R$ & LBCB.List(P) & Chr$(0)
    Next P
    
    If Len(R$) Then R$ = Left$(R$, Len(R$) - 1)
    
Quit:
    If Err Then R$ = ""
    ToString = R$
    
End Function
Public Function SaveToFile(LBCB As Object, _
                           ByVal File As String, _
                           SaveAs As rsCommand) As Boolean
    Dim R$, A() As String
    Dim P As Long
    
    Select Case SaveAs
        Case asTEXT
            For P = 0 To LBCB.ListCount - 1
                R$ = R$ & LBCB.List(P) & vbCrLf
            Next P
    
            If Len(R$) Then R$ = Left$(R$, Len(R$) - 2)
    
            SaveToFile = TXT.TextFile_WriteText(File, R$, True)
        Case asBINARY
            If ToArray(LBCB, A, False) Then _
                SaveToFile = ARR.SaveBinary(File, A)
    End Select
        
End Function
Public Function LoadFromFile(LBCB As Object, _
                             ByVal File As String, _
                             ReadAs As rsCommand) As Boolean
    Dim A() As String, R$
    
    On Local Error GoTo Quit
    
    Select Case ReadAs
        Case asTEXT
            R$ = TXT.TextFile_Read(File)
            A = Split(R$, vbCrLf)
        Case asBINARY
            A = ARR.ReadBinary(File)
    End Select
    
    LoadFromFile = AddArray(LBCB, A, True)

Quit:
    Erase A
    
End Function
Public Function CountItem(LBCB As Object, _
                          ByVal Item As String, _
                          Optional FindExact As Boolean = True) _
                                    As Long
    Dim P As Long, X As Long
        
    On Local Error GoTo Quit
    
    If Not FindExact Then Item = LCase$(Item)
    
    For P = 0 To LBCB.ListCount - 1
        If IIf(FindExact, LBCB.List(P), _
               LCase(LBCB.List(P))) = Item Then _
                    X = X + 1
    Next P
    
    CountItem = X
    
Quit:
    If Err.Number <> 0 Then CountItem = -1
    
End Function
Public Function ItemExist(LBCB As Object, _
                          ByVal ItemStr As String) _
                                As Boolean
    ItemExist = (FindItem(LBCB, ItemStr, True) > -1)
End Function
Public Function FindPhrase(LBCB As Object, _
                           ByVal Find As String, _
                           Optional FindExact As Boolean = True, _
                           Optional FindAtStart As Boolean = False) _
                                As Long
    Dim P As Long, Cmd As Long
    Dim Found As Boolean
        
    On Local Error GoTo Quit
    
    FindPhrase = -1
    If Not FindExact Then Find = LCase$(Find)
    
    If FindAtStart Then
        Cmd = Cis(LBCB.Hwnd, 7)
        If Cmd = -2 Then FindPhrase = Cmd: _
                         Exit Function
     
        FindPhrase = SearchLB(LBCB.Hwnd, Cmd, -1, Find)
    
        If Not FindExact Or FindPhrase = -1 Then Exit Function
    
        For P = FindPhrase To LBCB.ListCount - 1
            If InStr(LBCB.List(P), Find) = 1 Then FindPhrase = P: _
                                                  GoTo Quit
        Next P
        
        FindPhrase = -1
    Else
        For P = 0 To LBCB.ListCount - 1
            If InStr(IIf(FindExact, LBCB.List(P), _
                         LCase$(LBCB.List(P))), Find) <> 0 Then _
                                FindPhrase = P: _
                                GoTo Quit
        Next P
    End If
    
Quit:
    If Err.Number <> 0 Then FindPhrase = -3
    CBacWork = False
    
End Function
Public Function CB_AutoComplete(CB As Object, _
                                Optional CompleteFirstWordOnly As Boolean = False) _
                                        As Boolean
    Dim T$, fWORD As String, bWORD As String, Word As String
    Dim lWORD As String
    Dim P As Long, X As Long, Pos As Long
    
    Static CBacWork As Boolean
    
    On Local Error GoTo Quit
    
    If CBacWork Then Exit Function
    CBacWork = True
    
    T$ = CB.Text
    Pos = CB.SelStart
        
    On Local Error Resume Next
    
    'Wörter vorher filtern
    For P = Pos To 1 Step -1
        If Mid$(T$, P, 1) = " " Then fWORD = Mid$(T$, 1, P): _
                                     T$ = Mid$(T$, P + 1): _
                                     Exit For
    Next P
    
    'Hintere Wörter filtern
    For P = Pos To Len(T$)
        If Mid$(T$, P, 1) = " " Then bWORD = Mid$(T$, P): _
                                     T$ = Mid$(T$, 1, P - 1): _
                                     Exit For
    Next P
    
    'Eingegebenes Wort
    T$ = Trim$(T$)
    If Len(T$) <> 1 Then GoTo Ok
    
    'Versuchen Word zu finden
    X = LB.FindPhrase(CB, T$, True, True)
    If X = -1 Then GoTo Ok
    
    Word = CB.List(X)
    
    If Word = T$ Then GoTo Ok
    
    T$ = fWORD & Word & bWORD
    CB.Text = T$
    
    CB.SelStart = Pos
    CB.SelLength = Len(Word) - (Pos - Len(fWORD))
    
Ok:
    CB_AutoComplete = True
    
Quit:
    CBacWork = False
    
End Function
Public Function CB_SetDropDownHeight(ComboBox As Object, _
                                     ByVal newHeight As Long) _
                                            As Boolean
    Dim ParentFrm As Form
    Dim Result As Long
    
    On Local Error GoTo Quit
    
    With ComboBox
        Set ParentFrm = .Parent
        
        Result = MoveWindow(.Hwnd, _
                            ParentFrm.ScaleX(.Left, .Parent.ScaleMode, vbPixels), _
                            ParentFrm.ScaleY(.Top, .Parent.ScaleMode, vbPixels), _
                            ParentFrm.ScaleX(.Width, .Parent.ScaleMode, vbPixels), _
                            newHeight, 1)
    End With
    
    CB_SetDropDownHeight = (Result <> 0)

Quit:
    Set ParentFrm = Nothing
    
End Function
Public Function LB_ScrollBarIsVisible(Hwnd As Long, _
                                      ByVal ScrollBarIsVisible As Boolean) _
                                            As Boolean
    Dim Result As Long
    
    Const SB_VERT = 1
    
    On Local Error Resume Next

    Result = ShowScrollBar(Hwnd, SB_VERT, _
                           CLng(ScrollBarIsVisible))
                       
    LB_ScrollBarIsVisible = (Result <> 0)
    
End Function
Public Function LB_GetSelectedItems(ListBox As Object) As Long()
    Dim A() As Long, P As Long
    
    On Local Error GoTo Quit
    
    For P = 0 To ListBox.ListCount - 1
        If ListBox.Selected(P) Then _
            Call ARR.Add(A, CStr(P), ADDTOEND)
    Next P
    
    LB_GetSelectedItems = A
    
Quit:
    Erase A
    
    If Err.Number <> 0 Then Erase LB_GetSelectedItems
    
End Function
Public Function LB_GetSelectedItemsText(ListBox As Object) As String()
    Dim A() As Long, P As Long
    Dim R() As String
    
    On Local Error GoTo Quit
    
    A = LB_GetSelectedItems(ListBox)
    
    If ARR.isArray(A) Then
        ReDim R(UBound(A))
        
        For P = 0 To UBound(A)
            R(P) = ListBox.List(A(P))
        Next P
    End If
    
    LB_GetSelectedItemsText = R
    
Quit:
    Erase A
    Erase R
    
    If Err.Number <> 0 Then Erase LB_GetSelectedItemsText
    
End Function
Public Function LB_RemoveSelectedItems(ListBox As Object) As Boolean
    Dim P As Long
    
    On Local Error GoTo Quit
    
    P = 0
    
    Do While P < ListBox.ListCount
        If ListBox.Selected(P) Then
            ListBox.RemoveItem P
        Else
            P = P + 1
        End If
    Loop
    
    LB_RemoveSelectedItems = True
    
Quit:
End Function
Public Function LB_GetItemUnderMouse(LB As Object, _
                                     ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single) _
                                        As String
    Dim Index As Long
    
    On Local Error GoTo Quit
    
    Index = LB_GetIndexUnderMouse(LB, Button, Shift, X, Y)
    If Index > -1 Then LB_GetItemUnderMouse = LB.List(Index)

Quit:
End Function
Public Function LB_GetIndexUnderMouse(LB As Object, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single) _
                                            As Long
    Dim Index As Long, lX As Long, lY As Long, Param As Long
    
    On Local Error GoTo Quit
    
    lX = LB.Parent.ScaleX(X, vbTwips, vbPixels)
    lY = LB.Parent.ScaleY(Y, vbTwips, vbPixels)
    
    Param = CLng(lX) + &H10000 * CLng(lY)

    Index = SendMessage(LB.Hwnd, LB_ITEMFROMPOINT, 0, ByVal Param)

    If Index > LB.ListCount Then Index = -1
    
    LB_GetIndexUnderMouse = Index
    
Quit:
End Function
Public Function LB_ListEdit(LB As Object, _
                            Optional AddInputToLB As Boolean = False, _
                            Optional StaticEditBox As Boolean = False, _
                            Optional MaxInputLength As Long = 0, _
                            Optional AllowEmpty As Boolean = True, _
                            Optional CancelOnESC As Boolean = False) _
                                    As String
    Dim I As Long, nIndex As Long, Hwnd As Long
    Dim nTopIndex As Long, nItemHeight As Long
    Dim Work As Boolean
    Dim frm As Form
        
Start:
    On Local Error GoTo Quit
    
    Hwnd = LB.Parent.Hwnd

    nTopIndex = SendMessage(LB.Hwnd, LB_GETTOPINDEX, 0, 0)
    nItemHeight = SendMessage(LB.Hwnd, LB_GETITEMHEIGHT, 0, 0) * Screen.TwipsPerPixelY

    nIndex = LB.ListIndex

    If (nIndex > LB.ListCount) Or nIndex < 0 Then _
        LB_ListEdit = Chr$(0): _
        Exit Function
        
    If Not Work Then
        For I = 0 To Forms.Count - 1
            If Forms(I).Tag = "LB_ListEdit" & CStr(LB.Hwnd) Then
                If StaticEditBox Then
                    Call Set_ListIndex(LB.Hwnd, -1)
                    Call WIN.Set_Focus(Forms(I).uInput.Hwnd)
                Else
                    Forms(I).uInput.Tag = "new"
                End If
                
                LB_ListEdit = Chr$(0)
                Exit Function
            End If
        Next I
    
        Set frm = New MyControls
        Load frm
                
        frm.ListEdit = True
        frm.ListEditESC = CancelOnESC
        frm.ListEditEmpty = AllowEmpty
    End If
    
    Call Set_ListIndex(LB.Hwnd, -1)

    With frm
        .Tag = "LB_ListEdit" & CStr(LB.Hwnd)
        .ScaleMode = LB.Parent.ScaleMode

        With .uInput
            If Not Work Then
                Work = True
                
                If MaxInputLength > 0 Then .MaxLength = MaxInputLength
                
                .Font = LB.Font
                .FontSize = LB.FontSize
                .Width = LB.Width
            
                Call WIN.Set_Parent(.Hwnd, LB.Hwnd)
            
                .Left = 0
            
                .Height = nItemHeight
                .Visible = True
            End If
            
            .Top = (nIndex - nTopIndex) * nItemHeight
            .Text = LB.List(nIndex)
            
            .SelStart = Len(.Text)
            
            .ZOrder
            .Refresh
            
            Call WIN.Set_Focus(.Hwnd)
        End With
        
        Do
            If WIN.Get_TaskID(Hwnd) = 0 Then
                Exit Do
            Else
                If .uInput.Tag = "new" Then .uInput.Tag = "": _
                                            GoTo Start
            
                If .uStringWasSet Then
                    If AddInputToLB Then LB.List(nIndex) = .uString
        
                    LB_ListEdit = .uString
        
                    Exit Do
                End If
            End If
            
            DoEvents
        Loop
    End With
    
Quit:
    Unload frm
    Set frm = Nothing

End Function
