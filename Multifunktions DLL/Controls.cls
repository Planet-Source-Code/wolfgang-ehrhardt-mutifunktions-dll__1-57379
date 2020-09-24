VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Controls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function smSetColor Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Enum fAligmentNFO
    faCenterBottom = 5
    faCenterCenter = 4
    faCenterTop = 3
    faGeneral = 9
    faLeftBottom = 2
    faLeftCenter = 1
    faLeftTop = 0
    faRightBottom = 8
    faRightCenter = 7
    faRightTop = 6
End Enum

Public Enum fSortNFO
    fsGenericAscending = 1
    fsGenericDescending = 2
    fsSortNone = 0
    fsNumericAscending = 3
    fsNumericDescending = 4
    fsSortStringAscending = 7
    fsStringDescending = 8
    fsStringNoCaseAscending = 5
    fsStringNoCaseDescending = 6
End Enum

Public Enum sStyle
    ssGraficStyle = 0
    ssStandardStyle = 1
End Enum

Public Enum tbEVENTnfo
    tbKEYDOWN = 0
    tbCHANGE = 1
End Enum

Public Enum moButtonNFO
    moLEFTMOUSEBUTTON = 0
    moRIGHTMOUSEBUTTON = 1
End Enum

Public Enum flSortNFO
    flSORTBYCELL = 1
    flSOFTBYMLEXGRIDEWIDTH = 2
End Enum

Public Enum scCOLOR
    scRED = 1
    scGREEN = 2
    scBLUE = 3
End Enum

Private Const SRCINVERT = &H660046
Private Const SRCAND = &H8800C6

Private Const BS_PUSHLIKE = &H1000&

Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = &H20
Public Function CommandButton_SetAlignment(CommandButton As Object, _
                                           Alignment As BS_Alignment) _
                                                As Boolean
    Dim dwnewLong As Long
    
    On Local Error GoTo Quit
    
    dwnewLong = GetWindowLongA(CommandButton.hWnd, GWL_STYLE)

    Call SetWindowLong(CommandButton.hWnd, _
                       GWL_STYLE, dwnewLong Or Alignment)
    
    CommandButton.Refresh

    CommandButton_SetAlignment = True
    
Quit:
End Function
Public Function TextBox_AllowNumOnly(TextBox As Object, _
                                     tbEvent As tbEVENTnfo, _
                                     Optional DigitsAfterDecimal As Long = 0, _
                                     Optional KeyCode As Integer, _
                                     Optional HelpOnF1 As Boolean = True, _
                                     Optional SetToolTip As Boolean = True) _
                                            As Boolean
    Dim A As String, K As String, T As String, S$, H As String, O$
    Dim X As Long, P As Long, Y As Long
    Dim Found As Boolean
    
    On Local Error GoTo Quit
    
    TextBox_AllowNumOnly = True
    
    If SetToolTip Then _
        H = "In diesem Eingabefeld sind nur Zahlen erlaubt" & _
            IIf(HelpOnF1, " (F1 für Hilfe)", ""): _
        TextBox.ToolTipText = H
        
    A = IIf((DigitsAfterDecimal > 0), ",", "")
    
    T = TextBox.Text
    O$ = T
        
    Select Case tbEvent
        Case tbCHANGE
            For P = 1 To Len(T)
                Select Case Mid$(T, P, 1)
                    Case "0" To "9"
                        S$ = S$ & Mid$(T, P, 1)
                    Case A
                        If Not Found Then S$ = S$ & Mid$(T, P, 1): _
                                          Found = True
                End Select
            Next P
            
            T = S$
            
            Do While Left$(T, 1) = ","
                T = Mid$(T, 2)
            Loop
            
            X = InStr(T, ",")
            
            If X > 0 Then _
                If Len(T) - X > DigitsAfterDecimal Then _
                    T = Mid$(T$, 1, X + DigitsAfterDecimal)
            
            If O$ <> T Then TextBox.Text = T
        Case tbKEYDOWN
            If KeyCode = 46 Or KeyCode = 8 Or _
               KeyCode = 13 Or KeyCode = 37 Or _
               KeyCode = 38 Or KeyCode = 39 Or _
               KeyCode = 40 Then Exit Function
            
            If KeyCode = 27 Then TextBox.Text = "": _
                                 Exit Function
               
            If HelpOnF1 And KeyCode = 112 Then
                T = "Eingabefeld Hilfe"
                H = "In diesem Eingabefeld sind nur Zahlen zulässig." & vbCrLf & vbCrLf & _
                    "Folgende Zeichen sind erlaubt:" & vbCrLf & vbCrLf & Num
                
                If DigitsAfterDecimal > 0 Then _
                    H = H & " sowie ein "","" und " & DigitsAfterDecimal & " Ziffern nach dem "",""."
                
                H = H & vbCrLf & vbCrLf & _
                       "ESC löscht das Eingabefeld." & vbCrLf

                MsgBox H, vbInformation + vbOKOnly, T
                
                Exit Function
            End If
                         
            K = Chr$(KeyCode)
            X = InStr(T, ",")
            
            If InStr(Num & A, K) = 0 Then
                KeyCode = 0
                
                If (K = ",") And ((X <> 0) Or (Len(T) = 0)) Then _
                    KeyCode = 0
            End If
    End Select
      
    Exit Function
    
Quit:
    TextBox_AllowNumOnly = False
    
End Function
Public Function TextBox_AllowAlphaOnly(TextBox As Object, _
                                       tbEvent As tbEVENTnfo, _
                                       Optional KeyCode As Integer, _
                                       Optional LAlpha As Boolean = True, _
                                       Optional UAlpha As Boolean = True, _
                                       Optional Space As Boolean = True, _
                                       Optional HelpOnF1 As Boolean = True, _
                                       Optional SetToolTip As Boolean = True) _
                                            As Boolean
    Dim R$, S$, T As String, H As String
    Dim P As Long
    
    On Local Error GoTo Quit
    
    TextBox_AllowAlphaOnly = True
    
    If SetToolTip Then _
        H = "In diesem Eingabefeld sind nur Alphanumerische Eingaben zulässig." & _
            IIf(HelpOnF1, " (F1 für Hilfe)", ""): _
            TextBox.ToolTipText = H

    T = TextBox.Text
    
    If LAlpha Then R$ = lABC
    If UAlpha Then R$ = R$ & uABC
    If Space Then R$ = R$ & Chr$(32)

    Select Case tbEvent
        Case tbCHANGE
            For P = 1 To Len(T)
                If InStr(R$, Mid$(T, P, 1)) > 0 Then _
                    S$ = S$ & Mid$(T, P, 1)
            Next P
            
            If T <> S$ Then TextBox.Text = S$: _
                            TextBox.SelStart = Len(S$): _
                            TextBox.SelLength = 0
        Case tbKEYDOWN
            If KeyCode = 46 Or KeyCode = 8 Or _
               KeyCode = 13 Or KeyCode = 37 Or _
               KeyCode = 38 Or KeyCode = 39 Or _
               KeyCode = 40 Then Exit Function
            
            If KeyCode = 27 Then TextBox.Text = "": _
                                 Exit Function
               
            If HelpOnF1 And KeyCode = 112 Then
                T = "Eingabefeld Hilfe"
                H = "In diesem Eingabefeld sind nur Alphanumerische Eingaben zulässig." & vbCrLf & vbCrLf & _
                    "Folgende Zeichen sind erlaubt:" & vbCrLf & _
                    vbCrLf & R$ & vbCrLf
                    
                If Space Then H = H & "Sowie Freizeichen" & vbCrLf
                
                H = H & vbCrLf & "ESC löscht das Eingabefeld." & vbCrLf
                
                MsgBox H, vbInformation + vbOKOnly, T
                
                Exit Function
            End If

            If InStr(R$, Chr$(KeyCode)) = 0 Then KeyCode = 0
    End Select
    
    Exit Function
    
Quit:
    TextBox_AllowAlphaOnly = False

End Function
Public Sub Control_Save(Control As Object, _
                        ParentForm As Object, _
                        ByVal INIpath As String)
    Dim Sec As String
    Dim X As Long
    
    On Local Error Resume Next
    
    X = -1
    X = Control.Index
                
    Sec = ParentForm.Name & "." & Control.Name
    If X > -1 Then Sec = Sec & "(" & X & ")"
        
    Call INI.Remove_Section(INIpath, Sec)

    With Control
        Call INI.Write_INI(INIpath, Sec, "Alignment", .Alignment)
        Call INI.Write_INI(INIpath, Sec, "BorderStyle", .BorderStyle)
        Call INI.Write_INI(INIpath, Sec, "Checked", CLng(.Checked) * -1)
        Call INI.Write_INI(INIpath, Sec, "Style", .Style)
       
        Call INI.Write_INI(INIpath, Sec, "Top", .Top)
        Call INI.Write_INI(INIpath, Sec, "Left", .Left)
        Call INI.Write_INI(INIpath, Sec, "Height", .Height)
        Call INI.Write_INI(INIpath, Sec, "Width", .Width)

        Call INI.Write_INI(INIpath, Sec, "BackColor", .BackColor)
        Call INI.Write_INI(INIpath, Sec, "ForeColor", .ForeColor)
        
        Call INI.Write_INI(INIpath, Sec, "FontName", .FontName)
        Call INI.Write_INI(INIpath, Sec, "FontSize", .FontSize)
        Call INI.Write_INI(INIpath, Sec, "FontBold", CLng(.FontBold) * -1)
        Call INI.Write_INI(INIpath, Sec, "FontUnderline", CLng(.FontUnderline) * -1)
        Call INI.Write_INI(INIpath, Sec, "FontItalic", CLng(.FontItalic) * -1)

        Call INI.Write_INI(INIpath, Sec, "Caption", .Caption)
        Call INI.Write_INI(INIpath, Sec, "Text", .Text)
        
        Call INI.Write_INI(INIpath, Sec, "Enabled", CLng(.Enabled) * -1)
        Call INI.Write_INI(INIpath, Sec, "Visible", CLng(.Visible) * -1)
                
        Call INI.Write_INI(INIpath, Sec, "ToolTipText", .ToolTipText)
        Call INI.Write_INI(INIpath, Sec, "Tag", .Tag)
    End With

End Sub
Public Sub Control_Load(Control As Object, _
                        ParentForm As Object, _
                        ByVal INIpath As String)
    Dim Sec As String
    Dim X As Long
    
    On Local Error Resume Next
    
    X = -1
    X = Control.Index
        
    Sec = ParentForm.Name & "." & Control.Name
    If X > -1 Then Sec = Sec & "(" & X & ")"

    With Control
        .Alignment = INI.Read_INI(INIpath, Sec, "Alignment")
        .BorderStyle = INI.Read_INI(INIpath, Sec, "BorderStyle")
        .Checked = CBool(INI.Read_INI(INIpath, Sec, "Checked"))
        .Style = INI.Read_INI(INIpath, Sec, "Style")
        
        .Top = INI.Read_INI(INIpath, Sec, "Top")
        .Left = INI.Read_INI(INIpath, Sec, "Left")
        .Height = INI.Read_INI(INIpath, Sec, "Height")
        .Width = INI.Read_INI(INIpath, Sec, "Width")

        .BackColor = INI.Read_INI(INIpath, Sec, "BackColor")
        .ForeColor = INI.Read_INI(INIpath, Sec, "ForeColor")
        
        .FontName = INI.Read_INI(INIpath, Sec, "FontName")
        .FontSize = INI.Read_INI(INIpath, Sec, "FontSize")
        .FontBold = INI.Read_INI(INIpath, Sec, "FontBold")
        .FontUnderline = INI.Read_INI(INIpath, Sec, "FontUnderline")
        .FontItalic = INI.Read_INI(INIpath, Sec, "FontItalic")

        .Caption = INI.Read_INI(INIpath, Sec, "Caption")
        .Text = INI.Read_INI(INIpath, Sec, "Text")
        
        .Enabled = CBool(INI.Read_INI(INIpath, Sec, "Enabled"))
        .Visible = CBool(INI.Read_INI(INIpath, Sec, "Visible"))
            
        .ToolTipText = INI.Read_INI(INIpath, Sec, "ToolTipText")
        .Tag = INI.Read_INI(INIpath, Sec, "Tag")
        
        .Refresh
    End With
    
End Sub
Public Function Clipboard_GetFiles() As Variant
    Dim hDrop As Long, nFiles As Long, T As Long, Result As Long, X As Long
    Dim FileName As String, A() As String
    
    If IsClipboardFormatAvailable(CF_HDROP) Then
        Call OpenClipboard(0&)
        hDrop = GetClipboardData(CF_HDROP)
        
        nFiles = DragQueryFile(hDrop, -1&, "", 0)
        
        ReDim A(nFiles)

        For T = 0 To nFiles - 1
            FileName = Space(MAX_PATH)
            Result = DragQueryFile(hDrop, T, FileName, Len(FileName))
            A(T) = Left$(FileName, Result)
        Next T
    End If

    Call CloseClipboard
    
    Clipboard_GetFiles = A
    Erase A
    
End Function
Public Function CheckBox_SetStyle(CheckBox As Object, _
                                  Style As sStyle) As Boolean
    Dim CurStyle As Long

    On Local Error GoTo Quit

    CurStyle = GetWindowLong(CheckBox.hWnd, GWL_STYLE)

    CurStyle = IIf(Style = ssGraficStyle, _
                   CurStyle Or BS_PUSHLIKE, _
                   CurStyle And (Not BS_PUSHLIKE))
    
    If SetWindowLong(CheckBox.hWnd, GWL_STYLE, CurStyle) Then _
        CheckBox.Refresh: _
        CheckBox_SetStyle = True

Quit:
End Function
Public Function OptionButton_SetStyle(OptionButton As Object, _
                                      Style As sStyle) As Boolean
    OptionButton_SetStyle = CheckBox_SetStyle(OptionButton, Style)
End Function
Public Sub ProgressBar_SetColor(ByVal hWnd As Long, _
                                ByVal ForeColor As Long, _
                                ByVal BackColor As Long)
    
    Call smSetColor(hWnd, &H409, 0&, ForeColor)
    Call smSetColor(hWnd, &H2001, 0&, BackColor)
    
End Sub
Public Function Flex_SetColAligment(Flex As Object, _
                                    Aligment As fAligmentNFO, _
                                    Optional ColIndex As Long = -1, _
                                    Optional AllCols As Boolean = False) _
                                            As Boolean
    Dim Col As Long, fStart As Long, fEnd As Long
    
    On Local Error GoTo Quit
    
    If AllCols Then
        fStart = 0
        fEnd = Flex.cols - 1
    Else
        fStart = ColIndex
        fEnd = fStart
    End If
    
    For Col = fStart To fEnd
        Flex.ColAlignment(Col) = Aligment
    Next Col
    
    Flex.Refresh
    
    Flex_SetColAligment = True

Quit:
End Function
Public Function Flex_Sort(Flex As Object, _
                          Sort As fSortNFO, _
                          ByVal ColIndex As Long) _
                            As Boolean
    Dim oRow As Long, oCol As Long
    
    On Local Error Resume Next
    
    With Flex
        oRow = .Row
        oCol = .Col
    
        .Row = 1
        .Col = ColIndex
    
        .Sort = Sort
    
        .Row = oRow
        .Col = oCol
        
        .Refresh
    End With
    
    If Err.Number = 30009 Or _
       Err.Number = 0 Then Flex_Sort = True
    
End Function
Public Function Flex_ColorRow(Flex As Object, _
                              ByVal ForeColor As Long, _
                              ByVal BackColor As Long, _
                              Optional Row As Long = -1) _
                                    As Boolean
    Dim oRow As Long, oCol As Long, sRow As Long, sCol As Long
        
    On Local Error GoTo Quit
    
    With Flex
        oRow = .Row
        oCol = .Col
        
        .Row = IIf(Row = -1, 1, Row)
        .Col = 0
        
        .RowSel = .Rows - 1
        .ColSel = .cols - 1
    
        .CellBackColor = BackColor
        .CellForecolor = ForeColor
    
        .Row = oRow
        .Col = oCol
        
        .Refresh
    End With
    
    Flex_ColorRow = True
    
Quit:
End Function
Public Function Flex_Remove(Flex As Object, _
                            ByVal Row As Long) As Boolean

    On Local Error Resume Next
    
    If Row > Flex.Rows - 1 Then Exit Function
    
    Flex.RemoveItem Row
    
    If Err.Number <> 0 Then On Local Error GoTo Quit: _
                            Flex.Rows = 1
    
    Flex_Remove = True
    
Quit:
End Function
Public Function Flex_Find(Flex As Object, _
                          ByVal Find As String, _
                          Optional FindExact As Boolean = True, _
                          Optional Col As Long = -1, _
                          Optional Row As Long = 0, _
                          Optional SearchCompleteFlex As Boolean = False) _
                                As Long
    Dim R$
    Dim X As Long, sRow As Long, sCol As Long, Index As Long
    
    On Local Error GoTo Quit
    
    Flex_Find = -1
    
    If Not FindExact Then Find = LCase$(Find)
    
    With Flex
        If Not SearchCompleteFlex Then
            If Col > -1 Then
                For sRow = 1 To .Rows - 1
                    Index = ((sRow * .cols) + Col)
                    R$ = IIf(FindExact, .TextArray(Index), _
                                        LCase$(.TextArray(Index)))
                    If R$ = Find Then Flex_Find = sRow: _
                                      Exit Function
                Next sRow
            
                Exit Function
            End If
    
            If Row > 0 Then
                For sCol = 0 To .cols - 1
                    Index = ((Row * .cols) + sCol)
                    R$ = IIf(FindExact, .TextArray(Index), _
                                        LCase$(.TextArray(Index)))
                    If R$ = Find Then Flex_Find = sCol: _
                                      Exit Function
                Next sCol
            End If
        Else
            For sRow = 1 To .Rows - 1
                X = Flex_Find(Flex, Find, FindExact, , sRow, False)
                
                If X > -1 Then Flex_Find = X: _
                               Exit Function
            Next sRow
        End If
    End With
    
    Exit Function
    
Quit:
    If Err.Number <> 0 Then Flex_Find = -2
    
End Function
Public Function Flex_DeleteValue(Flex As Object, _
                                 ByVal Value As String, _
                                 Optional Row As Long = -1, _
                                 Optional FindExact As Boolean = True) _
                                        As Boolean
    Dim Index As Long, fStart As Long, fEnd As Long, sRow As Long
    
    On Local Error GoTo Quit
    
    fStart = IIf(Row = -1, 0, Row)
    fEnd = IIf(Row = -1, Flex.Rows, Row)
    
    For sRow = fStart To fEnd
        Index = Flex_Find(Flex, Value, FindExact, sRow, , False)
        If Index > -1 Then Call Flex_Remove(Flex, Index)
    Next sRow
    
    Flex_DeleteValue = True
    
Quit:
End Function
Public Function Flex_ItemExist(Flex As Object, _
                               ByVal Value As String, _
                               Optional FindExact As Boolean = True) _
                                    As Boolean
    Dim sRow As Long, Index As Long
    
    On Local Error GoTo Quit
    
    For sRow = 1 To Flex.Rows - 1
        If Flex_RowItemExist(Flex, Value, sRow, FindExact) Then _
            Flex_ItemExist = True: _
            Exit Function
    Next sRow
    
Quit:
End Function
Public Function Flex_ColItemExist(Flex As Object, _
                                  ByVal Item As String, _
                                  ByVal Col As Long, _
                                  Optional FindExact As Boolean = True) _
                                        As Boolean
    Dim Index As Long
    
    Index = Flex_Find(Flex, Item, FindExact, Col, , False)
    Flex_ColItemExist = (Index > -1)
    
End Function
Public Function Flex_RowItemExist(Flex As Object, _
                                  ByVal Value As String, _
                                  ByVal Row As Long, _
                                  Optional FindExact As Boolean = True) _
                                        As Boolean
    Dim Index As Long
    
    Index = Flex_Find(Flex, Value, FindExact, , Row, False)
    Flex_RowItemExist = (Index > -1)
    
End Function
Public Function PictureBox_FadingText(PictureBox As Object, _
                                      ByVal Text As String, _
                                      ByVal FadeTextColorStart As Long, _
                                      ByVal FadeTextColorEnd As Long, _
                                      Optional FontName As String = "Times New Roman", _
                                      Optional FontSize As Long = 36, _
                                      Optional FontBold As Boolean = True, _
                                      Optional FontItalic As Boolean = False) _
                                                As Boolean
    Dim frm As New MyControls
    Dim P As Long, DestCX As Long, DestCY As Long, TmpBackColor As Long
            
    On Local Error GoTo Quit
    
    Load frm
    
    With frm
        .ScaleMode = PictureBox.Parent.ScaleMode
    
        For P = 1 To 3
            .PicFade(P).AutoRedraw = True
            .PicFade(P).ScaleMode = PictureBox.ScaleMode
            .PicFade(P).Width = PictureBox.Width
            .PicFade(P).Height = PictureBox.Height
        
            .PicFade(P).ScaleMode = vbTwips
        Next P
        
        .ScaleMode = vbTwips
        
        .PicFade(1).Font.Name = FontName
        .PicFade(1).Font.SIZE = FontSize
        .PicFade(1).Font.Bold = FontBold
        .PicFade(1).Font.Italic = FontItalic
  
        Set .PicFade(3).Font = .PicFade(1).Font

        Call FRMS.Fade(.PicFade(2), FadeTextColorStart, FadeTextColorEnd)

        With .PicFade(1)
            .Cls
            .BackColor = vbWhite
            .ForeColor = vbBlack
            
            frm.PicFade(1).Print Text
        End With
     
        With .PicFade(3)
            DestCX = .Width / Screen.TwipsPerPixelX
            DestCY = .Height / Screen.TwipsPerPixelY

            .Cls
          
            TmpBackColor = .BackColor
            SetBkColor .hDC, frm.PicFade(1).BackColor
          
            BitBlt .hDC, 0, 0, DestCX, DestCY, frm.PicFade(2).hDC, 0, 0, SRCINVERT
            BitBlt .hDC, 0, 0, DestCX, DestCY, frm.PicFade(1).hDC, 0, 0, SRCAND
            BitBlt .hDC, 0, 0, DestCX, DestCY, frm.PicFade(2).hDC, 0, 0, SRCINVERT
          
            SetBkColor .hDC, TmpBackColor
          
            .Refresh
            
            Set PictureBox.Picture = .Image
            PictureBox.Refresh
        End With
    End With
    
    PictureBox_FadingText = True
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Sub Label_GetTextDimensions(ByRef Label As Object, _
                                   tHeight As Long, _
                                   tWidth As Long)
    Dim hDcT As Long
    Dim R As RECT
  
    On Local Error GoTo Quit
  
    hDcT = Label.Parent.hDC
    
    Call DrawText(hDcT, CStr(Label.Caption), -1, R, DT_CALCRECT)
    
    tHeight = R.Bottom
    tWidth = R.Right
        
Quit:
End Sub
Public Function PictureBox_ScrollText(PictureBox As Object, _
                                      ByVal Text As String, _
                                      FadeColor As scCOLOR) _
                                            As Boolean
    Dim frm As New frmCredits
    
    On Local Error Resume Next
    
    Load frm
    frm.Tag = "PictureBox_ScrollText"
    
    If frm.SetScroll(PictureBox, Text, FadeColor) Then
        PictureBox_ScrollText = True
    Else
        Unload frm
    End If
    
    Set frm = Nothing
    
End Function
Public Function PictureBox_Fade(PictureBox As Object, _
                                ByVal FadeFromColor As Long, _
                                ByVal FadeToColor As Long) _
                                        As Boolean
    PictureBox_Fade = FRMS.Fade(PictureBox, _
                                FadeFromColor, FadeToColor)
    
End Function
Public Function OptionButton_GetIndex(OptionButton As Object) As Long
    Dim Index As Long

    On Local Error GoTo Quit

    For Index = OptionButton.LBound To OptionButton.UBound
        If OptionButton(Index).Value Then _
            OptionButton_GetIndex = Index: _
            Exit For
    Next Index

Quit:
    If Err.Number <> 0 Then OptionButton_GetIndex = -1
    
End Function
Public Function TextBox_AllowResize(TextBox As Object, _
                                    Optional DeactivateResizing As Boolean = False) _
                                        As Boolean
    Dim Result As Long
    
    On Local Error GoTo Quit
    
    Result = GetWindowLong(TextBox.hWnd, GWL_STYLE) Or WS_THICKFRAME
    
    If DeactivateResizing Then Result = Result - WS_THICKFRAME
    
    Call SetWindowLong(TextBox.hWnd, GWL_STYLE, Result)
    Call SetWindowPos(TextBox.hWnd, TextBox.Parent.hWnd, 0, 0, 0, 0, SWP_NOZORDER _
                      Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME)

    TextBox_AllowResize = True
    
Quit:
End Function
Public Function ListBox_AllowResize(ListBox As Object, _
                                    Optional DeactivateResizing As Boolean = False) _
                                        As Boolean
    ListBox_AllowResize = TextBox_AllowResize(ListBox, _
                                              DeactivateResizing)
End Function
Public Function CheckBox_AsOptionButton(CheckBox As Object, _
                                        ByVal Index As Long) _
                                            As Boolean
    Dim P As Long
    
    On Local Error GoTo Quit
    
    If CheckBox(Index).Value Then
        For P = CheckBox.LBound To CheckBox.UBound
            If P <> Index Then CheckBox(P).Value = 0
        Next P
    End If
    
    CheckBox_AsOptionButton = True
    
Quit:
End Function
Public Function CheckBox_GetSelected(CheckBox As Object) _
                                        As Variant
    Dim P As Long
    Dim A() As Variant
        
    On Local Error GoTo Quit
    
    For P = CheckBox.LBound To CheckBox.UBound
        If CheckBox(P).Value Then Call ARR.Add(A, P, ADDTOEND)
    Next P
    
    If Not isArray(A) Then Erase A
    
Quit:
    If Err.Number <> 0 Then Erase A
    CheckBox_GetSelected = A

    Erase A
    
End Function
Public Function Control_GetTextSize(Control As Object, _
                                    Width As Long, _
                                    Height As Long) _
                                        As Long
    Dim hWnd As Long
    
    On Local Error Resume Next
    
    hWnd = Control.hWnd
    
    If Err.Number <> 0 Then Control_GetTextSize = -2: _
                            Exit Function
                            
    Control_GetTextSize = WIN.Get_TextSize(hWnd, Width, Height)
    
End Function
Public Function Control_AutoSizeByText(Control As Object) _
                                            As Boolean
    Dim H As Long, W As Long, oS As Long
    
    On Local Error Resume Next

    If Control_GetTextSize(Control, W, H) = 1 Then
        oS = Control.Parent.ScaleMode
        Control.Parent.ScaleMode = vbPixels
        
        Control.Height = H
        Control.Width = W
        
        Control.Parent.ScaleMode = oS
        
        Control_AutoSizeByText = (Err.Number = 0)
        
        Control.Refresh
    End If
    
End Function
Public Function WebBrowser_GetSource(WebBrowser As Object, _
                                     Optional WaitIfBusy As Boolean = False) _
                                        As String
    Dim P As Long
    Dim Source As String
    
    On Local Error Resume Next
    
    If WaitIfBusy Then
        Do While WebBrowser.Busy
            If Err.Number <> 0 Then Exit Function
            DoEvents
        Loop
    End If
    
    For P = 0 To WebBrowser.Document.All.Length
        Source = WebBrowser.Document.All(P).outerHTML
        If Len(Source) Then Exit For
    Next P
    
    WebBrowser_GetSource = Source

End Function
