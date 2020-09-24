VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Type CHARFORMAT2
    cbSize                  As Integer
    wPad1                   As Integer
    dwMask                  As Long
    dwEffects               As Long
    yHeight                 As Long
    yOffset                 As Long
    crTextColor             As Long
    bCharSet                As Byte
    bPitchAndFamily         As Byte
    szFaceName(0 To 32 - 1) As Byte
    wPad2                   As Integer
    wWeight                 As Integer
    sSpacing                As Integer
    crBackColor             As Long
    lLCID                   As Long
    dwReserved              As Long
    sStyle                  As Integer
    wKerning                As Integer
    bUnderlineType          As Byte
    bAnimation              As Byte
    bRevAuthor              As Byte
    bReserved1              As Byte
End Type

Public Enum TextBoxOperationNFO
    tbCUT = 1
    tbCOPY = 2
    tbPASTE = 3
    tbREPLACE = 4
    tbUNDO = 5
End Enum

Private Const EM_UNDO = &HC7
Private Const EM_REPLACESEL = &HC2
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETLINE = &HC4
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_SETCHARFORMAT = (&H400 + 68)

Private Const MAX_CHAR_PER_LINE = &H100

Private Const SCF_SELECTION = &H1&

Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Public Function AddColoredText(ByRef RTFBox As Object, _
                               ByVal Text As String, _
                               ByVal Color As Long) As Boolean
    Dim lngLength As Long, lngSelStart As Long
    
    On Local Error GoTo Quit
    
    lngLength = Len(Text)
    
    With RTFBox
        lngSelStart = .SelStart

        .SelLength = 0
        .SelText = Text
        .SelStart = lngSelStart
        .SelLength = lngLength
        .SelColor = Color
        .SelLength = 0
        .SelStart = lngSelStart + lngLength
    End With
    
    AddColoredText = True
    
Quit:
End Function
Public Function InsertPicture(ByRef RTF As Object, _
                              ByRef Pic As IPictureDisp) _
                                    As Boolean

    On Local Error GoTo Quit
    
    Clipboard.Clear
    Clipboard.SetData Pic

    Call SendMessage(RTF.hWnd, WM_PASTE, 0, 0)
    
    Clipboard.Clear
    
    InsertPicture = True
    
Quit:
End Function
Public Function Get_LineCount(ByRef RTFtxt As Object) As Long
    On Local Error Resume Next
    Get_LineCount = SendMessageB(RTFtxt.hWnd, EM_GETLINECOUNT, 0, 0&)
End Function
Public Sub Get_CursorPos(ByRef RTFtxt As Object, _
                         Z As Long, S As Long)
    On Local Error Resume Next
    
    Z = SendMessage(RTFtxt.hWnd, EM_LINEFROMCHAR, _
                    RTFtxt.SelStart, 0&) + 1
  
    S = RTFtxt.SelStart - SendMessage(RTFtxt.hWnd, _
                                      EM_LINEINDEX, Z - 1, 0&) + 1

End Sub
Public Function Get_Line(ByRef RTFtxt As Object, _
                         ByVal Line As Long) As String
    Dim Lo As Long, Hi As Long, Result As Long
    Dim Buffer As String

    On Local Error GoTo Quit
    
    Line = Line - 1

    Lo = MAX_CHAR_PER_LINE And &HFF
    Hi = Int(MAX_CHAR_PER_LINE / &H100)
    
    Buffer = Chr$(Lo) & Chr$(Hi) & Space$(MAX_CHAR_PER_LINE - 2)
    
    Result = SendMessageB(RTFtxt.hWnd, EM_GETLINE, Line, Buffer)
    Get_Line = Left$(Buffer, Result)

Quit:
End Function
Public Sub Get_TextDimensions(ByRef RTFtxt As Object, _
                              tHeight As Long, _
                              tWidth As Long)
    Dim hDcT As Long
    Dim R As RECT
  
    On Local Error GoTo Quit
  
    hDcT = RTFtxt.Parent.hDC
    
    Call DrawText(hDcT, CStr(RTFtxt.Text), -1, R, DT_CALCRECT)
    
    tHeight = R.Bottom
    tWidth = R.Right
        
Quit:
End Sub
Public Function RTFtxt_Operation(RTFtxtBox As Object, _
                                 Operation As TextBoxOperationNFO, _
                                 Optional TextToReplace As String) _
                                        As Boolean
    Dim hWnd As Long, wMsg As Long
    
    On Local Error GoTo Quit
    
    hWnd = RTFtxtBox.hWnd
    
    Select Case Operation
        Case tbREPLACE
            Call SendMessageString(hWnd, EM_REPLACESEL, 1, TextToReplace)
            RTFtxt_Operation = True
            Exit Function
        Case tbCUT:   wMsg = WM_CUT
        Case tbCOPY:  wMsg = WM_COPY
        Case tbPASTE: wMsg = WM_PASTE
        Case tbUNDO:  wMsg = EM_UNDO
    End Select
    
    Call SendMessageLong(hWnd, wMsg, 0, 0)

    RTFtxt_Operation = True
    
Quit:
End Function
Public Sub SetSelForeColor(RTF As Object, _
                           ByVal Color As Long)
    Dim I As Long, C As Long

    With RTF
        On Error Resume Next
         
        C = .SelColor
         
        If .SelLength = 0 Then Exit Sub
         
        I = .SelLength
        
        .SelColor = Color
        .SelLength = 0
        .SelStart = .SelStart + I
        .SelColor = C
    End With

End Sub
Public Sub SetSelBackColor(RTF As Object, _
                           ByVal Color As OLE_COLOR)
    Dim udtChar As CHARFORMAT2
    Dim RTFhwnd As Long
   
    Const CFM_BACKCOLOR = &H4000000
    Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR

    RTFhwnd = RTF.hWnd
   
    udtChar.dwMask = CFM_BACKCOLOR
    udtChar.crBackColor = TranslateColor(Color)
    udtChar.cbSize = Len(udtChar)
      
    Call SendMessage(RTFhwnd, EM_SETCHARFORMAT, SCF_SELECTION, udtChar)
    
End Sub
Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then _
        TranslateColor = -1
End Function
Public Function ColorizeWords(RTF As Object, _
                              ByVal ColorizeWord As String, _
                              ByVal ForeColor As Long, _
                              Optional BackColor As Long = -1, _
                              Optional DoTextCompare As Boolean = True, _
                              Optional OnlyWholeWords As Boolean = True) _
                                        As Boolean

    Dim I As Long, J As Long
    Dim SelStart As Long, WordLen As Long, TextLen As Long
    Dim DoIt As Boolean
    
    On Local Error GoTo Quit
    
    SelStart = RTF.SelStart
    WordLen = Len(ColorizeWord)
    TextLen = Len(RTF.Text)
    
    I = 1
      
    Do
        If DoTextCompare Then
            J = InStr(I, RTF.Text, ColorizeWord, vbTextCompare)
        Else
            J = InStr(I, RTF.Text, ColorizeWord, vbBinaryCompare)
        End If
        
        If J = 0 Then Exit Do
        
        DoIt = True
         
        If OnlyWholeWords Then
            If J > 1 Then
                If Mid$(RTF.Text, J - 1, 1) > Chr(32) Then DoIt = False
            End If
            
            If (J + WordLen) < TextLen Then
                If Mid$(RTF.Text, J + WordLen, 1) > Chr(32) Then DoIt = False
            End If
        End If
         
        RTF.SelStart = J - 1
        RTF.SelLength = WordLen
        
        If DoIt Then Call SetSelBackColor(RTF, BackColor): _
                     Call SetSelForeColor(RTF, ForeColor)
            
        I = J + WordLen
    Loop
      
    RTF.SelLength = 0
    RTF.SelStart = SelStart
            
    ColorizeWords = True
    
Quit:
End Function
Public Function sPrint(RTF As Object, _
                      Optional RightSpaceMM As Long = 25, _
                      Optional LeftSpaceMM As Long = 25) _
                            As Boolean
    
    On Local Error GoTo Quit
    
    Printer.Print
    
    With RTF
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelIndent = Printer.ScaleX(LeftSpaceMM, _
                                    vbMillimeters, _
                                    Printer.ScaleMode)
    
        .SelRightIndent = Printer.ScaleX(RightSpaceMM, _
                                         vbMillimeters, _
                                         Printer.ScaleMode)
    
        .SelPrint Printer.hDC, False
    
        .SelIndent = 0
        .SelRightIndent = 0
    End With
   
    Printer.EndDoc

    sPrint = True
    
Quit:
End Function
