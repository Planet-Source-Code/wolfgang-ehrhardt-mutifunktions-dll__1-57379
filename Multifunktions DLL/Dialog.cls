VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As LPCHOOSECOLOR) As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hWnd As Long, ByVal hIcon As Long, ByVal sPath As String, ByVal sTitle As String, ByVal sPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (lpcf As CHOOSEFONT_TYPE) As Long

Private Type LPCHOOSECOLOR
    lStructSize    As Long
    hWnd           As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As String
    Flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As String
End Type

Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName       As String * 32
End Type

Private Type CHOOSEFONT_TYPE
    lStructSize       As Long
    hwndOwner         As Long
    hDC               As Long
    lpLogFont         As Long
    iPointSize        As Long
    Flags             As Long
    rgbColors         As Long
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
    hInstance         As Long
    lpszStyle         As String
    nFontType         As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin          As Long
    nSizeMax          As Long
End Type


Private Const CC_SOLIDCOLOR = &H80
Private Const CC_FULLOPEN = &H2

Private Const CF_SCREENFONTS = &H1
Private Const CF_EFFECTS = &H100
Private Const CF_INITTOLOGFONTSTRUCT = &H40

Private Const FW_BOLD = 700

Private Const SCREEN_FONTTYPE = &H2000
Public Sub Execute(ByVal Title As String, _
                   Optional Prompt As String, _
                   Optional iCon As Long)
        
    If SYS.isWindowsNT Then _
        Title = StrConv(Title, vbUnicode): _
        Prompt = StrConv(Prompt, vbUnicode)
    
   Call SHRunDialog(0, iCon, 0, Title, Prompt, 0)
    
End Sub
Public Sub Find(Optional Path As String = "C:\")
    Call ShellExecute(0, "find", Path, "", "C:\", 1)
End Sub
Public Sub WindowsShutDown()
    Call SHShutDownDialog(0)
End Sub
Public Sub About()
    frmAbout.Show
End Sub
Public Function AboutBox(ByVal hWnd As Long, _
                         ByVal InfoAbout As String, _
                         Optional Copyright As String, _
                         Optional iCon As Long) As Boolean
    Dim Result As Long
    
    Result = ShellAbout(hWnd, InfoAbout, Copyright, iCon)
    AboutBox = (Result = 1)
    
End Function
Public Function Ask(ByVal hWnd As Long, _
                    ByVal Prompt As String, _
                    ByVal Title As String, _
                    Optional Button As VbMsgBoxStyle = vbYesNo + vbQuestion) _
                            As VbMsgBoxResult
    Ask = MsgBox(hWnd, Title, Prompt, Button)
End Function
Public Function InputBox(ByVal Titel As String, _
                         ByVal Prompt As String, _
                         Optional PromptColor As Variant, _
                         Optional PromptFontName As Variant, _
                         Optional PromptFontSize As Variant, _
                         Optional PromptIsBold As Boolean, _
                         Optional PromptIsUnderline As Boolean, _
                         Optional InputBoxBackColor As Variant, _
                         Optional MaxInputLenght As Variant, _
                         Optional PasswordChar As String = "", _
                         Optional Xpos As Variant, _
                         Optional Ypos As Variant, _
                         Optional AutoCloseSekTimer As Integer = 0, _
                         Optional ShowTimeLeftInTitle As Boolean) _
                                As String
    Dim frm As New MyControls
        
    Load frm
    Call frm.Set_InputBoxInfo(Titel, Prompt, PromptColor, _
                              PromptFontName, _
                              PromptFontSize, PromptIsBold, _
                              PromptIsUnderline, _
                              InputBoxBackColor, _
                              MaxInputLenght, PasswordChar, _
                              Xpos, Ypos, _
                              AutoCloseSekTimer, _
                              ShowTimeLeftInTitle)
    
    Do While Not frm.uStringWasSet
        DoEvents
    Loop
    
    InputBox = frm.uString
    
    Unload frm
    Set frm = Nothing
    
End Function
Public Function MsgBox(ByVal hWnd As Long, _
                       ByVal Title As String, _
                       ByVal Prompt As String, _
                       ByVal Button As VbMsgBoxStyle) _
                            As VbMsgBoxResult
    MsgBox = MessageBox(hWnd, Prompt, Title, CLng(Button))
End Function
Public Function PasswordDlg(Command As PassCommand, _
                            Optional Password As String, _
                            Optional MaxPasswordLenght As Long = 0) _
                                    As String
    
    Load frmPass
    
    Call frmPass.SetOptions(Command, Password, _
                            MaxPasswordLenght)
    
    Do While Not frmPass.UserInput
        DoEvents
    Loop

    PasswordDlg = frmPass.UserStr
    
    Unload frmPass
    Set frmPass = Nothing
    
End Function
Public Function ColorDlg(ByVal hWnd As Long, _
                         Optional ByVal ExtColors As Boolean = False) _
                                As Long
    Dim nCustColor(16) As Long
    Dim pColor As LPCHOOSECOLOR
        
    ColorDlg = -1
    
    With pColor
        .lStructSize = Len(pColor)
        .hWnd = hWnd
        .hInstance = CLng(0)
        
        .Flags = IIf(ExtColors, _
                     CC_SOLIDCOLOR Or CC_FULLOPEN, _
                     CC_SOLIDCOLOR)
        
        .lpCustColors = String$(16 * 4, 0)

        If ChooseColor(pColor) Then ColorDlg = .rgbResult
    End With

End Function
Public Function FontDlg(ByVal hWnd As Long, _
                        FontName As String, _
                        FontSize As Long, _
                        FontColor As Long, _
                        FontBold As Boolean, _
                        FontItalic As Boolean, _
                        FontUnderline As Boolean, _
                        FontStrikethrough As Boolean) _
                            As Boolean
    Dim Retval As Long
    Dim TmpFName As String
    
    Static LFnt As LOGFONT, CF_T As CHOOSEFONT_TYPE

    On Local Error GoTo Quit

    With CF_T
        .Flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
        .hwndOwner = hWnd
        .lStructSize = Len(CF_T)
        .lpLogFont = VarPtr(LFnt)
        .hInstance = 0
        .nFontType = SCREEN_FONTTYPE
        .rgbColors = 0
    End With
  
    If Trim$(LFnt.lfFaceName) = "" Then
        LFnt.lfFaceName = Space(31) & vbNullChar
    Else
        LFnt.lfFaceName = Left$(LFnt.lfFaceName, Len(LFnt.lfFaceName) - 1) & vbNullChar
    End If

    Retval = ChooseFont(CF_T)
  
    If Retval <> 0 Then
        TmpFName = StrConv(LFnt.lfFaceName, vbUnicode)
        FontName = Left$(TmpFName, InStr(1, TmpFName, vbNullChar) - 1)
        FontSize = CF_T.iPointSize / 10
        FontColor = CF_T.rgbColors
        FontBold = CBool(LFnt.lfWeight >= FW_BOLD)
        FontItalic = CBool(LFnt.lfItalic)
        FontUnderline = CBool(LFnt.lfUnderline)
        FontStrikethrough = CBool(LFnt.lfStrikeOut)
        
        FontDlg = True
        Exit Function
    End If
    
Quit:
    If Err.Number <> 0 Or Retval = 0 Then _
        FontName = "": _
        FontSize = 0: _
        FontBold = False: _
        FontItalic = False: _
        FontUnderline = False: _
        FontStrikethrough = False

End Function
