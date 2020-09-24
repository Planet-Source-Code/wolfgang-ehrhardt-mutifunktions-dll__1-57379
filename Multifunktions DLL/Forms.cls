VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Forms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function MoveForm Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCountA Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwnewLong As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32.dll" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long

Private Type TRIVERTEX
    X     As Long
    Y     As Long
    Red   As Integer
    Green As Integer
    Blue  As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type

Private Type rFormNFO
    xStep As Integer
    yStep As Integer
    xFrom As Long
    yFrom As Long
    xTo   As Long
    yTo   As Long
End Type

Public Enum cCenterNFO
    ccTOP = 1
    ccMIDDLE = 2
    ccBOTTOM = 3
    ccHWND = 4
End Enum

Public Enum SystemButtons
    sbCLOSEBOX = 0
    sbMINIMIZEBOX = 1
    sbMAXIMIZEBOX = 2
    sbRESIZEABLE = 3
    sbBORDER = 4
    sbCONTROLBOX = 5
    sbSHOWINTASKBAR = 7
    sbMOVEABLE = 8
End Enum

Private Const GCW_HCURSOR = -12

Private Const GRADIENT_FILL_RECT_H = &H0
Private Const GRADIENT_FILL_RECT_V = &H1
Private Const GRADIENT_FILL_TRIANGLE = &H2
Private Const GRADIENT_FILL_OP_FLAG = &HFF

Private Const HTCAPTION = 2

Private Const SM_CYCAPTION = &H4
Private Const SM_CXFRAME = &H20
Private Const SM_CYFRAME = &H21

Private Const WM_GETSYSMENU = &H313

Private Const WS_SIZEBOX = &H40000
Private Function CCol(ByVal Col As Byte) As Integer
    If Col > &H7F Then
        CCol = (Col * &H100&) - &H10000
    Else
        CCol = Col * &H100&
    End If
End Function
Public Function Fade(ByRef Form As Object, _
                     ByVal FadeFromColor As Long, _
                     ByVal FadeToColor As Long, _
                     Optional Alpha As Integer = 0, _
                     Optional Vertical As Boolean = False) _
                            As Boolean
    Dim OldSm As Integer
    Dim dwMode As Long, R As Long, g As Long, B As Long
    Dim Vertex(0 To 1) As TRIVERTEX
    Dim RECT As GRADIENT_RECT
    
    On Local Error GoTo Quit
    
    OldSm = Form.ScaleMode
    
    Form.ScaleMode = vbPixels
    Form.AutoRedraw = True
    
    dwMode = IIf(Vertical, GRADIENT_FILL_RECT_H, _
                           GRADIENT_FILL_RECT_V)
    
    Call GRAFX.Convert_ColorToRGB(FadeFromColor, R, g, B)
    
    With Vertex(0)
        .X = 0
        .Y = 0
        .Red = CCol(R)
        .Green = CCol(g)
        .Blue = CCol(B)
        .Alpha = Alpha
    End With

    Call GRAFX.Convert_ColorToRGB(FadeToColor, R, g, B)

    With Vertex(1)
        .X = Form.ScaleWidth
        .Y = Form.ScaleHeight
        .Red = CCol(R)
        .Green = CCol(g)
        .Blue = CCol(B)
        .Alpha = Alpha
    End With

    RECT.UpperLeft = 1
    RECT.LowerRight = 0
    
    Call GradientFillRect(Form.hDC, Vertex(0), 2, RECT, _
                          1, dwMode)
    
    Form.Refresh
    Form.ScaleMode = OldSm
    
    Fade = True

Quit:
End Function
Public Function Roll(Form As Object, _
                     ByVal WidthEnd As Long, _
                     ByVal HeightEnd As Long, _
                     Optional ByVal RollWidthFirst As Boolean = False, _
                     Optional ByVal Speed As Integer = 1) _
                            As Boolean
    Dim Px As Long, Py As Long
    Dim rNFO As rFormNFO
    
    On Local Error GoTo Quit
    
    If Speed = 0 Then _
        Form.Width = WidthEnd: _
        Form.Height = HeightEnd: _
        Form.Refresh: _
        Roll = True: _
        Exit Function
    
    If Speed < 0 Then Speed = Speed * (-1)

    Speed = Speed * 10
        
    With rNFO
        .xStep = 1
        .yStep = 1
    
        If RollWidthFirst Then
            If WidthEnd < Form.Width Then .xStep = -1
            If HeightEnd < Form.Height Then .yStep = -1
        
            .xFrom = Form.Width
            .xTo = WidthEnd
            .yFrom = Form.Height
            .yTo = HeightEnd
        Else
            If HeightEnd < Form.Height Then .xStep = -1
            If WidthEnd < Form.Width Then .yStep = -1
        
            .xFrom = Form.Height
            .xTo = HeightEnd
            .yFrom = Form.Width
            .yTo = WidthEnd
        End If
    
        .xStep = .xStep * Speed
        .yStep = .yStep * Speed
    
        For Px = .xFrom To .xTo Step .xStep
            If RollWidthFirst Then
                Form.Width = Px
            Else
                Form.Height = Px
            End If
            
            Form.Refresh
        Next Px
        
        For Py = .yFrom To .yTo Step .yStep
            If RollWidthFirst Then
                Form.Height = Py
            Else
                Form.Width = Py
            End If
        
            Form.Refresh
        Next Py
    End With
    
    Roll = True

Quit:
End Function
Public Function DeactivateX(Form As Object) As Boolean
    Dim hMenu As Long, nPosition As Long, menuItemCount As Long
    Dim M1 As Long, m2 As Long, m3 As Long
    
    On Local Error GoTo Quit

    hMenu = GetSystemMenu(Form.hWnd, 0)
   
    If hMenu Then
        Call MNU.Find_MenuItemByCaption(Form.hWnd, _
                                        "&Schließen" & Chr$(9) & "Alt+F4", _
                                        M1, m2, nPosition, True)
        If nPosition = 0 Then _
                Call MNU.Find_MenuItemByCaption(Form.hWnd, _
                                                "&Schließen", _
                                                M1, m2, nPosition, True)
        
        If RemoveMenuA(hMenu, nPosition + 1, _
                       MF_REMOVE Or MF_BYPOSITION) = 1 Then
            If RemoveMenuA(hMenu, nPosition, _
                           MF_REMOVE Or MF_BYPOSITION) = 1 Then _
                                DeactivateX = True
        End If
        
        Call DrawMenuBar(Form.hWnd)
        Call Update(Form)
    End If

Quit:
End Function
Public Function Move(Form As Object, _
                     Optional ByVal cx As Long = 0, _
                     Optional ByVal cy As Long = 0, _
                     Optional Speed As Integer = 1) As Boolean
    
    Dim P As Long
    Dim xStep As Integer, yStep As Integer
    
    On Local Error GoTo Quit
    
    xStep = 1
    yStep = 1
    
    If cx < Form.Left Then xStep = -1
    If cy < Form.Top Then yStep = -1
    
    If Speed < 0 Then Speed = Speed * (-1)
    
    xStep = xStep * Speed
    yStep = yStep * Speed
    
    If Speed = 0 Then _
        Form.Left = cx: _
        Form.Top = cy: _
        Move = True: _
        Exit Function
    
    For P = Form.Left To cx Step xStep
        Form.Left = P
        Form.Refresh
    Next P
    
    For P = Form.Top To cy Step yStep
        Form.Top = P
        Form.Refresh
    Next P
    
    Move = True

Quit:
End Function
Public Function Fade_Ranomized(Form As Object) As Boolean
            
    On Local Error Resume Next
        
    If Fade(Form, GRAFX.Get_RandomizedColor, _
            GRAFX.Get_RandomizedColor) Then _
                    Fade_Ranomized = True

End Function
Public Function Find_Form(pForms As Object, _
                          Optional FormName As String, _
                          Optional FormCaption As String, _
                          Optional FormTag As String, _
                          Optional FindExact As Boolean = True) _
                                As Long
    Dim P As Integer
    Dim MustBe As Integer, Ok As Integer
    Dim FC As String, ft As String
    
    MustBe = 0
    
    If FormName <> "" Then MustBe = 1
    If FormCaption <> "" Then MustBe = MustBe + 1
    If FormTag <> "" Then MustBe = MustBe + 1
    
    If Not FindExact Then _
        FormTag = LCase(FormTag): _
        FormCaption = LCase(FormCaption)
        
    FormName = LCase(FormName)

    For P = 0 To pForms.Count - 1
        Ok = 0
        
        If FormName <> "" Then
            If LCase(pForms(P).Name) = FormName Then
                Ok = 1
            Else
                GoTo GetNextForm
            End If
        End If
        
        If FormCaption <> "" Then
            FC = pForms(P).Caption
            If Not FindExact Then FC = LCase(FC)
            
            If FC = FormCaption Then
                Ok = Ok + 1
            Else
                GoTo GetNextForm
            End If
        End If
                    
        If FormTag <> "" Then
            ft = pForms(P).Tag
            If Not FindExact Then ft = LCase(ft)
            If ft = FormTag Then Ok = Ok + 1
        End If
        
GetNextForm:
        If Ok = MustBe Then _
            Find_Form = P: _
            Exit Function

    Next P
    
    Find_Form = -1
    
End Function
Public Function StayOnTop(Form As Object) As Boolean
    
    On Local Error GoTo Quit
    
    If WIN.StayOnTop(Form.hWnd, False) Then _
        Form.Refresh: _
        StayOnTop = True
        
Quit:
End Function
Public Function Center(Form As Object, _
                       CenterWhere As cCenterNFO, _
                       Optional hWnd As Long) As Boolean
    Dim L As Long, T As Long
    Dim rl As Long, rt As Long, rr As Long, rb As Long
            
    On Local Error GoTo Quit

    Select Case CenterWhere
        Case ccTOP
            T = 0
        Case ccMIDDLE
            T = (Screen.Height / 2) - (Form.Height / 2)
        Case ccBOTTOM
            Call DESK.Get_TaskbarPos(rl, rt, rr, rb)
            T = rt * vbContainerPosition
        Case ccHWND
            If hWnd < 1 Then Exit Function
            
            Call WIN.Get_Pos(hWnd, rl, rt, rr, rb)
            
            Form.Left = rl + (rr / 2)
            Form.Top = rt + (rb / 2)
            
            Center = True
            Exit Function
        Case Else: Exit Function
    End Select

    L = (Screen.Width / 2) - (Form.Width / 2)

    Form.Left = L
    Form.Top = T
    
    Form.Refresh
    
    Center = True
    
Quit:
End Function
Public Function Set_XPstyle(Form As Object, _
                            Optional ApplyStyle As Boolean = True) _
                                As Long
    Dim P As Long
    Dim C As Control, F As Form
    
    If Not SYS.XPstyleSupported Then Set_XPstyle = -1: _
                                     Exit Function
    
    On Local Error Resume Next

    If TypeOf Form Is Form Then
        Call WIN.Set_XPstyle(Form.hWnd, ApplyStyle)
        
        For Each C In Form
            Call WIN.Set_XPstyle(C.hWnd)
        Next C
       
        Set_XPstyle = 1
    Else
        Set_XPstyle = -2
    End If
    
End Function
Public Function Tile(Form As Object, Picture As StdPicture) _
                                As Boolean
    Dim X As Long, Y As Long, nWidth As Long, nHeight As Long
    
    On Local Error GoTo Quit
  
    With Form
        nWidth = .ScaleX(Picture.Width, 8, .ScaleMode)
        nHeight = .ScaleY(Picture.Height, 8, .ScaleMode)
    
        .AutoRedraw = True
        .Cls
    
        For X = 0 To .ScaleWidth Step nWidth
            For Y = 0 To .ScaleHeight Step nHeight
                .PaintPicture Picture, X, Y
            Next Y
        Next X
    
        .Picture = .Image
        .AutoRedraw = False
    End With
    
    Tile = True

Quit:
End Function
Public Function tFade(Form As Object, _
                      ByVal FadeStart As Byte, _
                      ByVal FadeStop As Byte, _
                      Optional StartAtCurrentTransparenz As Boolean = False, _
                      Optional FadeStep As Integer = 5, _
                      Optional Speed As Integer = 10) _
                            As Boolean
        
    On Local Error GoTo Quit
    
    tFade = WIN.tFade(Form.hWnd, FadeStart, FadeStop, _
                      StartAtCurrentTransparenz, _
                      FadeStep, Speed)

Quit:
End Function
Public Sub FormMove(Form As Object)
    Call ReleaseCapture
    Call SendMessage(Form.hWnd, WM_NCLBUTTONDOWN, _
                     HTCAPTION, 0&)
End Sub
Public Function CreateFormFromPicture(Form As Object, _
                                      Picture As Object, _
                                      Optional PictureFile As String = "") _
                                            As Boolean
    On Local Error Resume Next
    
    CreateFormFromPicture = CFFP(Form, Picture, PictureFile)
    
End Function
Public Sub FormPos_Save(Form As Object, _
                        INIpath As String)
        
    Call INI.Write_INI(INIpath, Form.Name, "Left", Form.Left)
    Call INI.Write_INI(INIpath, Form.Name, "Top", Form.Top)
    Call INI.Write_INI(INIpath, Form.Name, "Width", Form.Width)
    Call INI.Write_INI(INIpath, Form.Name, "Height", Form.Height)
    
End Sub
Public Sub FormPos_Load(Form As Object, _
                        INIpath As String)
        
    On Local Error Resume Next
    
    Form.Left = INI.Read_INI(INIpath, Form.Name, "Left")
    Form.Top = INI.Read_INI(INIpath, Form.Name, "Top")
    Form.Width = INI.Read_INI(INIpath, Form.Name, "Width")
    Form.Height = INI.Read_INI(INIpath, Form.Name, "Height")

End Sub
Public Function Form(ByVal sForm As Object, _
                     But As SystemButtons, _
                     Optional Enable As Boolean = True) _
                            As Boolean
    Dim M As String
    Dim Style As Long, lngStyle As Long, xStyle As Long
    Dim nVisible As Boolean
        
    xStyle = GWL_STYLE
    M = ""

    Select Case But
        Case sbCLOSEBOX
            Form = DeactivateX(sForm)
            Exit Function
        Case sbMAXIMIZEBOX
            Style = WS_MAXIMIZEBOX
            M = "Ma&ximieren"
        Case sbMINIMIZEBOX
            Style = WS_MINIMIZEBOX
            M = "Mi&nimieren"
        Case sbRESIZEABLE
            Style = WS_SIZEBOX
        Case sbBORDER
            Style = WS_CAPTION
        Case sbCONTROLBOX
            Style = WS_SYSMENU
        Case sbSHOWINTASKBAR
            Style = WS_EX_APPWINDOW
            xStyle = GWL_EXSTYLE
        Case sbMOVEABLE
            Form = MNU.Remove_MenuByCaption(sForm.hWnd, "&Verschieben", True)
            Exit Function
        Case Else
            Exit Function
    End Select

    lngStyle = GetWindowLongA(sForm.hWnd, xStyle)
    Call SetLastError(0)

    Select Case xStyle
        Case GWL_STYLE
            lngStyle = IIf(Enable, lngStyle Or Style, _
                                   lngStyle - Style)
                         
            Call SetWindowLong(sForm.hWnd, GWL_STYLE, lngStyle)
            
            If GetLastError = 0 Then
                If Len(M) Then
                    Form = MNU.Remove_MenuByCaption(sForm.hWnd, M, True)
                Else
                    Form = True
                End If
            End If
            
            Call Update(sForm)
        Case GWL_EXSTYLE
            nVisible = sForm.Visible
            sForm.Visible = False
                              
            lngStyle = IIf(Enable, lngStyle Or Style, _
                                   lngStyle And Not Style)
        
            Call SetWindowLong(sForm.hWnd, GWL_EXSTYLE, lngStyle)
            If GetLastError = 0 Then Form = True

            sForm.Refresh
            sForm.Visible = nVisible
    End Select

End Function
Public Function ShowControlsOnly(Form As Object) As Boolean
    Dim hWnd As Long, DstRgn As Long, TmpRgn As Long
    Dim MaxH As Long, MaxW As Long, H As Long
    Dim AlterScaleMode As Integer
    Dim CL As Long, Ct As Long, ch As Long, Cw As Long
    Dim FrameX As Long, FrameY As Long, Caption As Long
    Dim Cmd As Control
        
    On Local Error GoTo Quit
    
    hWnd = Form.hWnd
    
    DstRgn = CreateRectRgn(0, 0, 0, 0)
    
    AlterScaleMode = Form.ScaleMode
    Form.ScaleMode = vbPixels
  
    If Form.BorderStyle <> vbBSNone Then _
        FrameX = GetSystemMetrics(SM_CXFRAME): _
        FrameY = GetSystemMetrics(SM_CYFRAME): _
        Caption = GetSystemMetrics(SM_CYCAPTION)
    
    For Each Cmd In Form.Controls
        If Not TypeName(Cmd) = "Menu" Then
            H = 0
            
            On Local Error Resume Next
            H = GetParentA(Cmd.hWnd)
            Err.Clear
            On Local Error GoTo Quit
            
            If H = 0 Or H = hWnd Then
                CL = Cmd.Left + FrameY
                Ct = Cmd.Top + Caption + FrameX
                ch = Cmd.Height
                Cw = Cmd.Width
            
                TmpRgn = CreateRectRgn(CL, Ct, CL + Cw, Ct + ch)
            
                If MaxH < Ct + ch Then MaxH = Ct + ch
                If MaxW < CL + Cw Then MaxW = CL + Cw

                Call CombineRgn(DstRgn, DstRgn, TmpRgn, RGN_OR)
                Call DeleteObject(TmpRgn)
            End If
        End If
    Next
  
    Call SetWindowRgn(hWnd, DstRgn, True)
    Call DeleteObject(DstRgn)
    
    Form.ScaleMode = AlterScaleMode
    Form.Height = (MaxH + FrameY) * Screen.TwipsPerPixelY
    Form.Width = (MaxW + FrameX) * Screen.TwipsPerPixelX
    
    ShowControlsOnly = True
    
Quit:
End Function
Public Sub AllControlsSave(Form As Object, _
                           ByVal INIpath As String, _
                           Optional SaveForm As Boolean = False)
    Dim C As Control
        
    If SaveForm Then Call CRTL.Control_Save(Form, Form, INIpath)
    
    For Each C In Form
        Call CRTL.Control_Save(C, Form, INIpath)
    Next C
    
End Sub
Public Sub AllControlsLoad(Form As Object, _
                           ByVal INIpath As String, _
                           Optional LoadForm As Boolean = False)
    Dim C As Control
        
    If LoadForm Then Call CRTL.Control_Load(Form, Form, INIpath)

    For Each C In Form
        Call CRTL.Control_Load(C, Form, INIpath)
    Next C
    
End Sub
Public Sub FormSave(Form As Object, _
                    ByVal INIpath As String)
    Call CRTL.Control_Save(Form, Form, INIpath)
End Sub
Public Sub FormLoad(Form As Object, _
                    ByVal INIpath As String)
    Call CRTL.Control_Load(Form, Form, INIpath)
End Sub
Private Sub Update(Form As Object)
    Dim R As RECT
    
    Call GetWindowRect(Form.hWnd, R)
    Call SetWindowPos(Form.hWnd, 0, R.Left, R.Top, _
                      R.Right - R.Left, R.Bottom - R.Top, _
                      SWP_FRAMECHANGED)
                      
End Sub
Public Function ScrollCaption(Form As Object, _
                              Optional Speed As Long = 1) _
                                    As Boolean
                            
    On Local Error GoTo Quit
    
    Load MyControls
    Call MyControls.SetScrollCaption(Form, Speed)
    
    ScrollCaption = True
    
Quit:
    If Err.Number <> 0 Then _
        ScrollCaption = False: _
        If MyControls.Timer_ScrollCaption.UBound = 0 Then _
            Unload MyControls
    
End Function
Public Function IsModal(Form As Object) As Boolean
        
    On Error Resume Next
  
    Form.Show
    IsModal = (Err.Number = 401)
  
End Function
Public Function ShowAnimatedCursor(Form As Object, _
                                   ByVal CursorFile As String) _
                                        As Boolean
    Dim hcursor As Long
    
    On Local Error GoTo Quit
    
    hcursor = LoadCursorFromFile(CursorFile)
    
    If hcursor <> 0 Then _
        Form.MousePointer = 99: _
        Call SetClassLong(Form.hWnd, GCW_HCURSOR, hcursor): _
        ShowAnimatedCursor = True

Quit:
End Function
Public Sub ShowSysMenu(Form As Object)
    Dim IParam As Long
    Dim PA As POINTAPI
    
    On Local Error GoTo Quit
    
    Call MOUSE.Get_MousePos(PA.X, PA.Y)
    
    IParam = (PA.Y * &H10000) Or (PA.X And &HFFFF&)
    
    Call SendMessage(Form.hWnd, WM_GETSYSMENU, 0, ByVal IParam)

Quit:
End Sub
