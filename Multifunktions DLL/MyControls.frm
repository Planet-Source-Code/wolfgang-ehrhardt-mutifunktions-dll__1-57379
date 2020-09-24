VERSION 5.00
Begin VB.Form MyControls 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MyControls"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9540
   Visible         =   0   'False
   Begin VB.TextBox TextLB 
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicFade 
      Height          =   735
      Index           =   1
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox PicFade 
      Height          =   735
      Index           =   2
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   17
      Top             =   4680
      Width           =   735
   End
   Begin VB.PictureBox PicFade 
      Height          =   735
      Index           =   3
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   16
      Top             =   5400
      Width           =   735
   End
   Begin VB.PictureBox picCredits 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   4320
      ScaleHeight     =   975
      ScaleWidth      =   1575
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Timer Timer_ScrollCaption 
      Enabled         =   0   'False
      Index           =   0
      Left            =   1560
      Top             =   720
   End
   Begin VB.Timer Timer_cLabel 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   720
   End
   Begin VB.ListBox vLB 
      Height          =   255
      Index           =   0
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox vPic 
      Height          =   975
      Index           =   0
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox PicTMP 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   960
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Timer Timer_HyperLink 
      Enabled         =   0   'False
      Left            =   600
      Top             =   720
   End
   Begin VB.Timer Timer_AutoClose 
      Enabled         =   0   'False
      Left            =   120
      Top             =   720
   End
   Begin VB.TextBox uInput 
      Height          =   375
      Left            =   240
      MousePointer    =   3  'I-Cursor
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton But_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton But_Cancel 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "MyControls"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   7095
      Begin VB.ListBox LBtmp 
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bitmap Rotate && Zoom"
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4335
         Begin VB.PictureBox pd 
            AutoSize        =   -1  'True
            Height          =   960
            Left            =   2880
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   88
            TabIndex        =   5
            Top             =   240
            Width           =   1380
         End
         Begin VB.PictureBox ps 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   960
            Left            =   1440
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   85
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.PictureBox po 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   960
            Left            =   120
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   80
            TabIndex        =   3
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.ListBox sLB 
         Height          =   255
         Left            =   4560
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label ClickIcon 
         Caption         =   "ClickIcon"
         Height          =   255
         Left            =   4560
         MouseIcon       =   "MyControls.frx":0000
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Label LabelTMP 
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label LabelMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "MyControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'InputBox
Public uString  As Variant
Public uStringWasSet As Boolean
Dim AutoClose As Integer
Dim ShowTimeLeft As Boolean
Dim uTitle As String

'Hyperlink
Private Type HyperlinkNFO
    LabelColor          As Long
    LabelColorMouseOver As Long
    ParentForm          As Form
    Label               As Label
    OldValue            As Long
    hWnd                As Long
End Type

'cLabel
Private Type cLabelNFO
    Label       As Label
    moBackColor As Long
    moForeColor As Boolean
    moUnderline As Boolean
    moBold      As Boolean
    moFontname  As String
    moFontSize  As Long
    moCaption   As String
    moItalic    As Boolean

    BackColor As Long
    ForeColor As Boolean
    Underline As Boolean
    Bold      As Boolean
    FontName  As String
    FontSize  As Long
    Caption   As String
    Italic    As Boolean
End Type

Private Type ScrollNFO
    Form       As Form
    hWnd       As Long
    Speed      As Long
    sCaption   As String
    sText      As String
    L          As Long
    T          As Long
    CountDown  As Boolean
End Type

Private MyIcon As StdPicture
Private hLabel() As HyperlinkNFO
Private cLabel() As cLabelNFO
Private sScroll() As ScrollNFO

'GIF
Dim myPic As Object

'ListEdit
Public ListEdit As Boolean
Public ListEditESC As Boolean
Public ListEditEmpty As Boolean

Dim PtList(2) As POINTAPI

Private Sub But_Cancel_Click()
    Call SetEingabe(Chr$(0))
End Sub
Private Sub But_Ok_Click()
    Call SetEingabe(uInput.Text)
End Sub
Private Sub Form_Load()
    
    Me.Height = 0
    Me.Width = 0
    
    Me.Visible = False

    uStringWasSet = False
    hasuSting = False
    uString = Empty
    
End Sub
Public Function SortLB(LB As Object) As Boolean
    Dim P As Integer
    
    On Local Error GoTo Quit
    
    If LB.Sorted Then SortLB = False: _
                      Exit Function
    
    sLB.Clear
    
    For P = 0 To LB.ListCount - 1
        sLB.AddItem LB.List(P)
        sLB.ItemData(sLB.NewIndex) = LB.ItemData(P)
    Next P
    
    LB.Clear
    
    For P = 0 To sLB.ListCount - 1
        LB.AddItem sLB.List(P)
        LB.ItemData(LB.NewIndex) = sLB.ItemData(P)
    Next P
    
    SortLB = True
    
    Unload Me

Quit:
End Function
Public Function BitmapWork(PicBox As PictureBox, _
                           Optional Rotate As Variant, _
                           Optional Turn As Variant, _
                           Optional Flip As Variant, _
                           Optional zZoom As Variant) _
                                As StdPicture
    Dim X As Integer, NewX As Integer, NewY As Integer
    Dim SinAng1, CosAng1, SinAng2, SinAng3, Zoom
    
    ps.ScaleMode = PicBox.ScaleMode
    
    ps.Height = PicBox.Height
    ps.Width = PicBox.Width
    
    ps.ScaleMode = vbPixels
    
    ps.Picture = PicBox.Picture

    po.Width = ps.Width
    pd.Width = ps.Width
    po.Height = ps.Height
    pd.Height = ps.Height

    PtList(0).X = -(ps.ScaleWidth / 2)
    PtList(0).Y = -(ps.ScaleHeight / 2)
    PtList(1).X = ps.ScaleWidth / 2
    PtList(1).Y = -(ps.ScaleHeight / 2)
    PtList(2).X = -(ps.ScaleWidth / 2)
    PtList(2).Y = (ps.ScaleHeight / 2)
  
    If IsMissing(Rotate) Then Rotate = 0
    If IsMissing(Turn) Then Turn = 0
    If IsMissing(Flip) Then Flip = 0
    
    If IsMissing(zZoom) Then
        Zoom = Tan(45 * pi180)
    Else
        Zoom = Tan(zZoom * pi180)
    End If
    
    SinAng1 = Sin((Rotate + 90) * pi180)
    CosAng1 = Cos((Rotate + 90) * pi180)
    SinAng2 = Sin((Turn + 90) * pi180) * Zoom
    SinAng3 = Sin((Flip + 90) * pi180) * Zoom

    For X = 0 To 2
        NewX = (PtList(X).X * SinAng1 + PtList(X).Y * CosAng1) * SinAng2
        NewY = (PtList(X).Y * SinAng1 - PtList(X).X * CosAng1) * SinAng3
        
        PtList(X).X = NewX + (pd.ScaleWidth / 2)
        PtList(X).Y = NewY + (pd.ScaleHeight / 2)
    Next X
  
    po.Cls
    Call PlgBlt(po.hDC, PtList(0), ps.hDC, 0, 0, ps.ScaleWidth, ps.ScaleHeight, 0, 0, 0)
  
    pd.Picture = po.Image
    Set BitmapWork = pd.Picture
    
End Function
Private Sub Timer_AutoClose_Timer()
    
    Static Count As Integer
    
    Count = Count + 1
    
    If Count = AutoClose Then
        Call SetEingabe(uInput.Text)
    Else
        If ShowTimeLeft Then _
            Me.Caption = uTitle & " (" & _
                         AutoClose - Count & ")"
    End If
    
End Sub
Private Sub Timer_HyperLink_Timer()
    Dim P As Long
    Dim TimerWork As Boolean
            
    On Local Error GoTo Quit
    
    For P = LBound(hLabel) To UBound(hLabel)
        If hLabel(P).hWnd Then
            If WIN.Get_TaskID(hLabel(P).hWnd) Then
                TimerWork = True
                
                If MOUSE.isMouseOverControl(hLabel(P).Label) Then
                    If hLabel(P).OldValue <> hLabel(P).LabelColorMouseOver Then _
                        hLabel(P).OldValue = hLabel(P).LabelColorMouseOver: _
                        hLabel(P).Label.ForeColor = hLabel(P).LabelColorMouseOver: _
                        hLabel(P).Label.FontUnderline = True: _
                        hLabel(P).Label.MouseIcon = MyIcon: _
                        hLabel(P).Label.MousePointer = vbCustom: _
                        hLabel(P).Label.Refresh
                Else
                    If hLabel(P).OldValue <> hLabel(P).LabelColor Then _
                        hLabel(P).OldValue = hLabel(P).LabelColor: _
                        hLabel(P).Label.ForeColor = hLabel(P).LabelColor: _
                        hLabel(P).Label.FontUnderline = False: _
                        hLabel(P).Label.MousePointer = 0: _
                        hLabel(P).Label.Refresh
                End If
            Else
                hLabel(P).hWnd = 0
            End If
        End If
    Next P
    
Quit:
    If Not TimerWork Then
        Timer_HyperLink.Enabled = False
        Erase hLabel
        Unload Me
    Else
        DoEvents
    End If
    
End Sub
Private Sub Timer_ScrollCaption_Timer(Index As Integer)
    Dim P As Long
    
    If WIN.Get_TaskID(sScroll(Index).hWnd) = 0 Then
        Timer_ScrollCaption(Index).Enabled = False
        
        sScroll(Index).hWnd = 0
        
        For P = 1 To UBound(sScroll)
            If sScroll(P).hWnd <> 0 Then Exit Sub
        Next P
        
        Unload Me
    
        Exit Sub
    End If
    
    If sScroll(Index).sCaption = "" Then _
        sScroll(Index).sCaption = sScroll(Index).Form.Caption: _
        sScroll(Index).Form.Caption = " "
    
    If sScroll(Index).sText = "" Then _
        sScroll(Index).sText = sScroll(Index).sCaption: _
        sScroll(Index).L = sScroll(Index).Form.Width / 78: _
        sScroll(Index).CountDown = False
    
    If sScroll(Index).L = 1 Then
        If Len(sScroll(Index).sText) Then
            
            If Not sScroll(Index).CountDown Then _
                sScroll(Index).CountDown = True: _
                sScroll(Index).sText = sScroll(Index).Form.Caption
            
            sScroll(Index).sText = Mid$(sScroll(Index).sText, 2)
        Else
            sScroll(Index).sText = ""
        End If
    
        sScroll(Index).Form.Caption = sScroll(Index).sText
        sScroll(Index).T = sScroll(Index).T + 1
    Else
        sScroll(Index).Form.Caption = Space$(sScroll(Index).L) & sScroll(Index).sText
        sScroll(Index).L = sScroll(Index).L - 1
    End If

End Sub
Private Sub uInput_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 27
            If ListEdit Then
                If ListEditESC Then
                    KeyAscii = 0
                    uInput.Text = ""
                    
                    If ListEditEmpty Then Call But_Ok_Click
                End If
            End If
        Case 13
            KeyAscii = 0
            
            If Not ListEdit Then
                Call But_Ok_Click
            Else
                If ListEditEmpty Then
                    But_Ok_Click
                Else
                    If uInput.Text <> "" Then Call But_Ok_Click
                End If
            End If
    End Select

End Sub
Public Sub Set_InputBoxInfo(ByVal Titel As String, _
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
                            Optional ShowTimeLeftInTitle As Boolean)
        
    But_Cancel.Visible = True
    But_Ok.Visible = True
    LabelMessage.Visible = True
    uInput.Visible = True
    
    Me.Height = 2235
    Me.Width = 5565

    uTitle = Titel
    Me.Caption = Titel
    LabelMessage.Caption = Prompt
    
    If Not IsMissing(PromptColor) Then _
        LabelMessage.ForeColor = PromptColor
    If Not IsMissing(PromptFontName) Then _
        LabelMessage.FontName = PromptFontName
    If Not IsMissing(PromptFontSize) Then _
        LabelMessage.FontSize = PromptFontSize
    If Not IsMissing(InputBoxBackColor) Then _
        BackColor = InputBoxBackColor
    If Not IsMissing(MaxInputLenght) Then _
        uInput.MaxLength = MaxInputLenght
    
    If IsMissing(Xpos) And IsMissing(Ypos) Then
        Call FRMS.Center(Me, ccMIDDLE)
    Else
        If Xpos = Empty Then Xpos = 0
        If Ypos = Empty Then Ypos = 0
        
        Me.Left = Xpos
        Me.Top = Ypos
    End If
    
    LabelMessage.FontBold = PromptIsBold
    LabelMessage.FontUnderline = PromptIsUnderline
    
    If PasswordChar <> "" Then uInput.PasswordChar = PasswordChar
    
    ShowTimeLeft = ShowTimeLeftInTitle
    
    If ShowTimeLeft Then _
        Me.Caption = uTitle & " (" & AutoCloseSekTimer & ")"
    
    If AutoCloseSekTimer <> 0 Then _
        AutoClose = AutoCloseSekTimer: _
        Timer_AutoClose.Interval = 1000: _
        Timer_AutoClose.Enabled = True
        
    Call WIN.StayOnTop(Me.hWnd, False)
    
    Me.Visible = True
    
End Sub
Private Sub SetEingabe(Str As String)
    
    Me.Hide
    
    uString = Str
    uStringWasSet = True
    
    Timer_AutoClose.Enabled = False

End Sub
Public Function Create_Hyperlink(ParentForm As Form, _
                                 Label As Label, _
                                 LabelColor As Long, _
                                 LabelColorMouseOver As Long)
    Dim U As Integer
    
    On Local Error Resume Next
    
    ReDim Preserve hLabel(UBound(hLabel) + 1)
    If Err = 9 Then ReDim hLabel(0)

    U = UBound(hLabel)
    
    hLabel(U).LabelColor = LabelColor
    hLabel(U).LabelColorMouseOver = LabelColorMouseOver
    hLabel(U).hWnd = ParentForm.hWnd
    hLabel(U).OldValue = -1
    
    Set hLabel(U).ParentForm = ParentForm
    Set hLabel(U).Label = Label

    Set MyIcon = ClickIcon.MouseIcon

    Timer_HyperLink.Interval = 5
    Timer_HyperLink.Enabled = True

End Function
Public Sub SetScrollCaption(frm As Form, _
                            ByVal Speed As Long)
    Dim P As Long, Index As Long
        
    On Local Error Resume Next
    
    Index = -1
    Index = UBound(sScroll)
    
    If Index > -1 Then
        For P = 0 To UBound(sScroll)
            If sScroll(P).hWnd = frm.hWnd Then Index = P: _
                                               Exit For
        Next P
    End If

    If Index = -1 Then
        Err.Clear
                
        ReDim Preserve sScroll(UBound(sScroll) + 1)
        If Err.Number <> 0 Then ReDim sScroll(1): _
                                Err.Clear
        
        Index = UBound(sScroll)
    End If

    Set sScroll(Index).Form = frm
    sScroll(Index).hWnd = frm.hWnd
    sScroll(Index).Speed = (Speed * 50)
    sScroll(Index).sCaption = ""
    sScroll(Index).sText = ""
        
    Load Timer_ScrollCaption(Index)
    
    Timer_ScrollCaption(Index).Interval = sScroll(Index).Speed
    Timer_ScrollCaption(Index).Enabled = True
    
End Sub
