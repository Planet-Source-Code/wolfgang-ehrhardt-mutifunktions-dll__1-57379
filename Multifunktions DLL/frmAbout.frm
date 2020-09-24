VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":0000
   MousePointer    =   99  'Benutzerdefiniert
   ScaleHeight     =   5790
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LBnames 
      Height          =   255
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picCredits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      MouseIcon       =   "frmAbout.frx":030A
      MousePointer    =   99  'Benutzerdefiniert
      ScaleHeight     =   1875
      ScaleWidth      =   3675
      TabIndex        =   13
      Top             =   1560
      Width           =   3735
      Begin VB.Timer Timer_Fade 
         Enabled         =   0   'False
         Left            =   1080
         Top             =   1440
      End
      Begin VB.Timer CheckConnection 
         Enabled         =   0   'False
         Left            =   600
         Top             =   1440
      End
      Begin VB.Timer TimerLabel 
         Enabled         =   0   'False
         Left            =   120
         Top             =   1440
      End
      Begin VB.Label LabelCount 
         Alignment       =   1  'Rechts
         BackColor       =   &H000000FF&
         Caption         =   "30"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LabelUpdateText 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000000FF&
         Caption         =   "UpdateAction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin VB.CommandButton ButUpdate 
      Caption         =   "Nach Updates suchen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmAbout.frx":0614
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   10
      Top             =   4680
      Width           =   3495
   End
   Begin VB.TextBox TextVersion 
      Alignment       =   2  'Zentriert
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Version"
      Top             =   960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton BudSendBug 
      Caption         =   "Send Bug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MouseIcon       =   "frmAbout.frx":0766
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton ButHistory 
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmAbout.frx":08B8
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton ButExit 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      MouseIcon       =   "frmAbout.frx":0A0A
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   1
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label LabelSendNewMail 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Woeh@gmx.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Connectionstate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4395
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Height          =   855
      Left            =   120
      MouseIcon       =   "frmAbout.frx":0B5C
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   9
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label LabelHomepage 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Woeh.Tripod.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1110
      TabIndex        =   6
      Top             =   3960
      Width           =   1875
   End
   Begin VB.Label LabelDLLpath 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Path"
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmAbout.frx":0E66
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   5
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label LabelVersion 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label LabelProductName 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Multifunktions DLL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmAbout.frx":1170
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label LabelUpdate 
      BorderStyle     =   1  'Fest Einfach
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyVersion As String
Private Sub BudSendBug_Click()
        
    Call NET.SendNewMail("woeh@gmx.de", _
                         "Re: Bugreport (" & _
                         LabelVersion.Caption & ")", _
                         "Folgenden Bug habe ich gefunden: ")
    Unload Me

End Sub
Private Sub ButExit_Click()
    Unload Me
End Sub
Private Sub ButHistory_Click()
        
    Call SH.Execute(GetDir(App.Path) & "History.txt")
    
End Sub
Private Sub ButUpdate_Click()
    Dim LU As Label
    Dim Text As String, URL As String
    
    ButUpdate.Enabled = False
    LabelUpdate.Visible = True

    Set LU = LabelUpdateText
    
    LU.Caption = "Hole Versionsinformationen...."
    LU.Visible = True
    
    LabelUpdate.ZOrder
    LU.ZOrder
    
    LU.Refresh
    LabelUpdateText.Refresh
    
    URL = "http://woeh.tripod.com/Projekte/Version.txt"
    Text = NET.INetFile_Read(URL)
    
    If Text = "" Then
        LU.Caption = "Fehler beim übermittel der Versionsinformation."
        
        LabelCount.Caption = "30"
        LabelCount.Visible = True
        LabelCount.ZOrder
        
        TimerLabel.Interval = 1000
        TimerLabel.Enabled = True
    Else
        If CInt(Text) <= CInt(MyVersion) + 1 Then
            LU.Caption = "Keine neue Version verfügbar"
            
            LabelCount.Caption = "30"
            LabelCount.Visible = True
            LabelCount.ZOrder

            TimerLabel.Interval = 1000
            TimerLabel.Enabled = True
        Else
            If DLG.Ask(Me.hWnd, "Eine neuer Version ist verfügbar." & _
                        vbCrLf & "Soll die neue Version geladen werden ?", _
                        "Neue Version verfügbar") = vbYes Then
                LU.Caption = "Starte download..."
            
                URL = "http://woeh.tripod.com/Projekte/Multifunktions_DLL.zip"
            
                Call NET.INetFile_Save(URL, "", True)
                
                Me.Hide
                
                Me.Height = 0
                Me.Width = 0
                
                Unload Me
                
                Exit Sub
            End If
            
            LabelUpdate.Visible = False
            LU.Visible = False
        End If
    End If
    
End Sub
Private Sub CheckConnection_Timer()
        
    Static Old As Long
    
    If NET.isConnectedToInternet Then
        If Old < 2 Then
            Old = 2
            ButUpdate.Enabled = True
        
            Label4.Caption = "Online"
            Label4.ForeColor = vbGreen
        End If
    Else
        If Old = 2 Or Old = 0 Then
            Old = 1
            ButUpdate.Enabled = False
        
            Label4.Caption = "OFFLINE"
            Label4.ForeColor = vbRed
        End If
    End If
            
End Sub
Private Sub Form_Load()
    Dim Version As String, A As String, R$
    Dim mW As Long, mH As Long, P As Long
        
    Call WIN.StayOnTop(Me.hWnd)

    LabelProductName.Caption = App.ProductName
    
    Version = App.Major & "." & _
              App.Minor & "." & _
              App.Revision
    
    MyVersion = Replace(Version, ".", "")
    
    TextVersion.Text = Version
    
    LabelVersion.Caption = "Version " & Version
    
    LabelDLLpath.Caption = GetDir(App.Path) & _
                           App.ExeName & ".Dll"
    
    mH = 6135
    mW = 4065

    Me.Height = mH
    Me.Width = mW
    
    Call FRMS.Center(Me, ccMIDDLE)
    Call Timer_Fade_Timer
    
    Call NET.Hyperlink_Create(LabelSendNewMail, vbYellow, vbBlue)
    Call NET.Hyperlink_Create(LabelHomepage, vbYellow, vbBlue)
    
    Me.Height = 0
    Me.Width = 0
    
    Call CheckConnection_Timer
    
    CheckConnection.Interval = 1000
    CheckConnection.Enabled = True
    
    With LBnames
        .AddItem "David Ireland"
        .AddItem "Klaus Langbein (Klaus@Activevb.de)"
        .AddItem "Angie (as@vb-fun.de)"
        .AddItem "Götz Reinecke (Reinecke@activevb.de)"
        .AddItem "Benjamin Wilger (Benjamin@Activevb.de)"
        .AddItem "Helge Rex (Helge@Activevb.de)"
        .AddItem "Florian Rittmeier (Florian@Activevb.de)"
        .AddItem "LonelySuicide666"
        .AddItem "NoOne"
        .AddItem "Max Raskin"
        .AddItem "Jochen Wierum (JoWi@ActiveVB.de)"
        .AddItem "Kristof (Fleckenzwerge@aol.com)"
        .AddItem "Björn Kirsch"
        .AddItem "Dieter Otter (info@tools4vb.de)"
        .AddItem "Gerhard Kuklau (Gerhard.Kuklau@hdi.de)"
        .AddItem "Heinz Prelle (outa.space@t-online.de)"
        .AddItem "madmax (Majestic12@gmx.li)"
        .AddItem "Christoph von Wittich (Christoph@ActiveVB.de)"
        .AddItem "James..."
        .AddItem "R. Müller"
        .AddItem "rm (radeon_master@web.de)"
    End With
    
    A = App.ProductName & vbCrLf & _
        Version & vbCrLf & _
        vbCrLf & _
        "Zusammengetragen" & vbCrLf & _
        "& neu verfasst von" & vbCrLf & _
        "Wolfgang Ehrhardt" & vbCrLf & _
        vbCrLf & _
        "Mein besonderer Dank geht an" & vbCrLf & _
         "ActiveVB" & vbCrLf & _
         "und das" & vbCrLf & _
         "ActiveVB - Forum" & vbCrLf & _
         "vb@rchiv" & vbCrLf & _
         "VB Fun" & vbCrLf & _
         "& Planet Source Code" & vbCrLf & _
         vbCrLf & _
         "Alle Programmierer, die mir" & vbCrLf & _
         "in irgendeiner Weise geholfen" & vbCrLf & _
         "haben die DLL zu verwirklichen" & vbCrLf & _
         "--------------------------------------------------" & vbCrLf
    
    For P = 0 To LBnames.ListCount - 1
        A = A & LBnames.List(P) & vbCrLf
    Next P
    
    A = A & vbCrLf & vbCrLf & vbCrLf & _
        "Wenn du eine Funktion erkennst," & vbCrLf & _
        "die du geschrieben hast und" & vbCrLf & _
        "namendlich erwähnt werden" & vbCrLf & _
        "möchtest, schreibe mir eine" & vbCrLf & _
        "Mail mit Funktionsnamen" & vbCrLf & _
        "Vor- Nachname und evt. Mail" & vbCrLf & _
        vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        
    Call CRTL.PictureBox_ScrollText(picCredits, A, scBLUE)

    Me.Visible = True
    
    Call FRMS.Roll(Me, mW, mH, True, 6)
    Call Timer_Fade_Timer
        
    Timer_Fade.Interval = 1000
    Timer_Fade.Enabled = True

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
        
    Call FRMS.Roll(Me, 0, 0, True, 6)

End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub LabelDLLpath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub LabelHomepage_Click()
        
    Call NET.OpenURL("Http://Woeh.tripod.com")
    
End Sub
Private Sub LabelProductName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub LabelSendNewMail_Click()
        
    Call NET.SendNewMail("woeh@gmx.de", _
                         "Re: " & App.ProductName & _
                         " (" & LabelVersion.Caption & ")", _
                         "Hallo Wolfgang")
    
    Unload Me
    
End Sub
Private Sub LabelVersion_Click()
    Clipboard.Clear
    Clipboard.SetText TextVersion.Text
End Sub
Private Sub LED_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub LabelVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub picCredits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FRMS.FormMove(Me)
End Sub
Private Sub Timer_Fade_Timer()
    Dim zColor As Long
    
    Static Color As Long
    
    zColor = GRAFX.Get_RandomizedColor
    
    Call FRMS.Fade(Me, Color, zColor)
    
    Color = zColor

End Sub
Private Sub TimerLabel_Timer()
    
    Static Count As Long
    
    Count = Count + 1
    
    LabelCount.Caption = 30 - Count
    
    If Count > 29 Then Me.LabelCount.Visible = False: _
                       LabelUpdateText.Visible = False: _
                       LabelUpdateText.Caption = "": _
                       TimerLabel.Enabled = False
    
End Sub
