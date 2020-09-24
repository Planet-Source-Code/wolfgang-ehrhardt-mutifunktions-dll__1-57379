VERSION 5.00
Begin VB.Form ProgBar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Kein
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Sanduhr
   ScaleHeight     =   1410
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   200
      Min             =   100
      TabIndex        =   0
      Top             =   1920
      Value           =   100
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "FrameCaption"
      Height          =   1335
      Left            =   120
      MousePointer    =   11  'Sanduhr
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         MousePointer    =   11  'Sanduhr
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   343
         TabIndex        =   3
         Top             =   840
         Width           =   5175
         Begin VB.Label Label1 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "xxx%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   2160
            TabIndex        =   4
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Label Message 
         Alignment       =   2  'Zentriert
         Caption         =   "Message"
         Height          =   495
         Left            =   120
         MousePointer    =   11  'Sanduhr
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "ProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyMax As Integer
Dim MyMin As Integer

Dim Teiler As Long
Private Sub Form_Load()
    Me.Visible = False
End Sub
Private Sub ProgressBar(ByVal Prg As Integer, _
                        ByVal Min As Integer, _
                        ByVal Max As Integer)
    Dim Fx As Long
    
    If Prg < Min Or Prg > Max Or Max <= Min Then Exit Sub
    
    Prg = Int((100 / (Max - Min)) * (Prg - Min))
    
    If Prg Then
        With Picture1
            .Cls
        
            Fx = (.ScaleWidth - 2) / 100 * Prg
            Picture1.Line (0, 0)-(Fx + 1, .ScaleHeight - 1), _
                           &H8000000D, BF
            .CurrentX = Fx + 3
            .CurrentY = 0
        
            'Picture1.Print Trim$(CStr(Prg) & " %")
        End With
    End If
    
    Label1.Caption = CStr(Prg) & "%"
    Label1.Refresh
    
End Sub
Public Sub SetOption(ByVal Min As Long, _
                     ByVal Max As Long, _
                     ByVal Mesage As String, _
                     ProgressBarIs As Boolean)
    
    If Max > 32767 Then
        Teiler = 10
        
        Do While Max / Teiler > 32767
            Teiler = Teiler * 10
        Loop
    End If
    
    If Teiler = 0 Then Teiler = 1
    
    MyMax = Max / Teiler
    MyMin = Min
    
    HScroll.Min = MyMin
    HScroll.Max = MyMax
    HScroll.Value = IIf(MyMin - 1 < 0, 0, MyMin)
    
    Call ProgressBar(0, MyMin, MyMax)

    Frame1.Caption = Mesage
    Message.Caption = ""
        
    Me.Width = 5670
    
    If Not ProgressBarIs Then
        Me.Height = 810
        Frame1.Height = 735
        Picture1.Visible = False
        Message.Height = 375
    Else
        Me.Height = 1305
        Frame1.Height = 1215
        Picture1.Visible = True
        Message.Height = 495
    End If
    
    Call FRMS.Center(Me, ccMIDDLE)
    Call FRMS.StayOnTop(Me)

    Me.Visible = True
    
    Me.Refresh
    Call MISC.Sleep(0.01)
    Me.Refresh
    
End Sub
Public Sub SetValue(ByVal Value As Long)
    If CLng(Value / Teiler) <= HScroll.Max Then _
        HScroll.Value = Value / Teiler: _
        Call ProgressBar(Value / Teiler, MyMin, MyMax)
End Sub

