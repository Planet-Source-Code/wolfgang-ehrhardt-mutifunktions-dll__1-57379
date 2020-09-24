VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   0  'Kein
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   120
      Width           =   2925
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim myPic As PictureBox
Dim pHwnd As Long
Dim Fcolor As scCOLOR
Dim T() As String
Public Function SetScroll(PB As PictureBox, _
                          Text As String, _
                          FadeColor As scCOLOR) _
                                As Boolean
    Dim P As Long, oSM1 As Long, oSM2 As Long
    Dim tmpColor As Double
    
    On Local Error GoTo Quit
    
    T = Split(Text, vbCrLf)
    If Not IsArray(T) Then Exit Function
    
    Set myPic = PB
    pHwnd = myPic.Parent.Hwnd
    
    oSM1 = myPic.Parent.ScaleMode
    oSM2 = myPic.ScaleMode
    
    myPic.Parent.ScaleMode = Me.ScaleMode
    myPic.ScaleMode = picBuffer.ScaleMode
    
    Fcolor = FadeColor
    
    picBuffer.Width = PB.Width
    picBuffer.Height = PB.Height
        
    myPic.Parent.ScaleMode = oSM1
    myPic.ScaleMode = oSM2
        
    tmpColor = 255
        
    For P = 0 To picBuffer.ScaleHeight Step 5
        Select Case FadeColor
            Case scBLUE
                picBuffer.Line (-1, P - 1)-(picBuffer.ScaleWidth, P + 5), RGB(0, 0, tmpColor), BF
            Case scGREEN
                picBuffer.Line (-1, P - 1)-(picBuffer.ScaleWidth, P + 5), RGB(0, tmpColor, 0), BF
            Case scRED
                picBuffer.Line (-1, P - 1)-(picBuffer.ScaleWidth, P + 5), RGB(tmpColor, 0, 0), BF
            Case Else
                GoTo Quit
        End Select
            
        tmpColor = tmpColor + (5 * (0 - 255) / picBuffer.ScaleHeight)
    Next P
    
    picBuffer.Picture = picBuffer.Image
    PB.Picture = picBuffer.Picture
    
    Timer.Interval = 1
    Timer.Enabled = True
    
    SetScroll = True
    
Quit:
End Function
Private Sub Credits()
    Dim NumLines As Long, lX As Long, lY As Long, Zahl As Long
    Dim temp As Double
    
    With myPic
        .ForeColor = vbWhite
        .BackColor = vbWhite
        .AutoRedraw = True
        
        picBuffer.AutoRedraw = True

        NumLines = UBound(T)
        
        lX = .ScaleLeft
        lY = .ScaleHeight
        
        Do
            For Zahl = 0 To NumLines
                If WIN.Get_TaskID(pHwnd) = 0 Then _
                    Unload Me: _
                    Exit Sub
                
                If Zahl = 0 Then .Cls
                
                .CurrentY = lY + (Zahl * (.FontSize * 15) + (96 * Zahl))
                .CurrentX = (.ScaleWidth / 2) - (.TextWidth(T(Zahl)) / 2)
        
                If .CurrentY * 2 > 160 Then
                    temp = .CurrentY * 0.125
                Else
                    temp = 0
                End If
        
                Select Case Fcolor
                    Case scBLUE
                        .ForeColor = RGB(temp, temp, 255)
                    Case scGREEN
                        .ForeColor = RGB(temp, 255, temp)
                    Case scRED
                        .ForeColor = RGB(255, temp, temp)
                End Select
                
                If Zahl = NumLines And .CurrentY < -25 Then _
                    lX = .ScaleLeft: _
                    lY = .ScaleHeight: _
                    Zahl = 0
        
                myPic.Print T(Zahl)
            Next Zahl
    
            lY = lY - 15
    
            Call Wait(20)
        Loop
    End With

End Sub
Public Function Wait(ByVal Delay As Long)
    Dim TickCount As Long

    TickCount = GetTickCount
    
    While (TickCount + Delay) > GetTickCount
        DoEvents
    Wend
    
End Function
Private Sub Form_Unload(Cancel As Integer)
    'Set myPic = Nothing
End Sub
Private Sub Timer_Timer()
    Timer.Enabled = False
    Call Credits
    Unload Me
End Sub
