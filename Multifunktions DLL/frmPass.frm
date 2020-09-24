VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   300
      Left            =   1200
      Picture         =   "frmPass.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   285
      TabIndex        =   10
      Top             =   2880
      Width           =   285
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   285
      Left            =   2880
      Picture         =   "frmPass.frx":04F2
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   2880
      Width           =   300
   End
   Begin VB.CommandButton But_Ok 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton But_Cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   450
      Left            =   120
      Picture         =   "frmPass.frx":09A8
      ScaleHeight     =   450
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   8
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UserInput As Boolean
Public UserStr As String

Dim MyCommand As PassCommand
Dim MyPassword As String
Private Sub But_Cancel_Click()
    
    UserInput = True
    UserStr = Chr$(0)
    
    Me.Visible = False
    
End Sub
Private Sub But_Ok_Click()
    Dim Result As VbMsgBoxResult
    
    Select Case MyCommand
        Case psCHANGEPASSWORD
            If Text(0).Text <> MyPassword Then
                Result = MsgBox("Password was not changed !", _
                                 vbRetryCancel + vbCritical, _
                                 "Old Password Missmatch")
                If Result = vbCancel Then
                    Call UserOK(MyPassword)
                Else
                    Call Retry
                End If
            
                Exit Sub
            End If
            
            If Text(1).Text <> Text(2).Text Then
                Result = MsgBox("Password was not changed !", _
                                vbCritical + vbRetryCancel, _
                                "New Password missmatch")
                If Result = vbRetry Then
                    Text(1).Text = ""
                    Text(2).Text = ""
                    Text(1).SetFocus
                Else
                    Call UserOK("")
                End If
            
                Exit Sub
            End If
            
            If Text(1) = "" Then
                Result = MsgBox("This will disable your current Password !", _
                                vbQuestion + vbYesNo, _
                                "Password disable")
                If Result = vbNo Then _
                    Text(1).SetFocus: _
                    Exit Sub
            End If
            
            Call UserOK(Text(1).Text)
        Case psQUERYPASSWORD
            If MyPassword <> "" And Text(0).Text <> MyPassword Then
                Result = MsgBox("You are not logged in ! ", _
                                vbCritical + vbRetryCancel, _
                                "Wrong Password")
                If Result = vbRetry Then
                    Call Retry
                Else
                    Call But_Cancel_Click
                End If
                
                Exit Sub
            End If
                    
            Call UserOK(Text(0).Text)
        Case psCREATEPASSWORD
            If Text(0) <> Text(1) Then
                Result = MsgBox("No Password was set!", _
                         vbCritical + vbRetryCancel, _
                         "Password missmatch")
                
                If Result = vbRetry Then
                    Call Retry
                Else
                    Call But_Cancel_Click
                End If
            Else
                If Text(0) = "" Then
                    Result = MsgBox("This will disable your Password !", _
                                    vbInformation + vbYesNo, _
                                    "Password disabled")
                    If Result = vbNo Then _
                        Call Retry: _
                        Exit Sub
                End If
                
                Call UserOK(Text(0).Text)
            End If
    End Select
    
End Sub
Private Sub Retry()
    Text(0).Text = ""
    Text(1).Text = ""
    Text(2).Text = ""
    Text(0).SetFocus
End Sub
Private Sub UserOK(Str As String)
    UserInput = True
    UserStr = Str
    Me.Visible = False
End Sub
Private Sub Form_Load()
        
    Me.Visible = False
    Me.Width = 4710
    
    Call FRMS.Center(Me, ccMIDDLE)
    Call FRMS.StayOnTop(Me)

End Sub
Public Function SetOptions(Command As PassCommand, _
                           Password As Variant, _
                           Optional MaxPasswordLenght As Long = 0)
                           
    Dim P As Long, Modal As Long
    Dim Bh As Long, Ph As Long, mH As Long
    
    If MaxPasswordLenght > 0 Then
        For P = Text.LBound To Text.UBound
            Text(P).MaxLength = MaxPasswordLenght
        Next P
    End If
    
    MyCommand = Command
    MyPassword = Password
    
    Select Case MyCommand
        Case psCHANGEPASSWORD
            Label(0).Caption = "Enter old Password"
            Label(1).Caption = "Enter new Password"
            Label(2).Caption = "Confirm new Password"
            Me.Caption = "Change current Password"
            
            For P = 0 To 2
                Text(P).Visible = True
                Label(P).Visible = True
            Next P
            
            mH = 3855
            Bh = 2760
            Ph = 2880
        Case psCREATEPASSWORD
            Label(0).Caption = "Enter Password"
            Label(1).Caption = "Confirm Password"
            Me.Caption = "Create Password"
            
            Label(2).Visible = False
            Text(2).Visible = False
        
            mH = 2970
            Bh = 1920
            Ph = 2040
        Case psQUERYPASSWORD
            Label(0).Caption = "Enter Password"
            Me.Caption = "Password required"
            
            For P = 1 To Label.UBound
                Label(P).Visible = False
                Text(P).Visible = False
            Next P
            
            mH = 2100
            Bh = 1080
            Ph = 1200
        Case Else
            UserInput = True
            UserStr = ""
            Me.Visible = False
    End Select

    Me.Height = mH
    But_Cancel.Top = Bh
    But_Ok.Top = Bh
    Picture2.Top = Ph
    Picture3.Top = Ph
        
    Me.Show 1
    
End Function
Private Sub Picture2_Click()
    Call But_Cancel_Click
End Sub
Private Sub Picture3_Click()
    Call But_Ok_Click
End Sub
