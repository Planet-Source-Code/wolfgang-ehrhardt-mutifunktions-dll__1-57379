VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Schließen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox LBall 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub LBall_Click()
    Dim P As Long, X As Long
    Dim R$, C$, F$
        
    Me.Hide
    
    R$ = LBall.List(LBall.ListIndex)
    
    X = InStr(R$, ".")
    
    C$ = LCase$(Mid$(R$, 1, X - 1))
    F$ = LCase$(Mid$(R$, X + 1))
    
    With Main.TV
        For P = 1 To .Nodes.Count
            If LCase$(.Nodes(P).Text) = C$ Then .Nodes(P).Expanded = True
            
            If LCase$(.Nodes(P).Text) = F$ Then _
                .Nodes(P).Selected = True: _
                Call Main.TV_NodeClick(.Nodes(P)): _
                .SetFocus: _
                Unload Me: _
                Exit Sub
        Next P
    End With
    
End Sub
