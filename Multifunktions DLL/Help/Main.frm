VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Main 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Multifunktions DLL - Hilfe"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.Frame Frame3 
         Height          =   5655
         Left            =   2640
         TabIndex        =   5
         Top             =   120
         Width           =   9015
         Begin SHDocVwCtl.WebBrowser WB 
            Height          =   5295
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   8775
            ExtentX         =   15478
            ExtentY         =   9340
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.CheckBox CheckStay 
         Caption         =   "Immer oben"
         Height          =   255
         Left            =   10200
         TabIndex        =   6
         Top             =   5830
         Width           =   1215
      End
      Begin VB.ListBox LBcat 
         Height          =   255
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
         Begin VB.ListBox LBsub 
            Height          =   255
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   2040
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComctlLib.TreeView TV 
            Height          =   5295
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   9340
            _Version        =   393217
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            HotTracking     =   -1  'True
            SingleSel       =   -1  'True
            Appearance      =   1
         End
      End
      Begin VB.Label LabelCount 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   5830
         Width           =   480
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "          ?"
      Begin VB.Menu mnuAll 
         Caption         =   "Alle Funktionen"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Über"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '* ********************************************
    '* Dieses Hilfesystem ist selbstregistrierend *
    '**********************************************
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Const SYNCHRONIZE = &H100000

Dim File As String
Dim A() As String
Private Sub CheckStay_Click()
    Dim wAPI As New Window
    
    Call wAPI.StayOnTop(Me.hWnd, Not CBool(CheckStay.Value))
    
End Sub
Private Sub Form_Load()
    Dim P As Integer
    Dim Path As String
    Dim F As New Forms
    Dim wA As New Window
    Dim L As Long, T As Long
        
    Path = App.Path
    If Right(Path, 1) = "\" Then Path = Mid(Path, 1, Len(Path) - 1)
    
    For P = Len(Path) To 1 Step -1
        If Mid(Path, P, 1) = "\" Then Exit For
    Next P
    
    Path = Mid(Path, 1, P) & "WindowAPI.dll"
    
    If Not Exist(Path) Then
        MsgBox "Can't find '" & Path & "' !", vbCritical + vbOKOnly, "Error"
        End
    End If
    
    If Not RegisterComponents(Path) Then _
        MsgBox "Can't register '" & Path & _
               "' !", vbCritical + vbOKOnly, "Error": _
        End
    
    Form1.Show
    Call Init
    Unload Form1

    Call F.Center(Me, ccMIDDLE)
    
    L = Me.Left
    T = Me.Top
    
    Me.Top = 0 - Me.Height
    
    Me.Show
    
    TV.Nodes(1).Expanded = False

    Me.SetFocus
    
End Sub
Private Sub Init()
    Dim P As Long, X As Long, C As Long, Count As Long, y As Long
    Dim R$, File As String, mPath As String, Path As String, I As String
    Dim LB As New LB_CB
    Dim All As String

    Path = App.Path
    If Right(Path, 1) = "\" Then _
        Path = Mid(Path, 1, Len(Path) - 1)
    
    I = Path & "\Index.htm"
    
    TV.Nodes.Clear
    
    Call LB.AddDirsFromPath(LBcat, Path)
    
    For P = 0 To LBcat.ListCount - 1
        TV.Nodes.Add , , , LBcat.List(P)
        mPath = Path & "\" & LBcat.List(P)
        
        Form1.Label4.Caption = CInt(Form1.Label4.Caption) + 1
        Form1.Label4.Refresh
        
        Call LB.AddFilesFromPath(LBsub, mPath)
        y = 0
        
        For X = 0 To LBsub.ListCount - 1
            File = LBsub.List(X)
            
            If Right$(LCase$(File), 4) = ".htm" Then
                Count = Count + 1

                R$ = Replace(File, ".htm", "")
            
                All = All & LBcat.List(P) & "." & R$ & vbCrLf
            
                C = IIf(X - y, tvwNext, tvwChild)
            
                TV.Nodes.Add TV.Nodes.Count, C, , R$
                TV.Nodes(TV.Nodes.Count).Tag = mPath & "\" & File
        
                Form1.Label5.Caption = CInt(Form1.Label5.Caption) + 1
                Form1.Label5.Refresh
            Else
                y = y + 1
            End If
        Next X
    Next P
    
    LabelCount.Caption = Count & _
                         " Subs/Funktionen in " & _
                         LBcat.ListCount & _
                         " Kategorien"
                         
    'Clipboard.Clear
    'Clipboard.SetText All

    A = Split(All, vbCrLf)
    
    WB.Navigate I
    
End Sub
Private Sub mnuAbout_Click()
    Dim m As New Misc
    Call m.About
End Sub
Private Sub mnuAll_Click()
    Dim Frm As New Forms
    Dim LB As New LB_CB
    
    Load Form2
    
    Call Frm.Center(Form2, ccMIDDLE)
    Call LB.AddArray(Form2.LBall, A, True)
    
    Form2.Show 1
    
End Sub
Private Sub TV_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub
Public Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim P As Integer
    Dim R$
    
    If Node.Tag <> "" Then
        File = Node.Tag
        
        WB.Navigate File
        
        For P = Len(File) To 1 Step -1
            If Mid(File, P, 1) = "\" Then Exit For
        Next P
        
        R$ = Mid(File, P + 1)
        R$ = Replace(R$, ".htm", "")
    
        Frame3.Caption = Node.Parent.Text & " - " & R$
    End If
    
End Sub
Public Function RegisterComponents(ByVal ComponentPath As String, _
                                   Optional UnRegister As Boolean = False) _
                                        As Boolean
    Dim sProc As String
    Dim Lib As Long, r1 As Long, r2 As Long, Thread As Long

    On Local Error GoTo Quit

    Lib = LoadLibrary(ComponentPath)
  
    If Lib Then
        sProc = IIf(UnRegister, "DllUnregisterServer", _
                                "DllRegisterServer")
        r1 = GetProcAddress(Lib, sProc)
        
        If r1 Then
            Thread = CreateThread(ByVal 0, 0, ByVal r1, _
                                  ByVal 0, 0, r2)
            If Thread Then
                r2 = WaitForSingleObject(Thread, 10000)
                
                If r2 Then _
                    Call FreeLibrary(Lib): _
                    r2 = GetExitCodeThread(Thread, r2): _
                    Call ExitThread(r2): _
                    Exit Function
            
                Call CloseHandle(Thread)
            End If
        End If
        Call FreeLibrary(Lib)
    End If
  
    RegisterComponents = True
  
Quit:
End Function
Private Function Exist(ByVal File As String) As Boolean
    Dim ff As Integer

    On Local Error GoTo Quit
    
    ff = FreeFile
    
    Open File For Input As ff
    Close ff

    Exist = True
    
Quit:
End Function
Public Function TextFile_Read(ByVal TF_TextFile As String) As String
    Dim FN As Integer
    
    On Local Error GoTo Quit
    
    FN = FreeFile()
    
    Open TF_TextFile For Binary Access Read As #FN
        TextFile_Read = Space(LOF(FN))
        Get #FN, , TextFile_Read
    Close #FN
    
    If Right(TextFile_Read, 2) = vbCrLf Then _
        TextFile_Read = Left(TextFile_Read, Len(TextFile_Read) - 2)

    Exit Function
    
Quit:
    TextFile_Read = ""
    
End Function

