VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function MakePath Lib "imagehlp.dll" Alias "MakeSureDirectoryPathExists" (ByVal lpPath As String) As Long

Private Type BROWSEINFO
    hOwner         As Long
    pidlRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfn           As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type SHITEMID
    CB   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum DirectoryOperationNFO
    doCOPY = 1
    doMOVE = 2
    doRENAME = 3
    doDELETE = 4
End Enum

Private Const BIF_RETURNONLYFSDIRS = &H1
Public Function DirectoryDeleteFiles(ByVal Directory As String, _
                                     ByVal FileList As Variant, _
                                     Optional DeleteAllFiles As Boolean = False) _
                                            As Long
    Dim F() As Variant, Dl() As Variant
    Dim P As Long, D As String
    Dim File As New WindowAPI.File
    
    On Local Error GoTo Quit
    
    Directory = GetDir(Directory)
    
    If Not Exist(Directory) Then DirectoryDeleteFiles = 0: _
                                 Exit Function
    
    If DeleteAllFiles Then
        If Not ReadDirectory(Directory, Dl, F, True, True, True) Then
            DirectoryDeleteFiles = -1
            Exit Function
        Else
            FileList = F
        End If
    End If
        
    If Not ARR.isArray(FileList) Then
        ReDim F(0)
        F(0) = FileList
    Else
        F = FileList
    End If

    D = Dir(Directory & "*.*", vbDirectory + vbArchive _
                              + vbHidden + vbReadOnly _
                              + vbSystem)
    
    Do While Len(D)
        If D <> "." And D <> ".." Then
            If (GetAttr(Directory & D) And vbDirectory) <> vbDirectory Then
                If ARR.Exist(F, D) Then _
                    If File.fKill(Directory & D) < 1 Then _
                        DirectoryDeleteFiles = 2
            End If
        End If
        
        D = Dir
    Loop
    
    If DirectoryDeleteFiles = 0 Then DirectoryDeleteFiles = 1
    
Quit:
    If Err.Number <> 0 Then DirectoryDeleteFiles = Err.Number
    
End Function
Public Function Exist(ByVal Path As String, _
                      Optional sCreatePath As Boolean = False) _
                            As Boolean
    Dim fs As Object
   
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FolderExists(Path) Then
        Exist = True
    Else
        If sCreatePath Then Exist = CreatePath(Path)
    End If
        
End Function
Public Function BrowseForDirectory(ByVal hWnd As Long, _
                                   ByVal Title As String) _
                                        As String
    Dim pidl As Long, R As Long
    Dim bi As BROWSEINFO
    Dim idl As ITEMIDLIST
    Dim Pos As Integer
    Dim sPath As String
    
    bi.hOwner = hWnd
    bi.pidlRoot = 0&
    bi.lpszTitle = Title
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    
    pidl& = SHBrowseForFolder(bi)
    sPath$ = Space$(512)
    R = SHGetPathFromIDList(ByVal pidl&, ByVal sPath$)
    
    If R Then _
        Pos = InStr(sPath$, Chr$(0)): _
        BrowseForDirectory = Left$(sPath$, Pos - 1)
    
End Function
Public Function CreatePath(ByVal Path As String) As Boolean
    Call MakePath(GetDir(Path))
    If Exist(Path, False) Then CreatePath = True
End Function
Public Function Find(ByVal DirectoryToFind As String, _
                     ByVal StartSearchPath As String, _
                     Optional SearchAllDrives As Boolean = False, _
                     Optional ShowProgress As Boolean = False) _
                            As String
    Dim P As Long
    Dim sDRV As String, cDrv As String
    
    If StartSearchPath = "" And Not SearchAllDrives Then _
        Exit Function
    
    If ShowProgress Then pModule.LoadProgbar = True
    
    If SearchAllDrives Then
        sDRV = DRV.Get_AllDrives
        
        sDRV = Replace(sDRV, "A", "")
        sDRV = Replace(sDRV, "B", "")
        
        For P = 1 To Len(sDRV)
            cDrv = Mid$(sDRV, P, 1) & ":\"
            Find = GetAllFiles(cDrv, DirectoryToFind, _
                               True, ShowProgress)
            If Len(Find) Then Exit For
        Next P
    Else
        Find = GetAllFiles(StartSearchPath, DirectoryToFind, _
                           True, ShowProgress)
    End If
    
    If ShowProgress Then Unload pModule.PBar: _
                         Set pModule.PBar = Nothing

End Function
Public Function isDirectoryEmpty(ByVal Directory As String) _
                                        As Boolean
    isDirectoryEmpty = (PathIsDirectoryEmpty(Directory) = 1)
End Function
Public Function ReadDirectory(ByVal Directory As String, _
                              DirList As Variant, _
                              FileList As Variant, _
                              Optional SystemFiles As Boolean = True, _
                              Optional ReadOnlyFiles As Boolean = True, _
                              Optional HiddenFiles As Boolean = True) _
                                    As Boolean
    Dim File As String
    Dim X As Long, Y As Long
    Dim At As VbFileAttribute
    
    Erase DirList: Erase FileList
    
    If Not Exist(Directory) Then Exit Function
    
    At = vbDirectory
    
    If SystemFiles Then At = At + vbSystem
    If ReadOnlyFiles Then At = At + vbReadOnly
    If HiddenFiles Then At = At + vbHidden

    Directory = GetDir(Directory)
    
    File = Directory
    File = Dir(File & "*.*", At)
    
    On Local Error Resume Next
    
    Do While Len(File)
        If File <> "." And File <> ".." Then
            If (GetAttr(Directory & File) And vbDirectory) = vbDirectory Then
                If Err = 0 Then _
                    ReDim Preserve DirList(X): _
                    DirList(X) = File: _
                    X = X + 1
            Else
                ReDim Preserve FileList(Y)
                FileList(Y) = File
                Y = Y + 1
            End If
            Err.Clear
        End If
        File = Dir
    Loop
    
    ReadDirectory = True
    
Quit:
End Function
Public Function isDirectoryContainingDirs(ByVal Directory As String) _
                                             As Boolean
    Dim D() As Variant, F() As Variant
    
    If ReadDirectory(Directory, D, F) Then _
        If ARR.isArray(D) Then isDirectoryContainingDirs = True
    
End Function
Public Function isDirectoryContainingFiles(ByVal Directory As String) _
                                             As Boolean
    Dim D() As Variant, F() As Variant
    
    If ReadDirectory(Directory, D, F) Then _
        If ARR.isArray(F) Then isDirectoryContainingFiles = True
    
End Function
Public Function DirectroryOperation( _
                            Operation As DirectoryOperationNFO, _
                            ByVal Source As String, _
                            ByVal Target As String, _
                            Optional IncludingSubDirectorys As Boolean = False, _
                            Optional MoveToRecycleBin As Boolean = False) _
                                    As Boolean

    DirectroryOperation = SH_FileOperation(CLng(0), Operation, _
                                           Source, Target, _
                                           IncludingSubDirectorys, _
                                           False, False, False, _
                                           MoveToRecycleBin)
End Function
Public Function DirectroryOperationUsingShell(Operation As DirectoryOperationNFO, _
                                ByVal hWnd As Long, _
                                ByVal Source As String, _
                                ByVal Target As String, _
                                Optional IncludingSubDirectorys As Boolean = False, _
                                Optional Confirm As Boolean = True, _
                                Optional ShowDialogs As Boolean = True, _
                                Optional ShowProgress As Boolean = True, _
                                Optional MoveToRecycleBin As Boolean = False) _
                                        As Boolean
                                        
    DirectroryOperationUsingShell = SH_FileOperation( _
                                            hWnd, Operation, _
                                            Source, Target, _
                                            IncludingSubDirectorys, _
                                            Confirm, ShowDialogs, _
                                            ShowProgress, _
                                            MoveToRecycleBin)
End Function
Public Function PathIsDirectory(Path As String) As Boolean
    PathIsDirectory = Exist(Path)
End Function
Public Function Delete(ByVal Directory As String, _
                       Optional ByVal DeleteContent As Boolean = False) _
                            As Long
    Dim Ok As Boolean
    Dim fso As New FileSystemObject
    
    On Local Error Resume Next
    
    Do While Right$(Directory, 1) = "\"
        Directory = Mid$(Directory, 1, Len(Directory) - 1)
    Loop
    
    If Not Exist(Directory) Then Exit Function
    
    If DeleteContent Then Ok = DeleteTree(Directory): _
                          Delete = IIf(Ok, 1, -1): _
                          Exit Function
        
    Call SetAttr(Directory, vbNormal)
    Err.Clear
           
    Call fso.DeleteFolder(Directory, True)
    
    Delete = IIf(Err.Number = 0, 1, -1)
    
End Function
Public Sub DirectoryIsAccesable(ByVal Directory As String, _
                                CanRead As Boolean, _
                                CanOpenFiles As Boolean, _
                                CanWrite As Boolean, _
                                CanDelete As Boolean, _
                                CanCreateDirs As Boolean, _
                                CanDeleteDirs As Boolean)
    Dim Fl() As String, Dl() As String, File As String, Dir As String
    
    On Local Error GoTo Quit
    
    CanRead = False
    CanOpenFiles = False
    CanWrite = False
    CanDelete = False
    CanCreateDirs = False
    CanDeleteDirs = False
    
    Directory = GetDir(Directory)
    File = Directory & "test.tmp"
    Dir = Str.Get_RandomString(20, True, True, True, True, False)
    Dir = Directory & Dir
    
    If DI.PathIsDirectory(Directory) Then
        CanRead = DI.ReadDirectory(Directory, Dl, Fl, True, True, True)
        
        CanWrite = TXT.TextFile_WriteText(File, "test", True)
        CanOpenFiles = (TXT.TextFile_Read(File) = "test")
        CanDelete = (F.fKill(File) = 1)
        
        CanCreateDirs = CreatePath(Dir)
        CanDeleteDirs = (Delete(Dir) = 1)
    End If
    
Quit:
End Sub
Public Function DelTree(ByVal Directory As String) As Long
    Dim sCurrFile As String
    
    Static C As Long
    
    On Local Error Resume Next
    
    Directory = GetDir(Directory)
    
    If C = 0 Then _
        If Not Exist(Directory) Then DelTree = -2: _
                                     Exit Function
        
    C = C + 1
    
    sCurrFile = Dir(Directory & "*.*", vbDirectory + vbArchive _
                                     + vbHidden + vbReadOnly _
                                     + vbSystem)
     
    Do While Len(sCurrFile) > 0
        If sCurrFile <> "." And sCurrFile <> ".." Then
            If (GetAttr(Directory & sCurrFile) And vbDirectory) = vbDirectory Then
                Call DelTree(Directory & sCurrFile)
                sCurrFile = Dir(Directory & "*.*", vbDirectory)
            Else
                Call SetAttr(Directory & sCurrFile, vbNormal)
                Kill Directory & sCurrFile
                sCurrFile = Dir
            End If
        Else
            sCurrFile = Dir
        End If
    Loop
    
    RmDir Directory
    
    C = C - 1
    DelTree = 1
    
    If C = 0 Then If Exist(Directory) Then DelTree = -1
        
End Function
Public Function DirectoryDeleteSubDirs(ByVal Directory As String, _
                                       ByVal DirList As Variant, _
                                       Optional DeleteAllSubDirs As Boolean = False) _
                                            As Long
    Dim D() As Variant, Fl() As Variant
    Dim Df As String

    On Local Error GoTo Quit
    
    Directory = GetDir(Directory)
    
    If Not Exist(Directory) Then DirectoryDeleteSubDirs = 0: _
                                 Exit Function
    
    If DeleteAllSubDirs Then
        If Not ReadDirectory(Directory, D, Fl, True, True, True) Then
            DirectoryDeleteSubDirs = -1
            Exit Function
        Else
            DirList = D
        End If
    End If
        
    If Not isArray(DirList) Then
        ReDim D(0)
        D(0) = DirList
    Else
        D = DirList
    End If

StartLoop:
    Df = Dir(Directory & "*.*", vbDirectory + vbArchive _
                              + vbHidden + vbReadOnly _
                              + vbSystem)
    
    Do While Len(Df)
        If Df <> "." And Df <> ".." Then
            If (GetAttr(Directory & Df) And vbDirectory) = vbDirectory Then
                If ARR.Exist(D, Df) Then
                    If Not DeleteTree(Directory & Df) Then _
                        DirectoryDeleteSubDirs = 2
                    GoTo StartLoop
                End If
            End If
        End If
        
        Df = Dir
    Loop
    
    If DirectoryDeleteSubDirs = 0 Then _
        DirectoryDeleteSubDirs = 1
    
Quit:
    If Err.Number <> 0 Then DirectoryDeleteSubDirs = Err.Number
    
End Function
Public Sub DirectoryInfo(ByVal Directory As String, _
                         DirectorySize As Variant, _
                         FileCount As Long, _
                         DirectoryCount As Long, _
                         Optional InclSubDirectorys As Boolean = False)
                                        
    DirectorySize = GetAllPath(Directory, _
                               InclSubDirectorys, _
                               FileCount, DirectoryCount)

End Sub
