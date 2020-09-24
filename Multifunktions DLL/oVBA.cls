VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Office"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Word_Start() As Boolean
    Dim Word As Object
    
    On Local Error GoTo Quit
    
    Set Word = CreateObject("Word.Application")

    Word.WindowState = 1
    Word.Visible = True
  
    Word_Start = True
    
Quit:
    If Err.Number <> 0 Then Set Word = Nothing
    
End Function
Public Function Word_OpenDocument(ByVal DocPath As String, _
                                  ByVal WriteProtected As Boolean) _
                                        As Boolean
    Dim Word As Object
    
    On Local Error GoTo Quit
    
    If Not F.Exist(DocPath) Then Set Word = Nothing: _
                                 Exit Function
    
    Set Word = CreateObject("Word.Application")

    Word.WindowState = 1
    Word.Visible = True
  
    Word.Documents.Open DocPath, , WriteProtected

    Word_OpenDocument = True
    
Quit:
    If Err.Number <> 0 Then Set Word = Nothing
    
End Function
