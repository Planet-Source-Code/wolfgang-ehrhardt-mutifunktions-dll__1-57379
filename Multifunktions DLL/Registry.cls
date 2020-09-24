VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Any) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_Str Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValueA Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Enum RegRootNFO
    HKEY_CLASSES_ROOT_ = HKEY_CLASSES_ROOT
    HKEY_CURRENT_USER_ = HKEY_CURRENT_USER
    HKEY_LOCAL_MACHINE_ = HKEY_LOCAL_MACHINE
    HKEY_USERS_ = HKEY_USERS
    HKEY_PERFORMANCE_DATA_ = HKEY_PERFORMANCE_DATA
    HKEY_CURRENT_CONFIG_ = HKEY_CURRENT_CONFIG
    HKEY_DYN_DATA_ = HKEY_DYN_DATA
End Enum

Private Const BUFFER_SIZE As Long = 255

Private Const ERROR_NO_MORE_ITEMS = 259&

Private Const READ_CONTROL As Long = &H20000

Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ _
                        Or KEY_QUERY_VALUE _
                        Or KEY_ENUMERATE_SUB_KEYS _
                        Or KEY_NOTIFY) _
                       And (Not SYNCHRONIZE))


Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or _
                               KEY_SET_VALUE Or _
                               KEY_CREATE_SUB_KEY Or _
                               KEY_ENUMERATE_SUB_KEYS Or _
                               KEY_NOTIFY Or _
                               KEY_CREATE_LINK
Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7

Private Const REG_OPTION_NON_VOLATILE = &H0
Public Function RegKeyExist(ByVal Root As RegRootNFO, _
                            ByVal Key As String) As Boolean
    Dim hKey As Long
    
    If RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey) = 0 Then _
        Call RegCloseKey(hKey): _
        RegKeyExist = True

End Function
Public Function RegKeyCreate(ByVal Root As RegRootNFO, _
                             ByVal NewKey As String, _
                             Optional KeyValue As String = "") _
                                As Boolean
    Dim Result As Long, hKey As Long
  
    If RegCreateKey(Root, NewKey, hKey) = ERROR_SUCCESS Then
        If KeyValue = "" Then
            RegKeyCreate = True
        Else
            If RegSetValue(hKey, vbNullString, REG_SZ, KeyValue, 0) = ERROR_SUCCESS Then _
                RegKeyCreate = True
        End If
    End If
    
    Call RegCloseKey(hKey)

End Function
Public Function RegKeyDelete(ByVal Root As RegRootNFO, _
                             ByVal Key As String) _
                                As Boolean
    RegKeyDelete = (RegDeleteKey(Root, Key) = ERROR_SUCCESS)
End Function
Public Function RegFieldDelete(ByVal Root As RegRootNFO, _
                               ByVal Key As String, _
                               ByVal Field As String) _
                                    As Boolean
    Dim hKey As Long
    
    If RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, _
                    hKey) = ERROR_SUCCESS Then
        If RegDeleteValue(hKey, Field) = ERROR_SUCCESS Then _
            RegFieldDelete = True: _
            Call RegCloseKey(hKey)
    End If
    
End Function
Public Function RegValueSet(ByVal Root As RegRootNFO, _
                            ByVal Key As String, _
                            ByVal Field As String, _
                            ByVal Value As Variant, _
                            Optional CreateKeyPath As Boolean = False) _
                                As Boolean
    Dim Result As Long, hKey As Long, L As Long
    Dim P As Integer
    Dim S As String
    
    If CreateKeyPath Then
        If Not RegKeyExist(Root, Key) Then
            For P = 1 To Len(Key)
                If Mid$(Key, P, 1) = "\" Then _
                    Call RegKeyCreate(Root, Mid$(Key, 1, P - 1))
            Next P
            
            Call RegKeyCreate(Root, Key)
        End If
    End If
    
    If RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, _
                    hKey) = ERROR_SUCCESS Then
    
        Select Case VarType(Value)
            Case vbInteger, vbLong
                L = CLng(Value)
                Result = RegSetValueEx(hKey, Field, 0, _
                                       REG_DWORD, L, 4)
            Case vbString, vbBoolean
                S = CStr(Value)
                Result = RegSetValueEx_Str(hKey, Field, 0, _
                                           REG_SZ, S, _
                                           Len(S) + 1)
        End Select
        
        If Result = ERROR_SUCCESS Then RegValueSet = True
    End If
    
    RegCloseKey (hKey)

End Function
Public Function RegValueGet(ByVal Root As RegRootNFO, _
                            ByVal Key As String, _
                            ByVal Field As String) As Variant
    Dim Result As Long, hKey As Long, dwType As Long
    Dim Lng As Long, L As Long
    Dim Buffer As String
  
    If RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey) = ERROR_SUCCESS Then
        If RegQueryValueEx(hKey, Field, 0&, _
                           dwType, ByVal 0&, L) = ERROR_SUCCESS Then
        
            Select Case dwType
                Case REG_SZ
                    Buffer = Space$(L + 1)
                    Result = RegQueryValueEx(hKey, Field, _
                                             0&, dwType, _
                                             ByVal Buffer, L)
                    If Result = ERROR_SUCCESS Then _
                        RegValueGet = Left$(Buffer, Len(Buffer) - 2)
                Case REG_DWORD
                    Result = RegQueryValueEx(hKey, Field, _
                                             0&, dwType, _
                                             Lng, L)
                    If Result = ERROR_SUCCESS Then _
                        RegValueGet = Lng
            End Select
        End If
    End If
    
    Call RegCloseKey(hKey)
    
End Function
Public Function RegEnumKey(ByVal Root As RegRootNFO, _
                           ByVal Key As String) As String()
    Dim hKey As Long, cnt As Long, Ret As Long
    Dim sName As String, KeyList() As String
        
    On Local Error GoTo Quit
    
    Erase KeyList
    
    Ret = BUFFER_SIZE
    cnt = 0
    
    If RegOpenKey(Root, Key, hKey) <> 0 Then Exit Function
    
    sName = Space(BUFFER_SIZE)
                
    While RegEnumKeyEx(hKey, cnt, sName, Ret, ByVal 0&, _
                       vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
        ReDim Preserve KeyList(cnt)
        KeyList(cnt) = Left$(sName, Ret)
        
        cnt = cnt + 1

        sName = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
    Wend
    
    Call RegCloseKey(hKey)

Quit:
    If Err.Number <> 0 Then Erase KeyList
    
    RegEnumKey = KeyList
    
End Function
Public Function RegEnumValue(ByVal Root As RegRootNFO, _
                             ByVal Key As String, _
                             ItemList() As String, _
                             ValueList() As String) As Boolean
    Dim hKey As Long, cnt As Long, Ret As Long, RetData As Long
    Dim sName As String, sData As String, A() As String
    
    On Local Error GoTo Quit
    
    Erase ItemList
    Erase ValueList
    
    Ret = BUFFER_SIZE
    cnt = 0
    
    If RegOpenKey(Root, Key, hKey) Then Exit Function
    
    sName = Space(BUFFER_SIZE)
    sData = Space(BUFFER_SIZE)
        
    Ret = BUFFER_SIZE
    RetData = BUFFER_SIZE
        
    While RegEnumValueA(hKey, cnt, sName, Ret, 0, ByVal 0&, _
                        ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
        If RetData > 0 Then _
            ReDim Preserve ItemList(cnt): _
            ReDim Preserve ValueList(cnt): _
            ItemList(cnt) = Left$(sName, Ret): _
            ValueList(cnt) = Left$(sData, RetData - 1): _
            cnt = cnt + 1
            
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
            
        Ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
    Wend
    
    
    
    If (Not ARR.isArray(ItemList)) And (Not ARR.isArray(ItemList)) Then _
        Erase ItemList: _
        Erase ValueList
    
    Call RegCloseKey(hKey)
    RegEnumValue = True

Quit:
    Erase A
    
End Function
Public Function RegEnum(ByVal Root As RegRootNFO, _
                        ByVal Key As String, _
                        KeyList() As String, _
                        ItemList() As String, _
                        ValueList() As String) As Boolean
    Dim K() As String
    
    On Local Error GoTo Quit
    
    K = RegEnumKey(Root, Key)
    
    If ARR.isArray(K) Then _
        Call RegEnumValue(Root, Key, ItemList, ValueList): _
        RegEnum = True
    
Quit:
End Function
Private Function GetRootWord(Root As RegRootNFO) As String
    Dim R$
    
    Select Case Root
        Case HKEY_CLASSES_ROOT_: R$ = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER_: R$ = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE_: R$ = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS_: R$ = "HKEY_USERS"
        Case HKEY_PERFORMANCE_DATA_: R$ = "HKEY_PERFORMANCE_DATA"
        Case HKEY_CURRENT_CONFIG_: R$ = "HKEY_CURRENT_CONFIG"
        Case HKEY_DYN_DATA_: R$ = "HKEY_DYN_DATA"
        Case Else: R$ = ""
    End Select

    GetRootWord = R$

End Function
Public Function ExportRegKey(Root As RegRootNFO, _
                             ByVal Key As String, _
                             ByVal File As String) As Boolean
        
    On Local Error GoTo Quit
    
    If F.fKill(File) = -1 Then Exit Function
    
    If Not RegKeyExist(Root, Key) Then Exit Function
    
    If Not F.Create_File(File, True) Then Exit Function
    
    Call Shell("Regedit /e /s " & File & " " _
               & GetRootWord(Root) & "\" & Key, _
               vbHide)

    If FileLen(File) > 0 Then ExportRegKey = True
    
Quit:
End Function
