Attribute VB_Name = "pModule"
Option Explicit

Type ERRORNFO
    Number As String
    Des    As String
End Type

Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function PathisDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Private Const WS_EX_LAYERED = &H80000

Public LoadProgbar As Boolean
Public PBar As New ProgBar

'SFV
Dim m_CRC32Asm() As Byte
Dim m_CRC32 As Long
Dim m_CRC32Table(0 To 255) As Long
Public Function GetBitmapData(pBox As PictureBox) As String
    Dim BitmapHeight As Long
    Dim strBitmapHeight As String, BitmapData As String
    Dim BitmapWidth As Long, BitmapBytes As Long
    Dim BytesPerPixel As Integer
    Dim strBytesPerPixel As String, strBitmapWidth As String
    Dim bmp As BITMAP
    
    GetObject pBox.Image.Handle, Len(bmp), bmp
        
    BitmapHeight = bmp.bmHeight
    BitmapWidth = bmp.bmWidth
    BytesPerPixel = (bmp.bmBitsPixel \ 8)
    BitmapBytes = BitmapHeight * BitmapWidth * BytesPerPixel
    
    BitmapData = Space$(BitmapBytes)
    GetBitmapBits pBox.Image.Handle, BitmapBytes, ByVal BitmapData
    
    strBitmapHeight = Space$(4)
    CopyMemory ByVal strBitmapHeight, BitmapHeight, 4
    
    strBitmapWidth = Space$(4)
    CopyMemory ByVal strBitmapWidth, BitmapWidth, 4
    
    strBytesPerPixel = Space$(2)
    CopyMemory ByVal strBytesPerPixel, BytesPerPixel, 2
    
    GetBitmapData = strBitmapHeight & strBitmapWidth & strBytesPerPixel & BitmapData

End Function
Public Function GetAllFiles(ByVal Root As String, _
                            ByVal Such As String, _
                            Optional FindDirectory As Boolean, _
                            Optional ShowProgress As Boolean = False) _
                                    As String
    Dim File As String
    Dim hFile As Long
    Dim FD As WIN32_FIND_DATA
    
    Static frm As ProgBar

    If ShowProgress And LoadProgbar Then _
        Set frm = New ProgBar: _
        Load frm: _
        LoadProgbar = False: _
        Call frm.SetOption(0, 0, "Searching...", False)
                
    Root = GetDir(Root)
    
    If ShowProgress Then _
        frm.Message.Caption = Root: _
        frm.Message.Refresh
        
    hFile = FindFirstFile(Root & "*.*", FD)
    If hFile = 0 Then Exit Function
    
    Do
        File = Left$(FD.cFileName, InStr(FD.cFileName, Chr$(0)) - 1)
              
       If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            If (File <> ".") And (File <> "..") Then
                If FindDirectory And LCase(File) = LCase(Such) Then
                    GetAllFiles = Root & File
                    Unload frm
                    Set frm = Nothing
                    Exit Function
                Else
                    GetAllFiles = GetAllFiles(Root & File, Such, _
                                              FindDirectory, ShowProgress)
                End If
            End If
       Else
            If Not FindDirectory And LCase(File) = LCase(Such) Then _
                GetAllFiles = Root & File: _
                Unload frm: _
                Set frm = Nothing: _
                LoadProgbar = True: _
                Exit Function
        End If
    
    Loop While FindNextFile(hFile, FD) And GetAllFiles = ""
        
    Call FindClose(hFile)
        
End Function
Public Function toHex(ByVal Res As Byte, _
                      ByVal Str As String) As String
    Dim aa As String, bb As String, cc As String, ERg As String
    Dim X As Integer
    Dim aVal As Byte
  
    aa = Str
    
    X = Res - Len(aa) Mod Res
    If X <> Res Then aa = aa & Space(X)
    
    For X = 1 To Len(aa)
        bb = Mid$(aa, X, 1)
      
        If bb = Chr$(0) Then
            aVal = 0
        Else
            aVal = (Asc(bb))
        End If
      
        ERg = ERg & Hex(aVal)
      
        If X Mod Res <> 0 Then
            ERg = ERg & " "
        Else
            ERg = ERg & " "
            cc = ""
        End If
    Next X
    
    If Right$(ERg, 1) = " " Then ERg = Mid$(ERg, 1, Len(ERg) - 1)
    
    toHex = ERg

End Function
Public Function Pow(Number As Long, _
                    Power As Integer) As Long
    Dim X As Integer

    If Power > 1 Then
        For X = 2 To Power
            Number = Number * 2
        Next X

        Pow = Number
    Else
        If Power = 0 Then Pow = 1
        If Power = 1 Then Pow = Number
    End If

End Function
Public Function FindSysTray() As Long
    Dim h1 As Long, h2 As Long, h3 As Long
    Dim Wv As oS
        
    h1 = FindWindowA("Shell_TrayWnd", vbNullString)

    If h1 Then
        h2 = FindWindowEx(h1, 0, "TrayNotifyWnd", vbNullString)
    
        Wv = SYS.Get_WinVer
    
        If SYS.isWindowsNT Or Wv = osWindowsMillenium Then
            If Wv = osWindows2003 Or Wv = osWindows2000 Or Wv = osWindowsXP Then
                h2 = FindWindowEx(h2, 0, "SysPager", vbNullString)
        
                If h2 Then _
                    h3 = FindWindowEx(h2, 0, "ToolbarWindow32", vbNullString): _
                    FindSysTray = h3
            End If
        Else
            FindSysTray = h2
        End If
    End If

End Function
Public Sub DoTrans(ByVal hWnd As Long, ByVal Rate As Byte)
    Dim WinInfo As Long

    WinInfo = GetWindowLongA(hWnd, GWL_EXSTYLE)
    
    If Rate < 255 Then
        WinInfo = WinInfo Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, WinInfo
        SetLayeredWindowAttributes hWnd, 0, Rate, LWA_ALPHA
    Else
        WinInfo = WinInfo Xor WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, WinInfo
    End If

End Sub
Private Function CreatePoly(objBild As PictureBox) As Long
    Dim lngTransparenteFarbe As Long, lngStartLinie_X As Long
    Dim bolErsterBereich As Boolean, bolEingangsLinie As Boolean
    Dim hDC As Long, lngBildWeite As Long, lngBildHoehe As Long
    Dim lngX As Long, lngY As Long
    Dim lngGesamtBereich As Long, lngLinienBereich As Long

    
    hDC = objBild.hDC
    
    lngBildWeite = objBild.ScaleWidth
    lngBildHoehe = objBild.ScaleHeight
    
    bolErsterBereich = True
    bolEingangsLinie = False
    
    lngX = lngY = lngStartLinie_X = 0
    
    lngTransparenteFarbe = GetPixel(hDC, 0, 0)

    For lngY = 0 To lngBildHoehe - 1
        For lngX = 0 To lngBildWeite - 1
            If GetPixel(hDC, lngX, lngY) = lngTransparenteFarbe Or _
                                           lngX = lngBildWeite Then
                If bolEingangsLinie Then
                    bolEingangsLinie = False
                    lngLinienBereich = CreateRectRgn(lngStartLinie_X, lngY, lngX, lngY + 1)
    
                    If bolErsterBereich Then
                        lngGesamtBereich = lngLinienBereich
                        bolErsterBereich = False
                    Else
                        CombineRgn lngGesamtBereich, lngGesamtBereich, lngLinienBereich, RGN_OR
                        DeleteObject lngLinienBereich
                    End If
                End If
            Else
                If Not bolEingangsLinie Then
                    bolEingangsLinie = True
                    lngStartLinie_X = lngX
                End If
            End If
        Next lngX
    Next lngY
    
    CreatePoly = lngGesamtBereich

End Function
Public Function CFFP(Form As Form, _
                     Picture As Picture, File As String) _
                        As Boolean
    Dim FensterBereich As Long

    Load MyControls
    
    If File <> "" Then
        MyControls.PicTMP = LoadPicture(File)
    Else
        MyControls.PicTMP.Picture = Picture
    End If
    
    Form.Picture = MyControls.PicTMP.Picture

    With MyControls.PicTMP
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .AutoSize = True
        .BorderStyle = 0
        .Left = 0
        .Top = 0
        
        Form.Width = .Width
        Form.Height = .Height
  End With
  
    FensterBereich = CreatePoly(MyControls.PicTMP)
    
    If SetWindowRgn(Form.hWnd, FensterBereich, True) Then _
        CFFP = True

    Unload MyControls
    Form.Refresh

End Function
Public Function BubbleSort(Feld As Variant) As Boolean
    Dim LB As Long, UB As Long, Pos As Long, X As Long
    Dim temp As String
    
    On Local Error GoTo Quit
    
    LB = LBound(Feld)
    UB = UBound(Feld)
    
    While UB > LB
        Pos = LB
      
        For X = LB To UB - 1
            If Feld(X) > Feld(X + 1) Then _
                temp = Feld(X + 1): _
                Feld(X + 1) = Feld(X): _
                Feld(X) = temp: _
                Pos = X
        Next X
      
        UB = Pos
    Wend

    BubbleSort = True
    
Quit:
End Function
Public Function LBSortItem(ByRef hLBox As ListBox, _
                           ByRef hOrder As Boolean) As Boolean
    Dim rApi As Long
    Dim idx As Integer, maxElements As Integer
    Dim tmpLB() As String
    
    If hLBox.Sorted = True Then LBSortItem = False: _
                                Exit Function
    
    ReDim tmpLB(hLBox.ListCount - 1)
    maxElements = hLBox.ListCount
    
    For idx = 0 To hLBox.ListCount - 1
        tmpLB(idx) = hLBox.List(idx)
    Next idx
    
    Call SortArray2(tmpLB(), 0, maxElements - 1)
    
    hLBox.Clear
    
    For idx = 0 To maxElements - 1
        hLBox.AddItem tmpLB(idx), 0
    Next idx
    
End Function
Public Sub SortArray2(ByRef tmpArray() As String, ByVal idxLo As Long, ByVal idxHi As Long)
    Dim tmpString As String, tmpSwap As String
    Dim tmpLow As Long, tmpHi As Long
  
    tmpLow = idxLo
    tmpHi = idxHi
    tmpString = tmpArray((idxLo + idxHi) / 2)
  
    While (tmpLow <= tmpHi)
        While (tmpArray(tmpLow) < tmpString) And (tmpLow < idxHi)
            tmpLow = tmpLow + 1
        Wend
  
        While (tmpString < tmpArray(tmpHi)) And (tmpHi > idxLo)
            tmpHi = tmpHi - 1
        Wend
  
        If (tmpLow <= tmpHi) Then
            tmpSwap = tmpArray(tmpLow)
            tmpArray(tmpLow) = tmpArray(tmpHi)
            tmpArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
   
    If (idxLo < tmpHi) Then SortArray2 tmpArray(), idxLo, tmpHi
    If (tmpLow < idxHi) Then SortArray2 tmpArray(), tmpLow, idxHi

End Sub
' * SFV Start *
Public Function SFV(ByVal File As Variant) As String
    Dim I As Long, P As Long
    Dim lngChunkSize As Long, lngSize As Long
    Dim lngCRC32 As Long
    Dim bArrayFile() As Byte
    Dim F() As Variant
    Dim strFilePath As String, CalculateFile As String

    Const sASM As String = "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F464975F28B4D088901595B585E5F89EC5DC21000"
    
    On Local Error GoTo Quit
    
    If m_CRC32Table(1) = 0 Then
        m_CRC32Table(0) = &H0:          m_CRC32Table(1) = &H77073096
        m_CRC32Table(2) = &HEE0E612C:   m_CRC32Table(3) = &H990951BA
        m_CRC32Table(4) = &H76DC419:    m_CRC32Table(5) = &H706AF48F
        m_CRC32Table(6) = &HE963A535:   m_CRC32Table(7) = &H9E6495A3
        m_CRC32Table(8) = &HEDB8832:    m_CRC32Table(9) = &H79DCB8A4
        m_CRC32Table(10) = &HE0D5E91E:  m_CRC32Table(11) = &H97D2D988
        m_CRC32Table(12) = &H9B64C2B:   m_CRC32Table(13) = &H7EB17CBD
        m_CRC32Table(14) = &HE7B82D07:  m_CRC32Table(15) = &H90BF1D91
        m_CRC32Table(16) = &H1DB71064:  m_CRC32Table(17) = &H6AB020F2
        m_CRC32Table(18) = &HF3B97148:  m_CRC32Table(19) = &H84BE41DE
        m_CRC32Table(20) = &H1ADAD47D:  m_CRC32Table(21) = &H6DDDE4EB
        m_CRC32Table(22) = &HF4D4B551:  m_CRC32Table(23) = &H83D385C7
        m_CRC32Table(24) = &H136C9856:  m_CRC32Table(25) = &H646BA8C0
        m_CRC32Table(26) = &HFD62F97A:  m_CRC32Table(27) = &H8A65C9EC
        m_CRC32Table(28) = &H14015C4F:  m_CRC32Table(29) = &H63066CD9
        m_CRC32Table(30) = &HFA0F3D63:  m_CRC32Table(31) = &H8D080DF5
        m_CRC32Table(32) = &H3B6E20C8:  m_CRC32Table(33) = &H4C69105E
        m_CRC32Table(34) = &HD56041E4:  m_CRC32Table(35) = &HA2677172
        m_CRC32Table(36) = &H3C03E4D1:  m_CRC32Table(37) = &H4B04D447
        m_CRC32Table(38) = &HD20D85FD:  m_CRC32Table(39) = &HA50AB56B
        m_CRC32Table(40) = &H35B5A8FA:  m_CRC32Table(41) = &H42B2986C
        m_CRC32Table(42) = &HDBBBC9D6:  m_CRC32Table(43) = &HACBCF940
        m_CRC32Table(44) = &H32D86CE3:  m_CRC32Table(45) = &H45DF5C75
        m_CRC32Table(46) = &HDCD60DCF:  m_CRC32Table(47) = &HABD13D59
        m_CRC32Table(48) = &H26D930AC:  m_CRC32Table(49) = &H51DE003A
        m_CRC32Table(50) = &HC8D75180:  m_CRC32Table(51) = &HBFD06116
        m_CRC32Table(52) = &H21B4F4B5:  m_CRC32Table(53) = &H56B3C423
        m_CRC32Table(54) = &HCFBA9599:  m_CRC32Table(55) = &HB8BDA50F
        m_CRC32Table(56) = &H2802B89E:  m_CRC32Table(57) = &H5F058808
        m_CRC32Table(58) = &HC60CD9B2:  m_CRC32Table(59) = &HB10BE924
        m_CRC32Table(60) = &H2F6F7C87:  m_CRC32Table(61) = &H58684C11
        m_CRC32Table(62) = &HC1611DAB:  m_CRC32Table(63) = &HB6662D3D
        m_CRC32Table(64) = &H76DC4190:  m_CRC32Table(65) = &H1DB7106
        m_CRC32Table(66) = &H98D220BC:  m_CRC32Table(67) = &HEFD5102A
        m_CRC32Table(68) = &H71B18589:  m_CRC32Table(69) = &H6B6B51F
        m_CRC32Table(70) = &H9FBFE4A5:  m_CRC32Table(71) = &HE8B8D433
        m_CRC32Table(72) = &H7807C9A2:  m_CRC32Table(73) = &HF00F934
        m_CRC32Table(74) = &H9609A88E:  m_CRC32Table(75) = &HE10E9818
        m_CRC32Table(76) = &H7F6A0DBB:  m_CRC32Table(77) = &H86D3D2D
        m_CRC32Table(78) = &H91646C97:  m_CRC32Table(79) = &HE6635C01
        m_CRC32Table(80) = &H6B6B51F4:  m_CRC32Table(81) = &H1C6C6162
        m_CRC32Table(82) = &H856530D8:  m_CRC32Table(83) = &HF262004E
        m_CRC32Table(84) = &H6C0695ED:  m_CRC32Table(85) = &H1B01A57B
        m_CRC32Table(86) = &H8208F4C1:  m_CRC32Table(87) = &HF50FC457
        m_CRC32Table(88) = &H65B0D9C6:  m_CRC32Table(89) = &H12B7E950
        m_CRC32Table(90) = &H8BBEB8EA:  m_CRC32Table(91) = &HFCB9887C
        m_CRC32Table(92) = &H62DD1DDF:  m_CRC32Table(93) = &H15DA2D49
        m_CRC32Table(94) = &H8CD37CF3:  m_CRC32Table(95) = &HFBD44C65
        m_CRC32Table(96) = &H4DB26158:  m_CRC32Table(97) = &H3AB551CE
        m_CRC32Table(98) = &HA3BC0074:  m_CRC32Table(99) = &HD4BB30E2
        m_CRC32Table(100) = &H4ADFA541: m_CRC32Table(101) = &H3DD895D7
        m_CRC32Table(102) = &HA4D1C46D: m_CRC32Table(103) = &HD3D6F4FB
        m_CRC32Table(104) = &H4369E96A: m_CRC32Table(105) = &H346ED9FC
        m_CRC32Table(106) = &HAD678846: m_CRC32Table(107) = &HDA60B8D0
        m_CRC32Table(108) = &H44042D73: m_CRC32Table(109) = &H33031DE5
        m_CRC32Table(110) = &HAA0A4C5F: m_CRC32Table(111) = &HDD0D7CC9
        m_CRC32Table(112) = &H5005713C: m_CRC32Table(113) = &H270241AA
        m_CRC32Table(114) = &HBE0B1010: m_CRC32Table(115) = &HC90C2086
        m_CRC32Table(116) = &H5768B525: m_CRC32Table(117) = &H206F85B3
        m_CRC32Table(118) = &HB966D409: m_CRC32Table(119) = &HCE61E49F
        m_CRC32Table(120) = &H5EDEF90E: m_CRC32Table(121) = &H29D9C998
        m_CRC32Table(122) = &HB0D09822: m_CRC32Table(123) = &HC7D7A8B4
        m_CRC32Table(124) = &H59B33D17: m_CRC32Table(125) = &H2EB40D81
        m_CRC32Table(126) = &HB7BD5C3B: m_CRC32Table(127) = &HC0BA6CAD
        m_CRC32Table(128) = &HEDB88320: m_CRC32Table(129) = &H9ABFB3B6
        m_CRC32Table(130) = &H3B6E20C:  m_CRC32Table(131) = &H74B1D29A
        m_CRC32Table(132) = &HEAD54739: m_CRC32Table(133) = &H9DD277AF
        m_CRC32Table(134) = &H4DB2615:  m_CRC32Table(135) = &H73DC1683
        m_CRC32Table(136) = &HE3630B12: m_CRC32Table(137) = &H94643B84
        m_CRC32Table(138) = &HD6D6A3E:  m_CRC32Table(139) = &H7A6A5AA8
        m_CRC32Table(140) = &HE40ECF0B: m_CRC32Table(141) = &H9309FF9D
        m_CRC32Table(142) = &HA00AE27:  m_CRC32Table(143) = &H7D079EB1
        m_CRC32Table(144) = &HF00F9344: m_CRC32Table(145) = &H8708A3D2
        m_CRC32Table(146) = &H1E01F268: m_CRC32Table(147) = &H6906C2FE
        m_CRC32Table(148) = &HF762575D: m_CRC32Table(149) = &H806567CB
        m_CRC32Table(150) = &H196C3671: m_CRC32Table(151) = &H6E6B06E7
        m_CRC32Table(152) = &HFED41B76: m_CRC32Table(153) = &H89D32BE0
        m_CRC32Table(154) = &H10DA7A5A: m_CRC32Table(155) = &H67DD4ACC
        m_CRC32Table(156) = &HF9B9DF6F: m_CRC32Table(157) = &H8EBEEFF9
        m_CRC32Table(158) = &H17B7BE43: m_CRC32Table(159) = &H60B08ED5
        m_CRC32Table(160) = &HD6D6A3E8: m_CRC32Table(161) = &HA1D1937E
        m_CRC32Table(162) = &H38D8C2C4: m_CRC32Table(163) = &H4FDFF252
        m_CRC32Table(164) = &HD1BB67F1: m_CRC32Table(165) = &HA6BC5767
        m_CRC32Table(166) = &H3FB506DD: m_CRC32Table(167) = &H48B2364B
        m_CRC32Table(168) = &HD80D2BDA: m_CRC32Table(169) = &HAF0A1B4C
        m_CRC32Table(170) = &H36034AF6: m_CRC32Table(171) = &H41047A60
        m_CRC32Table(172) = &HDF60EFC3: m_CRC32Table(173) = &HA867DF55
        m_CRC32Table(174) = &H316E8EEF: m_CRC32Table(175) = &H4669BE79
        m_CRC32Table(176) = &HCB61B38C: m_CRC32Table(177) = &HBC66831A
        m_CRC32Table(178) = &H256FD2A0: m_CRC32Table(179) = &H5268E236
        m_CRC32Table(180) = &HCC0C7795: m_CRC32Table(181) = &HBB0B4703
        m_CRC32Table(182) = &H220216B9: m_CRC32Table(183) = &H5505262F
        m_CRC32Table(184) = &HC5BA3BBE: m_CRC32Table(185) = &HB2BD0B28
        m_CRC32Table(186) = &H2BB45A92: m_CRC32Table(187) = &H5CB36A04
        m_CRC32Table(188) = &HC2D7FFA7: m_CRC32Table(189) = &HB5D0CF31
        m_CRC32Table(190) = &H2CD99E8B: m_CRC32Table(191) = &H5BDEAE1D
        m_CRC32Table(192) = &H9B64C2B0: m_CRC32Table(193) = &HEC63F226
        m_CRC32Table(194) = &H756AA39C: m_CRC32Table(195) = &H26D930A
        m_CRC32Table(196) = &H9C0906A9: m_CRC32Table(197) = &HEB0E363F
        m_CRC32Table(198) = &H72076785: m_CRC32Table(199) = &H5005713
        m_CRC32Table(200) = &H95BF4A82: m_CRC32Table(201) = &HE2B87A14
        m_CRC32Table(202) = &H7BB12BAE: m_CRC32Table(203) = &HCB61B38
        m_CRC32Table(204) = &H92D28E9B: m_CRC32Table(205) = &HE5D5BE0D
        m_CRC32Table(206) = &H7CDCEFB7: m_CRC32Table(207) = &HBDBDF21
        m_CRC32Table(208) = &H86D3D2D4: m_CRC32Table(209) = &HF1D4E242
        m_CRC32Table(210) = &H68DDB3F8: m_CRC32Table(211) = &H1FDA836E
        m_CRC32Table(212) = &H81BE16CD: m_CRC32Table(213) = &HF6B9265B
        m_CRC32Table(214) = &H6FB077E1: m_CRC32Table(215) = &H18B74777
        m_CRC32Table(216) = &H88085AE6: m_CRC32Table(217) = &HFF0F6A70
        m_CRC32Table(218) = &H66063BCA: m_CRC32Table(219) = &H11010B5C
        m_CRC32Table(220) = &H8F659EFF: m_CRC32Table(221) = &HF862AE69
        m_CRC32Table(222) = &H616BFFD3: m_CRC32Table(223) = &H166CCF45
        m_CRC32Table(224) = &HA00AE278: m_CRC32Table(225) = &HD70DD2EE
        m_CRC32Table(226) = &H4E048354: m_CRC32Table(227) = &H3903B3C2
        m_CRC32Table(228) = &HA7672661: m_CRC32Table(229) = &HD06016F7
        m_CRC32Table(230) = &H4969474D: m_CRC32Table(231) = &H3E6E77DB
        m_CRC32Table(232) = &HAED16A4A: m_CRC32Table(233) = &HD9D65ADC
        m_CRC32Table(234) = &H40DF0B66: m_CRC32Table(235) = &H37D83BF0
        m_CRC32Table(236) = &HA9BCAE53: m_CRC32Table(237) = &HDEBB9EC5
        m_CRC32Table(238) = &H47B2CF7F: m_CRC32Table(239) = &H30B5FFE9
        m_CRC32Table(240) = &HBDBDF21C: m_CRC32Table(241) = &HCABAC28A
        m_CRC32Table(242) = &H53B39330: m_CRC32Table(243) = &H24B4A3A6
        m_CRC32Table(244) = &HBAD03605: m_CRC32Table(245) = &HCDD70693
        m_CRC32Table(246) = &H54DE5729: m_CRC32Table(247) = &H23D967BF
        m_CRC32Table(248) = &HB3667A2E: m_CRC32Table(249) = &HC4614AB8
        m_CRC32Table(250) = &H5D681B02: m_CRC32Table(251) = &H2A6F2B94
        m_CRC32Table(252) = &HB40BBE37: m_CRC32Table(253) = &HC30C8EA1
        m_CRC32Table(254) = &H5A05DF1B: m_CRC32Table(255) = &H2D02EF8D
    End If
    
    ReDim m_CRC32Asm(0 To Len(sASM) \ 2 - 1)
        
    For I = 1 To Len(sASM) Step 2
        m_CRC32Asm(I \ 2) = Val("&H" & Mid$(sASM, I, 2))
    Next I
        
    m_CRC32 = &HFFFFFFFF

    If Not isArray(File) Then
        ReDim F(0)
        F(0) = File
    Else
        F = File
    End If
    
    For P = 0 To UBound(F)
        strFilePath = F(P)
        lngSize = FileLen(strFilePath)
        lngChunkSize = 4096

        If lngSize <> 0 Then
            Open strFilePath For Binary Access Read As #1
                Do While Seek(1) < lngSize
                    If (lngSize - Seek(1)) > lngChunkSize Then
                        Do While Seek(1) < (lngSize - lngChunkSize)
                            ReDim bArrayFile(lngChunkSize - 1)
                            
                            Get #1, , bArrayFile()
                            
                            lngCRC32 = CalculateBytes(bArrayFile)
                            
                            DoEvents
                        Loop
                    Else
                        ReDim bArrayFile(lngSize - Seek(1))
                        
                        Get #1, , bArrayFile()
            
                        lngCRC32 = CalculateBytes(bArrayFile)
                    End If
                    
                    DoEvents
                Loop
            Close #1
    
            CalculateFile = Right$("00000000" & Hex$(lngCRC32), 8)
        Else
            CalculateFile = "00000000"
        End If
    Next P
    
    SFV = CalculateFile
    
Quit:
    If Err.Number <> 0 Then SFV = "ERROR " & Err.Number & " (" & Err.Description & ")"
    
End Function
Private Function CalculateBytes(ByteArray() As Byte) As Variant
    CalculateBytes = AddBytes(ByteArray)
End Function
Private Function AddBytes(ByteArray() As Byte) As Variant
    Dim ByteSize As Long
  
    On Local Error GoTo NoData
  
    ByteSize = UBound(ByteArray) - LBound(ByteArray) + 1
  
    On Local Error GoTo 0
  
    Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRC32Table(0)), ByteSize)
  
NoData:
    AddBytes = (Not m_CRC32)
  
' * SFV End *
End Function
Public Function sEnCrypt(ByVal Text, Key As String) As String
    Dim Z As String
    Dim I As Long, Position As Long
    Dim cptZahl As Long, orgZahl As Long
    Dim keyZahl As Long, cptString As String
    
    For I = 1 To Len(Text)
        Position = Position + 1
        If Position > Len(Key) Then Position = 1
        
        keyZahl = Asc(Mid$(Key, Position, 1))
        
        orgZahl = Asc(Mid$(Text, I, 1))
        cptZahl = orgZahl Xor keyZahl
        cptString = Hex(cptZahl)
            
        If Len(cptString) < 2 Then cptString = "0" & cptString
        
        Z = Z & cptString
    Next I
        
    sEnCrypt = Z

End Function
Public Function sDeCrypt(ByVal Text, Key As String) As String
    Dim Z As String
    Dim I As Integer, Position As Integer
    Dim cptZahl As Long, orgZahl As Long
    Dim keyZahl As Long, cptString As String
    
    For I = 1 To Len(Text)
        Position = Position + 1
        If Position > Len(Key) Then Position = 1
        
        keyZahl = Asc(Mid$(Key, Position, 1))
        
        If I > Len(Text) \ 2 Then Exit For
            
        cptZahl = CByte("&H" & Mid$(Text, I * 2 - 1, 2))
        orgZahl = cptZahl Xor keyZahl
            
        Z = Z & Chr$(orgZahl)
    Next I
     
    sDeCrypt = Z
    
End Function
Public Sub UpdateWin(hWnd As Long)
    Dim R As RECT
    
    Call GetWindowRect(hWnd, R)
    Call SetWindowPos(hWnd, 0, R.Left, R.Top, _
                      R.Right - R.Left, R.Bottom - R.Top, _
                      SWP_FRAMECHANGED)
End Sub
Public Function GetAllPath(ByVal Root As String, _
                           ByVal InclSub As Boolean, _
                           FileCount As Long, DirCount As Long) _
                                As Variant
    Dim File As String
    Dim hFile As Long
    Dim FD As WIN32_FIND_DATA
    
    Static Count As Long, FC As Long, DC As Long
    Static S As Variant
    
    On Local Error Resume Next
    
    If Count = 0 Then S = 0: _
                      FC = 0: _
                      DC = 0
    
    Count = Count + 1
    
    Root = GetDir(Root)
    
    hFile = FindFirstFile(Root & "*.*", FD)
    If hFile = 0 Then Exit Function
    
    Do
        File = Left$(FD.cFileName, InStr(FD.cFileName, Chr$(0)) - 1)
        
        If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            If File <> "." And File <> ".." Then _
                DC = DC + 1: _
                If InclSub Then Call GetAllPath(Root & File, True, FileCount, DirCount)
        Else
            FC = FC + 1
            S = S + CLng(FileLen(Root & File))
        End If
    
    Loop While FindNextFile(hFile, FD)
        
    Call FindClose(hFile)
        
    Count = Count - 1
    
    If Count = 0 Then _
        GetAllPath = S: _
        FileCount = FC: _
        DirCount = DC
    
End Function

