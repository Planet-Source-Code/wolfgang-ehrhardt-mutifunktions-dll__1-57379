VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Grafic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByRef lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public Enum IconSizeNfo
    is16x16 = 300
    is24x24 = 420
    is32x32 = 540
    is48x48 = 780
End Enum

Public Enum vImageListNfo
    vi16x16 = 300
    vi24x24 = 420
    vi32x32 = 540
    vi48x48 = 780
    viBYPICTURE = 1
End Enum

Public Enum vImageListCommand
    viADD = 1
    viGET = 2
    viDELETE = 3
    viCLEAR = 4
    viITEMCOUNT = 5
End Enum

Private Const DSTINVERT = &H550009

Private Const HIMETRIC_PER_INCH = 2540

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Dim abp As asmBmpPara
Dim pmd As Boolean
Dim pic() As StdPicture
Public Sub Convert_ColorToRGB(ByVal Color As Long, _
                              R As Long, _
                              G As Long, _
                              B As Long)
    R = Color And 255
    G = (Color And 65280) / 256
    B = (Color And 16711680) / 65536
End Sub
Public Function Convert_ColorToHTML(ByVal Color As Long) _
                                        As String
    Dim Shex As String
    
    On Local Error GoTo Quit
    
    Shex = Hex(Color)
    Shex = String(6 - Len(Shex), "0") & Shex
             
    Convert_ColorToHTML = "#" & Right$(Shex, 2) & _
                          Mid$(Shex, 3, 2) & Left$(Shex, 2)

Quit:
    If Err.Number <> 0 Then Convert_ColorToHTML = ""

End Function
Public Function Get_FileIconCount(ByVal File As String) As Long
    Get_FileIconCount = ExtractIcon(1, File, -1)
End Function
Public Function Get_IconFromFile(ByVal File As String, _
                                 ByVal IconNumber As Long) _
                                        As StdPicture
    Dim hIcon As Long
    Dim oAR As Boolean
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    Load frm
    frm.Tag = "Get_IconFromFile"
    
    With frm.PicTMP
        oAR = .AutoRedraw
        .AutoRedraw = True
    
        hIcon = ExtractIcon(0, File, IconNumber)
    
        If hIcon Then _
            If DrawIcon(.hDC, 0, 0, hIcon) Then _
                Set Get_IconFromFile = .Image
        
        .AutoRedraw = oAR
        .Refresh
    End With
    
Quit:
    Set frm = Nothing
    
End Function
Public Function Convert_HTMLtoLong(ByVal HTML As String) _
                                        As Long
    Dim R As Long, G As Long, B As Long, P As Long
    Dim T As String
    
    T = Replace(HTML, "#", "")

    If Not Str.isValidLong(MA.HexToDec(T)) Then _
        Convert_HTMLtoLong = -1: _
        Exit Function
    
    R = Val("&H" & Mid$(T, 1, 2))
    G = Val("&H" & Mid$(T, 3, 2))
    B = Val("&H" & Mid$(T, 5, 2))
    
    Convert_HTMLtoLong = RGB(R, G, B)
    
End Function
Public Function PictureToString(Picture As StdPicture) _
                                    As String
    Dim frm As New MyControls
    
    Load frm
    frm.Tag = "PictureToString"
    
    Set frm.PicTMP.Picture = Picture
    PictureToString = GetBitmapData(frm.PicTMP)
    
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureCompare(Picture1 As StdPicture, _
                               Picture2 As StdPicture) _
                                    As Boolean
    Dim PropBag1 As PropertyBag, PropBag2 As PropertyBag
    Dim strPic1 As String, strPic2 As String
    
    Set PropBag1 = New PropertyBag
    Set PropBag2 = New PropertyBag

    PropBag1.WriteProperty "Pic", Picture1
    PropBag2.WriteProperty "Pic", Picture2
    
    strPic1 = PropBag1.Contents
    strPic2 = PropBag2.Contents
    
    PictureCompare = (strPic1 = strPic2)
    
End Function
Public Function Convert_OLEtoRBG(ByVal OLEcolor As Long) _
                                        As Long
    Call OleTranslateColor(OLEcolor, 0, Convert_OLEtoRBG)
End Function
Public Function Convert_RGBtoLong(ByVal R As Long, _
                                  ByVal G As Long, _
                                  ByVal B As Long) As Long
    Convert_RGBtoLong = RGB(R, G, B)
End Function
Public Function Get_RandomizedColor() As Long
    Dim R As Integer, G As Integer, B As Integer
        
    R = MISC.Get_Randomize(0, 255)
    G = MISC.Get_Randomize(0, 255)
    B = MISC.Get_Randomize(0, 255)
    
    Get_RandomizedColor = RGB(R, G, B)

End Function
Public Function PictureBox_Tile(PictureBox As Object) _
                                    As Boolean
    
    PictureBox_Tile = FRMS.Tile(PictureBox, PictureBox)
    
End Function
Public Sub TwipsToPixel(X As Long, Y As Long)
    X = X / Screen.TwipsPerPixelX
    Y = Y / Screen.TwipsPerPixelY
End Sub
Public Sub PixelToTwips(X As Long, Y As Long)
    X = X * Screen.TwipsPerPixelX
    Y = Y * Screen.TwipsPerPixelY
End Sub
Public Function StretchToPicBox(PictureBox As Object) As Boolean
    
    On Local Error GoTo Quit
    
    With PictureBox
        .AutoRedraw = True
        .Cls
        .PaintPicture .Picture, 0, 0, .ScaleWidth, .ScaleHeight
        .AutoRedraw = False
    End With
    
    StretchToPicBox = True
    
Quit:
End Function
Public Function Stretch(Picture As StdPicture, _
                        ByVal newWidth As Long, _
                        ByVal newHeight As Long, _
                        Dimension As ScaleModeConstants) _
                            As StdPicture
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    Load frm
    frm.Tag = "Stretch"
    
    With frm
        .PicTMP.Picture = Picture
        .PicTMP.ScaleMode = Dimension
        
        .PicTMP.Width = newWidth
        .PicTMP.Height = newHeight
    
        If StretchToPicBox(.PicTMP) Then _
            Set Stretch = .PicTMP.Image
    End With
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureToIcon(Picture As StdPicture, _
                              IconSize As IconSizeNfo) _
                                    As StdPicture
    Dim frm As Form
    
    On Local Error GoTo Quit
    
    Set frm = New MyControls
    
    Load frm
    frm.Tag = "PictureToIcon"
    
    With frm
        .PicTMP.ScaleMode = vbPixels
        .PicTMP.AutoSize = False
    
        .PicTMP.Picture = Picture
    
        .PicTMP.Height = IconSize
        .PicTMP.Width = IconSize
    
        If StretchToPicBox(.PicTMP) Then _
            Set PictureToIcon = .PicTMP.Image
    End With
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function vListImage(Command As vImageListCommand, _
                           Index As Long, _
                           Optional Picture As StdPicture, _
                           Optional Dimension As vImageListNfo) _
                                As StdPicture
    Dim P As Long, X As Long
    Dim frm As Form
            
    On Local Error GoTo Quit

    X = -1
    
    Do While X = -1
        X = FRMS.Find_Form(Forms, "MyControls", , "vListImage", True)
        
        If X = -1 Then Set frm = New MyControls: _
                       Load frm: _
                       frm.Tag = "vListImage"
    Loop
            
    Set frm = Forms(X)
    
    Select Case Command
        Case viADD
            Index = frm.vPic.UBound + 1
            Load frm.vPic(Index)
            
            Select Case Dimension
                Case vi16x16
                    Set Picture = PictureToIcon(Picture, is16x16)
                Case vi24x24
                    Set Picture = PictureToIcon(Picture, is24x24)
                Case vi32x32
                    Set Picture = PictureToIcon(Picture, is32x32)
                Case vi48x48
                    Set Picture = PictureToIcon(Picture, is48x48)
                Case viBYPICTURE
                    '
            End Select
            
            Set frm.vPic(Index).Picture = Picture
            
            Index = Index - 1
        Case viDELETE
            X = Index + 1
            
            For P = (X + 1) To frm.vPic.UBound
                Set frm.vPic(P - 1).Picture = frm.vPic(P).Picture
            Next P
            
            Unload frm.vPic(frm.vPic.UBound)
        Case viGET
            Set vListImage = frm.vPic(Index + 1)
        Case viCLEAR
            Unload frm
            Set frm = Nothing
            Exit Function
        Case viITEMCOUNT
            Index = frm.vPic.UBound
    End Select
    
    If frm.vPic.UBound = 0 Then Unload frm
    
Quit:
    Set frm = Nothing
    If Err Then Index = -1
    
End Function
Public Function PictureColor(Picture As StdPicture, _
                             ByVal Color As Long) _
                                    As StdPicture
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    Load frm
    frm.Tag = "PictureColor"
    
    With frm.PicTMP
        .Left = 0
        .Top = 0
        
        .AutoSize = True
        .AutoRedraw = True
        .Picture = Picture
        .ZOrder
        
        .Picture = LoadPicture("")
        .BackColor = Color
        
        .PaintPicture Picture, 0, 0, , , , , , , vbSrcPaint
        .Refresh
    
        Set PictureColor = .Image
    End With
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureInvert(Picture As StdPicture) _
                                As StdPicture
    Dim frm As New MyControls
    
    Load frm
    frm.Tag = "PictureInvert"
    
    With frm.PicTMP
        .AutoRedraw = True
        .AutoSize = True
        
        Set frm.PicTMP = Picture
    
        If BitBlt(.hDC, 0, 0, .ScaleWidth, .ScaleHeight, _
                  .hDC, 0, 0, DSTINVERT) Then _
                       Set PictureInvert = .Image
    End With
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureRotate(Picture As StdPicture, _
                              ByVal Degrese As Long) _
                                    As StdPicture
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    If Degrese < -360 Or Degrese > 360 Then Exit Function
        
    Load frm
    frm.Tag = "PictureRotate"
    
    Set frm.PicTMP.Picture = Picture
    Set PictureRotate = frm.BitmapWork(frm.PicTMP.Picture, Degrese)
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureFlip(Picture As StdPicture, _
                            ByVal Degrese As Long) _
                                    As StdPicture
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    If Degrese < -360 Or Degrese > 360 Then Exit Function
        
    Load frm
    frm.Tag = "PictureFlip"
    
    Set frm.PicTMP.Picture = Picture
    Set PictureFlip = frm.BitmapWork(frm.PicTMP.Picture, , , Degrese)
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureTurn(Picture As StdPicture, _
                            ByVal Degrese As Long) _
                                    As StdPicture
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    If Degrese < -360 Or Degrese > 360 Then Exit Function
        
    Load frm
    frm.Tag = "PictureTurn"
    
    Set frm.PicTMP.Picture = Picture
    Set PictureTurn = frm.BitmapWork(frm.PicTMP.Picture, , Degrese)
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureZoom(Picture As StdPicture, _
                            ByVal ZoomFactor As Long) _
                                    As StdPicture
    Dim frm As New MyControls
    
    On Local Error GoTo Quit
    
    ZoomFactor = IIf(ZoomFactor > 0, 45 + ZoomFactor, _
                                     45 - ZoomFactor)
    
    If ZoomFactor < 0 Or ZoomFactor > 86 Then Exit Function
        
    Load frm
    frm.Tag = "PictureZoom"
    
    Set frm.PicTMP.Picture = Picture
    Set PictureZoom = frm.BitmapWork(frm.PicTMP, , , , ZoomFactor)
    
Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureBrightness(Picture As StdPicture, _
                                  ByVal Brightness As Long) _
                                        As StdPicture
    Dim lpAsmBmpPara As Long
    Dim bar(255) As Byte, src() As Byte, tgt() As Byte
    
    Dim frm As New MyControls
        
    Static ASM(23) As Long
        
    On Local Error GoTo Quit
    
    If Brightness < -255 Or Brightness > 255 Then Exit Function
    
    Load frm
    frm.Tag = "PictureBrightness"
    
    Call ClonePic(frm.PicTMP, Picture)
    
    If Not ReadDataBrightness(frm.PicTMP.Image, abp, src, tgt, pmd, False) Then _
        Exit Function
    
    WriteDataBrightness frm.PicTMP, abp.src, True
    
    lpAsmBmpPara = VarPtr(abp)
        
    If ASM(0) = 0 Then
        ASM(0) = &HEC8B5590:  ASM(1) = &H8B575653
        ASM(2) = &H758B107D:  ASM(3) = &HBC68B0C
        ASM(4) = &HF70779F6:  ASM(5) = &H1B8DE
        ASM(6) = &HB90000:    ASM(7) = &H2B000001
        ASM(8) = &HF73473F1:  ASM(9) = &H78899D6
        ASM(10) = &H2D60347:  ASM(11) = &H49F48AC6
        ASM(12) = &H5D8BF475: ASM(13) = &H14738B08
        ASM(14) = &H8B2C7B8B: ASM(15) = &HAF0F1C4B
        ASM(16) = &H5D8B204B: ASM(17) = &HC1D7AD10
        ASM(18) = &HC1D708C8: ASM(19) = &HC1D708C8
        ASM(20) = &H49AB10C8: ASM(21) = &H1BF9EF75
        ASM(22) = &H5B5E5FC0: ASM(23) = &H10C25D
    End If
     
    Call CallWindowProc(ASM(0), lpAsmBmpPara, Brightness, VarPtr(bar(0)), 0)
    
    Call WriteDataBrightness(frm.PicTMP, abp.tgt)
    
    Set PictureBrightness = frm.PicTMP.Picture

Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureEmboss(Picture As StdPicture, _
                              ByVal Emboss As Long) _
                                    As StdPicture
    Dim lpAsmBmpPara As Long
    Dim bar(255) As Byte, src() As Byte, tgt() As Byte
    
    Dim frm As New MyControls
        
    Static ASM(40) As Long
        
    On Local Error GoTo Quit
    
    If Emboss < 0 Or Emboss > 7 Then Exit Function
    
    Load frm
    frm.Tag = "PictureEmboss"
    
    Call ClonePic(frm.PicTMP, Picture)
    
    If Not ReadDataBrightness(frm.PicTMP.Image, abp, src, tgt, pmd, False) Then _
        Exit Function
    
    WriteDataBrightness frm.PicTMP, abp.src, True
    
    lpAsmBmpPara = VarPtr(abp)
    
    If ASM(0) = 0 Then
        ASM(0) = &HEC8B5590:  ASM(1) = &H8B575653
        ASM(2) = &H738B085D:  ASM(3) = &H2C7B8B14
        ASM(4) = &HF1C4B8B:   ASM(5) = &HF3204BAF
        ASM(6) = &HC4D8BA5:   ASM(7) = &H3307E180
        ASM(8) = &HE88B99C0:  ASM(9) = &HA8E0D240
        ASM(10) = &HB20874DD: ASM(11) = &H74C1A804
        ASM(12) = &HA8DAF702: ASM(13) = &H8B097477
        ASM(14) = &H7A80C6B:  ASM(15) = &HDDF70274
        ASM(16) = &H738BEA03: ASM(17) = &H49CE8B04
        ASM(18) = &H8B4D7E49: ASM(19) = &H4A4A0853
        ASM(20) = &HC146467E: ASM(21) = &HFE8B02E6
        ASM(22) = &H3147303:  ASM(23) = &H438B2C7B
        ASM(24) = &HF7DD8B0C: ASM(25) = &H565150DB
        ASM(26) = &H44B60F57: ASM(27) = &H42A0035
        ASM(28) = &HFFD48033: ASM(29) = &H7F04C422
        ASM(30) = &HC40AE41A: ASM(31) = &H46470788
        ASM(32) = &H3C6F7:    ASM(33) = &HE17B0000
        ASM(34) = &H75494746: ASM(35) = &H595E5FDC
        ASM(36) = &H3F80358:  ASM(37) = &HCD754AF0
        ASM(38) = &H5FC01BF9: ASM(39) = &HC25D5B5E
        ASM(40) = &H10
    End If
    
    Call CallWindowProc(ASM(0), lpAsmBmpPara, Emboss, 0, 0)

    Call WriteDataBrightness(frm.PicTMP, abp.tgt)
    
    Set PictureEmboss = frm.PicTMP.Picture

Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureToGray(Picture As StdPicture) As StdPicture
    Dim lpAsmBmpPara As Long
    Dim bar(255) As Byte, src() As Byte, tgt() As Byte
    Dim frm As New MyControls
        
    Static ASM(19) As Long
        
    On Local Error GoTo Quit
    
    Load frm
    
    Call ClonePic(frm.PicTMP, Picture)
    frm.PicTMP.AutoRedraw = False
    
    If Not ReadDataGray(frm.PicTMP.Image, abp, src, tgt, pmd, False) Then _
        Exit Function
        
    WriteDataGray frm.PicTMP, abp.src, True
    
    lpAsmBmpPara = VarPtr(abp)

    If ASM(0) = 0 Then
        ASM(0) = &HEC8B5590:  ASM(1) = &H68575653
        ASM(2) = &H3E8:       ASM(3) = &H8B085D8B
        ASM(4) = &H7B8B1473:  ASM(5) = &H1C4B8B2C
        ASM(6) = &H204BAF0F:  ASM(7) = &H8BDB33AD
        ASM(8) = &H8AD48AD3:  ASM(9) = &H55AF0FD8
        ASM(10) = &H5DAF0F10: ASM(11) = &HC1DA030C
        ASM(12) = &H65F710E8: ASM(13) = &HF7C30314
        ASM(14) = &HE08AF075: ASM(15) = &H8A08E0C1
        ASM(16) = &HD8E2ABC4: ASM(17) = &HC01B58F9
        ASM(18) = &H5D5B5E5F: ASM(19) = &H10C2
    End If
    
    Call CallWindowProc(ASM(0), lpAsmBmpPara, 71, 707, 222)

    Call WriteDataGray(frm.PicTMP, abp.tgt)
    
    Set PictureToGray = frm.PicTMP.Picture

Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function PictureContrast(Picture As StdPicture, _
                                ByVal Contrast As Long) _
                                    As StdPicture
    Dim lpAsmBmpPara As Long
    Dim bar(255) As Byte, src() As Byte, tgt() As Byte
    Dim frm As New MyControls

    Static ASM(51) As Long

    On Local Error GoTo Quit
    
    If Contrast < 0 Or Contrast > 100 Then Exit Function
    
    Load frm
    frm.Tag = "PictureContrast"
    
    Call ClonePic(frm.PicTMP, Picture)
    
    If Not ReadDataBrightness(frm.PicTMP.Image, abp, src, tgt, pmd, False) Then _
        Exit Function
    
    WriteDataBrightness frm.PicTMP, abp.src, True
    
    lpAsmBmpPara = VarPtr(abp)
    
    If ASM(0) = 0 Then
        ASM(0) = &HEC8B5590:  ASM(1) = &H6A575653
        ASM(2) = &H68146A1E:  ASM(3) = &H3D4CCCCD
        ASM(4) = &H8068&:     ASM(5) = &H10EC8300
        ASM(6) = &H7E8DF48B:  ASM(7) = &HC45DB08
        ASM(8) = &HDDF075DA:  ASM(9) = &H67E81E
        ASM(10) = &H7DD0000:  ASM(11) = &HD8EC75DA
        ASM(12) = &HC0D9E865: ASM(13) = &H4DDAE0D9
        ASM(14) = &H107D8BE4: ASM(15) = &HFFB9&
        ASM(16) = &HDBC1D800: ASM(17) = &H8B00B016
        ASM(18) = &HE45D031E: ASM(19) = &H8A4B0A7E
        ASM(20) = &H8EBC1C3:  ASM(21) = &HFFB00274
        ASM(22) = &H49470788: ASM(23) = &HD9DDE379
        ASM(24) = &H5D8B1EDD: ASM(25) = &H14738B08
        ASM(26) = &H8B2C7B8B: ASM(27) = &HAF0F1C4B
        ASM(28) = &H5D8B204B: ASM(29) = &HAAD7AC10
        ASM(30) = &H3C6F7:    ASM(31) = &HF57B0000
        ASM(32) = &HF17549A4: ASM(33) = &HF4658DF9
        ASM(34) = &H5E5FC01B: ASM(35) = &H10C25D5B
        ASM(36) = &HEC8B5500: ASM(37) = &H9B08EC83
        ASM(38) = &H8BFC7DD9: ASM(39) = &HCC80FC45
        ASM(40) = &H4589660F: ASM(41) = &HFE6DD9FE
        ASM(42) = &HDCEAD99B: ASM(43) = &HF855DB0E
        ASM(44) = &HD9F865DA: ASM(45) = &HF845DBF0
        ASM(46) = &HC2D8E8D9: ASM(47) = &HD9DDFDD9
        ASM(48) = &H1FDDD9DD: ASM(49) = &H9BFC6DD9
        ASM(50) = &HC908C483: ASM(51) = &HC3
    End If

    Call CallWindowProc(ASM(0), lpAsmBmpPara, Contrast, VarPtr(bar(0)), 0)

    Call WriteDataBrightness(frm.PicTMP, abp.tgt)
    
    Set PictureContrast = frm.PicTMP.Picture

Quit:
    Unload frm
    Set frm = Nothing
    
End Function
Public Function Picture_GetSize(ByVal File As String, _
                                Width As Long, _
                                Height As Long) As Long
    Dim picTemp As New StdPicture
    Dim hDC As Long

    On Local Error Resume Next
    
    Width = 0
    Height = 0
    
    If Not F.Exist(File) Then Picture_GetSize = -1: _
                              Exit Function
    
    Set picTemp = LoadPicture(File)
    
    If Err.Number <> 0 Then Picture_GetSize = -2: _
                            Exit Function
    
    hDC = GetDC(0)
    Width = MulDiv(GetDeviceCaps(hDC, LOGPIXELSX), picTemp.Width, HIMETRIC_PER_INCH)
    Height = MulDiv(GetDeviceCaps(hDC, LOGPIXELSY), picTemp.Height, HIMETRIC_PER_INCH)
        
    Picture_GetSize = 1
    
    Call ReleaseDC(0, hDC)
        
    Set picTemp = Nothing

End Function
