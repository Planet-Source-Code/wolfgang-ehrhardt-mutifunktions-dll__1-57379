Attribute VB_Name = "grafxBas"
Option Explicit
Option Base 0

Public Type PictDesc
    cbSizeofStruct As Long
    picType        As Long
    hImage         As Long
    xExt           As Long
    yExt           As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBM As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByRef lpvBits As Any, ByRef lpBmi As BITMAPINFO, ByVal uUsage As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, ByRef pbmi As BITMAPINFO, ByVal iUsage As Long, ByRef ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByVal pCLSID As Long) As Long
Private Declare Function CreatePic Lib "olepro32" Alias "OleCreatePictureIndirect" (ByRef lpPictDesc As PictDesc, ByVal riid As Long, ByVal fPictureOwnsHandle As Long, ByRef ipic As IPicture) As Long

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Public Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Public Type asmBmpPara
    src          As BITMAP
    tgt          As BITMAP
    srcExpansion As Long
End Type

Private Const BI_RGB            As Long = 0
Private Const DIB_RGB_COLORS    As Long = 0
Private Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Public Function ReadDataBrightness(pic As IPicture, abp As asmBmpPara, src() As Byte, tgt() As Byte, _
                                   Optional pmd As Boolean = True, Optional eMD As Boolean = True) As Boolean

    Dim bmp As BITMAP
    Dim bmi As BITMAPINFO
    Dim hDC As Long, ptr As Long, tmp As Long
    
    If GetObject(pic.Handle, Len(bmp), bmp) = 0 Then
        If eMD Then tmp = 0
    ElseIf pmd = True And bmp.bmBits <> 0 And bmp.bmBitsPixel = 32 Then
        ptr = bmp.bmBits
    Else
        With bmi.bmiHeader
            .biSize = Len(bmi.bmiHeader)
            .biCompression = BI_RGB
            .biHeight = bmp.bmHeight
            .biWidth = bmp.bmWidth
            .biPlanes = 1
            .biBitCount = 32
            .biSizeImage = .biWidth * 4 * .biHeight
            
            ReDim tgt(1 To .biWidth * 4, 1 To .biHeight)
            
            hDC = GetDC(0)
            
            If hDC = 0 Then
                If eMD Then tmp = 0
            ElseIf GetDIBits(hDC, pic.Handle, 0, .biHeight, tgt(1, 1), bmi, DIB_RGB_COLORS) Then
                ptr = VarPtr(tgt(1, 1))
            Else
                ReDim tgt(1 To UBound(src, 1), 1 To UBound(src, 2))
                If eMD Then tmp = 0
            End If
            
            If hDC Then ReleaseDC 0, hDC
        End With
    End If
    
    If ptr Then
        bmp.bmBitsPixel = 32
        bmp.bmWidthBytes = bmp.bmWidth * 4
        tmp = bmp.bmWidthBytes * bmp.bmHeight
        
        ReDim src(1 To bmp.bmWidthBytes, 1 To bmp.bmHeight)
        
        CopyMemory src(1, 1), ByVal ptr, tmp
        
        abp.src = bmp
        abp.src.bmBits = VarPtr(src(1, 1))
        
        abp.tgt = bmp
        abp.tgt.bmBits = ptr
        
        ReadDataBrightness = True
    End If
    
End Function
Public Function WriteDataBrightness(pbx As PictureBox, bmp As BITMAP, _
                                    Optional force As Boolean = False) As Boolean
    Dim pic As IPicture
    Dim tmp As BITMAP
    Dim bmi As BITMAPINFO
    Dim dsc As PictDesc
    Dim hDC As Long, ptr As Long, flg As Long
    Dim iid(15) As Byte
    
    flg = pbx.Picture.Handle And (force = False)
    If flg Then flg = GetObject(flg, Len(tmp), tmp)
    
    If bmp.bmBits = tmp.bmBits Then
        pbx.Refresh
        WriteDataBrightness = True
    ElseIf flg > 0 And tmp.bmBits <> 0 And _
           bmp.bmWidth = tmp.bmWidth And bmp.bmHeight = tmp.bmHeight And _
           bmp.bmBitsPixel = tmp.bmBitsPixel Then
                CopyMemory ByVal tmp.bmBits, ByVal bmp.bmBits, bmp.bmWidthBytes * bmp.bmHeight
                pbx.Refresh
        
                WriteDataBrightness = True
    Else
        dsc.cbSizeofStruct = Len(dsc)
        dsc.picType = vbPicTypeBitmap
        
        With bmi.bmiHeader
            .biSize = Len(bmi.bmiHeader)
            .biCompression = BI_RGB
            .biBitCount = 32
            .biHeight = bmp.bmHeight
            .biWidth = bmp.bmWidth
            .biPlanes = 1
            .biSizeImage = bmp.bmWidthBytes * bmp.bmHeight
        End With
        
        hDC = GetDC(0)
        
        If hDC Then dsc.hImage = CreateDIBSection(hDC, bmi, DIB_RGB_COLORS, ptr, 0, 0)
        
        If hDC = 0 Then
            '
        ElseIf dsc.hImage = 0 Or ptr = 0 Then
            '
        Else
            CopyMemory ByVal ptr, ByVal bmp.bmBits, bmp.bmWidthBytes * bmp.bmHeight
            
            If CLSIDFromString(StrPtr(IID_IPicture), VarPtr(iid(0))) <> S_OK Then
                '
            ElseIf CreatePic(dsc, VarPtr(iid(0)), True, pic) <> S_OK Then
                '
            Else
                Set pbx.Picture = Nothing
                Set pbx.Picture = pic
                
                dsc.hImage = 0
                
                WriteDataBrightness = True
            End If
        End If
        
        If hDC Then ReleaseDC 0, hDC
        If dsc.hImage Then DeleteObject dsc.hImage
    End If
    
End Function
Public Sub ClonePic(PictureBox As PictureBox, Picture As StdPicture)
    
    PictureBox.AutoRedraw = True
    PictureBox.AutoSize = False
    PictureBox.Width = Picture.Width
    PictureBox.Height = Picture.Height
    PictureBox.ZOrder
    
    Set PictureBox.Picture = Picture

End Sub
Function ReadDataGray(pic As IPicture, abp As asmBmpPara, src() As Byte, tgt() As Byte, _
                      Optional pmd As Boolean = True, Optional eMD As Boolean = True) As Boolean
    Dim bmp As BITMAP
    Dim bmi As BITMAPINFO
    Dim hDC As Long, ptr As Long, tmp As Long
    
    If GetObject(pic.Handle, Len(bmp), bmp) = 0 Then
        If eMD Then
            '
        End If
    ElseIf pmd = True And bmp.bmBits <> 0 And bmp.bmBitsPixel = 32 Then
        ptr = bmp.bmBits
    Else
        With bmi.bmiHeader
            .biSize = Len(bmi.bmiHeader)
            .biCompression = BI_RGB
            .biHeight = bmp.bmHeight
            .biWidth = bmp.bmWidth
            .biPlanes = 1
            .biBitCount = 32
            .biSizeImage = .biWidth * 4 * .biHeight
            
            ReDim tgt(1 To .biWidth * 4, 1 To .biHeight)
            hDC = GetDC(0)
            
            If hDC = 0 Then
                If eMD Then
                    '
                End If
            ElseIf GetDIBits(hDC, pic.Handle, 0, .biHeight, tgt(1, 1), bmi, DIB_RGB_COLORS) Then
                ptr = VarPtr(tgt(1, 1))
            Else
                ReDim tgt(1 To UBound(src, 1), 1 To UBound(src, 2))
                If eMD Then
                    '
                End If
            End If
            
            If hDC Then ReleaseDC 0, hDC
        End With
    End If
    
    If ptr Then
        bmp.bmBitsPixel = 32
        bmp.bmWidthBytes = bmp.bmWidth * 4
        tmp = bmp.bmWidthBytes * bmp.bmHeight
        
        ReDim src(1 To bmp.bmWidthBytes, 1 To bmp.bmHeight)
        CopyMemory src(1, 1), ByVal ptr, tmp
        
        abp.src = bmp
        abp.src.bmBits = VarPtr(src(1, 1))
        
        abp.tgt = bmp
        abp.tgt.bmBits = ptr
        
        ReadDataGray = True
    End If
    
End Function
Function WriteDataGray(pbx As PictureBox, bmp As BITMAP, _
                       Optional force As Boolean = False) As Boolean
    Dim pic As IPicture
    Dim tmp As BITMAP
    Dim bmi As BITMAPINFO
    Dim dsc As PictDesc
    Dim hDC As Long, ptr As Long, flg As Long
    Dim iid(15) As Byte
    
    flg = pbx.Picture.Handle And (force = False)
    If flg Then flg = GetObject(flg, Len(tmp), tmp)
    
    If bmp.bmBits = tmp.bmBits Then
        pbx.Refresh
        WriteDataGray = True
    ElseIf flg > 0 And tmp.bmBits <> 0 And bmp.bmWidth = tmp.bmWidth And bmp.bmHeight = tmp.bmHeight And bmp.bmBitsPixel = tmp.bmBitsPixel Then
        CopyMemory ByVal tmp.bmBits, ByVal bmp.bmBits, bmp.bmWidthBytes * bmp.bmHeight
        pbx.Refresh
        WriteDataGray = True
    Else
        dsc.cbSizeofStruct = Len(dsc)
        dsc.picType = vbPicTypeBitmap
        
        With bmi.bmiHeader
            .biSize = Len(bmi.bmiHeader)
            .biCompression = BI_RGB
            .biBitCount = 32
            .biHeight = bmp.bmHeight
            .biWidth = bmp.bmWidth
            .biPlanes = 1
            .biSizeImage = bmp.bmWidthBytes * bmp.bmHeight
        End With
        
        hDC = GetDC(0)
        
        If hDC Then dsc.hImage = CreateDIBSection(hDC, bmi, DIB_RGB_COLORS, ptr, 0, 0)
        
        If hDC = 0 Then
            '
        ElseIf dsc.hImage = 0 Or ptr = 0 Then
            '
        Else
            CopyMemory ByVal ptr, ByVal bmp.bmBits, bmp.bmWidthBytes * bmp.bmHeight
            
            If CLSIDFromString(StrPtr(IID_IPicture), VarPtr(iid(0))) <> S_OK Then
                '
            ElseIf CreatePic(dsc, VarPtr(iid(0)), True, pic) <> S_OK Then
                '
            Else
                Set pbx.Picture = Nothing
                Set pbx.Picture = pic
                
                dsc.hImage = 0
                WriteDataGray = True
            End If
        End If
        
        If hDC Then ReleaseDC 0, hDC
        If dsc.hImage Then DeleteObject dsc.hImage
    End If
    
End Function
